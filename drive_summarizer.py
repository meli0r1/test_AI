"""Create one summary from a local folder using OpenRouter.

The program reads files from a local directory, extracts text from common
document formats, sends photos and scanned PDF pages to a vision-capable
model, and writes a Markdown report.

Example:
    python drive_summarizer.py \
      --folder "C:\\Users\\forsu\\Desktop\\Задача [Трибониан AI]" \
      --model google/gemini-2.5-flash \
      --output summary.md
"""

from __future__ import annotations

import argparse
import base64
import csv
import io
import json
import mimetypes
import os
import re
import sys
import time
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Iterable, Optional
from urllib.error import HTTPError, URLError
from urllib.request import Request, urlopen

DEFAULT_FOLDER = r"C:\Users\forsu\Desktop\Задача [Трибониан AI]"
OPENROUTER_BASE_URL = "https://openrouter.ai/api/v1"
DEFAULT_MODEL = "google/gemini-2.5-flash"

TEXT_EXTENSIONS = {
    ".txt", ".md", ".markdown", ".rst", ".log", ".json", ".xml", ".html",
    ".htm", ".csv", ".tsv", ".rtf", ".yaml", ".yml", ".sql", ".py", ".js",
    ".ts", ".css", ".ini", ".cfg", ".conf",
}
IMAGE_EXTENSIONS = {".jpg", ".jpeg", ".png", ".webp", ".gif", ".bmp", ".tif", ".tiff"}
IGNORED_DIRECTORIES = {".git", ".venv", "venv", "__pycache__", "node_modules"}


@dataclass
class LocalFile:
    path: Path
    relative_path: str
    mime_type: str


@dataclass
class ExtractedDocument:
    file: LocalFile
    text: str


class OpenRouterClient:
    """Small dependency-free HTTP client for OpenRouter chat completions."""

    def __init__(self, api_key: str) -> None:
        self.api_key = api_key
        self.url = f"{OPENROUTER_BASE_URL}/chat/completions"

    def chat(self, model: str, messages: list[dict], retries: int = 3) -> str:
        payload = json.dumps({
            "model": model,
            "messages": messages,
            "temperature": 0.2,
        }).encode("utf-8")
        headers = {
            "Authorization": f"Bearer {self.api_key}",
            "Content-Type": "application/json",
            "HTTP-Referer": "http://localhost",
            "X-Title": "Local folder summarizer",
        }

        last_error = None
        for attempt in range(retries):
            try:
                request = Request(self.url, data=payload, headers=headers, method="POST")
                with urlopen(request, timeout=180) as response:
                    result = json.loads(response.read().decode("utf-8"))
                choices = result.get("choices") or []
                if not choices:
                    raise RuntimeError(f"OpenRouter вернул неожиданный ответ: {result}")
                content = choices[0].get("message", {}).get("content", "")
                if isinstance(content, list):
                    content = "\n".join(
                        part.get("text", "") for part in content if isinstance(part, dict)
                    )
                return str(content).strip()
            except HTTPError as exc:
                details = exc.read().decode("utf-8", errors="replace")
                last_error = RuntimeError(f"HTTP {exc.code}: {details[:1000]}")
            except (URLError, TimeoutError, OSError, json.JSONDecodeError, RuntimeError) as exc:
                last_error = exc
            if attempt < retries - 1:
                time.sleep(2 ** attempt)
        raise RuntimeError(f"ошибка OpenRouter: {last_error}")


def iter_local_files(root: Path, excluded_paths: set[Path]) -> Iterable[LocalFile]:
    root = root.resolve()
    for path in sorted(root.rglob("*"), key=lambda item: str(item).lower()):
        if not path.is_file() or path.resolve() in excluded_paths:
            continue
        if any(part in IGNORED_DIRECTORIES for part in path.parts):
            continue
        mime_type = mimetypes.guess_type(path.name)[0] or "application/octet-stream"
        yield LocalFile(path, str(path.relative_to(root)), mime_type)


def decode_text(data: bytes) -> str:
    for encoding in ("utf-8-sig", "utf-8", "cp1251", "latin-1"):
        try:
            return data.decode(encoding)
        except UnicodeDecodeError:
            continue
    return data.decode("utf-8", errors="replace")


def strip_html(text: str) -> str:
    text = re.sub(r"<script[^>]*>.*?</script>", " ", text, flags=re.I | re.S)
    text = re.sub(r"<style[^>]*>.*?</style>", " ", text, flags=re.I | re.S)
    text = re.sub(r"<[^>]+>", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def extract_pdf(data: bytes) -> str:
    from pypdf import PdfReader

    reader = PdfReader(io.BytesIO(data))
    pages = []
    for number, page in enumerate(reader.pages, 1):
        text = page.extract_text() or ""
        if text.strip():
            pages.append(f"[Страница {number}]\n{text}")
    return "\n\n".join(pages)


def extract_docx(data: bytes) -> str:
    from docx import Document

    document = Document(io.BytesIO(data))
    parts = [paragraph.text for paragraph in document.paragraphs if paragraph.text.strip()]
    for table_number, table in enumerate(document.tables, 1):
        rows = []
        for row in table.rows:
            rows.append(" | ".join(cell.text.replace("\n", " ") for cell in row.cells))
        parts.append(f"[Таблица {table_number}]\n" + "\n".join(rows))
    return "\n\n".join(parts)


def extract_pptx(data: bytes) -> str:
    from pptx import Presentation

    presentation = Presentation(io.BytesIO(data))
    slides = []
    for number, slide in enumerate(presentation.slides, 1):
        texts = []
        for shape in slide.shapes:
            if hasattr(shape, "text") and shape.text.strip():
                texts.append(shape.text.strip())
        if texts:
            slides.append(f"[Слайд {number}]\n" + "\n".join(texts))
    return "\n\n".join(slides)


def extract_xlsx(data: bytes) -> str:
    from openpyxl import load_workbook

    workbook = load_workbook(io.BytesIO(data), read_only=True, data_only=True)
    sheets = []
    for worksheet in workbook.worksheets:
        rows = []
        for row in worksheet.iter_rows(values_only=True):
            values = ["" if value is None else str(value) for value in row]
            if any(value.strip() for value in values):
                rows.append(" | ".join(values))
        if rows:
            sheets.append(f"[Лист: {worksheet.title}]\n" + "\n".join(rows))
    return "\n\n".join(sheets)


def extract_csv(data: bytes, delimiter: Optional[str] = None) -> str:
    text = decode_text(data)
    if delimiter is None:
        try:
            delimiter = csv.Sniffer().sniff(text[:4096]).delimiter
        except csv.Error:
            delimiter = ","
    rows = csv.reader(io.StringIO(text), delimiter=delimiter)
    return "\n".join(" | ".join(cell.strip() for cell in row) for row in rows)


def extract_text(file: LocalFile, data: bytes) -> str:
    suffix = file.path.suffix.lower()
    mime = file.mime_type.lower()

    if suffix == ".pdf" or mime == "application/pdf":
        return extract_pdf(data)
    if suffix == ".docx":
        return extract_docx(data)
    if suffix == ".pptx":
        return extract_pptx(data)
    if suffix == ".xlsx":
        return extract_xlsx(data)
    if suffix in {".csv", ".tsv"}:
        return extract_csv(data, "\t" if suffix == ".tsv" else None)
    if suffix in {".html", ".htm"}:
        return strip_html(decode_text(data))
    if suffix in TEXT_EXTENSIONS or mime.startswith("text/"):
        return decode_text(data)
    raise ValueError(f"формат не поддерживается: {file.path.suffix or file.mime_type}")


def prepare_image(data: bytes, mime_type: str, max_side: int) -> tuple[bytes, str]:
    """Resize large images before sending them as base64 to reduce cost."""
    try:
        from PIL import Image

        image = Image.open(io.BytesIO(data))
        image.thumbnail((max_side, max_side))
        output = io.BytesIO()
        if image.mode in {"RGBA", "LA"}:
            image.save(output, format="PNG", optimize=True)
            return output.getvalue(), "image/png"
        image.convert("RGB").save(output, format="JPEG", quality=88, optimize=True)
        return output.getvalue(), "image/jpeg"
    except Exception:  # noqa: BLE001 - use original bytes if optional image processing fails
        return data, mime_type if mime_type.startswith("image/") else "image/jpeg"


def render_pdf_pages(data: bytes, max_side: int) -> list[bytes]:
    """Render scanned PDF pages for vision analysis when no text layer exists."""
    import fitz  # PyMuPDF

    document = fitz.open(stream=data, filetype="pdf")
    rendered = []
    for page in document:
        rect = page.rect
        scale = min(1.8, max_side / max(rect.width, rect.height))
        pixmap = page.get_pixmap(matrix=fitz.Matrix(scale, scale), alpha=False)
        rendered.append(pixmap.tobytes("png"))
    return rendered


def split_text(text: str, max_chars: int) -> list[str]:
    text = text.strip()
    if not text:
        return []
    if len(text) <= max_chars:
        return [text]
    chunks = []
    start = 0
    while start < len(text):
        end = min(start + max_chars, len(text))
        if end < len(text):
            boundary = max(text.rfind("\n", start, end), text.rfind(". ", start, end))
            if boundary > start + max_chars // 2:
                end = boundary + (1 if text[boundary] == "\n" else 2)
        chunks.append(text[start:end].strip())
        start = max(end - 500, start + 1)
    return [chunk for chunk in chunks if chunk]


def describe_image(
    client: OpenRouterClient,
    model: str,
    file_name: str,
    data: bytes,
    mime_type: str,
    max_side: int,
) -> str:
    image_data, image_mime = prepare_image(data, mime_type, max_side)
    data_url = f"data:{image_mime};base64,{base64.b64encode(image_data).decode('ascii')}"
    prompt = (
        f"Это файл {file_name}. Проанализируй изображение максимально содержательно. "
        "Извлеки весь читаемый текст, числа, даты, названия, подписи, таблицы, "
        "схемы и важный визуальный контекст. Если часть текста неразборчива, "
        "отметь это явно. Не выдумывай отсутствующие сведения. Ответь по-русски."
    )
    return client.chat(
        client,
        model,
        [{
            "role": "user",
            "content": [
                {"type": "text", "text": prompt},
                {"type": "image_url", "image_url": {"url": data_url}},
            ],
        }],
    )


def summarize_document(client: OpenRouterClient, model: str, document: ExtractedDocument, max_chars: int) -> str:
    chunks = split_text(document.text, max_chars)
    if not chunks:
        return "Текст не найден."
    system = (
        "Ты аналитик документов. Суммируй только информацию из входного текста, "
        "не придумывай факты. Сохраняй числа, даты, названия, решения, задачи, "
        "ограничения и противоречия. Пиши по-русски, структурированно и сжато."
    )
    partials = []
    for index, chunk in enumerate(chunks, 1):
        partials.append(client.chat(model, [
            {"role": "system", "content": system},
            {"role": "user", "content": (
                f"Документ: {document.file.relative_path}\n"
                f"Часть {index}/{len(chunks)}:\n\n{chunk}\n\n"
                "Сделай содержательное резюме этой части с фактами и выводами."
            )},
        ]))
    if len(partials) == 1:
        return partials[0]
    return client.chat(model, [
        {"role": "system", "content": system},
        {"role": "user", "content": (
            "Объедини резюме частей одного документа в единое резюме без потери "
            "важных фактов. Удали дублирование.\n\n" + "\n\n---\n\n".join(partials)
        )},
    ])


def summarize_collection(client: OpenRouterClient, model: str, summaries: list[tuple[LocalFile, str]]) -> str:
    if not summaries:
        return "В папке не найдено файлов с извлекаемым текстом."
    digest = "\n\n".join(
        f"### {file.relative_path}\n{summary}" for file, summary in summaries
    )
    chunks = split_text(digest, 30000)
    system = (
        "Ты готовишь итоговый аналитический отчет по набору документов. Используй "
        "только переданные резюме. Если данных недостаточно, так и напиши. "
        "Сохраняй конкретику и не приписывай документам то, чего в них нет."
    )
    partials = [client.chat(model, [
        {"role": "system", "content": system},
        {"role": "user", "content": (
            "Сводка части набора документов:\n\n" + chunk + "\n\n"
            "Выдели общие темы, факты, выводы, решения, задачи, риски и противоречия."
        )},
    ]) for chunk in chunks]
    if len(partials) > 1:
        base = client.chat(model, [
            {"role": "system", "content": system},
            {"role": "user", "content": (
                "Объедини промежуточные аналитические сводки в одну, убери повторы:\n\n"
                + "\n\n---\n\n".join(partials)
            )},
        ])
    else:
        base = partials[0]
    return client.chat(model, [
        {"role": "system", "content": system},
        {"role": "user", "content": (
            "Подготовь финальное саммари всей папки в Markdown со следующими разделами:\n"
            "1. Краткий вывод для руководителя\n"
            "2. Основные темы и факты\n"
            "3. Важные числа и даты\n"
            "4. Решения и договоренности\n"
            "5. Задачи и следующие шаги\n"
            "6. Риски, пробелы и противоречия\n"
            "7. Резюме по документам\n\n"
            "Промежуточная сводка:\n" + base
        )},
    ])


def process_file(
    client: OpenRouterClient,
    model: str,
    file: LocalFile,
    image_max_side: int,
) -> str:
    data = file.path.read_bytes()
    suffix = file.path.suffix.lower()
    if suffix in IMAGE_EXTENSIONS or file.mime_type.startswith("image/"):
        return describe_image(client, model, file.relative_path, data, file.mime_type, image_max_side)

    text = extract_text(file, data).strip()
    if suffix == ".pdf" and not text:
        pages = render_pdf_pages(data, image_max_side)
        descriptions = []
        for number, page in enumerate(pages, 1):
            descriptions.append(describe_image(
                client, model, f"{file.relative_path}, страница {number}",
                page, "image/png", image_max_side,
            ))
        text = "\n\n".join(descriptions)
    if not text:
        raise ValueError("текст отсутствует")
    return text


def write_report(
    output: Path,
    root: Path,
    files: list[LocalFile],
    documents: list[ExtractedDocument],
    final_summary: str,
    skipped: list[tuple[LocalFile, str]],
) -> None:
    processed_ids = {document.file.path.resolve() for document in documents}
    lines = [
        "# Общее саммари локальной папки",
        "",
        f"- Источник: `{root}`",
        f"- Сформировано: `{datetime.now().astimezone().isoformat(timespec='seconds')}`",
        f"- Всего файлов: **{len(files)}**",
        f"- Обработано файлов: **{len(documents)}**",
        f"- Пропущено файлов: **{len(skipped)}**",
        "",
        "## Итоговое саммари",
        "",
        final_summary,
        "",
        "## Источники",
        "",
    ]
    for file in files:
        status = "обработан" if file.path.resolve() in processed_ids else "пропущен"
        lines.append(f"- `{file.relative_path}` — {status}")
    if skipped:
        lines += ["", "## Причины пропуска", ""]
        lines.extend(f"- `{file.relative_path}`: {reason}" for file, reason in skipped)
    output.write_text("\n".join(lines) + "\n", encoding="utf-8")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--folder", default=DEFAULT_FOLDER, help="Путь к локальной папке")
    parser.add_argument("--output", default="summary.md", type=Path)
    parser.add_argument(
        "--model",
        default=os.getenv("OPENROUTER_MODEL", DEFAULT_MODEL),
        help="OpenRouter model slug; для фотографий нужна vision-модель",
    )
    parser.add_argument("--chunk-chars", type=int, default=18000)
    parser.add_argument("--image-max-side", type=int, default=1600)
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    api_key = os.getenv("OPENROUTER_API_KEY")
    if not api_key:
        print("Ошибка: задайте переменную окружения OPENROUTER_API_KEY.", file=sys.stderr)
        return 2

    root = Path(args.folder).expanduser()
    if not root.is_dir():
        print(f"Ошибка: папка не найдена: {root}", file=sys.stderr)
        return 2

    output = args.output if args.output.is_absolute() else Path.cwd() / args.output
    client = OpenRouterClient(api_key)

    excluded_paths = {output.resolve()} if output.exists() else set()
    files = list(iter_local_files(root, excluded_paths))
    print(f"Найдено файлов: {len(files)}")

    documents: list[ExtractedDocument] = []
    skipped: list[tuple[LocalFile, str]] = []
    for number, file in enumerate(files, 1):
        print(f"[{number}/{len(files)}] {file.relative_path}")
        try:
            text = process_file(client, args.model, file, args.image_max_side)
            documents.append(ExtractedDocument(file, text))
        except Exception as exc:  # noqa: BLE001 - one bad file must not stop the folder
            reason = str(exc).replace("\n", " ")
            skipped.append((file, reason))
            print(f"  пропущен: {reason}", file=sys.stderr)

    per_file = []
    for number, document in enumerate(documents, 1):
        print(f"Саммари файла [{number}/{len(documents)}]: {document.file.relative_path}")
        per_file.append((
            document.file,
            summarize_document(client, args.model, document, args.chunk_chars),
        ))

    print("Формирую общее саммари...")
    final_summary = summarize_collection(client, args.model, per_file)
    write_report(output, root.resolve(), files, documents, final_summary, skipped)
    print(f"Готово: {output.resolve()}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
