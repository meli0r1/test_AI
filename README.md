# Саммари локальной папки

`drive_summarizer.py` читает локальную папку, извлекает данные из документов и анализирует текст, изображения и сканированные PDF через OpenRouter.

## 1. Установка

Нужен Python 3.10 или новее.

```powershell
py -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
```

## 2. OpenRouter

Создайте ключ OpenRouter и задайте его только в окружении текущего терминала:

```powershell
$env:OPENROUTER_API_KEY = "sk-or-v1-..."
```

Код использует endpoint OpenRouter `https://openrouter.ai/api/v1`. Для фотографий нужна модель с поддержкой image input, например `google/gemini-2.5-flash`.

## 3. Запуск

Путь по умолчанию уже установлен в коде:

```text
C:\Users\forsu\Desktop\Задача [Трибониан AI]
```

Запуск:

```powershell
python .\drive_summarizer.py `
  --folder "C:\Users\forsu\Desktop\Задача [Трибониан AI]" `
  --model google/gemini-2.5-flash `
  --output .\summary.md
```

Или можно запустить без `--folder`, поскольку этот путь используется по умолчанию:

```powershell
python .\drive_summarizer.py
```

Полезные параметры:

```powershell
python .\drive_summarizer.py `
  --folder "<локальный путь>" `
  --model google/gemini-2.5-flash `
  --chunk-chars 18000 `
  --image-max-side 1600
```

Результат появится в `summary.md`. Один поврежденный или неподдерживаемый файл не останавливает обработку: причина попадет в раздел «Причины пропуска».

## Важные замечания

- Саммари и анализ фотографий выполняются через OpenRouter, поэтому содержимое файлов отправляется внешнему API.
- API-ключ не записывается в исходный код. Не добавляйте его в Git и не публикуйте.
- Сканированные PDF автоматически переводятся в изображения и отправляются vision-модели постранично.
- Большие изображения уменьшаются перед отправкой, чтобы снизить расход токенов.
- Итоговый отчет является сжатым пересказом, а не заменой исходных файлов.
