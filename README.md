# HSE Case Management Telegram Bot

Archived prototype of a Telegram bot for searching HSE case-management data
from an Excel knowledge base.

The bot accepts a person name, department name, or department code, searches the
provided workbook, and returns the closest matching administrative records.

## Status

This repository is kept as a historical prototype. The real data file is not
included in the repository for privacy and security reasons.

## Tech Stack

- Python
- python-telegram-bot
- openpyxl
- Levenshtein distance search
- Docker

## Expected Data File

The application expects an Excel workbook mounted as `/app/Data.xlsx`.

At a high level, the workbook should contain:

- a `Полномочия` sheet with person-related authority records;
- an `Оргструктура` sheet with department names, department codes, and
  organization-structure metadata.

Do not commit real workbooks, exports, archives, tokens, or local `.env` files.

## Demo Data

You can generate a small synthetic workbook for local testing:

```shell
python scripts/create_sample_data.py
```

This creates `Data.xlsx` in the repository root. The file is ignored by Git and
contains only fake demo records.

## Configuration

Create a local `.env` file:

```env
BOT_TOKEN=your-telegram-bot-token
```

## Docker

Build the image:

```shell
docker build -t hse-case-management-telegram-bot .
```

Run the bot with a local workbook and environment file:

```shell
docker run --rm \
  --mount type=bind,source="$(pwd)/Data.xlsx",target=/app/Data.xlsx,readonly \
  --env-file .env \
  hse-case-management-telegram-bot
```

## Privacy Note

Historical data artifacts were intentionally removed from this repository. Keep
all real datasets outside Git.
