# VRTN Official Letterhead Skill (`vrtn-letterhead`)

Скилл для автоматизированного создания официальных документов на фирменном бланке бренда **VRTN Workshop / VERETENO workshop** (ИП Ананьев Максим Андреевич).

Генерирует документы полиграфического качества в формате **PDF** (на базе ReportLab) и редактируемые файлы **DOCX** (Microsoft Word) с соблюдением всех корпоративных стандартов:
* Фирменный логотип (`logo.png`)
* Типографика **Montserrat** (`Montserrat-Bold`, `Montserrat-Regular`)
* Двухколоночный блок реквизитов и тонкие линии-разделители
* Выделенный блок для ключевых положений (причина возврата, предмет обращения)
* Полные банковские реквизиты ИП Ананьев М.А. (Сбербанк)
* Корпоративный подвал: `CREATED IN KHAKASSIA | EST. 2015 | VRTN WORKSHOP`

---

## Структура папки скилла

```text
skills/vrtn-letterhead/
├── SKILL.md                 # Полная спецификация скилла для ИИ-агентов (Antigravity / Cursor)
├── README.md                # Описание и руководство пользователя
├── assets/                  # Корпоративные графические и шрифтовые ресурсы
│   ├── logo.png             # Оригинальный логотип VRTN Workshop
│   ├── Montserrat-Bold.ttf   # Шрифт Montserrat Bold
│   └── Montserrat-Regular.ttf# Шрифт Montserrat Regular
├── scripts/
│   └── generate_letter.py   # CLI-генератор документов PDF и DOCX
└── examples/
    └── sample_refund.json   # Пример конфигурации заявления на возврат средств
```

---

## Быстрый запуск генерации

```powershell
uv run --with reportlab --with python-docx python scripts/generate_letter.py --config examples/sample_refund.json
```

Параметры входного файла `JSON`:
* `output_basename` — базовое имя файла (без расширения)
* `output_dir` — директория для сохранения
* `formats` — список форматов `["pdf", "docx"]`
* `recipient_position`, `recipient_company`, `recipient_name`, `recipient_address`, `recipient_details` — блок адресата
* `title`, `subtitle`, `salutation` — заголовок, подзаголовок и обращение
* `paragraphs` — массив абзацев основного текста (поддерживаются теги `<b>...</b>`)
* `callout` — текст выделенного блока (например, причина возврата)
* `include_bank_details` — булево значение (`true`/`false`)
* `attachments` — список приложений
* `date` — дата документа
