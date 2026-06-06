# КАМІНХОЛ — Offers Manager

Внутрішній інструмент для управління пропозиціями маркетплейсу.

## Стек

- Next.js 14 (App Router)
- TypeScript + Tailwind CSS
- SheetJS (xlsx) — парсинг Excel у браузері
- GitHub API — публікація offers.json

## Встановлення

```bash
npm install
```

## Налаштування змінних середовища

Скопіюй `.env.example` → `.env.local` і заповни:

```
GITHUB_TOKEN=ghp_...          # GitHub Personal Access Token (repo scope)
GITHUB_OWNER=agency45tools
GITHUB_REPO=kaminhall-feed
GITHUB_BRANCH=main
NEXT_PUBLIC_APP_PASSWORD=...  # пароль для входу в додаток
```

## Запуск локально

```bash
npm run dev
```

## Деплой на Vercel

1. Залити код у GitHub репо (окреме від kaminhall-feed)
2. Підключити до Vercel
3. Додати всі env variables у Vercel Dashboard → Settings → Environment Variables
4. Deploy

## Логіка роботи

### Завантаження Excel

- **Broil King (.xls)**: Лист `TDSheet`, заголовки на рядку 2. Фільтруємо тільки `Broil King` (виключаємо Big Green Egg).
  - Артикул (col 0) → vendor_code матчинг
  - `Роздрібний з ПДВ постійний` (col 4) → price
  - Залишок (col 5) → stock

- **Weber (.xlsx)**: Лист `Weber`, заголовки на рядку 1. Пропускаємо рядки-категорії (де Артикул порожній).
  - Артикул (col 1) → vendor_code матчинг
  - Ціна (col 3) / Акційна ціна (col 4) → price / old_price
  - Залишок (col 5) → stock

### Логіка ціни (Weber)

- Якщо `Акційна ціна` відсутня: `price = Ціна`, `old_price = null`
- Якщо `Акційна ціна` є: `price = Акційна ціна`, `old_price = Ціна`

### Матчинг

`Excel.vendor_code` → `XML.<vendor_code>` → `XML.<code>` → `offers.json.code`

### Публікація

Кнопка "Опублікувати" → генерує `offers.json` → PUT до GitHub API →
Vercel автоматично деплоїть через webhook.

## Структура offers.json

```json
{
  "total": 192,
  "updatedAt": "2026-06-06T12:00:00Z",
  "data": [
    {
      "code": "G1-006",
      "price": 15999,
      "old_price": null,
      "availability": true,
      "stock": 5,
      "warranty_type": "manufacturer",
      "warranty_period": 24,
      "days_to_dispatch": 1,
      "max_pay_in_parts": 12,
      "delivery_methods": [...],
      "manufacture": { "country_code": "US", "year": null },
      "warehouses": [{ "id": "WH-01", "stock": 5 }]
    }
  ]
}
```
