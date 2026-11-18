# Финансовый планер — сайт

Личный дашборд: бюджет, подушка, цели, прогноз капитала, портфель.
Цифры хранятся в Google Таблице владельца. Сайт после входа через Google читает её (только чтение)
и строит графики. В репозитории нет никаких личных данных.

## Файлы
- `index.html` — сам сайт
- `config.js` — Client ID Google и ID таблицы

## Разовая настройка входа через Google (≈15 минут)
1. Открыть https://console.cloud.google.com/ → создать проект `finplan`.
2. «APIs & Services» → «Library» → найти **Google Sheets API** → Enable.
3. «APIs & Services» → «OAuth consent screen» (Google Auth Platform):
   - User type: **External**, имя приложения `Финансовый планер`, свой email.
   - Audience → **Test users** → добавить свой Gmail.
   - Data access → Add scope → `https://www.googleapis.com/auth/spreadsheets.readonly`.
4. «Clients» (Credentials) → Create client → **Web application**:
   - Authorized JavaScript origins: `https://korotkayadarina-png.github.io`
   - Скопировать **Client ID** (вида `1234-abc.apps.googleusercontent.com`) и вставить в `config.js`.
5. При первом входе Google покажет «Приложение не проверено» — это нормально для личного приложения:
   «Дополнительно» → «Перейти на сайт».

## Структура таблицы, которую читает сайт
- Бюджет: доходы A6:C15; расходы A20:D49
- Портфель: A5:E29
- Подушка: B4–B7
- Цели: A5:E24
- Прогноз: B4–B10
- История: A5:E64

Если менять расположение ячеек в таблице, нужно обновить и `index.html` (функция `fromWorkbook`).

Не является индивидуальной инвестиционной рекомендацией.
