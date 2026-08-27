<div align="center">

# QuickDock

Плавающая иконка-компаньон для Windows 10

[![Version](https://img.shields.io/badge/version-2.32.0-EE3239)]()
[![Platform](https://img.shields.io/badge/platform-Windows%2010%20x64-0078D4)]()
[![License](https://img.shields.io/badge/license-MIT-green)]()
[![Portable](https://img.shields.io/badge/portable-yes-success)]()

Маленькая иконка поверх всех окон. При наведении раскрывается панель с тремя разделами: буфер обмена, файлы → PDF, быстрые ответы.

</div>

---

## Возможности

### Иконка-персонаж
- **3 состояния**: `normal` (обычное), `blink` (моргает при копировании в буфер), `file` (когда тащите файл рядом)
- Перетаскивается в любое место экрана, прилипает к ближайшему краю
- Позиция запоминается между запусками
- Поддержка multi-monitor

### Раздел 1 — Буфер обмена
- Текущий + 4 предыдущих скопированных текста
- Длинные тексты сокращаются только в превью — при клике возвращается полный текст
- Копии картинок/файлов игнорируются (только текст)
- История очищается при перезапуске

### Раздел 2 — Файлы → PDF
- **Drag-and-drop** в обе стороны: тащите файлы на иконку или в dropzone, тащите файлы из списка наружу
- Форматы: PNG, JPG, JPEG, JFIF, BMP, GIF, WEBP, DOC, DOCX, XLSX, XLS, PDF
- **Word-файлы** (.doc, .docx) конвертируются через Microsoft Word COM — идеальное сохранение форматирования, изображений, таблиц
- **XLSX** — собственный рендер: только активные ячейки, поддержка merged cells, дат, стилей
- **PDF** — страницы копируются как есть (вектор сохраняется)
- Сборка в один PDF формата A4, fit-to-page (без обрезки, без переноса)
- Исходные файлы удаляются после сборки, PDF называется по первому файлу

### Раздел 3 — Быстрые ответы
- Свой список текстовых заготовок для копирования одним кликом
- Добавление, редактирование, удаление прямо в приложении
- Сохраняются между запусками (quickdock-replies.json)

### Tray-меню
- Показать / Скрыть
- Настройки → Размер иконки (Маленькая / Обычная / Крупная / Очень крупная)
- Поверх всех окон (fallback если панель застряла)
- Выход

## Дизайн
- **Nothing-style**: тёмный фон #0A0A0A, красный акцент #EE3239
- Шрифт лого: Manufacturing Consent (Google Fonts)
- Моноширинный: JetBrains Mono / Cascadia Mono
- Serif (заголовки): Cambria
- Скруглённые углы, декоративная угловая графика
- Без серых границ-разделителей, пунктир только у полей ввода и dropzone

## Установка

### Вариант 1 — Скачать готовую сборку
1. Скачайте `QuickDock.zip` (~123 МБ) со [страницы загрузки](https://preview-zai-web.space-z.ai/)
2. Распакуйте в любую папку
3. Запустите `QuickDock.exe`
4. Права администратора **не требуются**

### Вариант 2 — Собрать из исходников
```bash
# Клонировать репозиторий
git clone https://github.com/your-username/quickdock.git
cd quickdock

# Установить зависимости
cd quickdock
npm install
npx electron --version  # проверить установку

# Скачать Electron win32-x64 binary (для кросс-компиляции)
node -e "require('@electron/get').downloadArtifact({version:'33.4.11',platform:'win32',arch:'x64',artifactName:'electron',cacheRoot:require('os').homedir()+'/.cache/electron'}).then(p=>console.log(p))"

# Сгенерировать иконки (нужен Python + Pillow)
pip install Pillow
python ../scripts/make_icon.py

# Собрать portable .exe
node ../scripts/pack.js

# Результат: ../build-dist/QuickDock/
```

## Структура проекта

```
quickdock/
├── package.json              # Зависимости: electron, mammoth, pdf-lib, xlsx
├── build/                    # Иконки и графика
│   ├── icon-normal.png       # Обычное состояние иконки
│   ├── icon-normal-blink.png # Blink при копировании
│   ├── icon-file.png         # При перетаскивании файла
│   ├── corner-tl.png         # Декоративная графика (верх-лево)
│   ├── corner-tr.png         # (верх-право)
│   ├── corner-bl.png         # (низ-лево)
│   ├── tray.png              # Иконка трея
│   └── icon.ico              # Windows .ico (multi-size)
└── src/
    ├── main/
    │   ├── index.js          # Главный процесс: окна, tray, IPC, drag, hover
    │   ├── store.js          # Settings, replies, temp dir
    │   ├── clipboard.js      # Clipboard watcher (poll 500ms)
    │   └── pdf.js            # PDF builder: image/docx/xlsx/pdf → A4
    ├── preload/
    │   └── index.js          # contextBridge: IPC + webUtils.getPathForFile
    └── renderer/
        ├── icon.html         # Окно иконки
        ├── panel.html        # Окно панели (3 раздела)
        ├── css/
        │   ├── base.css      # CSS variables, кнопки, скроллбар
        │   ├── icon.css      # Иконка: центрирование PNG
        │   └── panel.css     # Панель: layout, tabs, lists, dropzone
        └── js/
            ├── icon.js       # Hover, drag-move, file-drop
            └── panel.js      # Tabs, clipboard, files, replies CRUD

scripts/
├── pack.js                   # Кросс-компиляция без wine (Linux → Windows)
└── make_icon.py              # Генерация иконок (Pillow)
```

## Файлы данных (создаются автоматически рядом с .exe)

| Файл | Описание | Стирается при закрытии? |
|------|----------|------------------------|
| `quickdock-temp/` | Временные файлы (копии добавленных файлов) | ✅ Да |
| `quickdock-replies.json` | Быстрые ответы | ❌ Нет |
| `quickdock-settings.json` | Размер иконки + позиция | ❌ Нет |

## Технические детали

- **Рантайм**: Electron 33 (Chromium 130 + Node 20)
- **Портативная сборка**: ничего не пишется в реестр, не устанавливается в систему
- **DPI-масштабирование**: `high-dpi-support` + `force-device-scale-factor=1`
- **Z-order**: иконка (`screen-saver`) выше панели (`floating`) — всегда кликабельна
- **Hover detection**: cursor polling (150ms) — не зависит от mouseenter/mouseleave
- **Прозрачность**: `backgroundColor: '#00000000'` + `transparent: true` для скруглённых углов
- **Drag-out**: `ipcRenderer.sendSync` → `webContents.startDrag` (синхронный IPC обязателен)
- **Файлы**: `webUtils.getPathForFile()` (Electron 33+ — `File.path` пустой при contextIsolation)

### Кросс-компиляция (Linux → Windows)
Сборка делается на Linux без wine. Скрипт `pack.js`:
1. Распаковывает `electron-v33.4.11-win32-x64.zip` (из кэша)
2. Переименовывает `electron.exe` → `QuickDock.exe`
3. Копирует `src/` + `build/` в `resources/app/`
4. Устанавливает production-зависимости
5. Удаляет `default_app.asar`

## Лицензия

MIT — используйте свободно.

## Ссылки

- [Скачать QuickDock.zip](https://preview-zai-web.space-z.ai/)
- [Конструктор UI](https://preview-zai-web.space-z.ai/theme-constructor)
