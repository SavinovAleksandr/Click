# Click v.5 - Исходный код (восстановленный)

Макрос для экспорта графики из файлов RastrWin3 (RG2) в документ Word.

## Язык и технологии

- **Язык:** Python 3.14 (оригинал), совместимо с Python 3.10+
- **Упаковка:** PyInstaller
- **GUI:** PyQt5
- **Платформа:** Windows (требует RastrWin3, Microsoft Word для SVG)

## Структура проекта

```
click_source/
├── Click_v.5.py      # Точка входа
├── Click_GUI.py      # Графический интерфейс (PyQt5)
├── Click_Rastr.py    # Интеграция с RastrWin3 (COM)
├── Click_word.py     # Вставка в Word (python-docx, COM)
├── Click_dxf2png.py  # Конвертация DXF -> PNG
├── Click_dxf2svg.py  # Конвертация DXF -> SVG
├── requirements.txt
└── README.md
```

## Установка

```bash
pip install -r requirements.txt
# На Windows дополнительно:
pip install pywin32
```

## Запуск

```bash
python Click_v.5.py
```

## Р7-Офис (вместо Microsoft Word)

Для автоматической вставки **SVG** в документы Р7-Офис:
1. Установите Р7-Офис с **Документ Конструктор** (docconstructor)
2. Включите опцию «Р7-Офис (вместо Microsoft Word)» в интерфейсе
3. Выберите формат SVG

Модуль `Click_word_r7.py` использует Document Builder API для создания output.docx с векторной графикой, совместимого с Р7-Офис.

## Требования

- Windows (RastrWin3, COM)
- RastrWin3 установлен и настроен
- Шаблон Word со стилями: 7.32 ЗАГОЛОВОК ПРИЛОЖЕНИЙ, 7.32 Иллюстрация и т.д.
- Для SVG + Word: Microsoft Word 2016+ с поддержкой SVG
- Для SVG + Р7-Офис: Р7-Офис с Документ Конструктор

## Примечание

Исходный код восстановлен из байткода Python 3.14 методом дизассемблирования и ручной реконструкции (декомпиляторы для Python 3.14 отсутствуют). Некоторые детали реализации могут отличаться от оригинала.
