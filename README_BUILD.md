# Инструкция по сборке приложения Disp

## Что было создано

1. **Disp.spec** - файл спецификации для PyInstaller
2. **Disp_Setup.iss** - скрипт для Inno Setup
3. **build_installer.bat** - автоматический скрипт сборки

## Требования

- Python с установленными зависимостями из requirements.txt
- PyInstaller: `pip install pyinstaller`
- Inno Setup (скачать с https://jrsoftware.org/isinfo.php)

## Способы сборки

### Способ 1: Автоматический (рекомендуется)

Просто запустите:
```bash
build_installer.bat
```

Этот скрипт:
1. Очистит предыдущие сборки
2. Создаст исполняемый файл с PyInstaller
3. Создаст инсталлер с Inno Setup

### Способ 2: Ручной

#### Шаг 1: Сборка исполняемого файла
```bash
cd C:\Users\User\Desktop\Disp
pyinstaller --clean --noconfirm .\Disp.spec
```

#### Шаг 2: Создание инсталлера
```bash
"C:\Program Files (x86)\Inno Setup 6\ISCC.exe" "Disp_Setup.iss"
```

## Результат

После сборки у вас будет:
- `dist\Disp.exe` - исполняемый файл приложения
- `installer\Disp_Setup.exe` - инсталлер для распространения

## Примечания

- Убедитесь, что Inno Setup установлен в стандартную папку
- Если путь к Inno Setup отличается, измените его в build_installer.bat
- Инсталлер включает все необходимые файлы и базы данных
- Создаются ярлыки в меню Пуск и на рабочем столе





