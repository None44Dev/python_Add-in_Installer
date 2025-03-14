import os
import shutil
import winreg
import sys
import win32com.client as Win_32

class InstallMacros:
    # Класс для установки макросов (VBA) в Excel.
    # Основные задачи:
    # 1. Копирование файла надстройки (.xlam) в папку AddIns.
    # 2. Настройка реестра Windows для автоматической загрузки надстройки в Excel.
    def __init__(self):
        # Инициализация объекта. Задаются основные пути и переменные:
        # - self.user_profile: Путь к профилю пользователя (например, C:\Users\Username).
        # - self.file: Имя файла надстройки (VBAProject_LMP.xlam).
        # - self.path_copy: Путь, куда будет скопирован файл надстройки (в папку AppData\Roaming\Microsoft\AddIns).
        # - self.exe_dir: Директория, откуда запускается программа (для упакованных приложений используется sys._MEIPASS).
        self.user_profile = os.environ['USERPROFILE']
        self.file = 'VBAProject_LMP.xlam'
        self.path_copy = fr'{self.user_profile}\AppData\Roaming\Microsoft\AddIns\{self.file}'
        self.exe_dir = sys._MEIPASS if getattr(sys, 'frozen', False) else os.path.dirname(os.path.abspath(__file__))

    def show_error(self, message):
        # Выводит ошибку и завершает выполнение программы.
        # :param message: Сообщение об ошибке.
        raise Exception(message)
    
    def get_excel_version(self):
        # Определяет версию Excel и задает путь в реестре для настройки надстройки.
        # - Использует win32com.client для получения версии Excel.
        # - Если Excel не установлен, выбрасывает ошибку.
        base_key = r"Software\Microsoft\Office"
        version = Win_32.Dispatch('Excel.Application').Version
        if version:
            self.path_winreg = fr'{base_key}\{version}\Excel\Options'
        else:
            self.show_error('Excel не установлен')

    def get_last_open(self):
        # Определяет следующий доступный ключ OPEN в реестре для добавления надстройки.
        # - Ищет существующие ключи OPEN в реестре.
        # - Если надстройка уже добавлена, использует существующий ключ.
        # - Если нет, создает новый ключ (например, OPEN, OPEN1, OPEN2 и т.д.).
        save_number = []
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, self.path_winreg) as key:
            i = 0
            try:
                while True:
                    name, value, _ = winreg.EnumValue(key, i)
                    if name.startswith('OPEN'):
                        if value == self.path_copy:
                            self.open = name
                            return
                        else:
                            save_number.append(int(name.replace('OPEN', "") or 0))
                    i += 1
            except OSError:
                pass

        if save_number:
            max_open = max(save_number)
            if max_open == 0:
                self.open = 'OPEN'
            else:
                self.open = f'OPEN{int(max_open)+1}'
        else:
            self.open = 'OPEN'
        return

    def copy_file(self):
        # Копирует файл надстройки из исходной директории в папку AddIns.
        # - Проверяет, существует ли файл надстройки.
        # - Если файл не найден, выбрасывает ошибку.
        # - Если копирование прошло успешно, файл будет доступен для Excel.
        addin_source_path = fr'{self.exe_dir}\{self.file}'
        if not os.path.exists(addin_source_path):
            self.show_error(fr'Файл не найден по пути: {addin_source_path}')
        try:
            shutil.copy2(addin_source_path, self.path_copy)
        except Exception as e:
            self.show_error(f'Ошибка при копировании надстройки: {e}')

    def set_registry_value(self):
        # Добавляет или обновляет значение в реестре для автоматической загрузки надстройки в Excel.
        # - Использует ключ OPEN, определенный в get_last_open.
        # - Если запись в реестр не удалась, выбрасывает ошибку.
        try:
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, self.path_winreg, 0, winreg.KEY_SET_VALUE) as reg_key:
                winreg.SetValueEx(reg_key, self.open, 0, winreg.REG_SZ, self.path_copy)
        except Exception as e:
            self.show_error(f'Ошибка при записи в реестр: {e}')