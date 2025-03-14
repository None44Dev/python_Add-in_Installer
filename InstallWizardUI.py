from tkinter import *
from tkinter import ttk, messagebox
from screeninfo import get_monitors
from PIL import Image, ImageTk
import Install as ins  # Импорт модуля для установки макросов

class MainWindow:
    def __init__(self):
        # Инициализация главного окна приложения
        self.root = Tk()
        self.root.title('Установщик')  # Заголовок окна

        # Определение директории, откуда запускается программа
        self.exe_dir = ins.sys._MEIPASS if getattr(ins.sys, 'frozen', False) else ins.os.path.dirname(ins.os.path.abspath(__file__))

        # Настройка геометрии окна (центрирование на экране)
        self.get_geometry()

        # Запрет изменения размеров окна
        self.root.resizable(False, False)

        # Установка фонового цвета окна
        self.root.config(bg='#e6e5e5')

        # Добавление элементов интерфейса
        self.add_frame()  # Добавление разделительной линии
        self.add_image()  # Добавление изображения
        self.add_lable()  # Добавление текстовых меток

        # Создание кнопок "Отмена" и "Установить"
        self.cancel_button = CancelButton(self)
        self.install_button = InstallButton(self)

        # Установка иконки приложения
        self.root.iconbitmap(fr'{self.exe_dir}\icon.ico')

        # Запуск главного цикла обработки событий
        self.root.mainloop()

    def get_geometry(self, window_width=640, window_height=480):
        # Определение размеров экрана и центрирование окна
        monitor = get_monitors()[0]
        screen_width = monitor.width
        screen_height = monitor.height
        x = (screen_width - window_width) // 2
        y = (screen_height - window_height) // 2
        self.root.geometry(f'{window_width}x{window_height}+{x}+{y}')

    def add_frame(self):
        # Добавление горизонтальной линии (разделителя)
        canvas = Canvas(self.root,
                        bg='#a1a1a1',
                        height=1,
                        width=640)
        canvas.place(x=-2, y=398, anchor=NW)

    def add_image(self):
        # Загрузка и отображение изображения (логотипа)
        img = Image.open(fr'{self.exe_dir}\logo.png').convert('RGBA')
        img = img.resize((200, 400), Image.LANCZOS)  # Изменение размера изображения
        self.photo = ImageTk.PhotoImage(img)
        self.label_img = Label(self.root,
                               image=self.photo,
                               bd=0,
                               highlightthickness=0,
                               bg='white')
        self.label_img.place(x=0, y=0, anchor=NW)

    def add_lable(self):
        # Добавление текстовых меток
        self.head_label = Label(self.root,
                                text='Установка надстройки',
                                font=('Times New Roman', 18, 'bold'),
                                bg='#e6e5e5')
        
        self.info_label = Label(self.root,
                                text='Эта программа установит надстройку в Excel.',
                                font=('Times New Roman', 10),
                                bg='#e6e5e5')
        
        # Размещение меток на окне
        self.head_label.place(x=220, y=15, anchor=NW)
        self.info_label.place(x=220, y=50, anchor=NW)


class InstallButton:
    def __init__(self, main_window):
        # Инициализация кнопки "Установить" и прогресс-бара
        self.main_window = main_window

        # Настройка стиля прогресс-бара
        style = ttk.Style()
        style.theme_use('default')
        style.configure('Custom.Horizontal.TProgressbar',
                        roughcolor="#d6d7d6",
                        bordercolor="#a1a1a1",
                        background="#54e51f",
                        lightcolor="#00BFFF",
                        darkcolor="#00BFFF",
                        thickness=20)

        # Создание прогресс-бара
        self.progress_bar = ttk.Progressbar(main_window.root,
                                            orient=HORIZONTAL,
                                            length=350,
                                            mode='determinate',
                                            style='Custom.Horizontal.TProgressbar')

        # Создание кнопки "Установить"
        self.button = Button(main_window.root, text='Установить',
                             padx=20, relief="raised",
                             bd=3,
                             highlightbackground="gray",
                             highlightcolor="#00BFFF",
                             activebackground="#00ffff",
                             activeforeground="#000000", 
                             command=self.open_progress_window)
        self.button.place(relx=0.81, rely=0.95, anchor=SE)

    def open_progress_window(self):
        # Обработчик нажатия на кнопку "Установить"
        self.button.config(text='Готово', state=DISABLED)  # Блокировка кнопки
        self.main_window.cancel_button.button.config(state=DISABLED)  # Блокировка кнопки "Отмена"

        # Обновление текста меток
        self.main_window.head_label.config(text='Установка...',
                                           font=('Times New Roman', 16, 'bold'))
        self.main_window.info_label.config(text='Пожалуйста, подождите, пока надстройка установится.',
                                           font=('Times New Roman', 10))
        
        # Запуск процесса установки
        self.start_install()

    def start_install(self):
        # Начало установки: отображение прогресс-бара и вызов первого этапа
        self.progress_bar.place(x=220, y=80)
        self.install_macros = ins.InstallMacros()
        self.next_stage(1)

    def next_stage(self, stage):
        # Обработка каждого этапа установки
        try:
            match stage:
                case 1: self.install_macros.get_excel_version(); self.update_progress(15, stage + 1)
                case 2: self.install_macros.get_last_open(); self.update_progress(40, stage + 1)
                case 3: self.install_macros.copy_file(); self.update_progress(75, stage + 1)
                case 4: self.install_macros.set_registry_value(); self.update_progress(100, stage + 1)
                case 5: self.update_progress(101, None)
        except Exception as e:
            # Обработка ошибок
            self.show_error(str(e))

    def update_progress(self, value, stage):
        # Обновление прогресс-бара и переход к следующему этапу
        if value <= 100:
            self.progress_bar['value'] = value
            self.main_window.root.after(15, self.next_stage, stage)
        else:
            # Завершение установки
            self.main_window.head_label.config(text='Готово!...',
                                               font=('Times New Roman', 16, 'bold'))
            self.main_window.info_label.config(text='Установка завершена!',
                                               font=('Times New Roman', 10))
            self.button.config(text='Завершить', state='active', command=self.main_window.root.destroy)

    def show_error(self, message):
        # Отображение сообщения об ошибке
        messagebox.showerror("Ошибка", message)
        self.main_window.root.destroy


class CancelButton:
    def __init__(self, main_window):
        # Инициализация кнопки "Отмена"
        self.button = Button(main_window.root,
                             text='Отмена',
                             padx=20,
                             relief="raised",
                             bd=3,
                             highlightbackground="gray",
                             highlightcolor="#00BFFF",
                             activebackground="#00ffff",
                             activeforeground="#000000",
                             command=main_window.root.destroy)
        self.button.place(relx=0.97, rely=0.95, anchor=SE)


if __name__ == '__main__':
    # Запуск приложения
    MainWindow()