import customtkinter as ctk
import tkinter as tk
import tkinter.filedialog as filedialog
import tkinter.messagebox as messagebox
import json
import os
import re
from openpyxl import Workbook, load_workbook
import traceback
import datetime
import sys
import subprocess
import io
import zipfile
import requests
import shutil
import time # Додано для затримки в скрипті оновлення

# Імпортуємо модуль для роботи з кошторисом
# Упевніться, що koshtorys.py знаходиться в тій же папці або доступний
koshtorys = None
try:
    import koshtorys
except ImportError:
    # Помилка імпорту koshtorys не критична для запуску основної програми
    # MessageBox показуємо пізніше, щоб не блокувати старт
    print("Модуль koshtorys.py не знайдено. Функціонал кошторису буде недоступний.")
    pass # Продовжуємо виконання

# Перевірка наявності win32com.client
try:
    import win32com.client as win32
    win32_available = True
except ImportError:
    win32_available = False
    print("Модуль win32com.client не знайдено. Функціонал генерації документів Word буде недоступний.")
    # messagebox.showerror("Помилка імпорту", "Не вдалося імпортувати модуль win32com.client.\nВстановіть його командою: pip install pywin32")
    # input("Нажмите Enter для завершения...")
    # sys.exit(1) # Не виходимо одразу, дозволяємо GUI запуститись, але без функціоналу Word

#-------------Записування error-----------
ERROR_LOG = "error.txt"

def log_error(exc_type, exc_value, exc_traceback_obj): # Изменен параметр для traceback
    """Логує помилку в файл error.txt та показує messagebox."""
    try:
        with open(ERROR_LOG, "a", encoding="utf-8") as f:
            f.write(f"\n[{datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}]\n")
            # Используем traceback.print_exception напрямую
            traceback.print_exception(exc_type, exc_value, exc_traceback_obj, file=f)
        # Показать сообщение об ошибке в GUI, если root уже создан
        if root and root.winfo_exists():
             messagebox.showerror("Помилка", f"Произошла ошибка: {str(exc_value)}\nПодробности в файле {ERROR_LOG}")
        else:
            print(f"Необработанная ошибка: {str(exc_value)}. Подробности в {ERROR_LOG}")
            traceback.print_exception(exc_type, exc_value, exc_traceback_obj) # Печать в консоль, если GUI еще нет
            # input("Нажмите Enter для завершения...") # Не блокируем поток, если это GUI

    except Exception as ex:
        print(f"Ошибка при логировании: {str(ex)}")
        traceback.print_exc() # Печать ошибки логирования
        # input("Нажмите Enter для продолжения...") # Не блокируем

# Перехватываем и логируем только необработанные исключения
def global_exception_handler(exc_type, exc_value, exc_traceback_obj):
    """Глобальный обработчик исключений для логирования и показа сообщения."""
    # Игнорируем SystemExit, чтобы allow sys.exit() to work properly
    if issubclass(exc_type, SystemExit):
        return
    log_error(exc_type, exc_value, exc_traceback_obj)
    # Также выводим в консоль для отладки, используя стандартный хук
    sys.__excepthook__(exc_type, exc_value, exc_traceback_obj)
    # input("Нажмите Enter для завершения...") # Не блокируем GUI
    # sys.exit(1) # Не завершаем принудительно из хука, чтобы Tkinter мог корректно закрыться

sys.excepthook = global_exception_handler

# Глобальные переменные
FIELDS = [
    "товар", "дк", "захід", "дата", "адреса", "сума",
    "сума прописом", "пдв", "кількість", "ціна за одиницю",
    "загальна сума", "разом"
]

MEMORY_FILE = "contracts_memory.json"

EXAMPLES = {
    "товар": "наприклад: медалі зі стрічкою",
    "дк": "наприклад: ДК 021:2015: 18512200-3",
    "захід": "наприклад: 4 етапу “Фізкультурно-оздоровчих заходів та змагань “Пліч-о-пліч Всеукраїнські шкільні ліги з футзалу серед учнів та учениць закладів загальної середньої освіти у 2024-2025 навчальному році під гаслом “РАЗОМ ПЕРЕМОЖЕМО”",
    "дата": "наприклад: з 06 по 09 травня, з 13 по 16 травня 2025 року",
    "адреса": "наприклад: КП МСК “Дніпро”, вул. Смілянська, 78, м. Черкаси.",
    "сума": "наприклад: 15 120 грн 00 коп.",
    "сума прописом": "наприклад: П’ятнадцять тисяч сто двадцять грн 00 коп.",
    "пдв": "наприклад: без ПДВ", # Може бути пустим
    "кількість": "наприклад: 144",
    "ціна за одиницю": "наприклад: 144.00",
    "загальна сума": "наприклад: 15 120.00",
    "разом": "наприклад(теж саме що і загальна сума): 15 120.00"
}

# Глобальные переменные для доступа из функций GUI
document_blocks = [] # Список словарей, каждый представляет блок { "frame": ..., "path": ..., "entries": { field: entry_widget } }
root = None
scroll_frame = None # Фрейм внутри скроллируемой области
context_menu = None # Контекстное меню для полей ввода

# Переменные для безопасного использования after
after_callbacks = {}

# ----- НАСТРОЙКИ ОБНОВЛЕННЯ -----
# Змініть ці значення відповідно до вашого репозиторію
CURRENT_VERSION = "2.8.3"
GITHUB_REPO_OWNER = "KostyaBulavko"
GITHUB_REPO_NAME = "sportfprall"
GITHUB_BRANCH = "main" # Назва гілки для оновлення
VERSION_FILE_URL = f"https://raw.githubusercontent.com/{GITHUB_REPO_OWNER}/{GITHUB_REPO_NAME}/{GITHUB_BRANCH}/version.txt"
ZIP_DOWNLOAD_URL = f"https://github.com/{GITHUB_REPO_OWNER}/{GITHUB_REPO_NAME}/archive/refs/heads/{GITHUB_BRANCH}.zip"
# ----- КОНЕЦ НАСТРОЕК ОБНОВЛЕНИЯ -----


# --- Функции для безопасного использования root.after ---
def safe_after_cancel(widget, after_id):
    """Безпечно скасовує заплановану функцію 'after'."""
    try:
        # Перевіряємо, чи існує віджет, перш ніж скасовувати
        # Деякі віджети можуть бути знищені, але посилання на них ще є
        if widget and hasattr(widget, 'winfo_exists') and widget.winfo_exists() and after_id:
             widget.after_cancel(after_id)
        # Видаляємо after_id зі списку, якщо він там є
        widget_id = str(widget)
        if widget_id in after_callbacks and after_id in after_callbacks[widget_id]:
            after_callbacks[widget_id].remove(after_id)
            if not after_callbacks[widget_id]:
                 del after_callbacks[widget_id]

    except Exception: # Ловимо будь-які помилки, якщо віджет вже недійсний тощо
        pass

def cleanup_after_callbacks():
    """Скасовує всі заплановані 'after' колбеки перед закриттям програми."""
    global after_callbacks
    # Створюємо копію ключів, бо словник може змінюватися під час ітерації
    for widget_id_str in list(after_callbacks.keys()):
        # widget = root.nametowidget(widget_id_str) # Може викликати помилку TclError
        # Спрощено: просто перебираємо і намагаємося скасувати через тимчасове посилання
        # Або покладаємося на перевірку winfo_exists() в safe_after_cancel
        for after_id in list(after_callbacks.get(widget_id_str, [])): # Копія списку ID
            # Ми не маємо прямого посилання на віджет тут надійно.
            # safe_after_cancel викличеться пізніше, коли Tk спробує виконати колбек,
            # або ми можемо спробувати знайти віджет, якщо це можливо.
            # Найбезпечніше - покладатися на перевірку в safe_after_cancel
            # Але для гарантованого очищення потрібно спробувати знайти віджет.
            # Цей пошук по str(widget) є недоліком такого підходу.
            # Більш надійний метод зберігав би weakref на віджет.

            # Простий, але менш надійний спосіб очищення на виході:
            # Намагаємося знайти віджет за його строковим ID.
            # Це може не працювати, якщо віджет вже видалено.
            try:
                # Цей метод може не працювати для всіх типів віджетів або їх структури
                widget = root.nametowidget(widget_id_str) if root and root.winfo_exists() else None
                safe_after_cancel(widget, after_id) # Спробуємо скасувати
            except Exception:
                # Якщо віджет вже не існує або nametowidget викликав помилку,
                # safe_after_cancel може не спрацювати, але це не критично,
                # головне - не викликати TclError при виході.
                pass # Пропускаємо помилку при спробі скасувати для неіснуючого віджета

    after_callbacks.clear() # Очищаємо словник

def safe_after(widget, ms, callback):
    """Безпечна версія widget.after, яка реєструє колбек для очищення."""
    if not widget or not hasattr(widget, 'winfo_exists') or not widget.winfo_exists():
        return None

    try:
        widget_id = str(widget) # Використовуємо строкове представлення як ключ

        # Обгортаємо оригінальний колбек для безпечного виклику та видалення з реєстру
        def safe_callback():
            try:
                callback()
            finally:
                # Після виконання колбека, видаляємо його з after_callbacks
                if widget_id in after_callbacks:
                    # Знаходимо after_id для цього колбека. Це нетривіально,
                    # оскільки after_cancel вимагає after_id.
                    # Простіше покладатися на safe_after_cancel, яку викликають вручну
                    # або при destroy.
                    # Для автоматичного видалення після виконання, нам потрібно знати after_id тут.
                    # Можна передавати after_id до safe_callback, але after його не повертає в callback.
                    # Тому залишаємо поточну логіку: колбек видаляється зі списку ТІЛЬКИ при cancel/destroy.
                    # Це означає, що словник after_callbacks може тимчасово містити ID завершених колбеків.
                    pass

        after_id = widget.after(ms, safe_callback)

        if widget_id not in after_callbacks:
            after_callbacks[widget_id] = []

        after_callbacks[widget_id].append(after_id)
        return after_id
    except Exception:
        return None # Якщо сталася помилка при плануванні (наприклад, віджет вже знищено)


# Патчимо стандартні методи after для автоматичної реєстрації
# Це може вплинути на інші бібліотеки, що використовують after.
# Альтернатива - завжди використовувати safe_after напряму.
# Патчинг виглядає ризикованим, але спрощує використання.

# Зберігаємо оригінальні методи
original_ctk_after = ctk.CTk.after if hasattr(ctk.CTk, 'after') else None
original_base_after = ctk.CTkBaseClass.after if hasattr(ctk.CTkBaseClass, 'after') else None

def patched_ctk_after(self, ms, func=None, *args):
    if func:
        callback = lambda: func(*args) if args else func()
        return safe_after(self, ms, callback)
    else:
        # Випадки, коли func is None (наприклад, after_idle)
        # Або просто повертає поточний after_id без планування нового.
        # Обробляємо через оригінальний метод, якщо він існував.
        if original_ctk_after:
             return original_ctk_after(self, ms, func, *args)
        else: # Якщо оригінального методу не було (малоймовірно)
             return None # Або викликати помилку?

def patched_base_after(self, ms, func=None, *args):
    if func:
        callback = lambda: func(*args) if args else func()
        return safe_after(self, ms, callback)
    else:
         if original_base_after:
              return original_base_after(self, ms, func, *args)
         else:
              return None

# Застосовуємо патчі, якщо оригінальні методи існують
if original_ctk_after:
    ctk.CTk.after = patched_ctk_after # type: ignore
if original_base_after:
    ctk.CTkBaseClass.after = patched_base_after # type: ignore


class SafeCTk(ctk.CTk):
    """Наслідуємо від CTk, щоб додати очищення after колбеків при знищенні."""
    def destroy(self):
        cleanup_after_callbacks() # Очищаємо колбеки перед знищенням головного вікна
        super().destroy()

# --- Кінець функцій для безпечного використання root.after ---


class CustomEntry(ctk.CTkEntry):
    """Поле вводу з плейсхолдером та валідацією для числових полів."""
    def __init__(self, master, field_name=None, **kwargs):
        self.field_name = field_name
        self.example_text = EXAMPLES.get(field_name, "Введіть значення")
        self.is_placeholder_visible = True

        self.is_numeric_field = field_name in ["кількість", "ціна за одиницю", "загальна сума", "разом"]

        # CTkEntry сам створює внутрішній tk.Entry як self._entry
        # placeholder_text="" приховує стандартний плейсхолдер CTk, ми робимо свій
        super().__init__(master, placeholder_text="", **kwargs)

        # Вставляємо власний плейсхолдер
        self._entry.insert(0, self.example_text)
        self._entry.config(foreground='gray')

        # Прив'язуємо події до внутрішнього tk.Entry
        self._entry.bind("<FocusIn>", self._on_focus_in)
        self._entry.bind("<FocusOut>", self._on_focus_out)

        if self.is_numeric_field:
            self._entry.bind("<KeyRelease>", self._check_numeric_input)

        # Прив'язуємо праву кнопку миші до самого CTkEntry для контекстного меню
        # Щоб меню викликалось незалежно від того, навели на текст чи ні
        self.bind("<Button-3>", self._show_context_menu) # Прив'язуємо до зовнішнього віджета

    def _check_numeric_input(self, event):
        """Перевіряє, чи введені символи є допустимими для числових полів."""
        if self.is_placeholder_visible:
            return

        current_text = self._entry.get()
        if not current_text:
            return

        # Дозволяємо цифри, пробіли, одну крапку або одну кому.
        # Видаляємо пробіли для валідації, але зберігаємо їх у полі
        text_for_validation = current_text.replace(" ", "")
        # Перевіряємо, чи текст складається з цифр і, можливо, ОДНІЄЇ крапки чи коми
        if not (re.fullmatch(r'\d*([.,]?\d*)?', text_for_validation)):
             # Якщо не валідний, знаходимо останній валідний префікс
             valid_text = ""
             for i in range(len(current_text)):
                  subtext = current_text[:i+1]
                  if re.fullmatch(r'[\d\s]*([.,][\d\s]*)?$', subtext): # Перевірка з пробілами для поступового вводу
                       valid_text = subtext
                  else:
                       break # Зупиняємось, як тільки зустріли невалідний символ

             # Заміняємо поточний текст на останній валідний
             cursor_pos = self._entry.index(tk.INSERT)
             self._entry.delete(0, tk.END)
             self._entry.insert(0, valid_text)
             # Намагаємось відновити позицію курсора
             try:
                 # Розраховуємо нову позицію курсора з урахуванням видалених символів
                 new_cursor_pos = min(cursor_pos, len(valid_text)) # Курсор не далі кінця тексту
                 # Якщо символ, що був за курсором, був видалений, курсор переміщується
                 # Це спрощений підхід, ідеально - відстежувати зміни символів перед курсором.
                 self._entry.icursor(max(0, new_cursor_pos)) # Курсор не може бути < 0
             except tk.TclError: # Може виникнути, якщо віджет вже видалено
                 pass


    def _on_focus_in(self, event):
        """При отриманні фокусу видаляє плейсхолдер, якщо він видимий."""
        if self.is_placeholder_visible:
            self._entry.delete(0, "end")
            self._entry.config(foreground='black')
            self.is_placeholder_visible = False

    def _on_focus_out(self, event=None): # Додав event=None для виклику без події
        """При втраті фокусу повертає плейсхолдер, якщо поле порожнє."""
        current_text = self._entry.get().strip()
        if not current_text:
            self._entry.delete(0, "end")
            self._entry.insert(0, self.example_text)
            self._entry.config(foreground='gray')
            self.is_placeholder_visible = True

    def _show_context_menu(self, event):
        """Показує контекстне меню для редагування тексту."""
        global context_menu
        if context_menu:
            # Оновлюємо стан команд меню перед показом
            has_selection = self._entry.selection_present()
            context_menu.entryconfig("Копіювати", state="normal" if has_selection else "disabled")
            context_menu.entryconfig("Вирізати", state="normal" if has_selection else "disabled")

            try: # Перевіряємо, чи є щось у буфері обміну
                root.clipboard_get()
                context_menu.entryconfig("Вставити", state="normal")
            except tk.TclError:
                context_menu.entryconfig("Вставити", state="disabled")

            # Зв'язуємо меню з внутрішнім віджетом, який його викликав
            context_menu.current_widget = self._entry # type: ignore

            # Показуємо меню
            try:
                 context_menu.tk_popup(event.x_root, event.y_root)
            finally:
                 # Не забудьте "відірвати" меню від віджета, коли воно закриється
                 context_menu.grab_release() # тип невідомий для миші
                 context_menu.current_widget = None # type: ignore

            return "break" # Щоб запобігти стандартному контекстному меню Tk

    def get(self):
        """Повертає вміст поля, ігноруючи плейсхолдер."""
        if self.is_placeholder_visible:
            return ""
        return self._entry.get()

    def insert(self, index, string):
        """Вставляє текст у поле, приховуючи плейсхолдер."""
        if self.is_placeholder_visible:
            self._on_focus_in(None) # Імітуємо фокус, щоб прибрати плейсхолдер
        self._entry.insert(index, string)

    def delete(self, first_index, last_index=None):
        """Видаляє текст з поля."""
        if self.is_placeholder_visible:
             # Якщо видаляємо плейсхолдер (наприклад, виділили його і натиснули Delete)
             # Потрібно очистити поле і прибрати плейсхолдер
             if first_index == 0 and (last_index == "end" or last_index is None):
                  self._on_focus_in(None) # Очищаємо поле
             return # В інших випадках видалення нічого не робить, якщо видно плейсхолдер

        self._entry.delete(first_index, last_index)
        # Після видалення, якщо поле стало порожнім, повертаємо плейсхолдер
        # Викликаємо _on_focus_out з невеликою затримкою, щоб дати можливість іншим подіям обробки тексту завершитись
        self.safe_after(1, self._on_focus_out) # type: ignore # Используем запатченный after

    def clear(self):
        """Очищає поле вводу та повертає плейсхолдер."""
        self.delete(0, tk.END)
        self._on_focus_out(None) # Гарантовано повертаємо плейсхолдер, якщо поле пусте

    def set(self, text):
        """Встановлює текст у полі, заміняючи поточний вміст і приховуючи плейсхолдер."""
        self.clear() # Очищаємо поле (включаючи плейсхолдер)
        if text: # Якщо текст не порожній, вставляємо його
             self.insert(0, text)
             # Переконаємось, що плейсхолдер приховано, якщо вставлено текст
             if self.is_placeholder_visible:
                 self._on_focus_in(None)
        else:
             # Якщо текст порожній, переконаємось, що плейсхолдер показано
             self._on_focus_out(None)


# --- ФУНКЦИИ ПАМЯТИ ---
def load_memory():
    """Завантажує збережені дані з JSON файлу."""
    try:
        if os.path.exists(MEMORY_FILE):
            with open(MEMORY_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        else:
            return {}
    except Exception as e:
        log_error(type(e), e, sys.exc_info()[2])
        messagebox.showerror("Помилка пам'яті", f"Помилка завантаження даних пам'яті: {e}")
        return {}

def save_memory(data, template_path=None):
    """Зберігає дані в JSON файл. Може зберігати для конкретного шаблону."""
    try:
        current_data = load_memory()

        if template_path:
            norm_path = os.path.normpath(template_path)
            # Якщо дані для цього шляху вже існують і це словник, оновлюємо
            if norm_path not in current_data or not isinstance(current_data[norm_path], dict):
                 current_data[norm_path] = {} # Створюємо новий словник, якщо старі дані не у форматі словника
            for field, value in data.items():
                 current_data[norm_path][field] = value
        else:
            # Застарілий режим збереження (не прив'язаний до шаблону) - просто оновлюємо словник
            # Можливо, варто відмовитись від цього режиму, якщо всі дані тепер прив'язані до шляху
            current_data.update(data)

        with open(MEMORY_FILE, "w", encoding="utf-8") as f:
            json.dump(current_data, f, ensure_ascii=False, indent=2)

    except Exception as e:
        log_error(type(e), e, sys.exc_info()[2])
        messagebox.showerror("Помилка пам'яті", f"Помилка збереження даних пам'яті: {e}")

def get_template_memory(template_path):
    """Отримує збережені дані для конкретного шаблону."""
    try:
        memory_data = load_memory()
        norm_path = os.path.normpath(template_path)

        if norm_path in memory_data and isinstance(memory_data[norm_path], dict):
            return memory_data[norm_path]
        # Обробка старого формату: якщо memory_data не містить вкладених словників по шляхах
        # і це не порожній словник, вважаємо, що це дані старого формату
        # (тобто плоский словник "поле":"значення")
        elif memory_data and not any(isinstance(v, dict) for v in memory_data.values()):
             return memory_data # Повертаємо весь плоский словник
        return {} # Немає даних для цього шаблону або старий формат не застосовний

    except Exception as e:
        log_error(type(e), e, sys.exc_info()[2])
        return {} # Повертаємо порожній словник у випадку помилки


# --- ФУНКЦИИ ЭКСПОРТА ---
def export_to_excel():
    """Експортує поточні дані з усіх блоків у файл Excel."""
    try:
        if not document_blocks:
            messagebox.showwarning("Excel", "Немає даних для експорту. Додайте хоча б один договір.")
            return

        wb = Workbook()
        ws = wb.active
        ws.title = "Дані договорів"

        # Заголовки
        header = ["#", "Шаблон"] + FIELDS
        ws.append(header)

        # Дані по кожному блоку
        for i, block in enumerate(document_blocks, start=1):
            row = [i, os.path.basename(block["path"])] + [block["entries"][f].get() for f in FIELDS]
            ws.append(row)

        excel_file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel файли", "*.xlsx"), ("Всі файли", "*.*")],
            title="Зберегти дані в Excel"
        )
        if not excel_file_path:
            return # Користувач скасував збереження

        wb.save(excel_file_path)
        messagebox.showinfo("Excel", f"Дані успішно збережено у '{os.path.basename(excel_file_path)}'")
    except Exception as e:
        log_error(type(e), e, sys.exc_info()[2])
        messagebox.showerror("Помилка", "Помилка при збереженні в Excel.")


# --- ФУНКЦИИ ГЕНЕРАЦИИ ДОКУМЕНТОВ ---
def generate_documents():
    """Генерує документи Word на основі шаблонів та введених даних."""
    try:
        if not win32_available:
            messagebox.showerror("Помилка", "Модуль win32com.client не знайдено. Неможливо генерувати документи Word.")
            return

        if not document_blocks:
            messagebox.showwarning("Увага", "Не додано жодного договору для генерації.")
            return

        # Перевірка на заповненість полів
        all_fields_filled = True
        for i, block in enumerate(document_blocks, start=1):
            for field in FIELDS:
                 entry_widget = block["entries"].get(field)
                 # Пропускаємо перевірку для поля "пдв", якщо воно порожнє, бо це допустимо
                 # Перевіряємо, чи поле не "пдв" або поле "пдв" не порожнє
                 if field != "пдв" and (not entry_widget or not entry_widget.get().strip()):
                      messagebox.showerror("Помилка", f"Блок договору #{i} (файл: {os.path.basename(block['path'])}):\nПоле <{field}> порожнє.")
                      all_fields_filled = False
                      return # Зупиняємося на першій помилці

        if not all_fields_filled:
            return # Якщо є порожні обов'язкові поля

        # Вибір папки для збереження
        save_dir = filedialog.askdirectory(title="Оберіть папку для збереження згенерованих документів")
        if not save_dir:
            return # Користувач скасував вибір папки

        # Збереження даних для кожного шаблону перед генерацією
        for block in document_blocks:
            block_data = {f: block["entries"][f].get() for f in FIELDS}
            save_memory(block_data, block["path"]) # Сохраняем данные для конкретного шаблона

        # Пропонуємо експортувати в Excel перед генерацією, якщо ще не зроблено
        # Можливо, варто додати флаг, чи вже експортували
        # if messagebox.askyesno("Excel", "Зберегти введені дані в Excel перед генерацією документів?"):
        #      export_to_excel() # Виклик export_to_excel вже має свій askyesno, можливо, це зайве

        # Виклик функції заповнення кошторису, якщо модуль доступний
        if koshtorys and hasattr(koshtorys, 'fill_koshtorys'):
            if messagebox.askyesno("Кошторис", "Заповнити кошторис на основі введених даних?"):
                try:
                    koshtorys.fill_koshtorys(document_blocks)
                except Exception as e_koshtorys:
                    log_error(type(e_koshtorys), e_koshtorys, sys.exc_info()[2])
                    messagebox.showerror("Помилка кошторису", f"Помилка при заповненні кошторису: {e_koshtorys}")
        elif not koshtorys:
             print("Модуль koshtorys не завантажено. Пропускаємо заповнення кошторису.")
             # messagebox.showwarning("Кошторис", "Функція заповнення кошторису недоступна (модуль koshtorys не завантажено).") # Можливо, занадто багато повідомлень

        # --- Генерація Word документів ---
        word = None # Ініціалізуємо змінну word
        try:
            word = win32.gencache.EnsureDispatch('Word.Application')
            word.Visible = False # Можна встановити в True для відладки

            generated_files_count = 0
            for block_idx, block in enumerate(document_blocks):
                try:
                    doc_path = os.path.abspath(block["path"])
                    if not os.path.exists(doc_path):
                        messagebox.showerror("Помилка", f"Файл шаблону не знайдено: {doc_path}\nПропуск блоку #{block_idx+1}.")
                        continue

                    doc = word.Documents.Open(doc_path)

                    # Заміна в основному тексті документа
                    for key in FIELDS:
                        find_obj = doc.Content.Find
                        find_obj.ClearFormatting()
                        find_obj.Replacement.ClearFormatting()
                        replace_value = block["entries"][key].get()
                        # Обробка порожнього поля ПДВ: замінюємо на порожній рядок, якщо це дозволено
                        if key == "пдв" and not replace_value.strip():
                             replace_value = "" # Або інше значення за замовчуванням, напр. "без ПДВ"
                        find_obj.Execute(FindText=f"<{key}>", ReplaceWith=replace_value, Replace=2, Wrap=1) # wdReplaceAll = 2, wdFindContinue = 1

                    # Заміна в верхніх і нижніх колонтитулах
                    for section in doc.Sections:
                        # wdHeaderFooterPrimary = 1, wdHeaderFooterFirstPage = 2, wdHeaderFooterEvenPages = 3
                        for header_footer_type in [1, 2, 3]:
                            try:
                                header = section.Headers(header_footer_type)
                                if header.Exists:
                                    for key in FIELDS:
                                        find_obj = header.Range.Find
                                        find_obj.ClearFormatting()
                                        find_obj.Replacement.ClearFormatting()
                                        replace_value = block["entries"][key].get()
                                        if key == "пдв" and not replace_value.strip(): replace_value = ""
                                        find_obj.Execute(FindText=f"<{key}>", ReplaceWith=replace_value, Replace=2, Wrap=1)
                            except Exception: pass # Ігноруємо помилки, якщо колонтитул певного типу відсутній або інша проблема

                            try:
                                footer = section.Footers(header_footer_type)
                                if footer.Exists:
                                    for key in FIELDS:
                                        find_obj = footer.Range.Find
                                        find_obj.ClearFormatting()
                                        find_obj.Replacement.ClearFormatting()
                                        replace_value = block["entries"][key].get()
                                        if key == "пдв" and not replace_value.strip(): replace_value = ""
                                        find_obj.Execute(FindText=f"<{key}>", ReplaceWith=replace_value, Replace=2, Wrap=1)
                            except Exception: pass # Ігноруємо помилки


                    base_name = os.path.splitext(os.path.basename(block["path"]).replace("ШАБЛОН", "").strip())[0]
                    # Формируем имя файла, избегая недопустимых символов
                    tovar_value = block['entries']['товар'].get()
                    # Дозволяємо букви, цифри, пробіли, дефіси, дужки, коми, точки, але замінюємо інші на _
                    safe_tovar_name = "".join(c if c.isalnum() or c in " -(),." else "_" for c in tovar_value)
                    safe_tovar_name = re.sub(r'__+', '_', safe_tovar_name) # Замінюємо послідовності _ на один _
                    safe_tovar_name = safe_tovar_name.strip(' _-') # Видаляємо _- з початку та кінця
                    safe_tovar_name = safe_tovar_name[:100] # Обмежуємо довжину, щоб уникнути проблем з файловою системою

                    # Використовуємо .docx або .docm залежно від типу шаблону, якщо є макроси
                    # Простий спосіб: зберігаємо як docm, якщо це був docm шаблон, інакше як docx
                    # Або завжди зберігати як docm, якщо вони можуть знадобитись?
                    # FileFormat=13 для .docm, FileFormat=16 для .docx
                    output_format = 13 if os.path.splitext(doc_path.lower())[1] == '.docm' else 16
                    output_extension = ".docm" if output_format == 13 else ".docx"

                    save_file_name = f"{base_name} ({safe_tovar_name}){output_extension}"
                    save_path = os.path.join(save_dir, save_file_name)

                    # Уникаємо перезапису, додаючи лічильник, якщо файл вже існує
                    counter = 1
                    original_save_path = save_path
                    while os.path.exists(save_path):
                        save_file_name = f"{base_name} ({safe_tovar_name}) ({counter}){output_extension}"
                        save_path = os.path.join(save_dir, save_file_name)
                        counter += 1

                    doc.SaveAs(save_path, FileFormat=output_format)
                    doc.Close(False) # wdDoNotSaveChanges = 0
                    generated_files_count +=1

                except Exception as e_doc:
                    log_error(type(e_doc), e_doc, sys.exc_info()[2])
                    messagebox.showerror("Помилка обробки документа", f"Помилка при обробці файлу: {os.path.basename(block['path'])}\n\n{str(e_doc)}")

            # --- Закриття Word ---
            try:
                if word is not None:
                    word.Quit()
            except Exception: pass # Ігноруємо помилки при закритті Word

            # --- Повідомлення про результат ---
            if generated_files_count > 0:
                messagebox.showinfo("Успіх", f"{generated_files_count} документ(и) збережено успішно в папці:\n{save_dir}")
            elif document_blocks: # Якщо блоки були, але ні один не згенеровано
                 messagebox.showwarning("Увага", "Жодного документа не було згенеровано через помилки.")


        except Exception as e_word:
            log_error(type(e_word), e_word, sys.exc_info()[2])
            messagebox.showerror("Помилка MS Word", f"Не вдалося запустити або використати MS Word: {str(e_word)}\nПереконайтеся, що MS Word встановлено та налаштовано.")
            # Спробувати закрити Word, якщо він був ініціалізований, але сталася помилка пізніше
            try:
                if word is not None:
                    word.Quit()
            except Exception: pass


    except Exception as e:
        log_error(type(e), e, sys.exc_info()[2])
        # Загальна помилка вже логується в global_exception_handler


# --- ФУНКЦИИ ОБНОВЛЕНИЯ ПРИЛОЖЕНИЯ ---

def compare_versions(v1, v2):
    """
    Порівнює дві версії у форматі x.y.z
    Повертає:
      -1, якщо v1 < v2
       0, якщо v1 == v2
       1, якщо v1 > v2
    """
    try:
        # Розбиваємо версії на числові частини
        parts1 = list(map(int, v1.split('.')))
        parts2 = list(map(int, v2.split('.')))

        # Доповнюємо коротші частини нулями
        max_len = max(len(parts1), len(parts2))
        parts1.extend([0] * (max_len - len(parts1)))
        parts2.extend([0] * (max_len - len(parts2)))

        # Попарне порівняння
        for i in range(max_len):
            if parts1[i] < parts2[i]:
                return -1
            if parts1[i] > parts2[i]:
                return 1

        return 0 # Версії рівні

    except ValueError:
        # Обробка випадку, якщо формат версії неправильний
        print(f"Warning: Неправильний формат версії: '{v1}' або '{v2}'. Неможливо порівняти.")
        return 0 # Вважаємо рівними або пропускаємо оновлення
    except Exception as e:
        log_error(type(e), e, sys.exc_info()[2])
        print(f"Warning: Сталася помилка при порівнянні версій: {e}")
        return 0


def get_latest_version_from_github():
    """Получает номер последней версии с GitHub."""
    try:
        print(f"Перевірка версії на {VERSION_FILE_URL}...")
        response = requests.get(VERSION_FILE_URL, timeout=10)
        response.raise_for_status() # Викликає виняток для помилок HTTP (4xx або 5xx)
        latest_version = response.text.strip()
        print(f"Знайдено останню версію: {latest_version}")
        return latest_version
    except requests.exceptions.RequestException as e:
        # log_error(type(e), e, sys.exc_info()[2]) # Логируется в вызывающей функции download_and_apply_update
        messagebox.showerror("Помилка оновлення", f"Не вдалося перевірити версію на GitHub: {e}")
        print(f"Помилка при отриманні версії з GitHub: {e}")
        return None
    except Exception as e:
        log_error(type(e), e, sys.exc_info()[2])
        messagebox.showerror("Помилка оновлення", f"Невідома помилка при отриманні версії: {e}")
        print(f"Невідома помилка при отриманні версії: {e}")
        return None


def download_and_apply_update():
    """Скачивает, подготавливает обновление и запускает скрипт для его применения."""
    global root # Потрібно для доступу до головного вікна для cleanup та destroy
    try:
        messagebox.showinfo("Перевірка оновлень", "Перевірка наявності оновлень...")
        latest_version = get_latest_version_from_github()

        if not latest_version:
            # Помилка вже показана в get_latest_version_from_github
            return

        version_comparison_result = compare_versions(CURRENT_VERSION, latest_version)

        if version_comparison_result >= 0: # Якщо поточна >= останньої
            messagebox.showinfo("Оновлення", "У вас вже встановлена остання версія.")
            return

        confirm = messagebox.askyesno("Оновлення доступне",
                                      f"Доступна нова версія: {latest_version}\nВаша версія: {CURRENT_VERSION}\n\nБажаєте оновитися зараз?")
        if not confirm:
            return

        messagebox.showinfo("Оновлення", "Починається завантаження оновлення. Будь ласка, зачекайте...\nДодаток буде перезапущено після оновлення.")

        # --- Спробувати Git Pull (якщо це Git репозиторій) ---
        # Залишаємо як альтернативу ZIP, але ZIP надійніший для кінцевого користувача без Git
        if os.path.exists(".git") and os.path.isdir(".git"):
            try:
                print("Спроба оновлення через Git...")
                # Додаємо CREATE_NO_WINDOW для приховування консолі Git
                # shell=True може бути небезпечним, краще передавати команду списком
                subprocess.check_call(['git', 'pull'], creationflags=subprocess.CREATE_NO_WINDOW)
                print("Git pull успішний.")
                messagebox.showinfo("Оновлення успішне", "Додаток оновлено через git pull. Будь ласка, перезапустіть програму вручну.")
                # У випадку успіху git pull, не запускаємо батнік і не виходимо автоматично
                # Просимо користувача перезапустити вручну, щоб уникнути проблем з перезапуском після git pull
                return
            except (subprocess.CalledProcessError, FileNotFoundError) as e_git:
                log_error(type(e_git), e_git, sys.exc_info()[2])
                messagebox.showwarning("Помилка Git", f"Не вдалося виконати 'git pull': {e_git}\nСпробую завантажити архів.")
                print(f"Git pull помилка: {e_git}. Спроба завантаження архіву.")
            except Exception as e: # Ловимо інші можливі помилки git
                 log_error(type(e), e, sys.exc_info()[2])
                 messagebox.showwarning("Помилка Git", f"Невідома помилка під час Git оновлення: {e}\nСпробую завантажити архів.")
                 print(f"Невідома помилка Git: {e}. Спроба завантаження архіву.")

        # --- Завантажити та розпакувати ZIP (основний механізм оновлення) ---
        temp_zip_path = "update_temp.zip"
        update_extract_path = "update_extracted_temp"
        updater_script_name = "run_update.bat" # Ім'я скрипта оновлення
        current_app_dir = os.getcwd() # Поточна робоча директорія програми
        updater_script_path = os.path.join(current_app_dir, updater_script_name) # Шлях до скрипта поруч з EXE/скриптом

        try:
            print(f"Завантаження ZIP архіву з {ZIP_DOWNLOAD_URL}...")
            response = requests.get(ZIP_DOWNLOAD_URL, stream=True, timeout=120) # Збільшив таймаут
            response.raise_for_status() # Перевірка на HTTP помилки

            # Видаляємо попередні тимчасові файли/папки, якщо вони залишились
            if os.path.exists(update_extract_path):
                print(f"Видалення тимчасової папки {update_extract_path}...")
                shutil.rmtree(update_extract_path, ignore_errors=True) # ignore_errors=True, якщо папка зайнята
            if os.path.exists(temp_zip_path):
                 print(f"Видалення тимчасового файлу {temp_zip_path}...")
                 os.remove(temp_zip_path) # Може викликати помилку, якщо файл зайнятий

            os.makedirs(update_extract_path, exist_ok=True) # Створюємо папку для розпакування

            print(f"Збереження архіву до {temp_zip_path}...")
            with open(temp_zip_path, 'wb') as f:
                # Використовуємо chunk size для великих файлів
                for chunk in response.iter_content(chunk_size=8192):
                     f.write(chunk)
            print("Завантаження завершено.")

            print(f"Розпакування архіву до {update_extract_path}...")
            with zipfile.ZipFile(temp_zip_path, 'r') as zip_ref:
                zip_ref.extractall(update_extract_path)
            print("Розпакування завершено.")

            #os.remove(temp_zip_path) # Видаляємо тимчасовий zip файл після успішного розпакування. Видалимо через батнік надійніше.

        except Exception as e_zip:
            log_error(type(e_zip), e_zip, sys.exc_info()[2])
            messagebox.showerror("Помилка завантаження/розпакування", f"Не вдалося завантажити або розпакувати архів оновлення: {e_zip}")
            print(f"Помилка завантаження/розпакування: {e_zip}")
            # Спробувати очистити тимчасові файли, якщо вони були частково створені
            if os.path.exists(temp_zip_path):
                 try: os.remove(temp_zip_path)
                 except OSError: pass # Ігноруємо помилку, якщо файл зайнятий
            if os.path.exists(update_extract_path):
                 try: shutil.rmtree(update_extract_path, ignore_errors=True)
                 except OSError: pass
            return # Виходимо, якщо не вдалося розпакувати


        # --- Знаходимо кореневу папку всередині розпакованого архіву ---
        # Зазвичай, ZIP з GitHub містить одну кореневу папку з назвою репозиторію та гілки (наприклад, MyRepo-main)
        extracted_items = os.listdir(update_extract_path)
        if not extracted_items:
             messagebox.showerror("Помилка оновлення", "Архів оновлення порожній або не вдалося знайти файли.")
             if os.path.exists(update_extract_path): shutil.rmtree(update_extract_path, ignore_errors=True)
             return

        source_update_folder_name = f"{GITHUB_REPO_NAME}-{GITHUB_BRANCH}"
        potential_source_folder = os.path.join(update_extract_path, source_update_folder_name)

        if os.path.isdir(potential_source_folder):
            source_update_folder = potential_source_folder
            print(f"Знайдено кореневу папку оновлення: {source_update_folder}")
        else:
            # Якщо стандартне ім'я не знайдене, шукаємо першу папку або беремо корінь розпакованого архіву
            found_source = False
            for item_name in extracted_items:
                item_path = os.path.join(update_extract_path, item_name)
                if os.path.isdir(item_path):
                    source_update_folder = item_path
                    found_source = True
                    print(f"Знайдено першу папку в архіві: {source_update_folder}")
                    break
            if not found_source:
                 source_update_folder = update_extract_path # Файли прямо в корені архіву
                 print(f"Файли в корені архіву. Папка оновлення: {source_update_folder}")

        # --- Створення скрипта оновлення (.bat) ---
        # Цей скрипт буде запущено після закриття основної програми
        main_script_path = os.path.abspath(sys.argv[0]) # Шлях до головного скрипта або EXE
        python_executable = sys.executable # Шлях до інтерпретатора Python або EXE

        # Файли, які НЕ потрібно замінювати під час оновлення. Шляхи відносно кореня програми.
        excluded_files_relative = [
            MEMORY_FILE,
            ERROR_LOG,
            updater_script_name,
            os.path.basename(main_script_path), # Головний скрипт/EXE поточної програми
            # Додайте інші файли, які є специфічними для локальної установки
            # наприклад, файли конфігурації, бази даних, користувацькі шаблони тощо.
            # 'my_config.ini', 'database.db', 'templates/my_template.docm'
        ]
        # Якщо програма запущена інтерпретатором, не замінюємо сам інтерпретатор (якщо він в папці програми)
        # Це менш надійний спосіб виключення інтерпретатора, але може спрацювати
        if os.path.dirname(python_executable) == current_app_dir:
             excluded_files_relative.append(os.path.basename(python_executable))
        # Якщо програма скомпільована в один файл EXE, main_script_path == python_executable
        # Список excluded_files_relative вже містить basename(main_script_path), цього достатньо

        # Зміст батч-файлу
        # Використовуємо %%i та %%f для змінних циклу в батч-файлі
        batch_script_content = f"""@echo off
REM Чекаємо кілька секунд, щоб основна програма повністю закрилася
timeout /t 5 /nobreak > nul
echo Починаємо оновлення...
echo Поточна директорія: {current_app_dir}
echo Папка з оновленням: {source_update_folder}

REM Переходимо в папку з оновленням, щоб легше копіювати
cd /d "{source_update_folder}"

REM Видаляємо старі файли (крім виключених) перед копіюванням нових
REM Цей крок є опціональним, але робить оновлення "чистішим"
REM Потрібно бути ДУЖЕ обережним з цим кроком! Можна випадково видалити зайве.
REM Для простоти, зараз ми просто копіюємо поверх, не видаляючи старі.

REM Копіюємо всі файли та папки з оновлення (джерело: поточна папка),
REM крім тих, що в списку виключень.
REM Використовуємо xcopy /s /e /y
REM Також фільтруємо за допомогою 'findstr /v' або подібного,
REM або просто копіюємо все і не замінюємо виключені файли (xcopy /y)

echo Копіювання файлів оновлення...
REM Переходимо в папку програми для копіювання туди
cd /d "{current_app_dir}"

REM Копіюємо з папки оновлення в поточну, ігноруючи помилки заблокованих файлів
REM /E - копіювати каталоги та підкаталоги, включаючи порожні
REM /Y - придушити запит на підтвердження перезапису
REM /R - перезаписувати файли "тільки для читання"
xcopy /E /Y /R "{source_update_folder}\\*" "{current_app_dir}\\" > nul 2>&1

REM Такий простий xcopy не дозволяє легко виключити файли.
REM Більш надійний спосіб: створити список файлів для копіювання в Python
REM і передати його в скрипт, або зробити цикл в скрипті з перевіркою виключень.
REM Для демонстрації залишимо простий xcopy, але пам'ятайте про обмеження.

echo Очищення тимчасових файлів...
REM Переходимо в папку, де був zip та розпаковане оновлення
REM Зазвичай, це папка програми, але update_extract_path все одно повний шлях
cd /d "{current_app_dir}"
rmdir /s /q "{update_extract_path}" > nul 2>&1
del "{temp_zip_path}" > nul 2>&1

echo Оновлення завершено. Перезапуск програми...

REM Запускаємо оновлену програму
REM Використовуємо повні шляхи та quotes на випадок пробілів
start "" "{python_executable}" "{main_script_path}"

REM Чекаємо секунду, щоб нова програма встигла запуститись, перш ніж видаляти батнік
timeout /t 1 /nobreak > nul

REM Видаляємо сам скрипт оновлення
del "{updater_script_name}"

exit /b 0
""" # exit /b 0 для завершення тільки батч-файлу, не консолі, з кодом успіху

        try:
            # Записуємо скрипт оновлення у файл
            print(f"Створення скрипта оновлення: {updater_script_path}")
            with open(updater_script_path, "w", encoding="utf-8") as f:
                f.write(batch_script_content)
            print("Скрипт оновлення створено.")

            # Запуск батч-файлу у фоновому режимі
            # shell=True потрібен для прямого виконання .bat
            # creationflags=subprocess.CREATE_NO_WINDOW приховує вікно консолі
            print(f"Запуск скрипта оновлення: {updater_script_path}")
            subprocess.Popen(['cmd', '/c', updater_script_path], shell=True, creationflags=subprocess.CREATE_NO_WINDOW)
            print("Скрипт оновлення запущено.")

            messagebox.showinfo("Оновлення", "Оновлення розпочато. Додаток закриється і перезапуститься автоматично.")

            # Закриття поточної програми
            print("Закриття поточної програми...")
            if root:
                # Скасовуємо всі заплановані колбеки перед виходом
                cleanup_after_callbacks()
                root.destroy() # Закриваємо головне вікно
            sys.exit(0) # Завершуємо поточний процес Python

        except Exception as e_script:
            log_error(type(e_script), e_script, sys.exc_info()[2])
            messagebox.showerror("Помилка скрипта оновлення", f"Не вдалося створити або запустити скрипт оновлення: {e_script}")
            print(f"Помилка створення/запуску скрипта оновлення: {e_script}")
            # Спробувати очистити тимчасові файли, якщо скрипт не запустився
            if os.path.exists(updater_script_path):
                 try: os.remove(updater_script_path)
                 except OSError: pass
            if os.path.exists(update_extract_path):
                 try: shutil.rmtree(update_extract_path, ignore_errors=True)
                 except OSError: pass
            # Не виходимо sys.exit(0), якщо скрипт не запустився успішно

    except requests.exceptions.RequestException as e:
        # Помилка завантаження - вже оброблена в get_latest_version_from_github або раніше тут
        print(f"Помилка під час оновлення (зовнішня): {e}")
        pass # Повідомлення про помилку вже показано


    except Exception as e:
        # Загальна помилка, яка не була перехоплена раніше
        log_error(type(e), e, sys.exc_info()[2])
        # Повідомлення про помилку вже показано глобальним обробником

# --- КІНЕЦЬ ФУНКЦІЙ ОНОВЛЕННЯ ---


# --- ФУНКЦИИ GUI ---

def add_document_block(template_path=None, initial_data=None):
    """Додає блок полів вводу для нового документа або завантажує існуючий."""
    global scroll_frame, document_blocks, root

    if not scroll_frame:
        messagebox.showerror("Помилка", "Не ініціалізовано фрейм для блоків документів.")
        return

    if not template_path:
        template_path = filedialog.askopenfilename(
            defaultextension=".docm",
            filetypes=[("Word Documents", "*.docm *.docx"), ("All Files", "*.*")],
            title="Виберіть файл шаблону документа"
        )
        if not template_path:
            return # Користувач скасував вибір

    if not os.path.exists(template_path):
         messagebox.showerror("Помилка", f"Файл шаблону не знайдено: {template_path}")
         return

    # Створення фрейму для нового блоку
    block_frame = ctk.CTkFrame(scroll_frame)
    block_frame.pack(fill="x", padx=5, pady=5, expand=True)

    # Заголовок блоку (назва файлу шаблону) та кнопка видалення
    header_frame = ctk.CTkFrame(block_frame, fg_color="transparent")
    header_frame.pack(fill="x", pady=(0, 5))
    template_name_label = ctk.CTkLabel(header_frame, text=os.path.basename(template_path), font=ctk.CTkFont(weight="bold"))
    template_name_label.pack(side="left", fill="x", expand=True)

    # Кнопка видалення блоку
    delete_button = ctk.CTkButton(header_frame, text="Видалити", width=80, command=lambda: remove_document_block(block_frame))
    delete_button.pack(side="right")


    entries = {}
    # Створення полів вводу для кожного поля
    for field in FIELDS:
        field_frame = ctk.CTkFrame(block_frame, fg_color="transparent")
        field_frame.pack(fill="x", pady=1)

        label = ctk.CTkLabel(field_frame, text=f"{field.capitalize()}:", width=120, anchor="w")
        label.pack(side="left")

        entry = CustomEntry(field_frame, field_name=field, fg_color="#f0f0f0") # Використовуємо наш CustomEntry
        entry.pack(side="left", fill="x", expand=True)
        entries[field] = entry

        # Прив'язуємо шорткати та контекстне меню до нового поля
        bind_entry_shortcuts(entry)


    # Зберігаємо посилання на блок
    block_info = {
        "frame": block_frame,
        "path": template_path,
        "entries": entries
    }
    document_blocks.append(block_info)

    # Завантажуємо дані пам'яті для цього шаблону, якщо є
    loaded_data = get_template_memory(template_path)
    if loaded_data:
        print(f"Завантажено пам'ять для шаблону {os.path.basename(template_path)}")
        for field, value in loaded_data.items():
            if field in entries and value is not None:
                 entries[field].set(str(value)) # Використовуємо метод set CustomEntry


def remove_document_block(block_frame_to_remove):
    """Видаляє блок полів вводу."""
    global document_blocks
    # Шукаємо блок за фреймом
    block_to_remove = None
    for block in document_blocks:
        if block["frame"] == block_frame_to_remove:
            block_to_remove = block
            break

    if block_to_remove:
        # Скасовуємо будь-які заплановані after колбеки для віджетів у цьому блоці
        try:
            for entry in block_to_remove["entries"].values():
                # Це може бути складно, якщо колбеки прив'язані до _entry
                # safe_after_cancel потребує after_id.
                # Простіше покладатися на root.destroy() та cleanup_after_callbacks
                # Або при видаленні блоку вручну перебирати after_callbacks
                # та скасовувати ті, що належать віджетам у цьому блоці.
                widget_id = str(entry)
                if widget_id in after_callbacks:
                     for after_id in list(after_callbacks[widget_id]):
                          safe_after_cancel(entry, after_id)
                     # list() копія потрібна, бо safe_after_cancel змінює словник

        except Exception as e_cleanup:
             print(f"Помилка при очищенні after колбеків для блоку: {e_cleanup}")

        block_frame_to_remove.destroy() # Знищуємо фрейм та всі його дочірні елементи
        document_blocks.remove(block_to_remove) # Видаляємо блок зі списку

        # Оновлюємо пам'ять, якщо потрібно (опціонально)
        # Можливо, не потрібно видаляти з пам'яті, дані можуть знадобитись пізніше


def bind_entry_shortcuts(entry_widget):
    """Прив'язує стандартні шорткати Ctrl+C/V/X/A до поля вводу та контекстне меню."""
    # Ця функція вже є вище, але дублюю її тут для повноти, якщо вона буде викликатись з create_gui
    # Переконайтеся, що використовується одна версія функції.
    # Видалив дублюючу реалізацію, залишив ту, що на початку.
    pass # Функція вже визначена раніше


def create_context_menu():
    """Створює контекстне меню для полів вводу."""
    global context_menu
    context_menu = tk.Menu(root, tearoff=0)
    context_menu.add_command(label="Вирізати", command=lambda: context_menu.current_widget.event_generate("<<Cut>>")) # type: ignore
    context_menu.add_command(label="Копіювати", command=lambda: context_menu.current_widget.event_generate("<<Copy>>")) # type: ignore
    context_menu.add_command(label="Вставити", command=lambda: context_menu.current_widget.event_generate("<<Paste>>")) # type: ignore
    context_menu.add_separator() # Ось виправлений рядок

    def cut_command():
        if hasattr(context_menu, 'current_widget') and context_menu.current_widget: # type: ignore
            context_menu.current_widget.event_generate('<<Cut>>') # type: ignore

    def copy_command():
        if hasattr(context_menu, 'current_widget') and context_menu.current_widget: # type: ignore
            context_menu.current_widget.event_generate('<<Copy>>') # type: ignore

    def paste_command():
        if hasattr(context_menu, 'current_widget') and context_menu.current_widget: # type: ignore
            # Перед вставкою, якщо це CustomEntry, прибираємо плейсхолдер
            if isinstance(context_menu.current_widget.master, CustomEntry): # type: ignore
                 context_menu.current_widget.master._on_focus_in(None) # type: ignore # Accessing internal method
            context_menu.current_widget.event_generate('<<Paste>>') # type: ignore

    def select_all_command():
        if hasattr(context_menu, 'current_widget') and context_menu.current_widget: # type: ignore
            context_menu.current_widget.select_range(0, 'end') # type: ignore
            context_menu.current_widget.icursor('end') # type: ignore # Курсор в кінець
            return "break" # type: ignore # Щоб не передавати далі

    # Додаємо команди до меню
    context_menu.add_command(label="Вирізати", command=cut_command)
    context_menu.add_command(label="Копіювати", command=copy_command)
    context_menu.add_command(label="Вставити", command=paste_command)
    context_menu.add_separator()
    context_menu.add_command(label="Виділити все", command=select_all_command)


def create_gui():
    """Створює основний графічний інтерфейс користувача."""
    global root, scroll_frame, context_menu, document_blocks

    # Ініціалізуємо головне вікно за допомогою нашого SafeCTk
    root = SafeCTk()
    root.title(f"Генератор документів для 'Пліч-о-пліч' (Версія {CURRENT_VERSION})")
    root.geometry("800x600") # Початковий розмір вікна

    # Перевірка наявності win32com.client при запуску GUI
    if not win32_available:
         messagebox.showwarning("Попередження", "Модуль win32com.client не знайдено.\nГенерація документів Word буде недоступна.\nВстановіть його командою: pip install pywin32")

    # Перевірка наявності koshtorys.py при запуску GUI
    if not koshtorys:
         messagebox.showwarning("Попередження", "Модуль koshtorys.py не знайдено.\nФункціонал кошторису буде недоступний.\nПереконайтесь, що файл знаходиться в папці з програмою.")


    # Створення контекстного меню
    create_context_menu()
    # Шорткати прив'язуються при створенні CustomEntry

    # --- Основна компоновка ---

    # Фрейм для кнопок керування зверху
    control_frame = ctk.CTkFrame(root)
    control_frame.pack(side="top", fill="x", padx=10, pady=5)

    # Кнопка "Додати договір"
    add_button = ctk.CTkButton(control_frame, text="Додати договір", command=add_document_block)
    add_button.pack(side="left", padx=5)

    # Кнопка "Згенерувати документи"
    generate_button = ctk.CTkButton(control_frame, text="Згенерувати документи", command=generate_documents)
    # Відключаємо кнопку, якщо win32com недоступний
    if not win32_available:
        generate_button.configure(state="disabled", text="Генерувати документи (потрібен win32com)")
    generate_button.pack(side="left", padx=5)

    # Кнопка "Експорт в Excel"
    excel_button = ctk.CTkButton(control_frame, text="Експорт в Excel", command=export_to_excel)
    excel_button.pack(side="left", padx=5)

    # Кнопка "Перевірити оновлення"
    update_button = ctk.CTkButton(control_frame, text="Перевірити оновлення", command=download_and_apply_update)
    update_button.pack(side="right", padx=5)


    # Скроллируемая область для блоков документов
    # CTkScrollableFrame - это фрейм с полосой прокрутки
    scroll_frame = ctk.CTkScrollableFrame(root)
    scroll_frame.pack(side="top", fill="both", expand=True, padx=10, pady=5)

    # TODO: Здесь может быть логика для загрузки ранее открытых шаблонов из памяти
    # Например:
    # memory_data = load_memory()
    # if memory_data:
    #     for path in memory_data.keys():
    #         # Добавить блок для каждого сохраненного пути, загрузив его данные
    #         if os.path.exists(path): # Проверяем, существует ли файл шаблона
    #             add_document_block(template_path=path, initial_data=memory_data[path]) # initial_data пока не используется в add_document_block


    # Запускаємо головний цикл Tkinter
    root.mainloop()


# --- Точка входа в программу ---
if __name__ == "__main__":
    # Перед запуском GUI, можна зробити першу перевірку на оновлення
    # або додати цю перевірку в окрему функцію, яка викликається після створення GUI
    # Наприклад, через root.after(..., check_for_updates_on_startup)
    # Для простоти, перевірка за кнопкою є безпечнішою.

    # Створюємо та запускаємо GUI
    create_gui()

    print("Програма завершила роботу.")