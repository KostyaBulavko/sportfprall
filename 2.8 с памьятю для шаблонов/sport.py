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

# Импортируем модуль для работы с кошторисом
import koshtorys

# Блок перехвата ошибок, чтобы консоль не закрывалась
try:
    # Попытка импорта win32com
    try:
        import win32com.client as win32
    except ImportError:
        messagebox.showerror("Ошибка импорта", "Не удалось импортировать модуль win32com.client.\nУстановите его командой: pip install pywin32")
        input("Нажмите Enter для завершения...")
        sys.exit(1)

    #-------------Записування error-----------
    ERROR_LOG = "error.txt"

    def log_error(exc_type, exc_value, exc_traceback):
        try:
            with open(ERROR_LOG, "a", encoding="utf-8") as f:
                f.write(f"\n[{datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}]\n")
                traceback.print_exception(exc_type, exc_value, exc_traceback, file=f)
            # Показать сообщение об ошибке
            messagebox.showerror("Ошибка", f"Произошла ошибка: {str(exc_value)}\nПодробности в файле {ERROR_LOG}")
        except Exception as ex:
            print(f"Ошибка при логировании: {str(ex)}")
            input("Нажмите Enter для продолжения...")

    sys.excepthook = log_error

    # Глобальные переменные
    FIELDS = [
        "товар", "дк", "захід", "дата", "адреса", "сума",
        "сума прописом", "пдв", "кількість", "ціна за одиницю",
        "загальна сума", "разом"
    ]
    
    # Изменено: новое название файла для хранения данных договоров
    MEMORY_FILE = "contracts_memory.json"
    
    EXAMPLES = {
    "товар": "наприклад: медалі зі стрічкою",
    "дк": "наприклад: ДК 021:2015: 18512200-3",
    "захід": "наприклад: 4 етапу “Фізкультурно-оздоровчих заходів та змагань “Пліч-о-пліч Всеукраїнські шкільні ліги з футзалу серед учнів та учениць закладів загальної середньої освіти у 2024-2025 навчальному році під гаслом “РАЗОМ ПЕРЕМОЖЕМО”",
    "дата": "наприклад: з 06 по 09 травня, з 13 по 16 травня 2025 року",
    "адреса": "наприклад: КП МСК “Дніпро”, вул. Смілянська, 78, м. Черкаси.",
    "сума": "наприклад: 15 120 грн 00 коп.",
    "сума прописом": "наприклад: П’ятнадцять тисяч сто двадцять грн 00 коп.",
    "пдв": "наприклад: без ПДВ",
    "кількість": "наприклад: 144",
    "ціна за одиницю": "наприклад: 144.00",
    "загальна сума": "наприклад: 15 120.00",
    "разом": "наприклад(теж саме що і загальна сума): 15 120.00"
    }

    # Глобальные переменные для доступа из функций
    document_blocks = []
    root = None
    scroll_frame = None
    context_menu = None
    
    # Словарь для хранения активных after ID
    after_callbacks = {}

    def safe_after_cancel(widget, after_id):
        """Безопасная отмена отложенного вызова"""
        try:
            if widget and after_id:
                widget.after_cancel(after_id)
        except Exception:
            pass  # Игнорировать ошибки при отмене

    def cleanup_after_callbacks():
        """Очистка всех отложенных вызовов"""
        global after_callbacks
        for widget_id in list(after_callbacks.keys()):
            for after_id in after_callbacks.get(widget_id, []):
                try:
                    widget = ctk._get_widget_by_id(widget_id)
                    if widget:
                        safe_after_cancel(widget, after_id)
                except Exception:
                    pass  # Игнорировать ошибки
        after_callbacks.clear()

    def safe_after(widget, ms, callback):
        """Безопасный вызов after с отслеживанием идентификаторов"""
        if not widget:
            return None
            
        try:
            widget_id = str(widget)
            after_id = widget.after(ms, callback)
            
            if widget_id not in after_callbacks:
                after_callbacks[widget_id] = []
                
            after_callbacks[widget_id].append(after_id)
            return after_id
        except Exception:
            return None

    # Переопределяем методы CustomTkinter
    original_after = ctk.CTk.after
    def patched_after(self, ms, func=None, *args):
        """Патч метода after для отслеживания вызовов"""
        if func:
            callback = lambda: func(*args) if args else func()
            return safe_after(self, ms, callback)
        else:
            return original_after(self, ms, func, *args)
            
    ctk.CTk.after = patched_after
    
    # Патч для CTkBaseClass
    original_base_after = ctk.CTkBaseClass.after
    def patched_base_after(self, ms, func=None, *args):
        """Патч метода after для отслеживания вызовов в базовых классах"""
        if func:
            callback = lambda: func(*args) if args else func()
            return safe_after(self, ms, callback)
        else:
            return original_base_after(self, ms, func, *args)
            
    ctk.CTkBaseClass.after = patched_base_after

    # Пользовательский класс окна CustomTkinter с защитой от after-ошибок
    class SafeCTk(ctk.CTk):
        def destroy(self):
            cleanup_after_callbacks()
            super().destroy()


    # Расширенный класс Entry с настоящими плейсхолдерами и проверкой числовых полей
    class CustomEntry(ctk.CTkEntry):
        def __init__(self, master, field_name=None, **kwargs):
            self.field_name = field_name
            self.example_text = EXAMPLES.get(field_name, "Введіть значення")
            self.is_placeholder_visible = True
            
            # Определяем числовые поля для проверки ввода
            self.is_numeric_field = field_name in ["кількість", "ціна за одиницю", "загальна сума", "разом"]
            
            # Инициализация без текста в placeholder_text
            super().__init__(master, placeholder_text="", **kwargs)
            
            # Вставляем пример как обычный текст серого цвета
            self._entry.insert(0, self.example_text)
            self._entry.config(foreground='gray')
            
            # Обработчики фокуса для автоматического обновления содержимого
            self._entry.bind("<FocusIn>", self._on_focus_in)
            self._entry.bind("<FocusOut>", self._on_focus_out)
            
            # Для числовых полей добавляем обработчик ввода вместо валидатора
            if self.is_numeric_field:
                self._entry.bind("<KeyRelease>", self._check_numeric_input)
        
        def _check_numeric_input(self, event):
            """Проверяет числовой ввод после нажатия клавиши"""
            if self.is_placeholder_visible:
                return
            
            current_text = self._entry.get()
            if not current_text:
                return
                
            # Если текст не соответствует числовому формату, найдем последний валидный подтекст
            if not (re.match(r'^[\d\s]+(\.\d*)?$', current_text) or 
                    re.match(r'^[\d\s]+(\,\d*)?$', current_text)):
                # Ищем последний валидный подтекст
                valid_text = ""
                for i in range(len(current_text)):
                    subtext = current_text[:i+1]
                    if (re.match(r'^[\d\s]+(\.\d*)?$', subtext) or 
                        re.match(r'^[\d\s]+(\,\d*)?$', subtext)):
                        valid_text = subtext
                        
                # Заменяем текущий текст на последний валидный
                if valid_text:
                    cursor_pos = self._entry.index(tk.INSERT)
                    self._entry.delete(0, tk.END)
                    self._entry.insert(0, valid_text)
                    try:
                        # Восстанавливаем позицию курсора, но не дальше конца текста
                        cursor_pos = min(cursor_pos, len(valid_text))
                        self._entry.icursor(cursor_pos)
                    except:
                        pass
        
        def _on_focus_in(self, event):
            # При получении фокуса убираем плейсхолдер, если он отображается
            if self.is_placeholder_visible:
                self._entry.delete(0, "end")
                self._entry.config(foreground='black')  # Меняем цвет на чёрный для ввода
                self.is_placeholder_visible = False
        
        def _on_focus_out(self, event):
            # При потере фокуса, если поле пустое, показываем плейсхолдер
            current_text = self._entry.get().strip()
            if not current_text:
                self._entry.delete(0, "end")  # Очищаем поле перед вставкой плейсхолдера
                self._entry.insert(0, self.example_text)
                self._entry.config(foreground='gray')  # Серый цвет для плейсхолдера
                self.is_placeholder_visible = True
        
        def get(self):
            # Переопределяем метод get, чтобы не возвращал текст плейсхолдера
            if self.is_placeholder_visible:
                return ""
            return self._entry.get()
        
        def insert(self, index, string):
            # Переопределяем метод insert, чтобы правильно вставлять текст
            if self.is_placeholder_visible:
                self._entry.delete(0, "end")
                self._entry.config(foreground='black')
                self.is_placeholder_visible = False
            self._entry.insert(index, string)
        
        def delete(self, first_index, last_index=None):
            # Переопределяем метод delete для правильной работы с плейсхолдером
            if self.is_placeholder_visible:
                return
            self._entry.delete(first_index, last_index)

    # ---------------- ФУНКЦИИ ----------------
    def load_memory():
        try:
            if os.path.exists(MEMORY_FILE):
                with open(MEMORY_FILE, "r", encoding="utf-8") as f:
                    return json.load(f)
            else:
                return {}  # Возвращаем пустой словарь, если файл не существует
        except Exception as e:
            log_error(type(e), e, traceback.extract_tb(sys.exc_info()[2]))
            return {}  # В случае ошибки тоже возвращаем пустой словарь

    def save_memory(data, template_path=None):
        """
        Сохраняет данные в файл памяти.
        
        Args:
            data: Словарь с данными полей для сохранения
            template_path: Путь к шаблону, если указан - сохраняем только для этого шаблона
        """
        try:
            # Загружаем текущие данные
            current_data = load_memory()
            
            if template_path:
                # Нормализуем путь для использования в качестве ключа
                norm_path = os.path.normpath(template_path)
                
                # Если сохраняем для конкретного шаблона
                if norm_path not in current_data:
                    current_data[norm_path] = {}
                
                # Обновляем данные для этого шаблона
                for field, value in data.items():
                    current_data[norm_path][field] = value
            else:
                # Сохраняем как общие данные (для обратной совместимости)
                for field, value in data.items():
                    if field not in current_data:
                        current_data[field] = {}
                    current_data[field] = value
            
            # Записываем обновленные данные в файл
            with open(MEMORY_FILE, "w", encoding="utf-8") as f:
                json.dump(current_data, f, ensure_ascii=False, indent=2)
                
        except Exception as e:
            log_error(type(e), e, traceback.extract_tb(sys.exc_info()[2]))

    def get_template_memory(template_path):
        """
        Получает сохраненные данные для конкретного шаблона
        
        Args:
            template_path: Путь к шаблону
            
        Returns:
            dict: Словарь с сохраненными данными или пустой словарь
        """
        try:
            memory_data = load_memory()
            norm_path = os.path.normpath(template_path)
            
            # Проверяем, есть ли данные для этого шаблона
            if norm_path in memory_data and isinstance(memory_data[norm_path], dict):
                return memory_data[norm_path]
            return {}
        except Exception as e:
            log_error(type(e), e, traceback.extract_tb(sys.exc_info()[2]))
            return {}

    def export_to_excel():
        try:
            wb = Workbook()
            ws = wb.active
            ws.append(["Шаблон"] + FIELDS)
            for block in document_blocks:
                row = [block["path"]] + [block["entries"][f].get() for f in FIELDS]
                ws.append(row)
            wb.save("заповнені_дані.xlsx")
            messagebox.showinfo("Excel", "Дані збережено у 'заповнені_дані.xlsx'")
        except Exception as e:
            log_error(type(e), e, traceback.extract_tb(sys.exc_info()[2]))
            messagebox.showerror("Помилка", "Помилка при збереженні в Excel.")

    def generate_documents():
        try:
            if not document_blocks:
                messagebox.showwarning("Увага", "Не додано жодного договору.")
                return

            # Проверка на заполненность полей
            for i, block in enumerate(document_blocks, start=1):
                for field in FIELDS:
                    entry_widget = block["entries"].get(field)
                    if not entry_widget or not entry_widget.get().strip():
                        messagebox.showerror("Помилка", f"Блок #{i}: поле <{field}> порожнє або відсутнє.")
                        return

            save_dir = filedialog.askdirectory(title="Оберіть папку для збереження документів")
            if not save_dir:
                return

            # Сохраняем данные для каждого блока по отдельности
            for block in document_blocks:
                block_data = {f: block["entries"][f].get() for f in FIELDS}
                save_memory(block_data, block["path"])
                
            export_to_excel()
            
            # Вызываем функцию заполнения кошториса из внешнего модуля
            koshtorys.fill_koshtorys(document_blocks)

            try:
                word = win32.gencache.EnsureDispatch('Word.Application')
                word.Visible = False

                for block in document_blocks:
                    try:
                        doc = word.Documents.Open(os.path.abspath(block["path"]))
                        for key in FIELDS:
                            doc.Content.Find.Execute(FindText=f"<{key}>", ReplaceWith=block["entries"][key].get(), Replace=2)
                        base_name = os.path.splitext(os.path.basename(block["path"]).replace("ШАБЛОН", "").strip())[0]
                        save_path = os.path.join(save_dir, f"{base_name} {block['entries']['товар'].get()}.docm")
                        doc.SaveAs(save_path, FileFormat=13)
                        doc.Close(False)
                    except Exception as e:
                        log_error(type(e), e, traceback.extract_tb(sys.exc_info()[2]))
                        messagebox.showerror("Помилка", f"Файл: {block['path']}\n\n{str(e)}")

                word.Quit()
                messagebox.showinfo("Успіх", f"{len(document_blocks)} документ(и) збережено успішно!")
            except Exception as e:
                log_error(type(e), e, traceback.extract_tb(sys.exc_info()[2]))
                messagebox.showerror("Помилка", f"Не удалось запустить MS Word: {str(e)}")

        except Exception as e:
            log_error(type(e), e, traceback.extract_tb(sys.exc_info()[2]))
            messagebox.showerror("Помилка", "Невідома помилка при генерації документів.")

    def bind_entry_shortcuts(entry):
        try:
            internal = entry._entry

            def select_all(event):
                internal.select_range(0, "end")
                return "break"  # Это важно для предотвращения стандартной обработки события

            def copy_text(event):
                if internal.selection_present():
                    internal.event_generate('<<Copy>>')
                return "break"

            def paste_text(event):
                internal.event_generate('<<Paste>>')
                return "break"

            def cut_text(event):
                if internal.selection_present():
                    internal.event_generate('<<Cut>>')
                return "break"

            def show_context_menu(event):
                context_menu.tk_popup(event.x_root, event.y_root)
                return "break"

            # Привязываем события напрямую к Entry виджету
            internal.bind("<Control-a>", select_all)
            internal.bind("<Control-A>", select_all)
            internal.bind("<Button-3>", show_context_menu)
            internal.bind("<Control-c>", copy_text)
            internal.bind("<Control-C>", copy_text)
            internal.bind("<Control-v>", paste_text)
            internal.bind("<Control-V>", paste_text)
            internal.bind("<Control-x>", cut_text)
            internal.bind("<Control-X>", cut_text)
            
            # Необходимо также привязать события к самому CTkEntry
            entry.bind("<Control-a>", select_all)
            entry.bind("<Control-A>", select_all)
            entry.bind("<Button-3>", show_context_menu)
            entry.bind("<Control-c>", copy_text)
            entry.bind("<Control-C>", copy_text)
            entry.bind("<Control-v>", paste_text)
            entry.bind("<Control-V>", paste_text)
            entry.bind("<Control-x>", cut_text)
            entry.bind("<Control-X>", cut_text)
            
        except Exception as e:
            log_error(type(e), e, traceback.extract_tb(sys.exc_info()[2]))

    # ---------------- ПАРОЛЬ ----------------
    APP_PASSWORD = "1234"
    def ask_password():
        def check():
            if entry.get() == APP_PASSWORD:
                cleanup_after_callbacks()
                password_window.destroy()
                launch_main_app()
            else:
                messagebox.showerror("Помилка", "Невірний пароль!")

        # силове завершення, щоб прибрати завислі after
        def on_close():
            try:
                cleanup_after_callbacks()
                password_window.destroy()
            except:
                pass
            sys.exit(0)

        try:
            password_window = SafeCTk()
            password_window.title("Введіть пароль")
            password_window.geometry("300x120")
            password_window.resizable(False, False)
            password_window.protocol("WM_DELETE_WINDOW", on_close)  # важливо

            ctk.CTkLabel(password_window, text="Введіть 4-значний пароль:").pack(pady=10)
            entry = ctk.CTkEntry(password_window, show="*", justify="center", font=("Arial", 18))
            entry.pack(pady=5)
            entry.focus()
            
            # Привязываем горячие клавиши для поля ввода пароля
            bind_entry_shortcuts(entry)

            ctk.CTkButton(password_window, text="Увійти", command=check).pack(pady=10)

            # Обработка нажатия Enter
            entry._entry.bind("<Return>", lambda event: check())

            password_window.mainloop()
        except Exception as e:
            log_error(type(e), e, traceback.extract_tb(sys.exc_info()[2]))
            print(f"Ошибка при создании окна пароля: {str(e)}")
            input("Нажмите Enter для завершения...")
            sys.exit(1)

    # ---------------- ОСНОВНИЙ ІНТЕРФЕЙС ----------------
    def launch_main_app():
        global root, scroll_frame, document_blocks, context_menu
        
        try:
            ctk.set_appearance_mode("light")
            ctk.set_default_color_theme("blue")

            root = SafeCTk()
            root.title("sportforall")
            root.geometry("1200x700")

            # Обработчик закрытия окна
            def on_root_close():
                try:
                    cleanup_after_callbacks()
                    root.destroy()
                    sys.exit(0)
                except:
                    sys.exit(0)
                    
            root.protocol("WM_DELETE_WINDOW", on_root_close)

            # Контекстное меню
            context_menu = tk.Menu(root, tearoff=0)
            context_menu.add_command(label="Копіювати", command=lambda: root.focus_get().event_generate('<<Copy>>'))
            context_menu.add_command(label="Вставити", command=lambda: root.focus_get().event_generate('<<Paste>>'))
            context_menu.add_command(label="Вирізати", command=lambda: root.focus_get().event_generate('<<Cut>>'))
            context_menu.add_command(label="Виділити все", command=lambda: select_all_text())
            
            def select_all_text():
                focused = root.focus_get()
                if hasattr(focused, "_entry"):  # Для CTkEntry
                    focused._entry.select_range(0, "end")
                elif hasattr(focused, "select_range"):  # Для обычных Entry
                    focused.select_range(0, "end")

            # Загрузка сохраненных данных (для обратной совместимости)
            memory_data = load_memory()
            document_blocks = []

            # Верхні кнопки
            top_frame = ctk.CTkFrame(root)
            top_frame.pack(pady=10)

            ctk.CTkButton(top_frame, text="➕ Додати договір", command=lambda: add_new_template(), fg_color="#2196F3").pack(side="left", padx=5)
            ctk.CTkButton(top_frame, text="📄 Згенерувати документи", command=generate_documents, fg_color="#4CAF50").pack(side="left", padx=5)
            
            # Добавляем кнопку для заполнения кошториса
            ctk.CTkButton(top_frame, text="💰 Заповнити кошторис", 
                         command=lambda: koshtorys.fill_koshtorys(document_blocks), 
                         fg_color="#FF9800").pack(side="left", padx=5)

            # Версія
            version_label = ctk.CTkLabel(top_frame, text="version 2.8", text_color="gray", font=("Arial", 12))
            version_label.pack(side="right", padx=10)

            # Скролюваний контейнер
            scroll_frame = ctk.CTkScrollableFrame(root, width=1100, height=600)
            scroll_frame.pack(padx=10, pady=10, fill="both", expand=True)

            def add_new_template():
                filepath = filedialog.askopenfilename(title="Оберіть шаблон договору", filetypes=[("Word Documents", "*.docm")])
                if filepath:
                    add_document_block(filepath)

            def add_document_block(filepath=None):
                try:
                    frame = ctk.CTkFrame(scroll_frame)
                    frame.pack(pady=10, fill="x")

                    path_var = ctk.StringVar(value=filepath or "")
                    ctk.CTkLabel(frame, textvariable=path_var, text_color="blue", anchor="w", width=1000).pack(anchor="w", pady=5)

                    entries = {}
                    
                    # Загружаем данные для этого конкретного шаблона
                    template_memory = {}
                    if filepath:
                        template_memory = get_template_memory(filepath)

                    for field in FIELDS:
                        row = ctk.CTkFrame(frame)
                        row.pack(fill="x", pady=4)

                        label = ctk.CTkLabel(row, text=f"<{field}>", anchor="w", width=150, font=("Arial", 12, "bold"))
                        label.pack(side="left", padx=5)

                        # Используем кастомный класс CustomEntry вместо стандартного CTkEntry
                        entry = CustomEntry(row, field_name=field, width=600)
                        entry.pack(side="left", padx=5, fill="x", expand=True)

                        hint_text = EXAMPLES.get(field, "")
                        
                        def make_show_hint(field_hint, field_name):
                            def show_hint():
                                if field_hint:
                                    messagebox.showinfo(f"Підказка для <{field_name}>", field_hint)
                                else:
                                    messagebox.showinfo("Підказка", "Приклад не вказано.")
                            return show_hint
                            
                        hint_btn = ctk.CTkButton(row, text="ℹ", width=30, height=28, font=("Arial", 14), 
                                                command=make_show_hint(hint_text, field))
                        hint_btn.pack(side="left", padx=5)

                        # Если есть сохраненные данные для этого шаблона, используем их
                        if field in template_memory and template_memory[field]:
                            entry.insert(0, template_memory[field])
                        # Иначе если есть данные в общей памяти, используем их (для обратной совместимости)
                        elif isinstance(memory_data.get(field), str) and memory_data[field]:
                            entry.insert(0, memory_data[field])

                        bind_entry_shortcuts(entry)
                        entries[field] = entry

                    block = {"path": path_var.get(), "entries": entries, "frame": frame}
                    document_blocks.append(block)

                    btns = ctk.CTkFrame(frame)
                    btns.pack(anchor="e", pady=5)

                    def clear_fields():
                        for entry in entries.values():
                            # Используем специальный метод clear, если он есть
                            if hasattr(entry, 'clear'):
                                entry.clear()
                            else:
                                # Стандартная очистка
                                entry.delete(0, 'end')
                                # Восстановление плейсхолдера происходит автоматически при потере фокуса
                                entry._on_focus_out(None)  # Имитируем потерю фокуса для активации плейсхолдера

                    def replace_template():
                        new_path = filedialog.askopenfilename(title="Оберіть новий шаблон", filetypes=[("Word Documents", "*.docm")])
                        if new_path:
                            path_var.set(new_path)
                            block["path"] = new_path

                    def remove_block():
                        document_blocks.remove(block)
                        frame.destroy()

                    ctk.CTkButton(btns, text="🧹 Очистити поля", command=clear_fields).pack(side="left", padx=3)
                    ctk.CTkButton(btns, text="Замінити шаблон", command=replace_template).pack(side="left", padx=3)
                    ctk.CTkButton(btns, text="🗑 Видалити", command=remove_block).pack(side="left", padx=3)
                    ctk.CTkButton(btns, text="💾 Зберегти в Excel", command=export_to_excel).pack(side="left", padx=3)
                except Exception as e:
                    log_error(type(e), e, traceback.extract_tb(sys.exc_info()[2]))
                    messagebox.showerror("Помилка", f"Ошибка при создании блока: {str(e)}")

            # Глобальные сочетания клавиш для всего приложения
            root.bind_all("<Control-s>", lambda event: export_to_excel())
            root.bind_all("<Control-g>", lambda event: generate_documents())

            root.mainloop()

        except Exception as e:
            log_error(type(e), e, traceback.extract_tb(sys.exc_info()[2]))
            messagebox.showerror("Помилка", "Виникла критична помилка.\nДеталі у файлі error.txt.")
            input("Нажмите Enter для завершения...")

except Exception as e:
    try:
        messagebox.showerror("Критична помилка", f"Виникла критична помилка: {str(e)}")
    except:
        print(f"Критична помилка: {str(e)}")
        print(traceback.format_exc())
        input("Нажмите Enter для завершения...")

# ВИКЛИК ПРОГРАММЫ
if __name__ == "__main__":
    try:
        ask_password()
    except Exception as e:
        print(f"Ошибка при запуске программы: {str(e)}")
        print(traceback.format_exc())
        input("Нажмите Enter для завершения...")