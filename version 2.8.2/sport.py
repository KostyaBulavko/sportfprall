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

# --- НОВЫЕ ИМПОРТЫ ДЛЯ ОБНОВЛЕНИЯ ---
import requests
import shutil
import subprocess
import io
import zipfile
# --- КОНЕЦ НОВЫХ ИМПОРТОВ ---

# Импортируем модуль для работы с кошторисом
# Убедитесь, что koshtorys.py находится в той же папке или доступен
try:
    import koshtorys
except ImportError:
    messagebox.showerror("Помилка імпорту", "Не вдалося імпортувати модуль koshtorys.py.\nПереконайтесь, що файл знаходиться в папці з програмою.")
    koshtorys = None # Заглушка, если импорт не удался

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

    def log_error(exc_type, exc_value, exc_traceback_obj): # Изменен параметр для traceback
        try:
            with open(ERROR_LOG, "a", encoding="utf-8") as f:
                f.write(f"\n[{datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}]\n")
                # Используем traceback.print_exception напрямую
                traceback.print_exception(exc_type, exc_value, exc_traceback_obj, file=f)
            # Показать сообщение об ошибке
            messagebox.showerror("Ошибка", f"Произошла ошибка: {str(exc_value)}\nПодробности в файле {ERROR_LOG}")
        except Exception as ex:
            print(f"Ошибка при логировании: {str(ex)}")
            # input("Нажмите Enter для продолжения...") # Не блокируем поток, если это GUI

    # Перехватываем и логируем только необработанные исключения
    def global_exception_handler(exc_type, exc_value, exc_traceback_obj):
        log_error(exc_type, exc_value, exc_traceback_obj)
        # Также выводим в консоль для отладки
        sys.__excepthook__(exc_type, exc_value, exc_traceback_obj)
        # input("Нажмите Enter для завершения...") # Не блокируем
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
    
    after_callbacks = {}

    # ----- НАСТРОЙКИ ОБНОВЛЕНИЯ -----
    CURRENT_VERSION = "2.8.2" 
    GITHUB_REPO_OWNER = "KostyaBulavko"  # ЗАМЕНИТЕ НА ВАШЕ ИМЯ ПОЛЬЗОВАТЕЛЯ GITHUB
    GITHUB_REPO_NAME = "sportfprall" # ЗАМЕНИТЕ НА ИМЯ ВАШЕГО РЕПОЗИТОРИЯ
    VERSION_FILE_URL = f"https://raw.githubusercontent.com/{GITHUB_REPO_OWNER}/{GITHUB_REPO_NAME}/main/version.txt"
    ZIP_DOWNLOAD_URL = f"https://github.com/{GITHUB_REPO_OWNER}/{GITHUB_REPO_NAME}/archive/refs/heads/main.zip"
    # ----- КОНЕЦ НАСТРОЕК ОБНОВЛЕНИЯ -----


    def safe_after_cancel(widget, after_id):
        try:
            if widget and widget.winfo_exists() and after_id: # Проверяем, существует ли виджет
                widget.after_cancel(after_id)
        except Exception:
            pass 

    def cleanup_after_callbacks():
        global after_callbacks
        for widget_id_str in list(after_callbacks.keys()):
            # Виджеты могут быть уже уничтожены, поэтому try-except
            try:
                # CTk._get_widget_by_id может быть недоступен или измениться
                # Проще просто пытаться отменить, полагаясь на widget.winfo_exists() в safe_after_cancel
                # Или хранить сами объекты виджетов, если это безопасно (но это сложнее с удалением)
                widget = None 
                # Попытка найти виджет (может быть ненадежно, если виджеты удалены)
                if root: # Предполагаем, что все виджеты - потомки root
                    for child in root.winfo_children(recursive=True): # type: ignore
                        if str(child) == widget_id_str:
                            widget = child
                            break
                
                for after_id in after_callbacks.get(widget_id_str, []):
                    if widget and widget.winfo_exists():
                       safe_after_cancel(widget, after_id)
                    # Если виджет не найден или уже удален, просто пропускаем
            except Exception:
                pass 
        after_callbacks.clear()

    def safe_after(widget, ms, callback):
        if not widget or not widget.winfo_exists():
            return None
            
        try:
            widget_id = str(widget) # Используем строковое представление как ключ
            after_id = widget.after(ms, callback)
            
            if widget_id not in after_callbacks:
                after_callbacks[widget_id] = []
                
            after_callbacks[widget_id].append(after_id)
            return after_id
        except Exception:
            return None

    original_after = ctk.CTk.after
    def patched_after(self, ms, func=None, *args):
        if func:
            callback = lambda: func(*args) if args else func()
            return safe_after(self, ms, callback)
        else:
            # Этот случай (когда func is None) обычно используется для `after_idle`
            # или для получения информации, а не для установки нового колбэка.
            # Обрабатываем его через оригинальный метод, но это может быть небезопасно.
            # Для большей безопасности, если func is None, лучше не использовать safe_after.
            return original_after(self, ms, func, *args) 
            
    ctk.CTk.after = patched_after # type: ignore
    
    original_base_after = ctk.CTkBaseClass.after
    def patched_base_after(self, ms, func=None, *args):
        if func:
            callback = lambda: func(*args) if args else func()
            return safe_after(self, ms, callback)
        else:
            return original_base_after(self, ms, func, *args)
            
    ctk.CTkBaseClass.after = patched_base_after # type: ignore

    class SafeCTk(ctk.CTk):
        def destroy(self):
            cleanup_after_callbacks() # Очищаем колбэки перед уничтожением главного окна
            super().destroy()


    class CustomEntry(ctk.CTkEntry):
        def __init__(self, master, field_name=None, **kwargs):
            self.field_name = field_name
            self.example_text = EXAMPLES.get(field_name, "Введіть значення")
            self.is_placeholder_visible = True
            
            self.is_numeric_field = field_name in ["кількість", "ціна за одиницю", "загальна сума", "разом"]
            
            super().__init__(master, placeholder_text="", **kwargs)
            
            self._entry.insert(0, self.example_text)
            self._entry.config(foreground='gray')
            
            self._entry.bind("<FocusIn>", self._on_focus_in)
            self._entry.bind("<FocusOut>", self._on_focus_out)
            
            if self.is_numeric_field:
                self._entry.bind("<KeyRelease>", self._check_numeric_input)
        
        def _check_numeric_input(self, event):
            if self.is_placeholder_visible:
                return
            
            current_text = self._entry.get()
            if not current_text:
                return
                
            # Разрешаем числа, пробелы, одну точку или одну запятую
            if not (re.match(r'^[\d\s]*([.,][\d\s]*)?$', current_text)):
                valid_text = ""
                for i in range(len(current_text)):
                    subtext = current_text[:i+1]
                    if (re.match(r'^[\d\s]*([.,][\d\s]*)?$', subtext)):
                        valid_text = subtext
                    else:
                        break # Прекращаем, как только символ невалиден
                
                # Заменяем текущий текст на последний валидный
                # Сохраняем позицию курсора
                cursor_pos = self._entry.index(tk.INSERT)
                self._entry.delete(0, tk.END)
                self._entry.insert(0, valid_text)
                try:
                    # Восстанавливаем позицию курсора, но не дальше конца текста
                    new_cursor_pos = min(cursor_pos - (len(current_text) - len(valid_text)), len(valid_text))
                    self._entry.icursor(max(0, new_cursor_pos)) # Курсор не может быть < 0
                except tk.TclError: # Может возникнуть, если виджет уже удален
                    pass
        
        def _on_focus_in(self, event):
            if self.is_placeholder_visible:
                self._entry.delete(0, "end")
                self._entry.config(foreground='black') 
                self.is_placeholder_visible = False
        
        def _on_focus_out(self, event=None): # Добавил event=None для вызова без события
            current_text = self._entry.get().strip()
            if not current_text:
                self._entry.delete(0, "end") 
                self._entry.insert(0, self.example_text)
                self._entry.config(foreground='gray') 
                self.is_placeholder_visible = True
        
        def get(self):
            if self.is_placeholder_visible:
                return ""
            return self._entry.get()
        
        def insert(self, index, string):
            if self.is_placeholder_visible:
                self._entry.delete(0, "end")
                self._entry.config(foreground='black')
                self.is_placeholder_visible = False
            self._entry.insert(index, string)
        
        def delete(self, first_index, last_index=None):
            if self.is_placeholder_visible:
                # Если плейсхолдер виден, удаление не должно ничего делать
                # Или, если нужно поведение как у обычного Entry, то:
                # self._on_focus_in(None) # Симулируем фокус, чтобы убрать плейсхолдер
                # self._entry.delete(first_index, last_index)
                # Но текущая логика - не удалять плейсхолдер напрямую, а только пользовательский ввод
                return 
            self._entry.delete(first_index, last_index)
            # Если после удаления поле стало пустым, _on_focus_out (если он сработает) вернет плейсхолдер
            # Можно добавить вызов _on_focus_out(None) здесь, если поле стало пустым, чтобы плейсхолдер появился сразу

        def clear(self): # Новый метод для очистки
            self.delete(0, tk.END)
            self._on_focus_out(None) # Восстановить плейсхолдер, если поле пустое


    # ---------------- ФУНКЦИИ ----------------
    def load_memory():
        try:
            if os.path.exists(MEMORY_FILE):
                with open(MEMORY_FILE, "r", encoding="utf-8") as f:
                    return json.load(f)
            else:
                return {}
        except Exception as e:
            log_error(type(e), e, sys.exc_info()[2])
            return {}

    def save_memory(data, template_path=None):
        try:
            current_data = load_memory()
            
            if template_path:
                norm_path = os.path.normpath(template_path)
                if norm_path not in current_data:
                    current_data[norm_path] = {}
                for field, value in data.items():
                    current_data[norm_path][field] = value
            else:
                # Общее сохранение (обратная совместимость) - перезаписываем целиком
                current_data.update(data) # Проще и корректнее для общего случая
                        
            with open(MEMORY_FILE, "w", encoding="utf-8") as f:
                json.dump(current_data, f, ensure_ascii=False, indent=2)
                
        except Exception as e:
            log_error(type(e), e, sys.exc_info()[2])

    def get_template_memory(template_path):
        try:
            memory_data = load_memory()
            norm_path = os.path.normpath(template_path)
            
            if norm_path in memory_data and isinstance(memory_data[norm_path], dict):
                return memory_data[norm_path]
            # Для обратной совместимости, если данные не по ключу пути, а старый формат
            elif not any(isinstance(v, dict) for v in memory_data.values()): # Проверяем, старый ли это формат
                 # Если все значения - строки (старый формат), возвращаем весь memory_data
                 # Это предполагает, что старый формат не смешивался с новым
                return memory_data # Возвращаем все, чтобы старая логика заполнения сработала
            return {}
        except Exception as e:
            log_error(type(e), e, sys.exc_info()[2])
            return {}

    def export_to_excel():
        try:
            if not document_blocks:
                messagebox.showwarning("Excel", "Немає даних для експорту. Додайте хоча б один договір.")
                return

            wb = Workbook()
            ws = wb.active
            ws.append(["Шаблон"] + FIELDS)
            for block in document_blocks:
                row = [block["path"]] + [block["entries"][f].get() for f in FIELDS]
                ws.append(row)
            
            excel_file_path = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel файли", "*.xlsx")],
                title="Зберегти дані в Excel"
            )
            if not excel_file_path:
                return # Пользователь отменил сохранение

            wb.save(excel_file_path)
            messagebox.showinfo("Excel", f"Дані збережено у '{excel_file_path}'")
        except Exception as e:
            log_error(type(e), e, sys.exc_info()[2])
            messagebox.showerror("Помилка", "Помилка при збереженні в Excel.")

    def generate_documents():
        try:
            if not document_blocks:
                messagebox.showwarning("Увага", "Не додано жодного договору.")
                return

            all_fields_filled = True
            for i, block in enumerate(document_blocks, start=1):
                for field in FIELDS:
                    entry_widget = block["entries"].get(field)
                    if not entry_widget or not entry_widget.get().strip():
                        # Пропускаем проверку для поля ПДВ, если оно может быть пустым по логике
                        if field == "пдв" and (not entry_widget or not entry_widget.get().strip()):
                             # Если ПДВ пусто, можно присвоить значение по умолчанию, например "без ПДВ"
                             # или просто проигнорировать, если это допустимо
                            pass # Или entry_widget.insert(0, "без ПДВ") если это нужно
                        else:
                            messagebox.showerror("Помилка", f"Блок договору #{i} (файл: {os.path.basename(block['path'])}):\nПоле <{field}> порожнє.")
                            all_fields_filled = False
                            return # Останавливаемся на первой ошибке
            if not all_fields_filled:
                return

            save_dir = filedialog.askdirectory(title="Оберіть папку для збереження документів")
            if not save_dir:
                return

            for block in document_blocks:
                block_data = {f: block["entries"][f].get() for f in FIELDS}
                save_memory(block_data, block["path"]) # Сохраняем данные для конкретного шаблона
            
            # Предлагаем экспортировать в Excel перед генерацией документов
            if messagebox.askyesno("Excel", "Зберегти введені дані в Excel перед генерацією документів?"):
                export_to_excel()
            
            if koshtorys and hasattr(koshtorys, 'fill_koshtorys'):
                # Спрашиваем пользователя перед заполнением кошториса
                if messagebox.askyesno("Кошторис", "Заповнити кошторис на основі введених даних?"):
                    koshtorys.fill_koshtorys(document_blocks)
            else:
                messagebox.showwarning("Кошторис", "Функція заповнення кошторису недоступна (модуль koshtorys не завантажено).")


            try:
                word = win32.gencache.EnsureDispatch('Word.Application')
                word.Visible = False # Можно установить в True для отладки
                
                generated_files_count = 0
                for block_idx, block in enumerate(document_blocks):
                    try:
                        doc_path = os.path.abspath(block["path"])
                        if not os.path.exists(doc_path):
                            messagebox.showerror("Помилка", f"Файл шаблону не знайдено: {doc_path}\nПропуск блоку #{block_idx+1}.")
                            continue

                        doc = word.Documents.Open(doc_path)
                        
                        # Замена в основном тексте документа
                        for key in FIELDS:
                            find_obj = doc.Content.Find
                            find_obj.ClearFormatting()
                            find_obj.Replacement.ClearFormatting()
                            find_obj.Execute(FindText=f"<{key}>", ReplaceWith=block["entries"][key].get(), Replace=2, Wrap=1) # wdReplaceAll = 2, wdFindContinue = 1

                        # Замена в верхних и нижних колонтитулах
                        for section in doc.Sections:
                            for header_footer_type in [1,2,3]: # wdHeaderFooterPrimary, wdHeaderFooterFirstPage, wdHeaderFooterEvenPages
                                try:
                                    header = section.Headers(header_footer_type)
                                    if header.Exists:
                                        for key in FIELDS:
                                            find_obj = header.Range.Find
                                            find_obj.ClearFormatting()
                                            find_obj.Replacement.ClearFormatting()
                                            find_obj.Execute(FindText=f"<{key}>", ReplaceWith=block["entries"][key].get(), Replace=2, Wrap=1)
                                except Exception: pass # Игнорируем ошибки, если колонтитул определенного типа отсутствует

                                try:
                                    footer = section.Footers(header_footer_type)
                                    if footer.Exists:
                                        for key in FIELDS:
                                            find_obj = footer.Range.Find
                                            find_obj.ClearFormatting()
                                            find_obj.Replacement.ClearFormatting()
                                            find_obj.Execute(FindText=f"<{key}>", ReplaceWith=block["entries"][key].get(), Replace=2, Wrap=1)
                                except Exception: pass


                        base_name = os.path.splitext(os.path.basename(block["path"]).replace("ШАБЛОН", "").strip())[0]
                        # Формируем имя файла, избегая недопустимых символов
                        tovar_value = block['entries']['товар'].get()
                        safe_tovar_name = "".join(c if c.isalnum() or c in " -" else "_" for c in tovar_value) # Оставляем буквы, цифры, пробелы, дефисы
                        safe_tovar_name = safe_tovar_name[:50] # Ограничиваем длину
                        
                        save_file_name = f"{base_name} ({safe_tovar_name}).docm"
                        save_path = os.path.join(save_dir, save_file_name)
                        
                        doc.SaveAs(save_path, FileFormat=13) # wdFormatXMLDocumentMacroEnabled = 13
                        doc.Close(False) # wdDoNotSaveChanges = 0
                        generated_files_count +=1
                    except Exception as e_doc:
                        log_error(type(e_doc), e_doc, sys.exc_info()[2])
                        messagebox.showerror("Помилка обробки документа", f"Помилка при обробці файлу: {block['path']}\n\n{str(e_doc)}")
                
                try: # Закрываем Word только если он был открыт
                    if 'word' in locals() and word is not None:
                        word.Quit()
                except Exception: pass # Игнорируем ошибки при закрытии Word

                if generated_files_count > 0:
                    messagebox.showinfo("Успіх", f"{generated_files_count} документ(и) збережено успішно в папці:\n{save_dir}")
                elif not document_blocks : # Если изначально не было блоков (хотя проверка выше должна это отловить)
                    pass # Сообщение уже было показано
                else: # Если были блоки, но ни один не сгенерирован из-за ошибок
                    messagebox.showwarning("Увага", "Жодного документа не було згенеровано через помилки.")

            except Exception as e_word:
                log_error(type(e_word), e_word, sys.exc_info()[2])
                messagebox.showerror("Помилка MS Word", f"Не вдалося запустити або використати MS Word: {str(e_word)}\nПереконайтеся, що MS Word встановлено та налаштовано.")

        except Exception as e:
            log_error(type(e), e, sys.exc_info()[2])
            # messagebox.showerror("Помилка", "Невідома помилка при генерації документів.") # Уже логируется в global_exception_handler

    def bind_entry_shortcuts(entry_widget): # Изменено имя параметра для ясности
        try:
            # Убеждаемся, что entry_widget это наш CustomEntry и у него есть _entry
            if not isinstance(entry_widget, CustomEntry) or not hasattr(entry_widget, '_entry'):
                return

            internal_tk_entry = entry_widget._entry # Это объект tk.Entry внутри CTkEntry

            def select_all(event):
                internal_tk_entry.select_range(0, "end")
                return "break" 

            def copy_text(event):
                if internal_tk_entry.selection_present():
                    internal_tk_entry.event_generate('<<Copy>>')
                return "break"

            def paste_text(event):
                # Перед вставкой, если это плейсхолдер, очищаем его
                if entry_widget.is_placeholder_visible:
                    entry_widget._on_focus_in(None) # Имитируем фокус, чтобы убрать плейсхолдер
                internal_tk_entry.event_generate('<<Paste>>')
                return "break"

            def cut_text(event):
                if internal_tk_entry.selection_present():
                    internal_tk_entry.event_generate('<<Cut>>')
                return "break"

            def show_context_menu(event):
                # Обновляем состояние команд меню перед показом
                if internal_tk_entry.selection_present():
                    context_menu.entryconfig("Копіювати", state="normal")
                    context_menu.entryconfig("Вирізати", state="normal")
                else:
                    context_menu.entryconfig("Копіювати", state="disabled")
                    context_menu.entryconfig("Вирізати", state="disabled")
                
                try: # Проверяем, есть ли что-то в буфере обмена
                    root.clipboard_get()
                    context_menu.entryconfig("Вставити", state="normal")
                except tk.TclError:
                    context_menu.entryconfig("Вставити", state="disabled")

                # Запоминаем, какой виджет вызвал меню
                context_menu.current_widget = internal_tk_entry # type: ignore
                context_menu.tk_popup(event.x_root, event.y_root)
                return "break"

            # Привязываем события к внутреннему tk.Entry
            internal_tk_entry.bind("<Control-a>", select_all)
            internal_tk_entry.bind("<Control-A>", select_all) # Для русской раскладки А
            internal_tk_entry.bind("<Button-3>", show_context_menu)
            internal_tk_entry.bind("<Control-c>", copy_text)
            internal_tk_entry.bind("<Control-C>", copy_text) # Для русской раскладки С
            internal_tk_entry.bind("<Control-v>", paste_text)
            internal_tk_entry.bind("<Control-V>", paste_text) # Для русской раскладки М (часто как V)
            internal_tk_entry.bind("<Control-x>", cut_text)
            internal_tk_entry.bind("<Control-X>", cut_text) # Для русской раскладки Ч

            # Привязка к самому CTkEntry (внешнему виджету) тоже может быть полезна
            # но основные действия с текстом происходят во внутреннем _entry
            entry_widget.bind("<Button-3>", show_context_menu) # Чтобы меню открывалось и по клику на сам CTkEntry

        except Exception as e:
            log_error(type(e), e, sys.exc_info()[2])
    
    # ---------------- ОБНОВЛЕНИЕ ПРИЛОЖЕНИЯ ----------------

    def get_latest_version_from_github():
        """Получает номер последней версии с GitHub."""
        try:
            response = requests.get(VERSION_FILE_URL, timeout=10)
            response.raise_for_status() 
            latest_version = response.text.strip()
            return latest_version
        except requests.exceptions.RequestException as e:
            # log_error(type(e), e, sys.exc_info()[2]) # Логируется в вызывающей функции
            messagebox.showerror("Помилка оновлення", f"Не вдалося перевірити версію на GitHub: {e}")
            return None

    def download_and_apply_update():
        """Скачивает и применяет обновление."""
        try:
            latest_version = get_latest_version_from_github()

            if not latest_version:
                return

            # Простое сравнение версий (можно усложнить, если версии типа 1.0.1)
            if latest_version == CURRENT_VERSION:
                messagebox.showinfo("Оновлення", "У вас вже встановлена остання версія.")
                return

            confirm = messagebox.askyesno("Оновлення доступне",
                                      f"Доступна нова версія: {latest_version}\nВаша версія: {CURRENT_VERSION}\n\nБажаєте оновитися зараз?")
            if not confirm:
                return

            messagebox.showinfo("Оновлення", "Починається завантаження оновлення. Будь ласка, зачекайте...\nДодаток буде перезапущено після оновлення.")
            
            # Проверяем, есть ли .git папка
            if os.path.exists(".git") and os.path.isdir(".git"):
                try:
                    subprocess.check_call(['git', 'pull'], creationflags=subprocess.CREATE_NO_WINDOW)
                    messagebox.showinfo("Оновлення успішне", "Додаток оновлено через git pull. Будь ласка, перезапустіть програму вручну.")
                    # Пробуем перезапустить, но 'git pull' может требовать ручного перезапуска
                    if root:
                        cleanup_after_callbacks()
                        root.destroy()
                    # Запуск нового экземпляра после git pull
                    python_executable = sys.executable
                    script_path = os.path.abspath(sys.argv[0])
                    subprocess.Popen([python_executable, script_path])
                    sys.exit(0) # Закрываем текущий
                    return 
                except subprocess.CalledProcessError as e_git:
                    log_error(type(e_git), e_git, sys.exc_info()[2])
                    messagebox.showwarning("Помилка Git", f"Не вдалося виконати 'git pull': {e_git}\nСпробуємо завантажити архів.")
                except FileNotFoundError:
                    messagebox.showwarning("Помилка Git", "'git' не знайдено в системі. Спробуємо завантажити архів.")
            
            # Если не git репозиторий или git pull не удался, скачиваем zip
            response = requests.get(ZIP_DOWNLOAD_URL, stream=True, timeout=60) # Увеличил таймаут
            response.raise_for_status()

            temp_zip_path = "update_temp.zip"
            with open(temp_zip_path, 'wb') as f:
                shutil.copyfileobj(response.raw, f) # Более эффективное копирование для stream
            
            extract_path = "update_extracted_temp"
            if os.path.exists(extract_path):
                shutil.rmtree(extract_path) 
            os.makedirs(extract_path, exist_ok=True)

            with zipfile.ZipFile(temp_zip_path, 'r') as zip_ref:
                zip_ref.extractall(extract_path)
            
            os.remove(temp_zip_path)

            extracted_items = os.listdir(extract_path)
            if not extracted_items:
                raise Exception("Архів порожній або не вдалося розпакувати.")
            
            # Папка внутри архива GitHub обычно REPO_NAME-BRANCH_NAME (e.g., MyRepo-main)
            source_update_folder_name = f"{GITHUB_REPO_NAME}-main" # Или другая ветка, если URL изменен
            potential_source_folder = os.path.join(extract_path, source_update_folder_name)

            if not os.path.isdir(potential_source_folder):
                # Если стандартное имя папки не найдено, берем первую папку из архива
                # Это менее надежно, но может сработать для архивов с другой структурой
                first_item_in_archive = os.path.join(extract_path, extracted_items[0])
                if os.path.isdir(first_item_in_archive):
                    source_update_folder = first_item_in_archive
                else: # Если и первый элемент не папка, значит файлы лежат прямо в корне архива
                    source_update_folder = extract_path
            else:
                source_update_folder = potential_source_folder
            
            # Собираем список файлов и папок для копирования
            # Пропускаем файлы, которые не должны обновляться (например, пользовательские данные)
            # В данном случае contracts_memory.json и error.txt
            excluded_files = [MEMORY_FILE, ERROR_LOG, "updater.bat"] 
            if os.path.basename(sys.executable).lower() == "python.exe" or os.path.basename(sys.executable).lower() == "pythonw.exe":
                excluded_files.append(os.path.basename(sys.executable)) # Не пытаемся заменить сам python.exe

            updater_script_commands = []
            for item_name in os.listdir(source_update_folder):
                s_item = os.path.join(source_update_folder, item_name)
                d_item = os.path.join(os.getcwd(), item_name) # Копируем в текущую рабочую директорию

                if item_name in excluded_files:
                    continue # Пропускаем исключенные файлы
                
                # Если это __pycache__ или .git, пропускаем
                if item_name == "__pycache__" or item_name == ".git":
                    continue
                
                if os.path.isfile(s_item):
                    # Удаляем старый файл перед копированием нового (move перезаписывает, но del надежнее для заблокированных)
                    updater_script_commands.append(f'if exist "{d_item}" ( del /F /Q "{d_item}" )')
                    updater_script_commands.append(f'move /Y "{s_item}" "{d_item}"')
                elif os.path.isdir(s_item):
                    # Удаляем старую папку рекурсивно перед копированием новой
                    updater_script_commands.append(f'if exist "{d_item}" ( rmdir /S /Q "{d_item}" )')
                    updater_script_commands.append(f'xcopy "{s_item}" "{d_item}\\" /E /I /Y /Q') # Xcopy для папок


            # Создаем скрипт-обновлятор (updater.bat)
            # Путь к текущему исполняемому файлу Python или скомпилированному .exe
            current_executable_path = sys.executable
            main_script_path = os.path.abspath(sys.argv[0])

            updater_script_content = f"""@echo off
chcp 65001 > nul
echo.
echo =====================================================
echo           ОНОВЛЕННЯ ДОДАТКУ - НЕ ЗАКРИВАЙТЕ ЦЕ ВІКНО!
echo =====================================================
echo.
echo Закриття попередньої версії додатку...
taskkill /IM "{os.path.basename(current_executable_path)}" /F > nul 2>&1
timeout /t 3 /nobreak > nul

echo.
echo Копіювання нових файлів...
"""
            updater_script_content += "\n".join(updater_script_commands)
            updater_script_content += f"""

echo.
echo Очищення тимчасових файлів...
rmdir /S /Q "{extract_path}" > nul 2>&1

echo.
echo Оновлення завершено! Запуск нової версії...
start "" "{current_executable_path}" "{main_script_path}"

echo.
echo Скрипт оновлення завершить роботу за 5 секунд...
timeout /t 5 /nobreak > nul
del "%~f0"
exit
"""
            updater_script_path = os.path.join(os.getcwd(), "updater.bat")
            try:
                with open(updater_script_path, "w", encoding="utf-8") as f: # utf-8 для chcp 65001
                    f.write(updater_script_content)
            except Exception as e_write_bat:
                log_error(type(e_write_bat), e_write_bat, sys.exc_info()[2])
                messagebox.showerror("Помилка оновлення", f"Не вдалося створити скрипт оновлення: {e_write_bat}")
                return


            # messagebox.showinfo("Оновлення", "Файли підготовлено. Додаток буде перезапущено для завершення оновлення.")
            
            if root:
                cleanup_after_callbacks()
                root.destroy()

            # Запускаем батник и выходим
            # CREATE_NEW_CONSOLE, чтобы батник открылся в новом окне и не блокировал
            # DETACHED_PROCESS, чтобы он работал независимо
            subprocess.Popen([updater_script_path], creationflags=subprocess.CREATE_NEW_CONSOLE | subprocess.DETACHED_PROCESS, close_fds=True)
            sys.exit(0)

        except requests.exceptions.Timeout:
            messagebox.showerror("Помилка завантаження", "Час очікування відповіді від сервера минув. Перевірте інтернет-з'єднання.")
        except requests.exceptions.ConnectionError:
            messagebox.showerror("Помилка завантаження", "Не вдалося з'єднатися з сервером оновлень. Перевірте інтернет-з'єднання.")
        except requests.exceptions.RequestException as e_req:
            log_error(type(e_req), e_req, sys.exc_info()[2])
            messagebox.showerror("Помилка завантаження", f"Не вдалося завантажити оновлення: {e_req}")
        except zipfile.BadZipFile as e_zip:
            log_error(type(e_zip), e_zip, sys.exc_info()[2])
            messagebox.showerror("Помилка розпакування", "Завантажений файл оновлення пошкоджено або не є ZIP-архівом.")
        except Exception as e_update:
            log_error(type(e_update), e_update, sys.exc_info()[2])
            # messagebox.showerror("Помилка оновлення", f"Сталася невідома помилка під час оновлення: {e_update}") # Логируется через sys.excepthook
        finally:
            # Очистка на случай, если что-то пошло не так до создания батника
            if os.path.exists("update_temp.zip"):
                try: os.remove("update_temp.zip")
                except: pass
            if os.path.exists("update_extracted_temp"): # Имя папки изменено
                try: shutil.rmtree("update_extracted_temp")
                except: pass
    
    # ---------------- ПАРОЛЬ ----------------
    APP_PASSWORD = "1234"
    def ask_password():
        global password_window_handle # Храним ссылку на окно пароля
        password_window_handle = None

        def check():
            if entry.get() == APP_PASSWORD:
                if password_window_handle:
                    try:
                        # Сначала отменяем колбэки, связанные с этим окном
                        # (хотя их тут нет, но для единообразия)
                        # cleanup_specific_window_callbacks(password_window_handle)
                        password_window_handle.destroy()
                    except tk.TclError: pass # Окно может быть уже уничтожено
                launch_main_app()
            else:
                messagebox.showerror("Помилка", "Невірний пароль!", parent=password_window_handle) # Указываем родителя
                entry.delete(0, tk.END) # Очищаем поле

        def on_close_password_window():
            if password_window_handle:
                try:
                    # cleanup_specific_window_callbacks(password_window_handle)
                    password_window_handle.destroy()
                except tk.TclError: pass
            sys.exit(0) # Завершаем приложение, если окно пароля закрыто

        try:
            # Используем SafeCTk для окна пароля тоже, если нужно
            password_window = SafeCTk() # или ctk.CTk() если SafeCTk не нужен здесь
            password_window_handle = password_window # Сохраняем ссылку

            password_window.title("Введіть пароль")
            password_window.geometry("300x150") # Немного больше для лучшего вида
            password_window.resizable(False, False)
            password_window.protocol("WM_DELETE_WINDOW", on_close_password_window) 
            
            # Центрирование окна
            password_window.eval('tk::PlaceWindow . center')


            ctk.CTkLabel(password_window, text="Введіть 4-значний пароль:").pack(pady=(15,5))
            entry = ctk.CTkEntry(password_window, show="*", justify="center", font=("Arial", 18), width=200)
            entry.pack(pady=5)
            entry.focus()
            
            # bind_entry_shortcuts(entry) # Для окна пароля контекстное меню и Ctr+A/C/V обычно не нужны

            ctk.CTkButton(password_window, text="Увійти", command=check).pack(pady=10)

            entry.bind("<Return>", lambda event: check()) # Используем внутренний _entry, если это CustomEntry
            # Если это ctk.CTkEntry, то entry.bind("<Return>", ...) должно работать

            password_window.mainloop()
        except Exception as e:
            # log_error(type(e), e, sys.exc_info()[2]) # Логируется глобальным обработчиком
            print(f"Ошибка при создании окна пароля: {str(e)}") # Вывод в консоль для отладки
            # input("Нажмите Enter для завершения...")
            sys.exit(1)

    # ---------------- ОСНОВНИЙ ІНТЕРФЕЙС ----------------
    def launch_main_app():
        global root, scroll_frame, document_blocks, context_menu
        
        try:
            ctk.set_appearance_mode("light")
            ctk.set_default_color_theme("blue")

            root = SafeCTk() # Используем SafeCTk
            root.title(f"sportforall v{CURRENT_VERSION}") # Добавляем версию в заголовок
            root.geometry("1200x700")
            root.minsize(800, 600) # Минимальный размер окна

            def on_root_close():
                # cleanup_after_callbacks() # Уже вызывается в root.destroy() из SafeCTk
                if root:
                    try:
                        root.destroy()
                    except tk.TclError: pass # Окно может быть уже уничтожено
                sys.exit(0) 
                
            root.protocol("WM_DELETE_WINDOW", on_root_close)

            # Контекстное меню
            context_menu = tk.Menu(root, tearoff=0)
            # Команды будут вызываться с current_widget, установленным в show_context_menu
            context_menu.add_command(label="Копіювати", command=lambda: context_menu.current_widget.event_generate('<<Copy>>') if hasattr(context_menu, 'current_widget') else None) # type: ignore
            context_menu.add_command(label="Вставити", command=lambda: context_menu.current_widget.event_generate('<<Paste>>') if hasattr(context_menu, 'current_widget') else None) # type: ignore
            context_menu.add_command(label="Вирізати", command=lambda: context_menu.current_widget.event_generate('<<Cut>>') if hasattr(context_menu, 'current_widget') else None) # type: ignore
            context_menu.add_separator()
            context_menu.add_command(label="Виділити все", command=lambda: context_menu.current_widget.select_range(0, "end") if hasattr(context_menu, 'current_widget') else None) # type: ignore
            
            document_blocks = [] # Очищаем при запуске основного приложения

            # Верхняя панель
            top_panel_frame = ctk.CTkFrame(root) # Новый общий фрейм для кнопок и версии
            top_panel_frame.pack(pady=10, padx=10, fill="x")

            # Фрейм для кнопок слева
            buttons_frame = ctk.CTkFrame(top_panel_frame) # Без master, значит потомок top_panel_frame
            buttons_frame.pack(side="left")

            ctk.CTkButton(buttons_frame, text="➕ Додати договір", command=lambda: add_new_template(), fg_color="#2196F3").pack(side="left", padx=5)
            ctk.CTkButton(buttons_frame, text="📄 Згенерувати документи", command=generate_documents, fg_color="#4CAF50").pack(side="left", padx=5)
            
            if koshtorys and hasattr(koshtorys, 'fill_koshtorys'): # Показываем кнопку только если модуль загружен
                ctk.CTkButton(buttons_frame, text="💰 Заповнити кошторис", 
                              command=lambda: koshtorys.fill_koshtorys(document_blocks) if koshtorys else None, 
                              fg_color="#FF9800").pack(side="left", padx=5)

            # Кнопка обновления
            ctk.CTkButton(buttons_frame, text="🔄 Оновити додаток", command=download_and_apply_update, fg_color="#00BCD4").pack(side="left", padx=5)

            # Версія справа
            version_label = ctk.CTkLabel(top_panel_frame, text=f"Версія: {CURRENT_VERSION}", text_color="gray", font=("Arial", 10))
            version_label.pack(side="right", padx=10)

            scroll_frame = ctk.CTkScrollableFrame(root) # Убрал размеры, чтобы fill/expand работали лучше
            scroll_frame.pack(padx=10, pady=(0,10), fill="both", expand=True)

            def add_new_template():
                filepath = filedialog.askopenfilename(
                    title="Оберіть шаблон договору (.docm)", 
                    filetypes=[("Word документи з макросами", "*.docm"), ("Всі файли", "*.*")]
                )
                if filepath:
                    # Проверка, не добавлен ли уже этот шаблон
                    for block in document_blocks:
                        if os.path.normpath(block["path"]) == os.path.normpath(filepath):
                            messagebox.showwarning("Увага", f"Шаблон '{os.path.basename(filepath)}' вже додано.", parent=root)
                            return
                    add_document_block(filepath)

            def add_document_block(filepath=None):
                try:
                    block_frame = ctk.CTkFrame(scroll_frame, border_width=1, border_color="gray70") # Добавил рамку
                    block_frame.pack(pady=10, padx=5, fill="x")

                    # Фрейм для пути и кнопки удаления/замены шаблона
                    path_frame = ctk.CTkFrame(block_frame) # Без master - потомок block_frame
                    path_frame.pack(fill="x", padx=5, pady=(5,0))

                    path_display_text = os.path.basename(filepath) if filepath else "Шлях не вказано"
                    path_label = ctk.CTkLabel(path_frame, text=path_display_text, text_color="blue", anchor="w", wraplength=700) # Ограничение ширины текста
                    path_label.pack(side="left", padx=(0,10), fill="x", expand=True)
                    
                    # Кнопки управления блоком справа от пути
                    block_control_frame = ctk.CTkFrame(path_frame) # Без master - потомок path_frame
                    block_control_frame.pack(side="right")


                    entries = {}
                    template_memory = {}
                    if filepath:
                        template_memory = get_template_memory(filepath)

                    # Создание полей ввода
                    fields_container = ctk.CTkFrame(block_frame) # Контейнер для полей
                    fields_container.pack(fill="x", padx=5, pady=5)


                    for i, field in enumerate(FIELDS):
                        row = ctk.CTkFrame(fields_container) # master=fields_container
                        row.pack(fill="x", pady=2)

                        label = ctk.CTkLabel(row, text=f"<{field}>:", anchor="w", width=160, font=("Arial", 12)) # Немного шире
                        label.pack(side="left", padx=5)

                        entry = CustomEntry(row, field_name=field) # master=row, width не задан, чтобы растягивался
                        entry.pack(side="left", padx=5, fill="x", expand=True)

                        def make_show_hint(field_hint, field_name_display): # field_name_display для корректного отображения в заголовке
                            def show_hint_func(): # Переименовал, чтобы не конфликтовать
                                if field_hint:
                                    messagebox.showinfo(f"Підказка для <{field_name_display}>", field_hint, parent=root)
                                else:
                                    messagebox.showinfo("Підказка", "Приклад не вказано.", parent=root)
                            return show_hint_func
                        
                        hint_btn = ctk.CTkButton(row, text="ℹ", width=28, height=28, font=("Arial", 14, "bold"),
                                                command=make_show_hint(EXAMPLES.get(field), field))
                        hint_btn.pack(side="left", padx=(0,5))
                        
                        # Заполнение из памяти
                        # Сначала приоритет у template_memory (данные для конкретного шаблона)
                        value_to_insert = template_memory.get(field, "")
                        # Если в template_memory для этого поля пусто, пробуем старую общую память
                        if not value_to_insert and isinstance(load_memory().get(field), str): # Проверяем старый формат
                             value_to_insert = load_memory().get(field, "")

                        if value_to_insert:
                            entry.insert(0, value_to_insert)
                        else: # Если ничего нет, показываем плейсхолдер по умолчанию
                            entry._on_focus_out(None)


                        bind_entry_shortcuts(entry)
                        entries[field] = entry
                    
                    current_block_data = {"path": filepath, "entries": entries, "frame": block_frame, "path_label": path_label}
                    document_blocks.append(current_block_data)

                    def clear_fields_for_block():
                        if messagebox.askyesno("Очистити поля", "Ви впевнені, що хочете очистити всі поля в цьому блоці?", parent=root):
                            for entry_widget in entries.values():
                                if hasattr(entry_widget, 'clear'):
                                    entry_widget.clear()
                                else:
                                    entry_widget.delete(0, 'end')
                                    if isinstance(entry_widget, CustomEntry):
                                        entry_widget._on_focus_out(None) 
                    
                    def replace_template_for_block():
                        new_path = filedialog.askopenfilename(title="Оберіть новий шаблон (.docm)", filetypes=[("Word документи з макросами", "*.docm")])
                        if new_path:
                            # Проверка, не выбран ли уже существующий шаблон из других блоков
                            for blk in document_blocks:
                                if blk != current_block_data and os.path.normpath(blk["path"]) == os.path.normpath(new_path):
                                    messagebox.showwarning("Увага", f"Шаблон '{os.path.basename(new_path)}' вже використовується в іншому блоці.", parent=root)
                                    return

                            current_block_data["path"] = new_path
                            current_block_data["path_label"].configure(text=os.path.basename(new_path))
                            # Опционально: очистить поля или загрузить память для нового шаблона
                            new_template_memory = get_template_memory(new_path)
                            for fld, entry_widget in entries.items():
                                entry_widget.clear() # Очищаем старое
                                val = new_template_memory.get(fld, "")
                                if val:
                                    entry_widget.insert(0, val)
                                else:
                                    entry_widget._on_focus_out(None) # Показываем плейсхолдер, если нет данных
                            messagebox.showinfo("Шаблон замінено", f"Шаблон оновлено на '{os.path.basename(new_path)}'.\nДані (якщо були) завантажено для нового шаблону.", parent=root)


                    def remove_block_from_ui():
                        if messagebox.askyesno("Видалити блок", "Ви впевнені, що хочете видалити цей блок договору?", parent=root):
                            document_blocks.remove(current_block_data)
                            block_frame.destroy()
                            # Если блоков не осталось, можно показать сообщение или обновить интерфейс
                            if not document_blocks and scroll_frame.winfo_exists(): # Проверка существования scroll_frame
                                # Можно добавить Label "Список порожній"
                               pass


                    ctk.CTkButton(block_control_frame, text="🧹 Очистити", command=clear_fields_for_block, width=100).pack(side="left", padx=3)
                    ctk.CTkButton(block_control_frame, text="🔄 Замінити шаблон", command=replace_template_for_block, width=150).pack(side="left", padx=3)
                    ctk.CTkButton(block_control_frame, text="🗑 Видалити блок", command=remove_block_from_ui, fg_color="tomato", hover_color="red", width=130).pack(side="left", padx=3)
                
                except Exception as e_add_block:
                    log_error(type(e_add_block), e_add_block, sys.exc_info()[2])
                    # messagebox.showerror("Помилка", f"Помилка при створенні блоку: {str(e_add_block)}", parent=root) # Уже логируется

            # Глобальные сочетания клавиш
            # root.bind_all("<Control-s>", lambda event: export_to_excel()) # Сохранение в Excel уже интегрировано в generate_documents
            # root.bind_all("<Control-g>", lambda event: generate_documents()) # "g" может конфликтовать с другими

            # Загрузка предыдущих блоков (если это необходимо при старте)
            # Это потребует сохранения структуры document_blocks в JSON, включая пути к шаблонам.
            # На данный момент, при каждом запуске список блоков пуст.
            # Если нужно восстанавливать - нужна отдельная логика сохранения/загрузки списка блоков.

            root.mainloop()

        except Exception as e_launch: # Ловим ошибки на этапе запуска основного приложения
            # log_error(type(e_launch), e_launch, sys.exc_info()[2]) # Уже логируется
            # messagebox.showerror("Критична помилка запуску", "Виникла критична помилка під час запуску основного інтерфейсу.\nДеталі у файлі error.txt.")
            # input("Нажмите Enter для завершения...") # Не блокируем
            sys.exit(1) # Выходим, если основной интерфейс не запустился

# Блок except верхнего уровня для отлова всего остального, что не поймано ранее
except Exception as e_global:
    # Попытка записать в лог, если log_error еще доступен
    try:
        # Создаем traceback объект вручную, если sys.exc_info() вернул None (маловероятно здесь)
        exc_type, exc_value, exc_tb = sys.exc_info()
        if exc_tb is None: # Если мы вне блока except, sys.exc_info() может вернуть (None, None, None)
                           # Но здесь мы должны быть внутри except e_global
            # В качестве запасного варианта, если sys.exc_info() не сработало
            import inspect
            exc_tb = inspect.trace()[-1] # Это не стандартный traceback объект, но лучше чем ничего

        log_error(type(e_global), e_global, exc_tb if exc_tb else sys.exc_info()[2])
    except Exception as e_log_final:
        print(f"Критична помилка при логуванні фінальної помилки: {str(e_log_final)}")

    # Попытка показать MessageBox, если Tkinter еще может работать
    try:
        if tk._default_root: # Проверяем, был ли инициализирован Tk
            messagebox.showerror("Непереборна критична помилка", f"Виникла непереборна критична помилка: {str(e_global)}\nДодаток буде закрито.\nДеталі у файлі error.txt (якщо вдалося записати).")
    except: # Если Tkinter не работает
        print(f"Непереборна критична помилка (не вдалося показати MessageBox): {str(e_global)}")
        print(traceback.format_exc()) # Выводим traceback в консоль
    
    # input("Натисніть Enter для завершення...") # Не блокируем, если это GUI
    sys.exit(1) # В любом случае выходим при такой ошибке


# ВИКЛИК ПРОГРАММЫ
if __name__ == "__main__":
    try:
        # Проверяем, запущен ли уже updater.bat (чтобы избежать бесконечного цикла перезапуска)
        # Это простая проверка, можно усложнить (например, через именованный мьютекс)
        if "updater_launched" not in os.environ:
            ask_password()
        else:
            # Если это перезапуск после обновления, просто выходим, т.к. новый экземпляр должен быть запущен батником
            # Или здесь можно показать сообщение "Обновление завершено, приложение перезапускается"
            # Но батник уже должен запустить новую копию.
            sys.exit(0) 
    except SystemExit: # Чтобы SystemExit из ask_password или launch_main_app не перехватывался следующим except
        pass
    except Exception as e_main:
        # Этот блок отловит ошибки, если они произошли до установки sys.excepthook
        # или если ask_password() выбросило исключение до того, как Tkinter был инициализирован
        print(f"Фатальна помилка під час ініціалізації програми: {str(e_main)}")
        print(traceback.format_exc())
        try:
            # Попытка записать в лог, если функция log_error определена
            log_error(type(e_main), e_main, sys.exc_info()[2])
        except NameError: # Если log_error еще не определена
            pass
        except Exception as e_log_init:
            print(f"Помилка при логуванні помилки ініціалізації: {e_log_init}")
        
        # input("Натисніть Enter для завершення...")
        sys.exit(1)