# -*- coding: utf-8 -*-
"""
Графический интерфейс приложения OsaiTransfer («Менеджер УП Нестинг»).
Двухпанельный менеджер для передачи УП (.anc, .cnc, .pgm, .txt, .nc)
между ПК оператора и стойкой ЧПУ OSAI COMPACT в обход ошибки Win32 58.
"""

import os
import posixpath
import shutil
import sys
import tempfile
import configparser
import threading
from datetime import datetime
import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog

try:
    from tkinterdnd2 import TkinterDnD, DND_FILES
    HAS_DND = True
except ImportError:
    HAS_DND = False

from smb_client import OsaiSMBClient


DEFAULT_EXTENSIONS = "*.anc;*.cnc;*.pgm;*.txt;*.nc"
PROTECTED_REMOTE_PATHS = {
    "", "/", "/user", "/oem", "/programs", "/macros", "/programs/work"
}


def get_app_dir():
    """Возвращает каталог рядом с исполняемым файлом (работает и в PyInstaller --onefile)."""
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def format_size(bytes_size):
    """Преобразование размера файла в понятный формат (Б, КБ, МБ)."""
    if bytes_size is None or bytes_size == 0:
        return "0 Б"
    for unit in ("Б", "КБ", "МБ", "ГБ"):
        if bytes_size < 1024.0:
            return f"{bytes_size:.1f} {unit}" if unit != "Б" else f"{int(bytes_size)} {unit}"
        bytes_size /= 1024.0
    return f"{bytes_size:.1f} ТБ"


def resolve_shortcut(lnk_path):
    """
    Разрешение Windows-ярлыка (.lnk) через WScript.Shell.
    Возвращает кортеж (target_path, is_dir, is_file) или (None, False, False).
    """
    if sys.platform != "win32" or not lnk_path.lower().endswith(".lnk"):
        return None, False, False
    try:
        import pythoncom
        import win32com.client

        pythoncom.CoInitialize()
        shell = None
        shortcut = None
        try:
            shell = win32com.client.Dispatch("WScript.Shell")
            shortcut = shell.CreateShortCut(lnk_path)
            target = shortcut.TargetPath
            if target:
                target = os.path.normpath(target)
                is_dir = os.path.isdir(target)
                is_file = (not is_dir) and os.path.isfile(target)
                return target, is_dir, is_file
        finally:
            shortcut = None
            shell = None
            pythoncom.CoUninitialize()
    except Exception:
        pass
    return None, False, False


def parse_extensions(ext_string):
    """Преобразование строки вида '*.anc;*.cnc' в кортеж расширений ('.anc', '.cnc')."""
    result = []
    for part in ext_string.replace(",", ";").split(";"):
        clean = part.strip().lstrip("*").lower()
        if clean:
            if not clean.startswith("."):
                clean = "." + clean
            result.append(clean)
    return tuple(result) if result else (".anc", ".cnc", ".pgm", ".txt", ".nc")


class OsaiTransferApp:
    """Главное окно двухпанельного менеджера передачи УП на ЧПУ OSAI."""

    def __init__(self, root):
        self.root = root
        self.root.title("OsaiTransfer — Передача УП на ЧПУ OSAI")
        self.root.geometry("1100x680")
        self.root.minsize(880, 520)

        # Чтение конфигурации рядом с exe/скриптом
        self.config_file = os.path.join(get_app_dir(), "config.ini")
        self.config = configparser.ConfigParser()
        self.load_config()

        # Инициализация клиента SMB
        self.smb_client = OsaiSMBClient(
            host=self.config.get("CNC", "host", fallback="192.168.0.2"),
            port=self.config.getint("CNC", "port", fallback=445),
            remote_name=self.config.get("CNC", "remote_name", fallback="COMPACT"),
            client_name=self.config.get("CNC", "client_name", fallback="OPERATOR-PC"),
            username=self.config.get("CNC", "username", fallback=""),
            password=self.config.get("CNC", "password", fallback="123"),
            timeout=self.config.getint("CNC", "timeout", fallback=5)
        )

        # Расширения УП из конфига
        ext_cfg = self.config.get("Local", "extensions", fallback=DEFAULT_EXTENSIONS)
        self.up_extensions = parse_extensions(ext_cfg)

        # Текущие состояния путей
        default_local = self.config.get("Local", "default_path", fallback="")
        if not default_local or not os.path.exists(default_local):
            default_local = os.path.expanduser("~/Desktop")
        self.current_local_path = default_local

        self.current_share = self.config.get("CNC", "share", fallback="USER")
        self.current_remote_path = self.config.get("CNC", "remote_dir", fallback="/PROGRAMS/WORK")
        self.current_remote_raw_path = self.current_remote_path
        self.is_connected = False
        self._busy = False

        # Словари метаданных элементов Treeview (item_id -> dict)
        self.local_items = {}
        self.remote_items = {}

        # Состояние сортировки таблиц
        self.local_sort = {"col": "name", "reverse": False}
        self.remote_sort = {"col": "name", "reverse": False}
        self._col_titles = {
            "name": "Имя файла / папки",
            "size": "Размер",
            "date": "Дата изменения",
        }

        # Состояние внутреннего перетаскивания (Drag & Drop между панелями)
        self._drag_state = None
        self._drag_tooltip = None

        # Настройка стилей и построение интерфейса
        self.setup_styles()
        self.build_ui()
        self.setup_drag_and_drop()

        # Первичная загрузка списков
        self.refresh_local()
        self.refresh_remote()

        # Запуск фонового монитора связи
        self.start_connection_monitor()

        # Временная папка для просмотра и системного меню файлов с ЧПУ
        self._preview_temp_dir = os.path.join(tempfile.gettempdir(), "OsaiTransferPreview")
        self._last_popup_pos = (200, 200)

        # Горячие клавиши и закрытие окна
        self.root.bind("<F5>", lambda event: self.refresh_all())
        self.right_tree.bind("<Delete>", lambda event: self.delete_remote_selected())
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

    def load_config(self):
        """Загрузка настроек из config.ini."""
        if os.path.exists(self.config_file):
            self.config.read(self.config_file, encoding="utf-8")
        if "CNC" not in self.config:
            self.config["CNC"] = {
                "host": "192.168.0.2",
                "port": "445",
                "remote_name": "COMPACT",
                "client_name": "OPERATOR-PC",
                "share": "USER",
                "remote_dir": "/PROGRAMS/WORK",
                "timeout": "5",
                "username": "",
                "password": "123"
            }
        if "Local" not in self.config:
            self.config["Local"] = {
                "default_path": os.path.expanduser("~/Desktop"),
                "extensions": DEFAULT_EXTENSIONS
            }
        if "App" not in self.config:
            self.config["App"] = {
                "theme": "light",
                "ping_interval": "5"
            }

    def save_config(self):
        """Сохранение изменений в config.ini."""
        try:
            self.config.set("Local", "default_path", self.current_local_path)
            self.config.set("CNC", "share", self.current_share)
            self.config.set("CNC", "remote_dir", self.current_remote_path)
            with open(self.config_file, "w", encoding="utf-8") as f:
                self.config.write(f)
        except Exception:
            pass

    def setup_styles(self):
        """Оформление внешнего вида интерфейса."""
        style = ttk.Style()
        try:
            style.theme_use("clam")
        except Exception:
            pass

        font_family = "Segoe UI" if sys.platform == "win32" else "Helvetica"
        self.base_font = (font_family, 10)
        self.header_font = (font_family, 10, "bold")

        style.configure("Treeview", font=self.base_font, rowheight=24)
        style.configure("Treeview.Heading", font=self.header_font)
        style.configure("TButton", font=self.base_font, padding=4)
        style.configure("Action.TButton", font=(font_family, 10, "bold"), padding=6)
        style.configure("TLabel", font=self.base_font)

    def build_ui(self):
        """Создание элементов двухпанельного интерфейса."""
        # Верхняя панель статуса
        top_bar = ttk.Frame(self.root, padding="8 6 8 6")
        top_bar.pack(side=tk.TOP, fill=tk.X)

        self.status_canvas = tk.Canvas(top_bar, width=16, height=16, highlightthickness=0)
        self.status_canvas.pack(side=tk.LEFT, padx=(0, 6))
        self.status_dot = self.status_canvas.create_oval(2, 2, 14, 14, fill="gray", outline="")

        self.status_label = ttk.Label(top_bar, text="Проверка связи с ЧПУ...", font=self.header_font)
        self.status_label.pack(side=tk.LEFT)

        btn_refresh_all = ttk.Button(top_bar, text="🔄 Обновить всё (F5)", command=self.refresh_all)
        btn_refresh_all.pack(side=tk.RIGHT, padx=4)

        btn_settings = ttk.Button(top_bar, text="⚙ Настройки", command=self.open_settings_dialog)
        btn_settings.pack(side=tk.RIGHT, padx=4)

        # Основной контейнер с панелями
        main_paned = ttk.PanedWindow(self.root, orient=tk.HORIZONTAL)
        main_paned.pack(fill=tk.BOTH, expand=True, padx=8, pady=4)

        # ================= ЛЕВАЯ ПАНЕЛЬ (ПК оператора) =================
        left_frame = ttk.Frame(main_paned, padding=4)
        main_paned.add(left_frame, weight=1)

        left_header = ttk.Frame(left_frame)
        left_header.pack(fill=tk.X, pady=(0, 4))
        ttk.Label(left_header, text="💻 Компьютер оператора", font=self.header_font).pack(side=tk.LEFT)

        loc_nav = ttk.Frame(left_frame)
        loc_nav.pack(fill=tk.X, pady=(0, 4))
        self.loc_path_entry = ttk.Entry(loc_nav, font=self.base_font)
        self.loc_path_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 4))
        self.loc_path_entry.insert(0, self.current_local_path)
        self.loc_path_entry.bind("<Return>", lambda e: self.on_local_path_enter())

        btn_up_local = ttk.Button(loc_nav, text="Вверх ⬆", width=8, command=self.navigate_local_up)
        btn_up_local.pack(side=tk.RIGHT, padx=(2, 0))

        btn_browse_local = ttk.Button(loc_nav, text="Обзор...", width=8, command=self.browse_local_folder)
        btn_browse_local.pack(side=tk.RIGHT)

        # Фильтр расширений
        loc_filter_bar = ttk.Frame(left_frame)
        loc_filter_bar.pack(fill=tk.X, pady=(0, 4))
        ttk.Label(loc_filter_bar, text="Фильтр:").pack(side=tk.LEFT, padx=(0, 4))

        ext_label = " ".join(f"*{e}" for e in self.up_extensions)
        self.up_filter_label = f"УП ({ext_label})"
        self.loc_filter_var = tk.StringVar(value=self.up_filter_label)
        self.loc_filter_cb = ttk.Combobox(
            loc_filter_bar,
            textvariable=self.loc_filter_var,
            values=[self.up_filter_label, "Все файлы (*.*)"],
            state="readonly",
            width=34
        )
        self.loc_filter_cb.pack(side=tk.LEFT)
        self.loc_filter_cb.bind("<<ComboboxSelected>>", lambda e: self.refresh_local())

        # Таблица локальных файлов
        left_tree_frame = ttk.Frame(left_frame)
        left_tree_frame.pack(fill=tk.BOTH, expand=True)

        self.left_tree = ttk.Treeview(
            left_tree_frame, columns=("name", "size", "date"),
            show="headings", selectmode="extended"
        )
        self.left_tree.heading("name", text="Имя файла / папки ▲", anchor=tk.W,
                               command=lambda: self.on_heading_click("left", "name"))
        self.left_tree.heading("size", text="Размер", anchor=tk.E,
                               command=lambda: self.on_heading_click("left", "size"))
        self.left_tree.heading("date", text="Дата изменения", anchor=tk.CENTER,
                               command=lambda: self.on_heading_click("left", "date"))

        self.left_tree.column("name", width=230, anchor=tk.W)
        self.left_tree.column("size", width=80, anchor=tk.E)
        self.left_tree.column("date", width=130, anchor=tk.CENTER)

        left_scroll = ttk.Scrollbar(left_tree_frame, orient=tk.VERTICAL, command=self.left_tree.yview)
        self.left_tree.configure(yscrollcommand=left_scroll.set)
        self.left_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        left_scroll.pack(side=tk.RIGHT, fill=tk.Y)

        self.left_tree.bind("<Double-1>", self.on_local_item_double_click)

        # Контекстное меню левой панели (fallback, если системное меню недоступно)
        self.local_menu = tk.Menu(self.root, tearoff=0)
        self.local_menu.add_command(label="Загрузить на ЧПУ ➔", command=self.upload_selected)
        self.local_menu.add_command(label="📄 Быстрый просмотр текста", command=self.preview_local_selected)
        self.local_menu.add_separator()
        self.local_menu.add_command(label="🔄 Обновить", command=self.refresh_local)
        self.left_tree.bind("<Button-3>", self.show_local_context_menu)

        # ================= ЦЕНТРАЛЬНАЯ ПАНЕЛЬ ДЕЙСТВИЙ =================
        mid_frame = ttk.Frame(main_paned, padding="6 10 6 10")
        main_paned.add(mid_frame, weight=0)

        spacer_top = ttk.Frame(mid_frame)
        spacer_top.pack(fill=tk.BOTH, expand=True)

        btn_upload = ttk.Button(mid_frame, text="Загрузить ➔", style="Action.TButton", command=self.upload_selected)
        btn_upload.pack(fill=tk.X, pady=6)

        btn_download = ttk.Button(mid_frame, text="⬅ Скачать", style="Action.TButton", command=self.download_selected)
        btn_download.pack(fill=tk.X, pady=6)

        ttk.Separator(mid_frame, orient=tk.HORIZONTAL).pack(fill=tk.X, pady=12)

        btn_mkdir_remote = ttk.Button(mid_frame, text="📁➕ Папка на ЧПУ", command=self.create_remote_folder)
        btn_mkdir_remote.pack(fill=tk.X, pady=4)

        btn_del_remote = ttk.Button(mid_frame, text="🗑 Удалить с ЧПУ", command=self.delete_remote_selected)
        btn_del_remote.pack(fill=tk.X, pady=4)

        spacer_bottom = ttk.Frame(mid_frame)
        spacer_bottom.pack(fill=tk.BOTH, expand=True)

        # ================= ПРАВАЯ ПАНЕЛЬ (Стойка ЧПУ OSAI) =================
        right_frame = ttk.Frame(main_paned, padding=4)
        main_paned.add(right_frame, weight=1)

        right_header = ttk.Frame(right_frame)
        right_header.pack(fill=tk.X, pady=(0, 4))
        ttk.Label(right_header, text="⚙ Стойка ЧПУ OSAI COMPACT", font=self.header_font).pack(side=tk.LEFT)
        ttk.Label(
            right_header,
            text="(можно перетаскивать файлы мышью)",
            foreground="#555555"
        ).pack(side=tk.RIGHT)

        rem_nav = ttk.Frame(right_frame)
        rem_nav.pack(fill=tk.X, pady=(0, 4))

        self.share_var = tk.StringVar(value=self.current_share)
        self.share_cb = ttk.Combobox(
            rem_nav, textvariable=self.share_var, values=["USER", "OEM"],
            state="readonly", width=7
        )
        self.share_cb.pack(side=tk.LEFT, padx=(0, 4))
        self.share_cb.bind("<<ComboboxSelected>>", self.on_share_change)

        self.rem_path_entry = ttk.Entry(rem_nav, font=self.base_font)
        self.rem_path_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 4))
        self.rem_path_entry.insert(0, self.current_remote_path)
        self.rem_path_entry.bind("<Return>", lambda e: self.on_remote_path_enter())

        btn_up_remote = ttk.Button(rem_nav, text="Вверх ⬆", width=8, command=self.navigate_remote_up)
        btn_up_remote.pack(side=tk.RIGHT)

        # Вторая строка навигации ЧПУ для кнопки создания папки и подсказки
        rem_tools_bar = ttk.Frame(right_frame)
        rem_tools_bar.pack(fill=tk.X, pady=(0, 4))
        ttk.Button(rem_tools_bar, text="📁➕ Новая папка", command=self.create_remote_folder).pack(side=tk.LEFT, padx=(0, 4))
        ttk.Button(rem_tools_bar, text="🔄 Обновить ЧПУ", command=self.refresh_remote).pack(side=tk.LEFT)

        # Таблица файлов ЧПУ
        right_tree_frame = ttk.Frame(right_frame)
        right_tree_frame.pack(fill=tk.BOTH, expand=True)

        self.right_tree = ttk.Treeview(
            right_tree_frame, columns=("name", "size", "date"),
            show="headings", selectmode="extended"
        )
        self.right_tree.heading("name", text="Имя файла / папки ▲", anchor=tk.W,
                                command=lambda: self.on_heading_click("right", "name"))
        self.right_tree.heading("size", text="Размер", anchor=tk.E,
                                command=lambda: self.on_heading_click("right", "size"))
        self.right_tree.heading("date", text="Дата изменения", anchor=tk.CENTER,
                                command=lambda: self.on_heading_click("right", "date"))

        self.right_tree.column("name", width=230, anchor=tk.W)
        self.right_tree.column("size", width=80, anchor=tk.E)
        self.right_tree.column("date", width=130, anchor=tk.CENTER)

        right_scroll = ttk.Scrollbar(right_tree_frame, orient=tk.VERTICAL, command=self.right_tree.yview)
        self.right_tree.configure(yscrollcommand=right_scroll.set)
        self.right_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        right_scroll.pack(side=tk.RIGHT, fill=tk.Y)

        self.right_tree.bind("<Double-1>", self.on_remote_item_double_click)

        # Контекстное меню правой панели
        self.remote_menu = tk.Menu(self.root, tearoff=0)
        self.remote_menu.add_command(label="👁 Открыть в программе Windows", command=self.open_remote_selected)
        self.remote_menu.add_command(label="📄 Быстрый просмотр текста (УП)", command=self.preview_remote_selected)
        self.remote_menu.add_command(label="🪟 Системное меню Windows...", command=self.show_remote_shell_menu)
        self.remote_menu.add_separator()
        self.remote_menu.add_command(label="⬅ Скачать на ПК", command=self.download_selected)
        self.remote_menu.add_command(label="📁➕ Создать папку...", command=self.create_remote_folder)
        self.remote_menu.add_separator()
        self.remote_menu.add_command(label="🗑 Удалить с ЧПУ", command=self.delete_remote_selected)
        self.remote_menu.add_separator()
        self.remote_menu.add_command(label="🔄 Обновить", command=self.refresh_remote)
        self.right_tree.bind("<Button-3>", self.show_remote_context_menu)

        # ================= СТРОКА СОСТОЯНИЯ (НИЗ) =================
        bottom_bar = ttk.Frame(self.root, padding="8 4 8 6")
        bottom_bar.pack(side=tk.BOTTOM, fill=tk.X)

        self.info_label = ttk.Label(bottom_bar, text="Готов к работе")
        self.info_label.pack(side=tk.LEFT, fill=tk.X, expand=True)

        self.progress_bar = ttk.Progressbar(bottom_bar, orient=tk.HORIZONTAL, mode="indeterminate", length=160)

    def on_close(self):
        """Очистка временного кэша просмотра и закрытие приложения."""
        try:
            shutil.rmtree(self._preview_temp_dir, ignore_errors=True)
        except Exception:
            pass
        try:
            self.smb_client.disconnect()
        except Exception:
            pass
        self.root.destroy()

    def _show_windows_shell_context_menu(self, paths, x_root, y_root, extra_actions=None):
        """
        Показывает нативное системное контекстное меню Windows (IContextMenu)
        для списка локальных файлов/папок с необязательными пунктами сверху.
        Возвращает True, если меню было показано (даже если пользователь закрыл его без выбора).
        """
        if sys.platform != "win32" or not paths:
            return False
        valid_paths = [os.path.normpath(p) for p in paths if p and os.path.exists(p)]
        if not valid_paths:
            return False

        try:
            import pythoncom
            import win32gui
            import win32con
            from win32com.shell import shell, shellcon
        except ImportError:
            return False

        pythoncom.CoInitialize()
        hmenu = None
        cm = None
        parent = None
        desktop = None
        try:
            hwnd = win32gui.GetParent(self.root.winfo_id()) or self.root.winfo_id()
            pidls = []
            for p in valid_paths:
                pidl, _ = shell.SHParseDisplayName(p, 0)
                if pidl:
                    pidls.append(pidl)
            if not pidls:
                return False

            parent_pidl = pidls[0][:-1]
            # Shell требует элементы из одного родительского каталога
            same_parent = [p for p in pidls if p[:-1] == parent_pidl]
            child_pidls = [[p[-1]] for p in (same_parent or [pidls[0]])]

            desktop = shell.SHGetDesktopFolder()
            parent = desktop.BindToObject(parent_pidl, None, shell.IID_IShellFolder) if parent_pidl else desktop
            _, cm = parent.GetUIObjectOf(hwnd, child_pidls, shell.IID_IContextMenu)

            hmenu = win32gui.CreatePopupMenu()
            insert_pos = 0
            actions = extra_actions or []
            for idx, (label, _) in enumerate(actions, start=1):
                win32gui.InsertMenu(
                    hmenu, insert_pos, win32con.MF_BYPOSITION | win32con.MF_STRING, idx, label
                )
                insert_pos += 1
            if actions:
                win32gui.InsertMenu(hmenu, insert_pos, win32con.MF_BYPOSITION | win32con.MF_SEPARATOR, 0, "")
                insert_pos += 1

            id_first = 1000
            cm.QueryContextMenu(
                hmenu, insert_pos, id_first, 0x7FFF,
                shellcon.CMF_NORMAL | shellcon.CMF_EXPLORE
            )

            win32gui.SetForegroundWindow(hwnd)
            cmd = win32gui.TrackPopupMenu(
                hmenu,
                win32con.TPM_LEFTALIGN | win32con.TPM_RETURNCMD | win32con.TPM_RIGHTBUTTON,
                int(x_root),
                int(y_root),
                0,
                hwnd,
                None,
            )

            if 1 <= cmd <= len(actions):
                _, cb = actions[cmd - 1]
                self.root.after(0, cb)
            elif cmd >= id_first:
                cm.InvokeCommand((0, hwnd, cmd - id_first, None, None, win32con.SW_SHOWNORMAL, 0, 0))
                self.root.after(450, self.refresh_local)
            return True
        except Exception:
            return False
        finally:
            if hmenu:
                try:
                    win32gui.DestroyMenu(hmenu)
                except Exception:
                    pass
            cm = None
            parent = None
            desktop = None
            pythoncom.CoUninitialize()

    def show_local_context_menu(self, event):
        """Показ нативного системного контекстного меню Windows по локальному списку."""
        row_id = self.left_tree.identify_row(event.y)
        if row_id and row_id not in self.left_tree.selection():
            self.left_tree.selection_set(row_id)
            self.left_tree.focus(row_id)

        selected = self.left_tree.selection()
        target_paths = []
        single_file = False
        for iid in selected:
            meta = self.local_items.get(iid)
            if meta and not meta.get("is_up") and meta.get("path"):
                target_paths.append(meta["path"])
        if len(target_paths) == 1 and os.path.isfile(target_paths[0]):
            single_file = True

        if target_paths:
            extra = [("Загрузить на ЧПУ ➔", self.upload_selected)]
            if single_file:
                extra.append(("Быстрый просмотр текста", self.preview_local_selected))
            if self._show_windows_shell_context_menu(target_paths, event.x_root, event.y_root, extra_actions=extra):
                return

        self.local_menu.tk_popup(event.x_root, event.y_root)

    def show_remote_context_menu(self, event):
        """Показ контекстного меню правой кнопкой мыши по списку ЧПУ."""
        self._last_popup_pos = (event.x_root, event.y_root)
        row_id = self.right_tree.identify_row(event.y)
        if row_id and row_id not in self.right_tree.selection():
            self.right_tree.selection_set(row_id)
            self.right_tree.focus(row_id)
        self.remote_menu.tk_popup(event.x_root, event.y_root)

    # ------------------ Просмотр и открытие файлов ------------------
    def _open_path_in_windows(self, file_path):
        """Открывает локальный файл в ассоциированной программе Windows."""
        try:
            if sys.platform == "win32":
                os.startfile(file_path)
            else:
                import subprocess
                opener = "open" if sys.platform == "darwin" else "xdg-open"
                subprocess.Popen([opener, file_path])
        except Exception as e:
            messagebox.showerror("Ошибка открытия", f"Не удалось открыть файл:\n{file_path}\n\n{e}")

    def _open_text_preview_window(self, title, file_path):
        """Открывает встроенное окно быстрого просмотра текста/УП с автоопределением кодировки."""
        max_bytes = 1024 * 1024  # 1 МБ для мгновенного отображения
        try:
            file_size = os.path.getsize(file_path)
            with open(file_path, "rb") as f:
                raw = f.read(max_bytes)
        except Exception as e:
            messagebox.showerror("Ошибка чтения", f"Не удалось прочитать файл:\n{e}")
            return

        is_binary = b"\x00" in raw[:4096]
        if is_binary:
            if messagebox.askyesno(
                "Двоичный файл",
                f"Файл «{title}» не является текстовым (например, изображение или PDF).\n\n"
                f"Открыть его в программе Windows по умолчанию?"
            ):
                self._open_path_in_windows(file_path)
            return

        text = None
        for enc in ("utf-8-sig", "cp1251", "utf-8"):
            try:
                text = raw.decode(enc)
                break
            except UnicodeDecodeError:
                continue
        if text is None:
            text = raw.decode("utf-8", errors="replace")

        truncated = file_size > max_bytes

        win = tk.Toplevel(self.root)
        win.title(f"Просмотр: {title}")
        win.geometry("780x560")

        toolbar = ttk.Frame(win, padding="8 6 8 4")
        toolbar.pack(side=tk.TOP, fill=tk.X)

        info_text = f"{title} ({format_size(file_size)})"
        if truncated:
            info_text += " — показан первый 1 МБ"
        ttk.Label(toolbar, text=info_text, font=self.header_font).pack(side=tk.LEFT)

        def copy_all():
            win.clipboard_clear()
            win.clipboard_append(text)
            self.info_label.config(text=f"Содержимое «{title}» скопировано в буфер обмена")

        ttk.Button(
            toolbar,
            text="Открыть в программе Windows",
            command=lambda p=file_path: self._open_path_in_windows(p)
        ).pack(side=tk.RIGHT, padx=(4, 0))
        ttk.Button(toolbar, text="Копировать всё", command=copy_all).pack(side=tk.RIGHT, padx=4)

        text_frame = ttk.Frame(win, padding="8 0 8 8")
        text_frame.pack(fill=tk.BOTH, expand=True)

        mono_font = ("Consolas", 10) if sys.platform == "win32" else ("Menlo", 11)
        txt_widget = tk.Text(text_frame, wrap=tk.NONE, font=mono_font, undo=False)
        y_scroll = ttk.Scrollbar(text_frame, orient=tk.VERTICAL, command=txt_widget.yview)
        x_scroll = ttk.Scrollbar(text_frame, orient=tk.HORIZONTAL, command=txt_widget.xview)
        txt_widget.configure(yscrollcommand=y_scroll.set, xscrollcommand=x_scroll.set)

        y_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        x_scroll.pack(side=tk.BOTTOM, fill=tk.X)
        txt_widget.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        txt_widget.insert("1.0", text)
        txt_widget.configure(state="disabled")

    def preview_local_selected(self):
        """Встроенный текстовый просмотр выбранного локального файла."""
        for iid in self.left_tree.selection():
            meta = self.local_items.get(iid)
            if meta and not meta.get("is_up") and not meta.get("is_dir") and os.path.isfile(meta["path"]):
                self._open_text_preview_window(meta["name"], meta["path"])
                return
        messagebox.showinfo("Просмотр", "Выберите один файл в левой панели.")

    def _get_selected_remote_single_file(self):
        """Возвращает метаданные выбранного одиночного файла на ЧПУ или None."""
        for iid in self.right_tree.selection():
            meta = self.remote_items.get(iid)
            if meta and not meta.get("is_up"):
                if meta.get("is_dir"):
                    messagebox.showinfo("Просмотр файла", "Выберите файл на ЧПУ (для папок откройте её двойным кликом).")
                    return None
                return meta
        messagebox.showinfo("Просмотр файла", "Выберите файл в правой панели ЧПУ.")
        return None

    def _fetch_remote_to_temp_and_run(self, meta, action_cb):
        """Скачивает одиночный файл с ЧПУ во временный кэш и выполняет действие в GUI-потоке."""
        if self._busy:
            messagebox.showwarning("Подождите", "Дождитесь завершения текущей операции.")
            return

        raw_name = meta.get("raw_name", meta["name"])
        rem_path = f"{self.current_remote_raw_path.rstrip('/')}/{raw_name}"
        share = self.current_share
        safe_sub = os.path.join(self._preview_temp_dir, share)
        os.makedirs(safe_sub, exist_ok=True)
        local_temp = os.path.join(safe_sub, meta["name"])

        self._set_busy(True, f"Получение файла со станка для просмотра: {meta['name']}...")

        def worker():
            err_msg = None
            try:
                self.smb_client.download_file(share, rem_path, local_temp)
            except Exception as e:
                err_msg = str(e)

            def finish():
                self._set_busy(False)
                if err_msg:
                    self.info_label.config(text=f"Ошибка чтения файла с ЧПУ: {err_msg}")
                    messagebox.showerror("Ошибка чтения с ЧПУ", f"Не удалось получить файл «{meta['name']}»:\n{err_msg}")
                else:
                    self.info_label.config(text=f"Файл «{meta['name']}» загружен во временный кэш просмотра")
                    action_cb(local_temp)

            self.root.after(0, finish)

        threading.Thread(target=worker, daemon=True).start()

    def open_remote_selected(self):
        """Открывает выбранный файл со станка в штатной программе Windows (фото, PDF, блокнот)."""
        meta = self._get_selected_remote_single_file()
        if meta:
            self._fetch_remote_to_temp_and_run(meta, self._open_path_in_windows)

    def preview_remote_selected(self):
        """Быстрый встроенный текстовый просмотр файла со станка."""
        meta = self._get_selected_remote_single_file()
        if meta:
            self._fetch_remote_to_temp_and_run(
                meta,
                lambda p, title=meta["name"]: self._open_text_preview_window(title, p)
            )

    def show_remote_shell_menu(self):
        """Показывает нативное системное меню Windows для временной копии файла со станка."""
        meta = self._get_selected_remote_single_file()
        if not meta:
            return
        x_root, y_root = self._last_popup_pos

        def open_menu(temp_path):
            self.info_label.config(
                text=f"Системное меню Windows для временной копии «{meta['name']}» (удаление на ЧПУ — через кнопку «Удалить с ЧПУ»)"
            )
            if not self._show_windows_shell_context_menu([temp_path], x_root, y_root):
                self._open_path_in_windows(temp_path)

        self._fetch_remote_to_temp_and_run(meta, open_menu)

    # ------------------ Сортировка таблиц ------------------
    def on_heading_click(self, side, col):
        """Переключение сортировки по нажатию на заголовок столбца."""
        sort_cfg = self.local_sort if side == "left" else self.remote_sort
        if sort_cfg["col"] == col:
            sort_cfg["reverse"] = not sort_cfg["reverse"]
        else:
            sort_cfg["col"] = col
            sort_cfg["reverse"] = False
        self._apply_tree_sort(side)

    def _update_heading_indicators(self, side):
        """Обновление стрелок ▲/▼ в заголовках столбцов."""
        tree = self.left_tree if side == "left" else self.right_tree
        sort_cfg = self.local_sort if side == "left" else self.remote_sort
        for col, base_title in self._col_titles.items():
            if col == sort_cfg["col"]:
                arrow = " ▼" if sort_cfg["reverse"] else " ▲"
                tree.heading(col, text=base_title + arrow)
            else:
                tree.heading(col, text=base_title)

    def _apply_tree_sort(self, side):
        """Сортирует строки таблицы на месте (папка '..' всегда сверху, папки перед файлами)."""
        tree = self.left_tree if side == "left" else self.right_tree
        items_map = self.local_items if side == "left" else self.remote_items
        sort_cfg = self.local_sort if side == "left" else self.remote_sort
        col = sort_cfg["col"]
        reverse = sort_cfg["reverse"]

        self._update_heading_indicators(side)

        up_rows = []
        dir_rows = []
        file_rows = []

        for item_id in tree.get_children():
            meta = items_map.get(item_id, {})
            if meta.get("is_up"):
                up_rows.append((item_id, meta))
            elif meta.get("is_dir"):
                dir_rows.append((item_id, meta))
            else:
                file_rows.append((item_id, meta))

        def dir_key(pair):
            _, meta = pair
            if col == "date":
                return (meta.get("sort_mtime", 0.0), meta.get("sort_name", ""))
            return (meta.get("sort_name", ""),)

        def file_key(pair):
            _, meta = pair
            if col == "size":
                return (meta.get("sort_size", 0), meta.get("sort_name", ""))
            if col == "date":
                return (meta.get("sort_mtime", 0.0), meta.get("sort_name", ""))
            return (meta.get("sort_name", ""),)

        dir_rows.sort(key=dir_key, reverse=reverse)
        file_rows.sort(key=file_key, reverse=reverse)

        ordered = up_rows + dir_rows + file_rows
        for idx, (item_id, _) in enumerate(ordered):
            tree.move(item_id, "", idx)

    # ------------------ Drag & Drop ------------------
    def setup_drag_and_drop(self):
        """Регистрация зон перетаскивания (внешний D&D из Проводника + внутренний между панелями)."""
        # Внутренний Drag & Drop между левой и правой панелями
        self.left_tree.bind("<ButtonPress-1>", lambda e: self._on_tree_drag_start("left", e), add="+")
        self.left_tree.bind("<B1-Motion>", self._on_tree_drag_motion, add="+")
        self.left_tree.bind("<ButtonRelease-1>", self._on_tree_drag_release, add="+")

        self.right_tree.bind("<ButtonPress-1>", lambda e: self._on_tree_drag_start("right", e), add="+")
        self.right_tree.bind("<B1-Motion>", self._on_tree_drag_motion, add="+")
        self.right_tree.bind("<ButtonRelease-1>", self._on_tree_drag_release, add="+")

        # Внешний Drag & Drop из Проводника Windows (tkinterdnd2)
        if not HAS_DND:
            return
        try:
            self.right_tree.drop_target_register(DND_FILES)
            self.right_tree.dnd_bind("<<Drop>>", self.on_drop_to_remote)
            self.left_tree.drop_target_register(DND_FILES)
            self.left_tree.dnd_bind("<<Drop>>", self.on_drop_to_local)
        except Exception as e:
            self.info_label.config(text=f"Предупреждение Drag & Drop: {e}")

    def _is_point_in_widget(self, widget, x_root, y_root):
        """Проверяет, находятся ли экранные координаты курсора внутри области виджета."""
        if not widget.winfo_ismapped():
            return False
        wx = widget.winfo_rootx()
        wy = widget.winfo_rooty()
        ww = widget.winfo_width()
        wh = widget.winfo_height()
        return wx <= x_root <= wx + ww and wy <= y_root <= wy + wh

    def _show_drag_tooltip(self, x_root, y_root, text):
        """Показывает плавающую подсказку перетаскивания рядом с курсором."""
        if self._drag_tooltip is None:
            tip = tk.Toplevel(self.root)
            tip.overrideredirect(True)
            tip.attributes("-topmost", True)
            lbl = tk.Label(
                tip,
                text=text,
                bg="#e8f0fe",
                fg="#1a3d7c",
                relief="solid",
                borderwidth=1,
                padx=8,
                pady=4,
                font=self.base_font,
            )
            lbl.pack()
            tip._lbl = lbl
            self._drag_tooltip = tip
        else:
            self._drag_tooltip._lbl.config(text=text)
        self._drag_tooltip.geometry(f"+{x_root + 14}+{y_root + 14}")

    def _hide_drag_tooltip(self):
        """Скрывает плавающую подсказку перетаскивания."""
        if self._drag_tooltip is not None:
            try:
                self._drag_tooltip.destroy()
            except Exception:
                pass
            self._drag_tooltip = None
        self.root.config(cursor="")

    def _collect_drag_items(self, side):
        """Собирает выбранные объекты для перетаскивания из указанной панели."""
        tree = self.left_tree if side == "left" else self.right_tree
        items_map = self.local_items if side == "left" else self.remote_items
        result = []
        for item_id in tree.selection():
            meta = items_map.get(item_id)
            if not meta or meta.get("is_up"):
                continue
            if side == "left":
                path = meta["path"]
                if meta.get("is_dir") and os.path.isdir(path):
                    result.append({"path": path, "is_dir": True, "name": os.path.basename(path)})
                elif os.path.isfile(path):
                    result.append({"path": path, "is_dir": False, "name": os.path.basename(path)})
            else:
                raw_name = meta.get("raw_name", meta["name"])
                rem_raw_path = f"{self.current_remote_raw_path.rstrip('/')}/{raw_name}"
                result.append({
                    "name": meta["name"],
                    "raw_name": raw_name,
                    "is_dir": meta["is_dir"],
                    "remote_path": rem_raw_path,
                })
        return result

    def _on_tree_drag_start(self, side, event):
        """Фиксирует начало возможного перетаскивания строки таблицы."""
        tree = self.left_tree if side == "left" else self.right_tree
        items_map = self.local_items if side == "left" else self.remote_items

        region = tree.identify_region(event.x, event.y)
        if region in ("heading", "separator"):
            self._drag_state = None
            return

        row_id = tree.identify_row(event.y)
        if not row_id:
            self._drag_state = None
            return

        meta = items_map.get(row_id)
        if not meta or meta.get("is_up"):
            self._drag_state = None
            return

        has_mod = bool(event.state & 0x0005)  # Shift (0x1) или Control (0x4)
        current_sel = tree.selection()
        collapse_on_release = False
        if (not has_mod) and (row_id in current_sel) and (len(current_sel) > 1):
            # Сохраняем множественное выделение до выяснения: клик это или перетаскивание
            tree.focus(row_id)
            collapse_on_release = True

        self._drag_state = {
            "side": side,
            "start_x": event.x_root,
            "start_y": event.y_root,
            "row_id": row_id,
            "active": False,
            "items": None,
            "collapse_on_release": collapse_on_release,
        }
        if collapse_on_release:
            return "break"

    def _on_tree_drag_motion(self, event):
        """Обработка перемещения мыши с зажатой кнопкой для внутреннего Drag & Drop."""
        state = self._drag_state
        if not state:
            return

        dx = abs(event.x_root - state["start_x"])
        dy = abs(event.y_root - state["start_y"])
        if not state["active"]:
            if dx < 8 and dy < 8:
                return
            tree = self.left_tree if state["side"] == "left" else self.right_tree
            if state["row_id"] not in tree.selection():
                tree.selection_set(state["row_id"])
            items = self._collect_drag_items(state["side"])
            if not items:
                self._drag_state = None
                return
            state["active"] = True
            state["items"] = items

        target_tree = self.right_tree if state["side"] == "left" else self.left_tree
        over_target = self._is_point_in_widget(target_tree, event.x_root, event.y_root)

        count = len(state["items"])
        label = state["items"][0]["name"] if count == 1 else f"объектов: {count}"
        if state["side"] == "left":
            prefix = "➔ Загрузить на ЧПУ: " if over_target else "Перетащите на правую панель ЧПУ: "
        else:
            prefix = "⬅ Скачать на ПК: " if over_target else "Перетащите на левую панель ПК: "

        self._show_drag_tooltip(event.x_root, event.y_root, prefix + label)
        self.root.config(cursor="hand2" if over_target else "arrow")

    def _on_tree_drag_release(self, event):
        """Завершение внутреннего Drag & Drop при отпускании кнопки мыши."""
        state = self._drag_state
        self._drag_state = None
        self._hide_drag_tooltip()

        if not state:
            return

        if not state["active"]:
            if state.get("collapse_on_release"):
                tree = self.left_tree if state["side"] == "left" else self.right_tree
                tree.selection_set(state["row_id"])
                tree.focus(state["row_id"])
            return

        target_tree = self.right_tree if state["side"] == "left" else self.left_tree
        if not self._is_point_in_widget(target_tree, event.x_root, event.y_root):
            return

        items = state.get("items") or []
        if not items:
            return

        if state["side"] == "left":
            self._confirm_and_upload_items(items)
        else:
            self._confirm_and_download_items(items)

    def on_drop_to_remote(self, event):
        """Перетаскивание файлов или папок из Проводника Windows на панель ЧПУ."""
        if not event.data:
            return
        try:
            raw_paths = self.root.tk.splitlist(event.data)
        except Exception:
            raw_paths = [event.data]

        items = []
        for p in raw_paths:
            p = os.path.normpath(p)
            if p.lower().endswith(".lnk"):
                target, is_dir, is_file = resolve_shortcut(p)
                if target and (is_dir or is_file):
                    items.append({"path": target, "is_dir": is_dir, "name": os.path.basename(target)})
            elif os.path.isdir(p):
                items.append({"path": p, "is_dir": True, "name": os.path.basename(p)})
            elif os.path.isfile(p):
                items.append({"path": p, "is_dir": False, "name": os.path.basename(p)})

        if not items:
            return

        self._confirm_and_upload_items(items)

    def on_drop_to_local(self, event):
        """Перетаскивание папки из Проводника на левую панель — быстрый переход в неё."""
        if not event.data:
            return
        try:
            raw_paths = self.root.tk.splitlist(event.data)
        except Exception:
            raw_paths = [event.data]

        if not raw_paths:
            return

        first = os.path.normpath(raw_paths[0])
        if first.lower().endswith(".lnk"):
            target, is_dir, is_file = resolve_shortcut(first)
            if target:
                first = target if is_dir else os.path.dirname(target)
        elif os.path.isfile(first):
            first = os.path.dirname(first)

        if os.path.isdir(first):
            self._set_local_path(first)

    # ------------------ Мониторинг связи ------------------
    def start_connection_monitor(self):
        """Запуск фоновой проверки доступности порта 445 стойки ЧПУ."""
        def monitor_loop():
            is_alive = self.smb_client.is_alive()
            self.root.after(0, lambda alive=is_alive: self.update_connection_status(alive))
            interval = max(2, self.config.getint("App", "ping_interval", fallback=5))
            self.root.after(interval * 1000, self.start_connection_monitor)

        t = threading.Thread(target=monitor_loop, daemon=True)
        t.start()

    def update_connection_status(self, is_alive):
        """Обновление индикатора состояния связи со станком."""
        self.is_connected = is_alive
        if is_alive:
            self.status_canvas.itemconfig(self.status_dot, fill="#28a745")
            self.status_label.config(
                text=f"🟢 Станок на связи ({self.smb_client.host}:{self.smb_client.port})",
                foreground="#155724"
            )
        else:
            self.status_canvas.itemconfig(self.status_dot, fill="#dc3545")
            self.status_label.config(
                text=f"🔴 Нет связи со станком ({self.smb_client.host})",
                foreground="#721c24"
            )

    def _set_busy(self, busy, status_text=None):
        """Включение/выключение индикатора загрузки."""
        self._busy = busy
        if status_text is not None:
            self.info_label.config(text=status_text)
        if busy:
            self.progress_bar.pack(side=tk.RIGHT, padx=4)
            self.progress_bar.start(10)
        else:
            self.progress_bar.stop()
            self.progress_bar.pack_forget()

    # ------------------ Локальная панель ------------------
    def _set_local_path(self, new_path):
        """Переход в локальный каталог с обновлением поля пути и конфига."""
        self.current_local_path = os.path.normpath(new_path)
        self.loc_path_entry.delete(0, tk.END)
        self.loc_path_entry.insert(0, self.current_local_path)
        self.refresh_local()
        self.save_config()

    def navigate_local_up(self):
        parent_dir = os.path.dirname(self.current_local_path)
        if parent_dir and parent_dir != self.current_local_path:
            self._set_local_path(parent_dir)

    def browse_local_folder(self):
        folder = filedialog.askdirectory(initialdir=self.current_local_path, title="Выберите локальную папку УП")
        if folder:
            self._set_local_path(folder)

    def on_local_path_enter(self):
        new_path = self.loc_path_entry.get().strip()
        if os.path.isdir(new_path):
            self._set_local_path(new_path)
        else:
            messagebox.showwarning("Ошибка", f"Папка не существует или недоступна:\n{new_path}")

    def refresh_local(self):
        """Чтение локальной папки (включая ярлыки .lnk) и обновление таблицы."""
        for item in self.left_tree.get_children():
            self.left_tree.delete(item)
        self.local_items.clear()

        if not os.path.exists(self.current_local_path):
            return

        # Родительская папка
        parent_dir = os.path.dirname(self.current_local_path)
        if parent_dir and parent_dir != self.current_local_path:
            item_id = self.left_tree.insert("", tk.END, values=("📁 .. [На уровень вверх]", "", ""))
            self.local_items[item_id] = {"is_up": True, "is_dir": True, "path": parent_dir}

        dirs = []
        shortcuts_dir = []
        files = []

        filter_mode = self.loc_filter_var.get()
        exts = self.up_extensions if "УП" in filter_mode else None

        try:
            with os.scandir(self.current_local_path) as it:
                for entry in it:
                    if entry.is_dir():
                        try:
                            mtime_ts = entry.stat().st_mtime
                            mtime_dt = datetime.fromtimestamp(mtime_ts)
                        except OSError:
                            mtime_ts, mtime_dt = 0.0, None
                        dirs.append((entry.name, entry.path, mtime_dt, mtime_ts))
                    elif entry.is_file():
                        low_name = entry.name.lower()
                        if low_name.endswith(".lnk"):
                            target, is_dir, is_file = resolve_shortcut(entry.path)
                            display_lnk = os.path.splitext(entry.name)[0]
                            if target and is_dir:
                                try:
                                    mtime_ts = os.stat(target).st_mtime
                                    mtime_dt = datetime.fromtimestamp(mtime_ts)
                                except OSError:
                                    mtime_ts, mtime_dt = 0.0, None
                                shortcuts_dir.append((display_lnk, target, mtime_dt, mtime_ts))
                            elif target and is_file:
                                if exts is None or target.lower().endswith(exts):
                                    try:
                                        st = os.stat(target)
                                        files.append((
                                            f"🔗 {display_lnk}",
                                            target,
                                            st.st_size,
                                            datetime.fromtimestamp(st.st_mtime),
                                            st.st_mtime,
                                            True
                                        ))
                                    except OSError:
                                        pass
                            elif exts is None:
                                # Показываем неразрешённый ярлык только в режиме «Все файлы»
                                st = entry.stat()
                                files.append((
                                    f"🔗 {entry.name}",
                                    entry.path,
                                    st.st_size,
                                    datetime.fromtimestamp(st.st_mtime),
                                    st.st_mtime,
                                    True
                                ))
                        else:
                            if exts is None or low_name.endswith(exts):
                                st = entry.stat()
                                files.append((
                                    entry.name,
                                    entry.path,
                                    st.st_size,
                                    datetime.fromtimestamp(st.st_mtime),
                                    st.st_mtime,
                                    False
                                ))
        except Exception as e:
            winerr = getattr(e, "winerror", None)
            if winerr == 58:
                self.info_label.config(text="Сетевой диск ЧПУ заблокирован Windows (Win32 58) — используйте правую панель ЧПУ!")
                messagebox.showwarning(
                    "Сетевой диск заблокирован Windows",
                    f"Windows 11 блокирует просмотр этого диска (ошибка Win32 58):\n{self.current_local_path}\n\n"
                    f"Для работы с папками станка используйте правую панель «Стойка ЧПУ OSAI COMPACT»."
                )
            else:
                self.info_label.config(text=f"Ошибка чтения локальной папки: {e}")
            return

        for name, full_path, mtime_dt, mtime_ts in dirs:
            date_str = mtime_dt.strftime("%d.%m.%Y %H:%M") if mtime_dt else ""
            item_id = self.left_tree.insert("", tk.END, values=(f"📁 {name}", "", date_str))
            self.local_items[item_id] = {
                "is_dir": True,
                "is_up": False,
                "path": full_path,
                "name": name,
                "sort_name": name.lower(),
                "sort_size": 0,
                "sort_mtime": mtime_ts,
            }

        for name, target_path, mtime_dt, mtime_ts in shortcuts_dir:
            date_str = mtime_dt.strftime("%d.%m.%Y %H:%M") if mtime_dt else ""
            item_id = self.left_tree.insert("", tk.END, values=(f"📁🔗 {name}", "", date_str))
            self.local_items[item_id] = {
                "is_dir": True,
                "is_up": False,
                "path": target_path,
                "name": name,
                "is_shortcut": True,
                "sort_name": name.lower(),
                "sort_size": 0,
                "sort_mtime": mtime_ts,
            }

        for disp_name, full_path, f_size, f_date, f_ts, is_lnk in files:
            date_str = f_date.strftime("%d.%m.%Y %H:%M") if f_date else ""
            item_id = self.left_tree.insert("", tk.END, values=(disp_name, format_size(f_size), date_str))
            base_name = os.path.basename(full_path)
            self.local_items[item_id] = {
                "is_dir": False,
                "is_up": False,
                "path": full_path,
                "name": base_name,
                "is_shortcut": is_lnk,
                "sort_name": disp_name.lstrip("🔗 ").lower(),
                "sort_size": f_size or 0,
                "sort_mtime": f_ts or 0.0,
            }

        self._apply_tree_sort("left")

    def on_local_item_double_click(self, event):
        item_id = self.left_tree.focus()
        meta = self.local_items.get(item_id)
        if not meta:
            return
        if meta.get("is_up") or meta.get("is_dir"):
            target = meta["path"]
            if os.path.isdir(target):
                self._set_local_path(target)
            else:
                messagebox.showwarning("Папка недоступна", f"Не удалось открыть путь:\n{target}")

    # ------------------ Удалённая панель ЧПУ ------------------
    def on_share_change(self, event=None):
        self.current_share = self.share_var.get()
        self.current_remote_path = "/PROGRAMS/WORK" if self.current_share == "USER" else "/"
        self.current_remote_raw_path = self.current_remote_path
        self.rem_path_entry.delete(0, tk.END)
        self.rem_path_entry.insert(0, self.current_remote_path)
        self.refresh_remote()
        self.save_config()

    def on_remote_path_enter(self):
        new_path = self.rem_path_entry.get().strip().replace("\\", "/")
        if not new_path.startswith("/"):
            new_path = "/" + new_path
        if len(new_path) > 1:
            new_path = new_path.rstrip("/")
        self.current_remote_path = new_path
        self.current_remote_raw_path = new_path
        self.refresh_remote()
        self.save_config()

    def navigate_remote_up(self):
        cur = self.current_remote_path.rstrip("/").replace("\\", "/") or "/"
        parent = posixpath.dirname(cur) or "/"
        cur_raw = self.current_remote_raw_path.rstrip("/").replace("\\", "/") or "/"
        parent_raw = posixpath.dirname(cur_raw) or "/"
        self.current_remote_path = parent
        self.current_remote_raw_path = parent_raw
        self.rem_path_entry.delete(0, tk.END)
        self.rem_path_entry.insert(0, self.current_remote_path)
        self.refresh_remote()
        self.save_config()

    def refresh_remote(self):
        """Асинхронное получение списка файлов со стойки ЧПУ."""
        for item in self.right_tree.get_children():
            self.right_tree.delete(item)
        self.remote_items.clear()

        self._set_busy(True, "Загрузка списка файлов со стойки ЧПУ...")
        share = self.current_share
        raw_path = self.current_remote_raw_path

        def worker():
            try:
                resolved_raw = self.smb_client.resolve_remote_path(share, raw_path, must_exist=True)
                dirs, files = self.smb_client.list_directory(share, resolved_raw)
                self.root.after(0, lambda d=dirs, f=files, r=resolved_raw: self._on_remote_loaded(d, f, r))
            except Exception as e:
                err = str(e)
                self.root.after(0, lambda msg=err: self._on_remote_error(msg))

        threading.Thread(target=worker, daemon=True).start()

    def _on_remote_loaded(self, dirs, files, resolved_raw=None):
        self._set_busy(False)
        if resolved_raw:
            self.current_remote_raw_path = resolved_raw

        if self.current_remote_path not in ("", "/"):
            item_id = self.right_tree.insert("", tk.END, values=("📁 .. [На уровень вверх]", "", ""))
            self.remote_items[item_id] = {"is_up": True, "is_dir": True}

        for d in dirs:
            mtime_dt = d.get("mtime")
            date_str = mtime_dt.strftime("%d.%m.%Y %H:%M") if mtime_dt else ""
            mtime_ts = mtime_dt.timestamp() if mtime_dt else 0.0
            item_id = self.right_tree.insert("", tk.END, values=(f"📁 {d['name']}", "", date_str))
            self.remote_items[item_id] = {
                "is_up": False,
                "is_dir": True,
                "name": d["name"],
                "raw_name": d.get("raw_name", d["name"]),
                "sort_name": d["name"].lower(),
                "sort_size": 0,
                "sort_mtime": mtime_ts,
            }

        for f in files:
            mtime_dt = f.get("mtime")
            date_str = mtime_dt.strftime("%d.%m.%Y %H:%M") if mtime_dt else ""
            mtime_ts = mtime_dt.timestamp() if mtime_dt else 0.0
            item_id = self.right_tree.insert("", tk.END, values=(f["name"], format_size(f["size"]), date_str))
            self.remote_items[item_id] = {
                "is_up": False,
                "is_dir": False,
                "name": f["name"],
                "raw_name": f.get("raw_name", f["name"]),
                "size": f["size"],
                "sort_name": f["name"].lower(),
                "sort_size": f["size"] or 0,
                "sort_mtime": mtime_ts,
            }

        self._apply_tree_sort("right")
        self.info_label.config(
            text=f"ЧПУ [{self.current_share}:{self.current_remote_path}]: {len(dirs)} папок, {len(files)} файлов"
        )

    def _on_remote_error(self, err_msg):
        self._set_busy(False, f"Ошибка соединения с ЧПУ: {err_msg}")
        self.update_connection_status(False)

    def on_remote_item_double_click(self, event):
        item_id = self.right_tree.focus()
        meta = self.remote_items.get(item_id)
        if not meta:
            return

        if meta.get("is_up"):
            self.navigate_remote_up()
        elif meta.get("is_dir"):
            cur = self.current_remote_path.rstrip("/")
            cur_raw = self.current_remote_raw_path.rstrip("/")
            self.current_remote_path = f"{cur}/{meta['name']}"
            self.current_remote_raw_path = f"{cur_raw}/{meta.get('raw_name', meta['name'])}"
            self.rem_path_entry.delete(0, tk.END)
            self.rem_path_entry.insert(0, self.current_remote_path)
            self.refresh_remote()
            self.save_config()

    def refresh_all(self):
        """Обновление обеих панелей."""
        self.refresh_local()
        self.refresh_remote()

    # ------------------ Передача файлов и папок ------------------
    def upload_selected(self):
        """Загрузка выбранных локальных файлов и/или папок в текущую папку ЧПУ."""
        selected_items = self.left_tree.selection()
        if not selected_items:
            messagebox.showinfo("Загрузка на станок", "Выберите файлы или папки в левой панели для загрузки.")
            return

        items_to_upload = []
        for item_id in selected_items:
            meta = self.local_items.get(item_id)
            if not meta or meta.get("is_up"):
                continue
            path = meta["path"]
            if meta.get("is_dir") and os.path.isdir(path):
                items_to_upload.append({"path": path, "is_dir": True, "name": os.path.basename(path)})
            elif os.path.isfile(path):
                items_to_upload.append({"path": path, "is_dir": False, "name": os.path.basename(path)})

        if not items_to_upload:
            messagebox.showinfo("Загрузка на станок", "Среди выбранных элементов нет файлов или папок для передачи.")
            return

        self._confirm_and_upload_items(items_to_upload)

    def _confirm_and_upload_items(self, items_to_upload):
        """Подтверждение и запуск асинхронной загрузки файлов/папок на ЧПУ."""
        if self._busy:
            messagebox.showwarning("Подождите", "Дождитесь завершения текущей операции.")
            return

        dirs_cnt = sum(1 for x in items_to_upload if x["is_dir"])
        files_cnt = len(items_to_upload) - dirs_cnt
        summary_parts = []
        if files_cnt:
            summary_parts.append(f"файлов: {files_cnt}")
        if dirs_cnt:
            summary_parts.append(f"папок: {dirs_cnt}")

        confirm_text = (
            f"Загрузить на стойку ЧПУ ({', '.join(summary_parts)}) в каталог:\n"
            f"[{self.current_share}] {self.current_remote_path}?"
        )
        if not messagebox.askyesno("Подтверждение загрузки", confirm_text):
            return

        self._set_busy(True, "Загрузка на стойку ЧПУ...")
        share = self.current_share
        remote_dir = self.current_remote_raw_path

        def worker():
            success_count = 0
            err_msg = None
            for item in items_to_upload:
                try:
                    if item["is_dir"]:
                        self.smb_client.upload_directory(
                            item["path"],
                            share,
                            remote_dir,
                            status_callback=lambda s: self.root.after(0, lambda txt=s: self.info_label.config(text=txt))
                        )
                    else:
                        self.root.after(0, lambda n=item["name"]: self.info_label.config(text=f"Загрузка файла: {n}"))
                        self.smb_client.upload_file(item["path"], share, remote_dir)
                    success_count += 1
                except Exception as e:
                    err_msg = str(e)
                    break

            self.root.after(0, lambda ok=success_count, tot=len(items_to_upload), err=err_msg: self._on_upload_complete(ok, tot, err))

        threading.Thread(target=worker, daemon=True).start()

    def _on_upload_complete(self, success_count, total_count, err_msg):
        self._set_busy(False)
        if err_msg:
            self.info_label.config(text=f"Сбой загрузки: {err_msg}")
            messagebox.showerror("Ошибка загрузки", f"Передано {success_count} из {total_count} объектов.\nОшибка: {err_msg}")
            if success_count > 0:
                self.refresh_remote()
        else:
            self.refresh_remote()
            self.info_label.config(text=f"Успешно загружено объектов на ЧПУ: {success_count}")

    def download_selected(self):
        """Скачивание выбранных файлов и/или папок с ЧПУ в локальную папку."""
        if self._busy:
            messagebox.showwarning("Подождите", "Дождитесь завершения текущей операции.")
            return

        selected_items = self.right_tree.selection()
        if not selected_items:
            messagebox.showinfo("Скачивание со станка", "Выберите файл(ы) или папку в правой панели для скачивания.")
            return

        items_to_download = []
        for item_id in selected_items:
            meta = self.remote_items.get(item_id)
            if not meta or meta.get("is_up"):
                continue
            raw_name = meta.get("raw_name", meta["name"])
            rem_path = f"{self.current_remote_raw_path.rstrip('/')}/{raw_name}"
            items_to_download.append({
                "name": meta["name"],
                "raw_name": raw_name,
                "is_dir": meta["is_dir"],
                "remote_path": rem_path
            })

        if not items_to_download:
            messagebox.showinfo("Скачивание со станка", "Среди выбранных элементов нет объектов для скачивания.")
            return

        self._start_download_items(items_to_download)

    def _confirm_and_download_items(self, items_to_download):
        """Подтверждение (для Drag & Drop со станка) и запуск скачивания."""
        if self._busy:
            messagebox.showwarning("Подождите", "Дождитесь завершения текущей операции.")
            return

        dirs_cnt = sum(1 for x in items_to_download if x["is_dir"])
        files_cnt = len(items_to_download) - dirs_cnt
        summary_parts = []
        if files_cnt:
            summary_parts.append(f"файлов: {files_cnt}")
        if dirs_cnt:
            summary_parts.append(f"папок: {dirs_cnt}")

        confirm_text = (
            f"Скачать со стойки ЧПУ ({', '.join(summary_parts)}) в папку:\n"
            f"{self.current_local_path}?"
        )
        if not messagebox.askyesno("Подтверждение скачивания", confirm_text):
            return
        self._start_download_items(items_to_download)

    def _start_download_items(self, items_to_download):
        """Асинхронное скачивание списка объектов со стойки ЧПУ."""
        self._set_busy(True, "Скачивание со стойки ЧПУ...")
        share = self.current_share
        local_dir = self.current_local_path

        def worker():
            success_count = 0
            err_msg = None
            for item in items_to_download:
                try:
                    if item["is_dir"]:
                        self.smb_client.download_directory(
                            share,
                            item["remote_path"],
                            local_dir,
                            status_callback=lambda s: self.root.after(0, lambda txt=s: self.info_label.config(text=txt))
                        )
                    else:
                        self.root.after(0, lambda n=item["name"]: self.info_label.config(text=f"Скачивание: {n}"))
                        loc_dest = os.path.join(local_dir, item["name"])
                        self.smb_client.download_file(share, item["remote_path"], loc_dest)
                    success_count += 1
                except Exception as e:
                    err_msg = str(e)
                    break

            self.root.after(0, lambda ok=success_count, tot=len(items_to_download), err=err_msg: self._on_download_complete(ok, tot, err))

        threading.Thread(target=worker, daemon=True).start()

    def _on_download_complete(self, success_count, total_count, err_msg):
        self._set_busy(False)
        if err_msg:
            self.info_label.config(text=f"Сбой скачивания: {err_msg}")
            messagebox.showerror("Ошибка скачивания", f"Скачано {success_count} из {total_count} объектов.\nОшибка: {err_msg}")
            if success_count > 0:
                self.refresh_local()
        else:
            self.refresh_local()
            self.info_label.config(text=f"Успешно скачано объектов на ПК: {success_count}")

    # ------------------ Операции с папками и удалением на ЧПУ ------------------
    def create_remote_folder(self):
        """Создание новой папки в текущем каталоге ЧПУ."""
        if self._busy:
            messagebox.showwarning("Подождите", "Дождитесь завершения текущей операции.")
            return

        folder_name = simpledialog.askstring(
            "Новая папка на ЧПУ",
            f"Создать папку в [{self.current_share}] {self.current_remote_path}:",
            parent=self.root
        )
        if not folder_name or not folder_name.strip():
            return

        folder_name = folder_name.strip()
        invalid_chars = '<>:"/\\|?*'
        if any(ch in folder_name for ch in invalid_chars):
            messagebox.showerror("Ошибка имени", f"Имя папки содержит недопустимые символы:\n{invalid_chars}")
            return

        target_path = f"{self.current_remote_raw_path.rstrip('/')}/{folder_name}"
        share = self.current_share
        self._set_busy(True, f"Создание папки '{folder_name}'...")

        def worker():
            err_msg = None
            try:
                self.smb_client.create_directory(share, target_path)
            except Exception as e:
                err_msg = str(e)
            self.root.after(0, lambda name=folder_name, err=err_msg: self._on_mkdir_complete(name, err))

        threading.Thread(target=worker, daemon=True).start()

    def _on_mkdir_complete(self, folder_name, err_msg):
        self._set_busy(False)
        if err_msg:
            messagebox.showerror("Ошибка создания папки", f"Не удалось создать папку '{folder_name}':\n{err_msg}")
            self.info_label.config(text=f"Ошибка создания папки: {err_msg}")
        else:
            self.info_label.config(text=f"Создана папка: {folder_name}")
            self.refresh_remote()

    def delete_remote_selected(self):
        """Удаление выбранных файлов и/или папок со стойки ЧПУ."""
        if self._busy:
            messagebox.showwarning("Подождите", "Дождитесь завершения текущей операции.")
            return

        selected_items = self.right_tree.selection()
        if not selected_items:
            messagebox.showinfo("Удаление", "Выберите файл или папку на ЧПУ для удаления.")
            return

        items_to_delete = []
        for item_id in selected_items:
            meta = self.remote_items.get(item_id)
            if not meta or meta.get("is_up"):
                continue
            disp_path = f"{self.current_remote_path.rstrip('/')}/{meta['name']}"
            norm_check = disp_path.replace("\\", "/").rstrip("/").lower()
            if norm_check in PROTECTED_REMOTE_PATHS:
                messagebox.showerror("Защита ЧПУ", f"Запрещено удалять системный каталог станка:\n{disp_path}")
                return
            raw_name = meta.get("raw_name", meta["name"])
            rem_path = f"{self.current_remote_raw_path.rstrip('/')}/{raw_name}"
            items_to_delete.append({
                "name": meta["name"],
                "is_dir": meta["is_dir"],
                "remote_path": rem_path
            })

        if not items_to_delete:
            return

        if len(items_to_delete) == 1:
            one = items_to_delete[0]
            if one["is_dir"]:
                msg = (
                    f"Удалить папку со стойки ЧПУ вместе со ВСЕМ содержимым?\n\n"
                    f"📁 {one['name']}\n\nОперация необратима!"
                )
            else:
                msg = f"Вы действительно хотите удалить файл со станка:\n\n📄 {one['name']}?"
        else:
            dirs_cnt = sum(1 for x in items_to_delete if x["is_dir"])
            files_cnt = len(items_to_delete) - dirs_cnt
            msg = (
                f"Удалить выбранные объекты со стойки ЧПУ?\n"
                f"(папок: {dirs_cnt}, файлов: {files_cnt})\n\n"
                f"Удаление папок вместе со всем содержимым необратимо!"
            )

        if not messagebox.askyesno("Подтверждение удаления", msg, icon="warning"):
            return

        self._set_busy(True, "Удаление объектов со стойки ЧПУ...")
        share = self.current_share

        def worker():
            deleted_count = 0
            err_msg = None
            for item in items_to_delete:
                try:
                    if item["is_dir"]:
                        self.smb_client.delete_directory_recursive(share, item["remote_path"])
                    else:
                        self.smb_client.delete_file(share, item["remote_path"])
                    deleted_count += 1
                except Exception as e:
                    err_msg = str(e)
                    break
            self.root.after(0, lambda ok=deleted_count, tot=len(items_to_delete), err=err_msg: self._on_delete_complete(ok, tot, err))

        threading.Thread(target=worker, daemon=True).start()

    def _on_delete_complete(self, deleted_count, total_count, err_msg):
        self._set_busy(False)
        if err_msg:
            self.info_label.config(text=f"Ошибка удаления: {err_msg}")
            messagebox.showerror("Ошибка удаления", f"Удалено {deleted_count} из {total_count} объектов.\nОшибка: {err_msg}")
            if deleted_count > 0:
                self.refresh_remote()
        else:
            self.refresh_remote()
            self.info_label.config(text=f"Удалено объектов со станка: {deleted_count}")

    # ------------------ Настройки ------------------
    def open_settings_dialog(self):
        """Диалог настроек подключения и фильтра расширений."""
        dlg = tk.Toplevel(self.root)
        dlg.title("Настройки подключения к ЧПУ")
        dlg.geometry("430x330")
        dlg.resizable(False, False)
        dlg.transient(self.root)
        dlg.grab_set()

        frame = ttk.Frame(dlg, padding=16)
        frame.pack(fill=tk.BOTH, expand=True)

        ttk.Label(frame, text="IP-адрес стойки ЧПУ:").grid(row=0, column=0, sticky=tk.W, pady=6)
        ip_entry = ttk.Entry(frame)
        ip_entry.grid(row=0, column=1, sticky=tk.EW, pady=6)
        ip_entry.insert(0, self.config.get("CNC", "host", fallback="192.168.0.2"))

        ttk.Label(frame, text="Порт TCP SMB:").grid(row=1, column=0, sticky=tk.W, pady=6)
        port_entry = ttk.Entry(frame)
        port_entry.grid(row=1, column=1, sticky=tk.EW, pady=6)
        port_entry.insert(0, self.config.get("CNC", "port", fallback="445"))

        ttk.Label(frame, text="Имя ресурса по умолч.:").grid(row=2, column=0, sticky=tk.W, pady=6)
        share_entry = ttk.Entry(frame)
        share_entry.grid(row=2, column=1, sticky=tk.EW, pady=6)
        share_entry.insert(0, self.config.get("CNC", "share", fallback="USER"))

        ttk.Label(frame, text="Пароль SMB (WinCE):").grid(row=3, column=0, sticky=tk.W, pady=6)
        pass_entry = ttk.Entry(frame, show="*")
        pass_entry.grid(row=3, column=1, sticky=tk.EW, pady=6)
        pass_entry.insert(0, self.config.get("CNC", "password", fallback="123"))

        ttk.Label(frame, text="Фильтр расширений УП:").grid(row=4, column=0, sticky=tk.W, pady=6)
        ext_entry = ttk.Entry(frame)
        ext_entry.grid(row=4, column=1, sticky=tk.EW, pady=6)
        ext_entry.insert(0, self.config.get("Local", "extensions", fallback=DEFAULT_EXTENSIONS))

        frame.columnconfigure(1, weight=1)

        def save_and_close():
            host = ip_entry.get().strip()
            try:
                port = int(port_entry.get().strip())
            except ValueError:
                messagebox.showerror("Ошибка", "Порт должен быть целым числом.", parent=dlg)
                return

            share_val = share_entry.get().strip().upper() or "USER"
            password = pass_entry.get().strip()
            ext_val = ext_entry.get().strip() or DEFAULT_EXTENSIONS

            self.config.set("CNC", "host", host)
            self.config.set("CNC", "port", str(port))
            self.config.set("CNC", "share", share_val)
            self.config.set("CNC", "password", password)
            self.config.set("Local", "extensions", ext_val)

            if share_val != self.current_share:
                self.current_share = share_val
                self.share_var.set(share_val)
                self.current_remote_path = "/PROGRAMS/WORK" if share_val == "USER" else "/"
                self.current_remote_raw_path = self.current_remote_path
                self.rem_path_entry.delete(0, tk.END)
                self.rem_path_entry.insert(0, self.current_remote_path)
            self.up_extensions = parse_extensions(ext_val)

            ext_label_new = " ".join(f"*{e}" for e in self.up_extensions)
            new_up_label = f"УП ({ext_label_new})"
            was_up = (self.loc_filter_var.get() == self.up_filter_label)
            self.up_filter_label = new_up_label
            self.loc_filter_cb.configure(values=[self.up_filter_label, "Все файлы (*.*)"])
            if was_up:
                self.loc_filter_var.set(self.up_filter_label)

            self.save_config()

            self.smb_client.host = host
            self.smb_client.port = port
            self.smb_client.password = password
            self.smb_client.disconnect()

            dlg.destroy()
            self.refresh_all()

        btn_box = ttk.Frame(frame)
        btn_box.grid(row=5, column=0, columnspan=2, pady=16)
        ttk.Button(btn_box, text="Сохранить", command=save_and_close, style="Action.TButton").pack(side=tk.LEFT, padx=6)
        ttk.Button(btn_box, text="Отмена", command=dlg.destroy).pack(side=tk.LEFT, padx=6)


def main():
    root = TkinterDnD.Tk() if HAS_DND else tk.Tk()
    OsaiTransferApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
