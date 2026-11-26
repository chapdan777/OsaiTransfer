# -*- coding: utf-8 -*-
"""
Модуль сетевого взаимодействия со стойкой ЧПУ OSAI по протоколу SMB 1.0.
Работает строго в пространстве пользователя (Userspace TCP сокеты) поверх pysmb.
Исключает ошибку Win32 58 и решает особенности контроллера Windows CE:
- Удаление файлов и папок с обязательным ведущим слэшем (\)
- Поддержка как имён UTF-16LE (Unicode), так и устаревших имён CP1251-в-Latin-1
- Автоматический сброс соединения при сбойных пакетах ошибок WinCE (ProtocolError)
- Рекурсивная загрузка, скачивание и удаление папок
"""

import os
import posixpath
import time
import socket
import threading
from datetime import datetime
from smb.SMBConnection import SMBConnection
from smb.base import NotConnectedError, SMBTimeout, _PendingRequest
from smb.smb_structs import (
    SMBMessage,
    ComDeleteRequest,
    ComCreateDirectoryRequest,
    ComDeleteDirectoryRequest,
    ProtocolError,
    OperationFailure,
)


PROTECTED_WINCE_PATHS = {"", "\\", "\\user", "\\oem", "\\programs", "\\macros"}

NTSTATUS_MESSAGES = {
    0xC000003A: "Путь не найден на ЧПУ (STATUS_OBJECT_PATH_NOT_FOUND)",
    0xC0000034: "Файл или папка не найдены на ЧПУ (STATUS_OBJECT_NAME_NOT_FOUND)",
    0xC0000022: "Отказано в доступе на ЧПУ (STATUS_ACCESS_DENIED)",
    0xC0000043: "Файл занят другим процессом на ЧПУ (STATUS_SHARING_VIOLATION)",
    0xC000007F: "Недостаточно места на диске ЧПУ (STATUS_DISK_FULL)",
    0xC0000035: "Объект с таким именем уже существует (STATUS_OBJECT_NAME_COLLISION)",
    0xC0000101: "Папка не пуста (STATUS_DIRECTORY_NOT_EMPTY)",
}


def _has_cyrillic(text: str) -> bool:
    """Проверка наличия символов кириллицы (или знака №) в строке."""
    return any(("\u0400" <= ch <= "\u04FF") or ch == "№" for ch in text)


def decode_smb_name(name: str) -> str:
    """
    Восстанавливает читаемое имя файла/папки с WinCE.
    Если контроллер отдал байты CP1251, декодированные pysmb как Latin-1,
    преобразует их обратно в кириллицу. Настоящие Unicode-имена не меняет.
    """
    if not name:
        return ""
    clean = name.rstrip("\x00")
    try:
        raw_bytes = clean.encode("iso-8859-1")
        decoded = raw_bytes.decode("cp1251")
        if _has_cyrillic(decoded):
            return decoded
    except (UnicodeEncodeError, UnicodeDecodeError):
        pass
    return clean


def encode_smb_name(name: str) -> str:
    """
    Обратное преобразование кириллицы в CP1251-в-Latin-1 (для совместимости).
    Для новых файлов/папок на WinCE используется прямой UTF-16LE без этого вызова.
    """
    if not name or not _has_cyrillic(name):
        return name or ""
    try:
        raw_bytes = name.encode("cp1251")
        return raw_bytes.decode("iso-8859-1")
    except (UnicodeEncodeError, UnicodeDecodeError):
        return name


def _normalize_posix_path(path: str) -> str:
    """Приводит путь к формату '/a/b' без завершающего слэша (кроме '/')."""
    norm = (path or "/").replace("\\", "/").strip()
    if not norm.startswith("/"):
        norm = "/" + norm
    if len(norm) > 1:
        norm = norm.rstrip("/")
    return norm or "/"


def _format_protocol_error(err: ProtocolError) -> str:
    """Извлекает код NTSTATUS из усечённого пакета ответа WinCE и формирует понятное сообщение."""
    if len(err.args) >= 2 and isinstance(err.args[1], (bytes, bytearray)):
        raw = err.args[1]
        if len(raw) >= 9 and raw[:4] == b"\xffSMB":
            nt_status = int.from_bytes(raw[5:9], "little")
            known = NTSTATUS_MESSAGES.get(nt_status)
            if known:
                return known
            return f"Ошибка протокола SMB ЧПУ (код {hex(nt_status)})"
    return f"Ошибка протокола SMB ЧПУ: {err}"


class OsaiSMBClient:
    """Пользовательский клиент SMB1 для контроллера OSAI COMPACT (Windows CE)."""

    def __init__(self, host="192.168.0.2", port=445, remote_name="COMPACT",
                 client_name="OPERATOR-PC", username="", password="123", timeout=5):
        self.host = host
        self.port = int(port)
        self.remote_name = remote_name
        self.client_name = client_name
        self.username = username
        self.password = password
        self.timeout = int(timeout)
        self._conn = None
        self._lock = threading.RLock()

    def _create_connection(self):
        """Создание нового экземпляра SMB-соединения."""
        return SMBConnection(
            username=self.username,
            password=self.password,
            my_name=self.client_name,
            remote_name=self.remote_name,
            is_direct_tcp=(self.port == 445),
        )

    def connect(self):
        """Установка соединения с контроллером."""
        with self._lock:
            try:
                self.disconnect()
                conn = self._create_connection()
                success = conn.connect(self.host, self.port, timeout=self.timeout)
                if success:
                    self._conn = conn
                    return True
                return False
            except Exception:
                self._conn = None
                return False

    def disconnect(self):
        """Корректное закрытие соединения."""
        with self._lock:
            if self._conn:
                try:
                    self._conn.close()
                except Exception:
                    pass
                self._conn = None

    def is_alive(self):
        """Быстрая проверка доступности порта 445 (таймаут 1.5 сек)."""
        try:
            sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            sock.settimeout(1.5)
            res = sock.connect_ex((self.host, self.port))
            sock.close()
            return res == 0
        except Exception:
            return False

    def _ensure_connection(self):
        """Проверка соединения и автоматическое переподключение."""
        if self._conn is None:
            if not self.connect():
                raise ConnectionError(f"Нет связи с ЧПУ {self.host}:{self.port}")

    def _ensure_tree(self, share):
        """Подключение дерева сетевого ресурса (USER, OEM)."""
        self._ensure_connection()
        if share not in self._conn.connected_trees:
            self._conn.listPath(share, "/", timeout=self.timeout)
        return self._conn.connected_trees[share]

    def _run_with_reconnect(self, fn):
        """
        Выполняет SMB-операцию с однократным переподключением при обрыве сокета
        и обязательным сбросом соединения при ProtocolError (чтобы не отравлять буфер pysmb).
        """
        self._ensure_connection()
        try:
            return fn()
        except (NotConnectedError, SMBTimeout, socket.error, EOFError, OSError):
            if not self.connect():
                raise ConnectionError(f"Нет связи с ЧПУ {self.host}:{self.port}")
            try:
                return fn()
            except ProtocolError as pe:
                self.disconnect()
                raise RuntimeError(_format_protocol_error(pe)) from pe
            except OperationFailure as oe:
                raise RuntimeError(str(oe)) from oe
        except ProtocolError as pe:
            self.disconnect()
            raise RuntimeError(_format_protocol_error(pe)) from pe
        except OperationFailure as oe:
            raise RuntimeError(str(oe)) from oe

    def _list_raw(self, share, raw_path):
        """Низкоуровневое перечисление каталога без перекодирования пути."""
        norm_raw = _normalize_posix_path(raw_path)
        raw_entries = self._run_with_reconnect(
            lambda: self._conn.listPath(share, norm_raw, timeout=self.timeout)
        )

        dirs = []
        files = []
        for entry in raw_entries:
            clean_raw = entry.filename.rstrip("\x00")
            if clean_raw.strip() in (".", "..", ""):
                continue

            display_name = decode_smb_name(clean_raw)
            entry_info = {
                "name": display_name,
                "raw_name": clean_raw,
                "is_dir": entry.isDirectory,
                "size": entry.file_size if not entry.isDirectory else 0,
                "mtime": datetime.fromtimestamp(entry.last_write_time) if entry.last_write_time else None,
            }
            if entry.isDirectory:
                dirs.append(entry_info)
            else:
                files.append(entry_info)

        dirs.sort(key=lambda x: x["name"].lower())
        files.sort(key=lambda x: x["name"].lower())
        return dirs, files

    def _find_entry_raw_name(self, entries, segment):
        """Поиск точного raw_name сегмента среди записей каталога (Unicode или CP1251-в-Latin-1)."""
        for item in entries:
            if item["raw_name"] == segment:
                return item["raw_name"]
        for item in entries:
            if item["name"] == segment:
                return item["raw_name"]
        seg_low = segment.lower()
        for item in entries:
            if item["name"].lower() == seg_low or item["raw_name"].lower() == seg_low:
                return item["raw_name"]
        return None

    def resolve_remote_path(self, share, path, must_exist=True):
        """
        Преобразует путь (отображаемый Unicode или смешанный) в точный raw_path контроллера,
        корректно обрабатывая как Unicode-папки ('В РАБОТУ'), так и CP1251-папки ('338 фанера 20').
        """
        norm = _normalize_posix_path(path)
        if norm == "/" or norm.isascii():
            return norm

        if not must_exist:
            parent_dir, leaf = posixpath.split(norm)
            raw_parent = self.resolve_remote_path(share, parent_dir, must_exist=True)
            if not leaf or leaf.isascii():
                return f"{raw_parent.rstrip('/')}/{leaf}"
            dirs, files = self._list_raw(share, raw_parent)
            matched = self._find_entry_raw_name(dirs + files, leaf)
            return f"{raw_parent.rstrip('/')}/{matched or leaf}"

        parts = [p for p in norm.split("/") if p]
        cur_raw = "/"
        for idx, seg in enumerate(parts):
            if seg.isascii():
                cur_raw = f"{cur_raw.rstrip('/')}/{seg}"
                continue
            dirs, files = self._list_raw(share, cur_raw)
            pool = dirs if idx < len(parts) - 1 else (dirs + files)
            matched = self._find_entry_raw_name(pool, seg)
            if not matched:
                # На случай, если искомый элемент (папка/файл) находится в другом списке
                matched = self._find_entry_raw_name(dirs + files, seg)
            if not matched:
                raise FileNotFoundError(f"Объект не найден на ЧПУ: {norm}")
            cur_raw = f"{cur_raw.rstrip('/')}/{matched}"
        return cur_raw

    def list_directory(self, share, path="/"):
        """
        Получение списка файлов и папок в указанном каталоге ЧПУ.
        Автоматически определяет корректное кодирование каждого сегмента пути.
        """
        with self._lock:
            norm = _normalize_posix_path(path)
            dirs, files = self._list_raw(share, norm)

            # Если каталог с не-ASCII символами вернулся пустым, проверяем,
            # не хранится ли его имя на WinCE в виде CP1251-байтов.
            if not dirs and not files and not norm.isascii() and norm != "/":
                resolved = self.resolve_remote_path(share, norm, must_exist=True)
                if resolved != norm:
                    dirs, files = self._list_raw(share, resolved)

            return dirs, files

    def _send_raw_smb_request(self, share, payload_builder, op_desc, allow_collision=False):
        """Общий метод отправки прямых SMB1-запросов (создание/удаление) на Windows CE."""
        def attempt():
            tid = self._ensure_tree(share)
            result = {}

            def cb(msg, **kw):
                result["ok"] = True
                result["status"] = msg.status

            def eb(failure):
                result["err"] = failure

            m = SMBMessage(payload_builder())
            m.tid = tid
            self._conn._sendSMBMessage(m)
            self._conn.pending_requests[m.mid] = _PendingRequest(
                m.mid, int(time.time()) + 10, cb, eb
            )

            start_t = time.time()
            while not result:
                if time.time() - start_t > 10:
                    raise TimeoutError(f"Таймаут операции ({op_desc})")
                self._conn._pollForNetBIOSPacket(1)

            if result.get("err"):
                raise RuntimeError(f"Сетевая ошибка ({op_desc}): {result['err']}")

            status = result.get("status")
            if status and status.hasError:
                code = status.internal_value
                if allow_collision and code == 0xC0000035:
                    return True
                known = NTSTATUS_MESSAGES.get(code)
                if known:
                    raise RuntimeError(f"Ошибка ({op_desc}): {known}")
                raise RuntimeError(f"Ошибка ({op_desc}) — код {hex(code)}")
            return True

        return self._run_with_reconnect(attempt)

    def create_directory(self, share, remote_path):
        """
        Создание папки на контроллере ЧПУ.
        Windows CE обязательно требует ведущий слэш (\\).
        """
        with self._lock:
            resolved = self.resolve_remote_path(share, remote_path, must_exist=False)
            win_path = resolved.replace("/", "\\")
            if not win_path.startswith("\\"):
                win_path = "\\" + win_path
            return self._send_raw_smb_request(
                share,
                lambda: ComCreateDirectoryRequest(path=win_path),
                f"создание папки {win_path}",
                allow_collision=True,
            )

    def delete_file(self, share, remote_path):
        """
        Удаление файла на контроллере ЧПУ.
        Windows CE требует обязательный ведущий слэш (\\) и поисковые атрибуты 0x16.
        """
        with self._lock:
            resolved = self.resolve_remote_path(share, remote_path, must_exist=True)
            win_path = resolved.replace("/", "\\")
            if not win_path.startswith("\\"):
                win_path = "\\" + win_path
            if win_path.lower() in PROTECTED_WINCE_PATHS:
                raise ValueError("Запрещено удалять базовые системные каталоги станка.")
            return self._send_raw_smb_request(
                share,
                lambda: ComDeleteRequest(filename_pattern=win_path, search_attributes=0x16),
                f"удаление файла {win_path}",
            )

    def delete_directory(self, share, remote_path):
        """Удаление пустой папки на контроллере ЧПУ."""
        with self._lock:
            resolved = self.resolve_remote_path(share, remote_path, must_exist=True)
            win_path = resolved.replace("/", "\\")
            if not win_path.startswith("\\"):
                win_path = "\\" + win_path
            if win_path.lower() in PROTECTED_WINCE_PATHS:
                raise ValueError("Запрещено удалять базовые системные каталоги станка.")
            return self._send_raw_smb_request(
                share,
                lambda: ComDeleteDirectoryRequest(path=win_path),
                f"удаление папки {win_path}",
            )

    def delete_directory_recursive(self, share, remote_path):
        """Рекурсивное удаление папки со всем содержимым на ЧПУ."""
        with self._lock:
            raw_root = self.resolve_remote_path(share, remote_path, must_exist=True)

            def _walk_delete(cur_raw):
                dirs, files = self._list_raw(share, cur_raw)
                for f in files:
                    self._send_raw_smb_request(
                        share,
                        lambda p=f"{cur_raw.rstrip('/')}/{f['raw_name']}".replace("/", "\\"): ComDeleteRequest(
                            filename_pattern=p, search_attributes=0x16
                        ),
                        f"удаление файла {f['name']}",
                    )
                for d in dirs:
                    _walk_delete(f"{cur_raw.rstrip('/')}/{d['raw_name']}")
                win_dir = cur_raw.replace("/", "\\")
                if not win_dir.startswith("\\"):
                    win_dir = "\\" + win_dir
                if win_dir.lower() in PROTECTED_WINCE_PATHS:
                    raise ValueError("Запрещено удалять базовые системные каталоги станка.")
                self._send_raw_smb_request(
                    share,
                    lambda: ComDeleteDirectoryRequest(path=win_dir),
                    f"удаление папки {cur_raw}",
                )

            _walk_delete(raw_root)
            return True

    def upload_file(self, local_file_path, share, remote_dir, callback=None):
        """Загрузка одиночного файла на ЧПУ."""
        with self._lock:
            file_name = os.path.basename(local_file_path)
            raw_dir = self.resolve_remote_path(share, remote_dir, must_exist=True)
            target_path = f"{raw_dir.rstrip('/')}/{file_name}"

            # Если файл с таким именем уже существует в CP1251-кодировке, перезаписываем его точный raw_name
            if not file_name.isascii():
                _, existing_files = self._list_raw(share, raw_dir)
                matched = self._find_entry_raw_name(existing_files, file_name)
                if matched:
                    target_path = f"{raw_dir.rstrip('/')}/{matched}"

            file_size = os.path.getsize(local_file_path)

            def do_store():
                with open(local_file_path, "rb") as f:
                    return self._conn.storeFile(share, target_path, f, timeout=60)

            bytes_sent = self._run_with_reconnect(do_store)
            if callback:
                callback(file_size, file_size)
            return bytes_sent

    def download_file(self, share, remote_path, local_save_path, callback=None):
        """Скачивание одиночного файла с ЧПУ."""
        with self._lock:
            raw_remote = self.resolve_remote_path(share, remote_path, must_exist=True)

            def do_retrieve():
                with open(local_save_path, "wb") as f:
                    return self._conn.retrieveFile(share, raw_remote, f, timeout=60)

            _, bytes_received = self._run_with_reconnect(do_retrieve)
            if callback:
                callback(bytes_received, bytes_received)
            return bytes_received

    def upload_directory(self, local_dir_path, share, remote_parent_dir, status_callback=None):
        """Рекурсивная загрузка всей локальной папки на ЧПУ."""
        with self._lock:
            raw_parent = self.resolve_remote_path(share, remote_parent_dir, must_exist=True)
            base_name = os.path.basename(os.path.normpath(local_dir_path))
            target_remote_dir = f"{raw_parent.rstrip('/')}/{base_name}"
            self.create_directory(share, target_remote_dir)

            total_files = sum(len(files) for _, _, files in os.walk(local_dir_path))
            copied = 0
            for root, dirs, files in os.walk(local_dir_path):
                rel_path = os.path.relpath(root, local_dir_path)
                if rel_path == ".":
                    cur_remote_dir = target_remote_dir
                else:
                    cur_remote_dir = f"{target_remote_dir}/{rel_path.replace(os.sep, '/')}"

                for d in dirs:
                    self.create_directory(share, f"{cur_remote_dir}/{d}")

                for f in files:
                    loc_file = os.path.join(root, f)
                    if status_callback:
                        status_callback(f"Загрузка ({copied + 1}/{total_files}): {f}")
                    self.upload_file(loc_file, share, cur_remote_dir)
                    copied += 1
            return copied

    def download_directory(self, share, remote_dir_path, local_parent_dir, status_callback=None):
        """Рекурсивное скачивание всей папки с ЧПУ на локальный диск ПК."""
        with self._lock:
            raw_root = self.resolve_remote_path(share, remote_dir_path, must_exist=True)
            dir_name = decode_smb_name(raw_root.rstrip("/").split("/")[-1])
            target_local_dir = os.path.join(local_parent_dir, dir_name)
            os.makedirs(target_local_dir, exist_ok=True)

            copied = 0

            def _walk_remote(rem_raw_dir, loc_dir):
                nonlocal copied
                dirs, files = self._list_raw(share, rem_raw_dir)
                for f in files:
                    rem_f = f"{rem_raw_dir.rstrip('/')}/{f['raw_name']}"
                    loc_f = os.path.join(loc_dir, f["name"])
                    if status_callback:
                        status_callback(f"Скачивание: {f['name']}")
                    self.download_file(share, rem_f, loc_f)
                    copied += 1

                for d in dirs:
                    sub_rem = f"{rem_raw_dir.rstrip('/')}/{d['raw_name']}"
                    sub_loc = os.path.join(loc_dir, d["name"])
                    os.makedirs(sub_loc, exist_ok=True)
                    _walk_remote(sub_rem, sub_loc)

            _walk_remote(raw_root, target_local_dir)
            return copied
