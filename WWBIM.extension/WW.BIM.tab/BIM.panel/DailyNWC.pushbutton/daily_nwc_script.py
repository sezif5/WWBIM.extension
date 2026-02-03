# -*- coding: utf-8 -*-
"""
Daily NWC Export
"""

import os
import sys
import datetime
import glob
import threading
import time
import codecs
import re

script_dir = os.path.dirname(os.path.abspath(__file__))
lib_dir = os.path.join(os.path.dirname(script_dir), "..", "..", "..", "lib")
lib_dir = os.path.normpath(lib_dir)

if os.path.isdir(lib_dir) and lib_dir not in sys.path:
    sys.path.insert(0, lib_dir)

from pyrevit import script, coreutils, forms
from Autodesk.Revit.UI import ExternalEvent, IExternalEventHandler, ExternalEventRequest

from openbg import open_in_background
from closebg import close_with_policy
from nwc_export_utils import export_rvt_to_nwc

OBJECT_FOLDER_CONFIG = "Object_folder_path.txt"
DAILY_EXPORT_LIST = "Ежедневная выгрузка.txt"
LOG_FOLDER = "logs"
AUTO_MODE_FLAG = "--auto"
LOCK_FILE = ".export_lock"
TASK_NAME = "Daily NWC Export"
BAT_FILE = "run_daily_export.bat"
HOST_MODEL_CONFIG = "Host_Model.txt"
LOG_RETENTION_DAYS = 7
SERVICE_INTERVAL_SECONDS = 300
SERVICE_PERSISTENT_KEY = "daily_nwc_service"
EXPORT_ENABLED = True
EXPORT_TIME_CONFIG = "Export_Time.txt"
SERVICE_CHECK_INTERVAL_MS = 60000  # 1 minute check interval
EXPORT_TIME_WINDOW_SECONDS = 300  # 5 minute window for export trigger
SERVICE_DIAG_MODE = True  # True: only log queue steps, no export

out = script.get_output()
# Note: Removed out.close_others() - it conflicts with service logs


def _norm(p):
    return os.path.normpath(os.path.abspath(p)) if p else p


def _module_dir():
    try:
        return os.path.dirname(os.path.abspath(__file__))
    except Exception:
        return os.getcwd()


def _find_scripts_root(start_dir):
    cur = _norm(start_dir)
    for _ in range(0, 8):
        if os.path.basename(cur).lower() == "scripts":
            return cur
        parent = os.path.dirname(cur)
        if not parent or parent == cur:
            break
        cur = parent
    return _norm(os.path.join(start_dir, os.pardir, os.pardir))


def read_txt_file(filepath):
    if not os.path.exists(filepath):
        return None
    try:
        with open(filepath, "rb") as f:
            for raw in f:
                try:
                    line = raw.decode("utf-8").strip()
                except Exception:
                    try:
                        line = raw.decode("cp1251").strip()
                    except Exception:
                        line = raw.strip()
                if line:
                    return line
    except Exception as e:
        out.print_md(":x: Ошибка чтения файла `{}`: {}".format(filepath, e))
        return None


def write_txt_file(filepath, content):
    try:
        with open(filepath, "wb") as f:
            f.write(content.encode("utf-8"))
        return True
    except Exception as e:
        out.print_md(":x: Ошибка записи в файл `{}`: {}".format(filepath, e))
        return False


def get_script_root():
    module_dir = _module_dir()
    _env_root = os.environ.get("WW_SCRIPTS_ROOT")
    if _env_root and os.path.isdir(_env_root):
        return _norm(_env_root)
    return _find_scripts_root(module_dir)


def select_object_folder_ui():
    folder = forms.pick_folder(title="Выберите папку Object")
    if folder:
        return _norm(folder)
    return None


def read_object_folder_path(script_root):
    config_path = os.path.join(script_root, OBJECT_FOLDER_CONFIG)
    return read_txt_file(config_path)


def save_object_folder_path(object_folder_path, script_root):
    config_path = os.path.join(script_root, OBJECT_FOLDER_CONFIG)
    return write_txt_file(config_path, object_folder_path)


def read_export_time(object_folder_path):
    """Read export time from config file, returns default '00:00' if not set."""
    config_path = os.path.join(object_folder_path, EXPORT_TIME_CONFIG)
    time_str = read_txt_file(config_path)
    if time_str:
        # Validate format HH:MM
        try:
            parts = time_str.split(":")
            if len(parts) == 2:
                h, m = int(parts[0]), int(parts[1])
                if 0 <= h <= 23 and 0 <= m <= 59:
                    return time_str
        except ValueError:
            pass
    return "00:00"


def save_export_time(export_time, object_folder_path):
    """Save export time to config file."""
    config_path = os.path.join(object_folder_path, EXPORT_TIME_CONFIG)
    return write_txt_file(config_path, export_time)


def read_object_config(object_name, object_folder_path):
    result = {
        "rvt_paths": [],
        "nwc_folder": None,
        "rvt_exists": False,
        "nwc_folder_exists": False,
    }

    rvt_file = os.path.join(object_folder_path, object_name + ".txt")
    rvt_paths = []
    try:
        with codecs.open(rvt_file, "r", encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if line:
                    line = normalize_revit_server_path(line)
                    rvt_paths.append(line)
    except Exception:
        try:
            with codecs.open(rvt_file, "r", encoding="cp1251") as f:
                for line in f:
                    line = line.strip()
                    if line:
                        line = normalize_revit_server_path(line)
                        rvt_paths.append(line)
        except Exception:
            pass

    result["rvt_paths"] = rvt_paths

    if rvt_paths and any(p.upper().startswith("RSN:") for p in rvt_paths):
        out.print_md("[DEBUG] Found {} RVT paths (normalized)".format(len(rvt_paths)))
        for i, p in enumerate(rvt_paths, 1):
            out.print_md("[DEBUG]   [{}] {}".format(i, p))

    def path_exists(p):
        """
        Проверяет существование пути.
        Для Revit Server всегда возвращает True (проверка будет при открытии).
        Для локальных файлов проверяет os.path.exists().
        """
        if not p:
            return False

        try:
            if p.upper().startswith("RSN:"):
                out.print_md("[DEBUG] Revit Server path assumed valid: {}".format(p))
                return True
            exists = os.path.exists(p)
            if exists:
                out.print_md("[DEBUG] File exists: {}".format(p))
            else:
                out.print_md("[DEBUG] File NOT found: {}".format(p))
            return exists
        except Exception as e:
            out.print_md("[DEBUG] Path exists check error: {} - {}".format(p, str(e)))
            return False

    out.print_md("[DEBUG] Checking if any RVT file exists...")
    rvt_exists = any(path_exists(p) for p in rvt_paths) if rvt_paths else False
    out.print_md("[DEBUG] RVT exists result: {}".format(rvt_exists))
    result["rvt_exists"] = rvt_exists

    nwc_file = os.path.join(object_folder_path, object_name + "_NWC.txt")
    nwc_folder = read_txt_file(nwc_file)
    if nwc_folder:
        result["nwc_folder"] = _norm(nwc_folder)
        result["nwc_folder_exists"] = os.path.isdir(result["nwc_folder"])

    return result


def get_available_objects(object_folder_path):
    if not os.path.isdir(object_folder_path):
        return []

    objects = []
    try:
        files = os.listdir(object_folder_path)
        for file in files:
            if file.lower().endswith(".txt") and not file.lower().endswith("_nwc.txt"):
                if file not in [
                    OBJECT_FOLDER_CONFIG,
                    DAILY_EXPORT_LIST,
                    BAT_FILE,
                    LOCK_FILE,
                ]:
                    objects.append(file[:-4])
    except Exception as e:
        out.print_md(":x: Ошибка чтения папки `{}`: {}".format(object_folder_path, e))

    out.print_md("[DEBUG] Available objects: {}".format(sorted(objects)))
    return sorted(objects)


def read_export_list(object_folder_path):
    list_path = os.path.join(object_folder_path, DAILY_EXPORT_LIST)
    if not os.path.exists(list_path):
        return []

    objects = []
    try:
        with codecs.open(list_path, "r", encoding="utf-8") as f:
            content = f.read()
            lines = content.splitlines()
            for line_num, line in enumerate(lines, 1):
                line = line.strip()
                if line:
                    if "," in line:
                        parts = line.split(",")
                        for part in parts:
                            part = part.strip()
                            if part:
                                objects.append(part)
                    else:
                        objects.append(line)

        out.print_md("[DEBUG] File content ({} lines):".format(len(lines)))
        for i, line in enumerate(lines[:10], 1):
            out.print_md("[DEBUG]   Line {}: '{}'".format(i, line))
        if len(lines) > 10:
            out.print_md("[DEBUG]   ... and {} more lines".format(len(lines) - 10))
        out.print_md("[DEBUG] Parsed objects ({}): {}".format(len(objects), objects))
    except Exception as e:
        out.print_md(":x: Ошибка чтения списка `{}`: {}".format(list_path, e))

    return objects


def save_export_list(objects, object_folder_path):
    list_path = os.path.join(object_folder_path, DAILY_EXPORT_LIST)
    content = "\n".join(objects)
    return write_txt_file(list_path, content)


def select_objects_ui(object_folder_path):
    available = get_available_objects(object_folder_path)
    out.print_md("[DEBUG] select_objects_ui: folder = {}".format(object_folder_path))
    out.print_md("[DEBUG] available = {}".format(available))

    if not available:
        forms.alert(
            "В папке Object нет доступных объектов.\n"
            "Создайте файлы вида 'ИмяОбъекта.txt' с путями к RVT файлам.",
            ok=False,
            exitscript=True,
        )

    selected = forms.SelectFromList.show(
        available,
        title="Выберите объекты для ежедневной выгрузки",
        multiselect=True,
        button_name="Выбрать",
    )

    if selected:
        save_export_list(list(selected), object_folder_path)
        return list(selected)
    else:
        return None


def normalize_revit_server_path(path):
    """
    Нормализует путь к файлу на Revit Server.

    Примеры:
    - rsn:\\server\path -> RSN://server/path
    - rsn:\\\\server\path -> RSN://server/path
    - rsn:/server/path -> RSN://server/path
    - RSN:\\\\server\path -> RSN://server/path
    """
    if not path:
        return path

    import re

    path_upper = path.upper()
    if not path_upper.startswith("RSN:"):
        return path

    path = path.replace("rsn:", "RSN:")
    path = path.replace("RSN:", "RSN:")

    path = path.replace("\\\\", "\\")
    path = path.replace("\\\\", "\\")
    path = path.replace("\\", "/")
    path = path.replace("RSN:/", "RSN://")
    path = path.replace("RSN:///", "RSN://")

    return path


def _dotnet_datetime_to_py(dt_value):
    """Конвертирует System.DateTime в datetime.datetime."""
    if dt_value is None:
        return None
    try:
        return datetime.datetime(
            dt_value.Year,
            dt_value.Month,
            dt_value.Day,
            dt_value.Hour,
            dt_value.Minute,
            dt_value.Second,
        )
    except Exception:
        return None


def get_file_modification_date(filepath):
    """
    Получает дату изменения файла.

    Для локальных файлов использует os.path.getmtime().
    Для файлов на Revit Server (RSN://) использует BasicFileInfo.Read().
    """
    if not filepath:
        return None

    if filepath.upper().startswith("RSN:"):
        try:
            from Autodesk.Revit.DB import ModelPathUtils, BasicFileInfo

            mp = ModelPathUtils.ConvertUserVisiblePathToModelPath(filepath)
            if mp is None:
                out.print_md("[DEBUG] Cannot convert to ModelPath: {}".format(filepath))
                return None

            try:
                info = BasicFileInfo.Read(mp)
                if info is not None and info.LastModifiedTime is not None:
                    result = _dotnet_datetime_to_py(info.LastModifiedTime)
                    out.print_md(
                        "[DEBUG] Got BasicFileInfo.Read for {}: {}".format(
                            os.path.basename(filepath), result
                        )
                    )
                    return result
            except Exception as e:
                out.print_md(
                    "[DEBUG] BasicFileInfo.Read error: {}".format(str(e)[:100])
                )

            try:
                info = BasicFileInfo.GetBasicFileInfo(mp)
                if info is not None and info.LastModifiedTime is not None:
                    result = _dotnet_datetime_to_py(info.LastModifiedTime)
                    out.print_md(
                        "[DEBUG] Got BasicFileInfo.GetBasicFileInfo for {}: {}".format(
                            os.path.basename(filepath), result
                        )
                    )
                    return result
            except Exception as e:
                out.print_md(
                    "[DEBUG] BasicFileInfo.GetBasicFileInfo error: {}".format(
                        str(e)[:100]
                    )
                )

            return None
        except Exception as e:
            out.print_md(
                "[DEBUG] Revit Server info error for {}: {}".format(
                    filepath, str(e)[:100]
                )
            )
            return None
    else:
        if not os.path.exists(filepath):
            return None
        try:
            timestamp = os.path.getmtime(filepath)
            return datetime.datetime.fromtimestamp(timestamp)
        except Exception:
            return None


def get_file_path_type(filepath):
    """Определяет тип пути: 'local', 'revit_server' или 'unknown'."""
    if not filepath:
        return "unknown"
    if filepath.startswith(("RSN://", "rsn://", "rsn:/", "RSN:/")):
        return "revit_server"
    return "local"


def check_need_export(rvt_path, nwc_folder, object_name, app=None, revit=None):
    path_type = get_file_path_type(rvt_path)

    if path_type == "revit_server":
        rvt_filename = os.path.splitext(os.path.basename(rvt_path))[0]

        nwc_path1 = os.path.join(nwc_folder, rvt_filename + ".nwc")
        nwc_date1 = get_file_modification_date(nwc_path1)

        nwc_path2 = None
        nwc_date2 = None
        match = re.search(r"_R(\d+)$", rvt_filename)
        if match:
            suffix = match.group(0)
            version_num = match.group(1)
            new_version_num = str(int(version_num) + 1)
            new_suffix = "_N" + str(new_version_num)
            nwc_filename2 = rvt_filename.replace(suffix, new_suffix)
            nwc_path2 = os.path.join(nwc_folder, nwc_filename2 + ".nwc")
            nwc_date2 = get_file_modification_date(nwc_path2)

        nwc_date = None
        nwc_used_path = None

        if nwc_date1 and nwc_date2:
            nwc_date = max(nwc_date1, nwc_date2)
            nwc_used_path = nwc_path1 if nwc_date1 > nwc_date2 else nwc_path2
        elif nwc_date1:
            nwc_date = nwc_date1
            nwc_used_path = nwc_path1
        elif nwc_date2:
            nwc_date = nwc_date2
            nwc_used_path = nwc_path2

        return {
            "need_export": True,
            "reason": "Revit Server (всегда экспортировать)",
            "rvt_date": None,
            "nwc_date": nwc_date,
            "nwc_exists": nwc_date is not None,
            "nwc_used_path": nwc_used_path,
            "rvt_filename": rvt_filename,
            "nwc_path1": nwc_path1,
            "nwc_date1": nwc_date1,
            "nwc_path2": nwc_path2,
            "nwc_date2": nwc_date2,
            "path_type": path_type,
        }

    rvt_date = get_file_modification_date(rvt_path)

    rvt_filename = os.path.splitext(os.path.basename(rvt_path))[0]

    nwc_path1 = os.path.join(nwc_folder, rvt_filename + ".nwc")
    nwc_date1 = get_file_modification_date(nwc_path1)

    nwc_path2 = None
    nwc_date2 = None
    match = re.search(r"_R(\d+)$", rvt_filename)
    if match:
        suffix = match.group(0)
        version_num = match.group(1)
        new_version_num = str(int(version_num) + 1)
        new_suffix = "_N" + str(new_version_num)
        nwc_filename2 = rvt_filename.replace(suffix, new_suffix)
        nwc_path2 = os.path.join(nwc_folder, nwc_filename2 + ".nwc")
        nwc_date2 = get_file_modification_date(nwc_path2)

    nwc_date = None
    nwc_used_path = None

    if nwc_date1 and nwc_date2:
        nwc_date = max(nwc_date1, nwc_date2)
        nwc_used_path = nwc_path1 if nwc_date1 > nwc_date2 else nwc_path2
    elif nwc_date1:
        nwc_date = nwc_date1
        nwc_used_path = nwc_path1
    elif nwc_date2:
        nwc_date = nwc_date2
        nwc_used_path = nwc_path2

    need_export = True
    reason = ""

    if nwc_date:
        if rvt_date and rvt_date <= nwc_date:
            need_export = False
            reason = "NWC актуален"
        elif rvt_date and rvt_date > nwc_date:
            need_export = True
            reason = "RVT обновлён"
        else:
            need_export = True
            reason = "Не удалось определить даты"
    else:
        need_export = True
        reason = "NWC не существует"

    return {
        "need_export": need_export,
        "reason": reason,
        "rvt_date": rvt_date,
        "nwc_date": nwc_date,
        "nwc_exists": nwc_date is not None,
        "nwc_used_path": nwc_used_path,
        "rvt_filename": rvt_filename,
        "nwc_path1": nwc_path1,
        "nwc_date1": nwc_date1,
        "nwc_path2": nwc_path2,
        "nwc_date2": nwc_date2,
        "path_type": path_type,
    }


def init_logger(object_folder_path):
    log_dir = os.path.join(object_folder_path, LOG_FOLDER)
    if not os.path.exists(log_dir):
        try:
            os.makedirs(log_dir)
        except Exception as e:
            out.print_md(":x: Ошибка создания папки логов `{}`: {}".format(log_dir, e))
            return None

    cleanup_old_logs(log_dir)

    today = datetime.datetime.now().strftime("%Y-%m-%d")
    log_path = os.path.join(log_dir, "export_log_{}.txt".format(today))

    try:
        with codecs.open(log_path, "a", encoding="utf-8") as f:
            timestamp = datetime.datetime.now().strftime("%H:%M:%S")
            f.write("=" * 40 + "\n")
            f.write(
                "Daily NWC Export - {} {}\n".format(
                    datetime.datetime.now().strftime("%Y-%m-%d"), timestamp
                )
            )
            f.write("=" * 40 + "\n\n")
        return log_path
    except Exception as e:
        out.print_md(":x: Ошибка создания файла лога `{}`: {}".format(log_path, e))
        return None


def cleanup_old_logs(log_dir):
    try:
        if not os.path.isdir(log_dir):
            return

        now = datetime.datetime.now()
        cutoff = now - datetime.timedelta(days=LOG_RETENTION_DAYS)

        for log_file in glob.glob(os.path.join(log_dir, "export_log_*.txt")):
            try:
                file_time = datetime.datetime.fromtimestamp(os.path.getmtime(log_file))
                if file_time < cutoff:
                    os.remove(log_file)
            except Exception:
                pass
    except Exception:
        pass


def log_message(log_path, message):
    if not log_path:
        return
    try:
        with codecs.open(log_path, "a", encoding="utf-8") as f:
            timestamp = datetime.datetime.now().strftime("[%H:%M:%S]")
            f.write("{} {}\n".format(timestamp, message))
            f.flush()  # Force write to disk
            os.fsync(f.fileno())  # Force OS to flush
    except Exception:
        pass


def log_export_start(log_path, object_name, rvt_path):
    msg = "Object: {}".format(object_name)
    log_message(log_path, msg)
    msg = "  - RVT: {}".format(rvt_path)
    log_message(log_path, msg)


def log_export_success(log_path, object_name, elapsed_time, file_size_mb):
    msg = "  - Export: SUCCESS (time: {}, size: {:.1f} MB)".format(
        elapsed_time, float(file_size_mb) if file_size_mb is not None else 0
    )
    log_message(log_path, msg)


def log_export_error(log_path, object_name, error_message):
    msg = "  - Error: {}".format(error_message)
    log_message(log_path, msg)


def log_export_skipped(
    log_path, object_name, reason, rvt_date=None, nwc_date=None, nwc_used_path=None
):
    msg = "  - NWC: {}".format(reason)
    log_message(log_path, msg)
    if rvt_date:
        msg = "  - RVT date: {}".format(rvt_date.strftime("%Y-%m-%d %H:%M:%S"))
        log_message(log_path, msg)
    if nwc_date:
        msg = "  - NWC date: {}".format(nwc_date.strftime("%Y-%m-%d %H:%M:%S"))
        log_message(log_path, msg)
    if nwc_used_path:
        msg = "  - NWC used: {}".format(nwc_used_path)
        log_message(log_path, msg)


def log_summary(log_path, total, exported, skipped, errors, total_time):
    log_message(log_path, "")
    log_message(log_path, "=" * 40)
    msg = "TOTAL: {}, Exported: {}, Skipped: {}, Errors: {}".format(
        total, exported, skipped, errors
    )
    log_message(log_path, msg)
    msg = "Total time: {}".format(total_time)
    log_message(log_path, msg)
    log_message(log_path, "=" * 40)


class NWCExportHandler(IExternalEventHandler):
    def __init__(self):
        self.export_request = None
        self.export_result = None
        self.export_event = threading.Event()

    def set_export_request(self, object_name, rvt_path, nwc_folder, mode="real"):
        self.export_request = {
            "object_name": object_name,
            "rvt_path": rvt_path,
            "nwc_folder": nwc_folder,
            "mode": mode,
        }
        self.export_result = None
        self.export_event.clear()

    def Execute(self, uiapp):
        if self.export_request is None:
            return

        object_name = self.export_request["object_name"]
        rvt_path = self.export_request["rvt_path"]
        nwc_folder = self.export_request["nwc_folder"]
        mode = self.export_request["mode"]

        try:
            result = export_rvt_to_nwc(rvt_path, nwc_folder, uiapp.Application, uiapp)
            self.export_result = result
        except Exception as e:
            self.export_result = {
                "success": False,
                "error": str(e),
                "exported_file": None,
                "file_size_mb": None,
            }
        finally:
            self.export_event.set()

    def GetName(self):
        return "NWC Export Handler"


class NWCExportTask:
    def __init__(self):
        self.handler = NWCExportHandler()
        self.external_event = ExternalEvent.Create(self.handler)

    def export(
        self, object_name, rvt_path, nwc_folder, mode="real", timeout=600, log_path=None
    ):
        self.handler.set_export_request(object_name, rvt_path, nwc_folder, mode)
        request = self.external_event.Raise()

        if request == ExternalEventRequest.Accepted:
            if log_path:
                log_message(
                    log_path, "  - ExternalEvent accepted, waiting for export..."
                )
            self.handler.export_event.wait(timeout=timeout)
            if log_path:
                log_message(log_path, "  - ExternalEvent completed")
            return self.handler.export_result
        else:
            if log_path:
                log_message(
                    log_path,
                    "  - ExternalEvent request was not accepted: {}".format(request),
                )
            return {
                "success": False,
                "error": "ExternalEvent request was not accepted: {}".format(request),
                "exported_file": None,
                "file_size_mb": None,
            }


class NWCExportQueueHandler(IExternalEventHandler):
    def __init__(self):
        self.task_queue = []
        self.current_index = 0
        self.is_exporting = False
        self.log_path = None
        self.object_folder_path = None
        self.app = None
        self.revit = None
        self.service_manager = None
        self.export_enabled = False
        self.queue_event = None

    def set_queue(
        self,
        tasks,
        log_path,
        object_folder_path,
        app,
        revit,
        service_manager,
        export_enabled=False,
    ):
        self.task_queue = tasks
        self.current_index = 0
        self.is_exporting = True
        self.log_path = log_path
        self.object_folder_path = object_folder_path
        self.app = app
        self.revit = revit
        self.service_manager = service_manager
        self.export_enabled = export_enabled

    def set_queue_event(self, event):
        self.queue_event = event

    def Execute(self, uiapp):
        if not self.is_exporting or self.current_index >= len(self.task_queue):
            self._cleanup()
            return

        log_message(
            self.log_path,
            "[Queue] Execute entered. Index: {} of {}".format(
                self.current_index + 1, len(self.task_queue)
            ),
        )

        task = self.task_queue[self.current_index]
        object_name = task["object_name"]
        rvt_path = task["rvt_path"]
        nwc_folder = task["nwc_folder"]

        log_message(
            self.log_path,
            "[Queue] Processing task {}/{}: {} - {}".format(
                self.current_index + 1,
                len(self.task_queue),
                object_name,
                os.path.basename(rvt_path),
            ),
        )

        if SERVICE_DIAG_MODE:
            log_message(
                self.log_path,
                "[Queue] Diagnostic mode - skipping export logic",
            )
            self._move_to_next_task()
            return

        try:
            path_type = get_file_path_type(rvt_path)
            path_exists = False

            if path_type == "revit_server":
                path_exists = get_file_modification_date(rvt_path) is not None
            else:
                path_exists = os.path.exists(rvt_path)

            if not path_exists:
                error = "RVT not found: {}".format(rvt_path)
                log_export_error(self.log_path, object_name, error)
                self._move_to_next_task()
                return

            check_result = check_need_export(
                rvt_path,
                nwc_folder,
                object_name,
                self.app,
                self.revit,
            )

            log_message(
                self.log_path, "[Queue] Check result: {}".format(check_result["reason"])
            )

            if not check_result["need_export"]:
                log_export_skipped(
                    self.log_path,
                    object_name,
                    check_result["reason"],
                    check_result.get("rvt_date"),
                    check_result.get("nwc_date"),
                    check_result.get("nwc_used_path"),
                )
                self._move_to_next_task()
                return

            if self.export_enabled:
                log_message(
                    self.log_path,
                    "[Queue] Starting export for {}...".format(object_name),
                )

                result = export_rvt_to_nwc(
                    rvt_path,
                    nwc_folder,
                    uiapp.Application,
                    uiapp,
                )

                if result and result.get("success"):
                    file_size = result.get("file_size_mb", 0)
                    elapsed = result.get("time_export", "0s")
                    log_export_success(
                        self.log_path,
                        object_name,
                        elapsed,
                        file_size,
                    )
                else:
                    error = (
                        result.get("error", "Unknown error")
                        if result
                        else "Unknown error"
                    )
                    log_export_error(self.log_path, object_name, error)
            else:
                log_message(
                    self.log_path,
                    "[Queue] Export disabled (dry-run mode) for {}".format(object_name),
                )

        except Exception as e:
            error = "Export error for {}: {}".format(object_name, str(e))
            log_export_error(self.log_path, object_name, error)

        self._move_to_next_task()

    def _move_to_next_task(self):
        self.current_index += 1

        if self.current_index < len(self.task_queue):
            if self.queue_event:
                self.queue_event.Raise()
        else:
            self._cleanup()

    def _cleanup(self):
        log_message(
            self.log_path,
            "[Queue] Queue completed. Total tasks: {}".format(len(self.task_queue)),
        )
        self.is_exporting = False
        self.task_queue = []
        self.current_index = 0

        if self.service_manager:
            self.service_manager._on_queue_finished()

        remove_lock_file(self.object_folder_path)

    def GetName(self):
        return "NWC Export Queue Handler"


class ServiceManager:
    """
    Service manager for scheduled NWC exports.

    Uses WPF DispatcherTimer instead of threading to ensure all Revit API calls
    (including ExternalEvent) happen on the UI thread.
    """

    def __init__(
        self,
        object_folder_path,
        export_list,
        log_path,
        app,
        revit,
        export_enabled=False,
        export_time="00:00",
    ):
        self.object_folder_path = object_folder_path
        self.export_list = export_list
        self.log_path = log_path
        self.app = app
        self.revit = revit
        self.running = False
        self.timer = None
        self.queue_handler = NWCExportQueueHandler()
        self.queue_event = ExternalEvent.Create(self.queue_handler)
        self.queue_handler.set_queue_event(self.queue_event)
        self.last_error = None
        self.last_error_time = None
        self.export_enabled = export_enabled
        self.export_time = export_time
        self.last_export_date = None
        self._is_exporting = False

    def start(self):
        """Start the service using DispatcherTimer (UI thread safe)."""
        if self.running:
            return False

        try:
            from System.Windows.Threading import DispatcherTimer
            from System import TimeSpan, EventHandler

            self.running = True
            self.timer = DispatcherTimer()
            self.timer.Interval = TimeSpan.FromMilliseconds(SERVICE_CHECK_INTERVAL_MS)
            self.timer.Tick += EventHandler(self._on_timer_tick)
            self.timer.Start()

            log_message(
                self.log_path,
                "[Service] Started. Check interval: {}ms, Export time: {}".format(
                    SERVICE_CHECK_INTERVAL_MS, self.export_time
                ),
            )
            log_message(
                self.log_path,
                "[Service] Timer created and started",
            )
            return True
        except Exception as e:
            self.last_error = str(e)
            self.last_error_time = datetime.datetime.now()
            log_message(self.log_path, "[Service] Failed to start: {}".format(e))
            return False

    def stop(self):
        """Stop the service."""
        self.running = False
        if self.timer:
            try:
                self.timer.Stop()
                self.timer = None
            except Exception:
                pass
        log_message(self.log_path, "[Service] Stopped.")

    def _on_timer_tick(self, sender, args):
        """
        Timer tick handler - called on UI thread.
        Safe to use ExternalEvent here.
        """
        if not self.running or self._is_exporting:
            return

        log_message(
            self.log_path,
            "[Service] Timer tick at {}".format(
                datetime.datetime.now().strftime("%H:%M:%S")
            ),
        )

        try:
            if self._should_export_now():
                log_message(
                    self.log_path,
                    "[Service] Time to export: {} (scheduled: {})".format(
                        datetime.datetime.now().strftime("%H:%M:%S"), self.export_time
                    ),
                )
                self._is_exporting = True
                try:
                    self._check_and_export()
                    self.last_export_date = datetime.datetime.now().strftime("%Y-%m-%d")
                finally:
                    self._is_exporting = False
            else:
                log_message(
                    self.log_path,
                    "[Service] Not time to export yet",
                )
        except Exception as e:
            self.last_error = str(e)
            self.last_error_time = datetime.datetime.now()
            log_message(self.log_path, "[Service Error] {}".format(e))

    def _should_export_now(self):
        """
        Check if export should run now.

        Robust daily logic:
        - If now >= scheduled time AND not already exported today → trigger
        - This ensures export happens even if Revit was busy/frozen
          or computer was asleep past the scheduled time.
        """
        now = datetime.datetime.now()
        current_date = now.strftime("%Y-%m-%d")

        # Already exported today?
        if self.last_export_date == current_date:
            return False

        # Parse scheduled time
        try:
            hour, minute = map(int, self.export_time.split(":"))
            scheduled = now.replace(hour=hour, minute=minute, second=0, microsecond=0)
        except ValueError:
            log_message(
                self.log_path,
                "[Service] Invalid export_time format: {}".format(self.export_time),
            )
            return False

        # Check if current time >= scheduled time
        if now >= scheduled:
            diff_seconds = int((now - scheduled).total_seconds())
            log_message(
                self.log_path,
                "[Service] Time to export: {} seconds after scheduled time {}".format(
                    diff_seconds, self.export_time
                ),
            )
            return True

        return False

    def _check_and_export(self):
        """Check each object and export if needed using queue handler."""
        # Use lock file to prevent concurrent runs
        if check_lock_file(self.object_folder_path):
            log_message(
                self.log_path,
                "[Service] Skipping - another export is in progress (lock file exists)",
            )
            return

        log_message(self.log_path, "[Service] Building export queue")
        # Build task queue
        tasks = []
        for object_name in self.export_list:
            if not self.running:
                break

            config = read_object_config(object_name, self.object_folder_path)

            if not config["rvt_exists"]:
                error = "RVT file not found: {}".format(config["rvt_paths"])
                log_export_error(self.log_path, object_name, error)
                continue

            if not config["nwc_folder_exists"]:
                error = "NWC folder not found: {}".format(config["nwc_folder"])
                log_export_error(self.log_path, object_name, error)
                continue

            for rvt_path in config["rvt_paths"]:
                tasks.append(
                    {
                        "object_name": object_name,
                        "rvt_path": rvt_path,
                        "nwc_folder": config["nwc_folder"],
                    }
                )

        if tasks:
            log_message(
                self.log_path,
                "[Service] Starting export queue with {} tasks".format(len(tasks)),
            )
            create_lock_file(self.object_folder_path)
            self.queue_handler.set_queue(
                tasks,
                self.log_path,
                self.object_folder_path,
                self.app,
                self.revit,
                self,
                self.export_enabled,
            )
            raise_result = self.queue_event.Raise()
            log_message(
                self.log_path,
                "[Service] ExternalEvent.Raise returned: {}".format(raise_result),
            )
        else:
            log_message(self.log_path, "[Service] No tasks to export")

    def get_status(self):
        """Get current service status."""
        status = {
            "running": self.running,
            "last_error": self.last_error,
            "last_error_time": self.last_error_time,
            "last_export_date": self.last_export_date,
            "export_time": self.export_time,
            "is_exporting": self._is_exporting,
        }
        return status

    def _on_queue_finished(self):
        """Callback when export queue finishes."""
        log_message(self.log_path, "[Service] Export queue finished")
        self.last_export_date = datetime.datetime.now().strftime("%Y-%m-%d")
        self._is_exporting = False


def get_service_manager():
    """
    Получает менеджер службы из AppDomain.

    Использует System.AppDomain.CurrentDomain для хранения ссылки на ServiceManager.
    Это работает независимо от версии pyRevit и держит объект в памяти
    пока работает процесс Revit.
    """
    try:
        import System

        return System.AppDomain.CurrentDomain.GetData(SERVICE_PERSISTENT_KEY)
    except Exception:
        return None


def set_service_manager(manager):
    """
    Сохраняет менеджер службы в AppDomain.

    AppDomain.CurrentDomain живёт пока работает Revit, что гарантирует
    что ServiceManager и его DispatcherTimer останутся в памяти.
    """
    try:
        import System

        System.AppDomain.CurrentDomain.SetData(SERVICE_PERSISTENT_KEY, manager)
        return True
    except Exception as e:
        out.print_md(":x: Failed to save service manager: {}".format(e))
        return False


def clear_service_manager():
    """Очищает менеджер службы из AppDomain."""
    try:
        import System

        System.AppDomain.CurrentDomain.SetData(SERVICE_PERSISTENT_KEY, None)
        return True
    except Exception:
        return False


def export_single_object(
    object_name,
    rvt_path,
    nwc_folder,
    log_path,
    auto_mode=False,
    export_task=None,
    export_enabled=False,
):
    log_export_start(log_path, object_name, rvt_path)

    check_result = check_need_export(rvt_path, nwc_folder, object_name)

    path_type = check_result.get("path_type", "unknown")
    log_message(log_path, "  - Path type: {}".format(path_type.upper()))
    log_message(log_path, "  - Check result: {}".format(check_result["reason"]))

    rvt_date = check_result.get("rvt_date")
    if rvt_date:
        log_message(
            log_path, "  - RVT date: {}".format(rvt_date.strftime("%Y-%m-%d %H:%M:%S"))
        )
    else:
        log_message(log_path, "  - RVT date: NOT FOUND")

    nwc_used_path = check_result.get("nwc_used_path")
    nwc_date = check_result.get("nwc_date")
    if nwc_used_path and nwc_date:
        log_message(
            log_path,
            "  - NWC used: {} ({})".format(
                nwc_date.strftime("%Y-%m-%d %H:%M:%S"), nwc_used_path
            ),
        )
    elif nwc_used_path:
        log_message(log_path, "  - NWC used: {} (NOT FOUND)".format(nwc_used_path))
    else:
        log_message(log_path, "  - NWC used: NONE (no NWC files found)")

    if not check_result["need_export"]:
        log_export_skipped(
            log_path,
            object_name,
            check_result["reason"],
            check_result.get("rvt_date"),
            check_result.get("nwc_date"),
            check_result.get("nwc_used_path"),
        )
        if not auto_mode:
            out.print_md(
                ":white_check_mark: {}: **{}**".format(
                    object_name, check_result["reason"]
                )
            )
        return True, False, None

    if not auto_mode:
        out.print_md(":hourglass: Exporting: **{}**".format(object_name))

    if export_enabled:
        if export_task:
            export_result = export_task.export(
                object_name,
                rvt_path,
                nwc_folder,
                "real" if export_enabled else "dry",
                log_path=log_path,
            )
        else:
            temp_task = NWCExportTask()
            export_result = temp_task.export(
                object_name,
                rvt_path,
                nwc_folder,
                "real" if export_enabled else "dry",
                log_path=log_path,
            )

        if export_result and export_result.get("success"):
            file_size = export_result.get("file_size_mb", 0)
            elapsed = export_result.get("time_export", "0s")
            log_export_success(log_path, object_name, elapsed, file_size)

            if not auto_mode:
                msg = ":white_check_mark: {}: SUCCESS ({} MB)".format(
                    object_name, round(file_size, 2)
                )
                out.print_md(msg)
            return True, True, None
        else:
            error = (
                export_result.get("error", "Unknown error")
                if export_result
                else "Unknown error"
            )
            log_export_error(log_path, object_name, error)

            if not auto_mode:
                out.print_md(":x: {}: **{}**".format(object_name, error))
            return False, False, error
    else:
        log_message(log_path, "  - Export disabled (dry-run mode)")
        if not auto_mode:
            out.print_md(
                ":white_check_mark: {}: **Export disabled (dry-run)** - {}".format(
                    object_name, check_result["reason"]
                )
            )
        return True, False, None


def export_all_objects(
    object_folder_path, export_list, auto_mode=False, export_enabled=False
):
    log_path = init_logger(object_folder_path)
    if not log_path:
        return None

    t_all = coreutils.Timer()

    total = 0
    exported = 0
    skipped = 0
    errors = 0

    if not auto_mode:
        out.print_md("## EXPORT NWC ({})".format(len(export_list)))
        out.print_md("Export folder: **{}**".format(object_folder_path))
        out.print_md("Objects: {}".format(export_list))
        out.print_md("___")
        out.update_progress(0, len(export_list))

    log_message(log_path, "Export list: {}".format(export_list))

    for i, object_name in enumerate(export_list):
        total += 1

        config = read_object_config(object_name, object_folder_path)

        if not config["rvt_exists"]:
            error = "RVT file not found: {}".format(config["rvt_paths"])
            log_export_error(log_path, object_name, error)
            if not auto_mode:
                out.print_md(":x: {}: **{}**".format(object_name, error))
            errors += 1
            if not auto_mode:
                out.update_progress(i + 1, len(export_list))
            continue

        if not config["nwc_folder_exists"]:
            error = "NWC folder not found: {}".format(config["nwc_folder"])
            log_export_error(log_path, object_name, error)
            if not auto_mode:
                out.print_md(":x: {}: **{}**".format(object_name, error))
            errors += 1
            if not auto_mode:
                out.update_progress(i + 1, len(export_list))
            continue

        object_exported = 0
        object_skipped = 0
        object_errors = 0

        for rvt_idx, rvt_path in enumerate(config["rvt_paths"], 1):
            path_type = get_file_path_type(rvt_path)
            out.print_md(
                "[DEBUG] RVT {} - path type: {} - path: {}".format(
                    rvt_idx, path_type, rvt_path
                )
            )
            path_exists = False

            if path_type == "revit_server":
                path_exists = True
                out.print_md(
                    "[DEBUG] RVT {} - Revit Server path, assuming valid".format(rvt_idx)
                )
            else:
                path_exists = os.path.exists(rvt_path)

            if path_exists:
                success, was_exported, error = export_single_object(
                    object_name + "_" + str(rvt_idx),
                    rvt_path,
                    config["nwc_folder"],
                    log_path,
                    auto_mode,
                    export_task=None,
                    export_enabled=export_enabled,
                )

                if success:
                    if was_exported:
                        object_exported += 1
                    else:
                        object_skipped += 1
                else:
                    object_errors += 1
            else:
                error = "RVT not found: {}".format(rvt_path)
                log_export_error(log_path, object_name + "_" + str(rvt_idx), error)
                object_errors += 1

        exported += object_exported
        skipped += object_skipped
        errors += object_errors

        if not auto_mode:
            out.print_md(
                "Object {}: {} RVT files, {} exported, {} skipped, {} errors".format(
                    object_name,
                    len(config["rvt_paths"]),
                    object_exported,
                    object_skipped,
                    object_errors,
                )
            )
            out.update_progress(i + 1, len(export_list))

    all_s = str(datetime.timedelta(seconds=int(t_all.get_time())))

    log_summary(log_path, total, exported, skipped, errors, all_s)

    if not auto_mode:
        out.print_md("___")
        out.print_md(
            "**Done. Total: {}, Exported: {}, Skipped: {}, Errors: {}**".format(
                total, exported, skipped, errors
            )
        )
        out.print_md("**Total time: {}**".format(all_s))
        out.print_md("**Log: `{}`**".format(log_path))

    return {
        "total": total,
        "exported": exported,
        "skipped": skipped,
        "errors": errors,
        "total_time": all_s,
        "log_path": log_path,
    }


def check_lock_file(object_folder_path):
    lock_path = os.path.join(object_folder_path, LOCK_FILE)
    if os.path.exists(lock_path):
        try:
            lock_time = datetime.datetime.fromtimestamp(os.path.getmtime(lock_path))
            if datetime.datetime.now() - lock_time > datetime.timedelta(hours=24):
                try:
                    os.remove(lock_path)
                except Exception:
                    pass
                return False
            else:
                return True
        except Exception:
            return True
    return False


def create_lock_file(object_folder_path):
    lock_path = os.path.join(object_folder_path, LOCK_FILE)
    try:
        with open(lock_path, "w") as f:
            f.write(str(datetime.datetime.now()))
        return True
    except Exception:
        return False


def remove_lock_file(object_folder_path):
    lock_path = os.path.join(object_folder_path, LOCK_FILE)
    try:
        if os.path.exists(lock_path):
            os.remove(lock_path)
    except Exception:
        pass


def get_pyrevit_cli_path():
    possible_paths = [
        os.path.join(
            os.path.expanduser("~"),
            "AppData",
            "Roaming",
            "pyRevit",
            "Revit",
            "addin",
            "bin",
            "pyrevit.exe",
        ),
        os.path.join(
            os.path.expanduser("~"), "AppData", "Local", "pyRevit", "bin", "pyrevit.exe"
        ),
        os.path.join(
            os.path.expanduser("~"),
            "AppData",
            "Roaming",
            "pyRevit",
            "2024",
            "bin",
            "pyrevit.exe",
        ),
        os.path.join(
            os.path.expanduser("~"),
            "AppData",
            "Roaming",
            "pyRevit",
            "2023",
            "bin",
            "pyrevit.exe",
        ),
        os.path.join(
            os.path.expanduser("~"),
            "AppData",
            "Roaming",
            "pyRevit",
            "2022",
            "bin",
            "pyrevit.exe",
        ),
        os.path.join(
            os.environ.get("PROGRAMDATA", ""), "pyRevit", "bin", "pyrevit.exe"
        ),
        os.path.join(
            os.environ.get("PROGRAMFILES", ""), "pyRevit-Master", "bin", "pyrevit.exe"
        ),
        os.path.join(
            os.environ.get("PROGRAMFILES(X86)", ""),
            "pyRevit-Master",
            "bin",
            "pyrevit.exe",
        ),
        os.path.join(
            os.path.expanduser("~"),
            "AppData",
            "Roaming",
            "pyRevit-Master",
            "bin",
            "pyrevit.exe",
        ),
    ]

    out.print_md(":information_source: Searching for pyrevit.exe...")
    for path in possible_paths:
        exists = os.path.exists(path)
        out.print_md(
            "  Checking: {} {}".format(
                path, ":white_check_mark: found" if exists else ":x: not found"
            )
        )
        if exists:
            out.print_md(":white_check_mark: Using: `{}`".format(path))
            return path

    out.print_md(":x: pyrevit.exe not found in standard paths")
    return None


def create_bat_file(object_folder_path, export_time):
    pyrevit_path = get_pyrevit_cli_path()

    if not pyrevit_path:
        return None

    bat_path = os.path.join(object_folder_path, BAT_FILE)

    script_path = os.path.abspath(__file__)
    script_dir = os.path.dirname(script_path)
    revit_year = "2022"

    host_model = read_txt_file(os.path.join(object_folder_path, HOST_MODEL_CONFIG))
    if not host_model:
        host_model = ""
        models_param = ""
        models_option = ""
    else:
        models_param = ' --models="{}"'.format(host_model)
        models_option = '"%MODEL%"'

    bat_lines = [
        "@echo off",
        "setlocal enabledelayedexpansion",
        "",
        'set "PYREVIT={}"'.format(pyrevit_path),
        'set "SCRIPT_DIR={}"'.format(script_dir),
        'set "REVIT_YEAR={}"'.format(revit_year),
        'set "EXPORT_TIME={}"'.format(export_time),
        'set "MODEL={}"'.format(host_model),
        'set "PYREVIT_DAILY_AUTO=1"',
        'set "PYREVIT_DAILY_TIME=%EXPORT_TIME%"',
        "",
        "echo ======================================",
        "echo Daily NWC Export",
        "echo ======================================",
        "echo Date: %date%",
        "echo Time: %time%",
        "echo Script Dir: %SCRIPT_DIR%",
        "echo PyRevit: %PYREVIT%",
        "echo Revit: %REVIT_YEAR%",
        "echo Model: %MODEL%",
        "echo ======================================",
        "echo.",
        "",
        "echo [1] Launching pyRevit with --import...",
        '"%PYREVIT%" run "daily_nwc_script" --revit=%REVIT_YEAR% --import="%SCRIPT_DIR%"{}'.format(
            models_param
        ),
        "",
        "echo.",
        "echo [2] PyRevit exit code: %ERRORLEVEL%",
        "",
        "echo.",
        "if %ERRORLEVEL% EQU 0 (",
        "    echo [3] SUCCESS: PyRevit executed successfully",
        ") else (",
        "    echo [3] ERROR: PyRevit exited with code %ERRORLEVEL%",
        "    echo.",
        "    echo Check:",
        "    echo   - Revit must be OPEN when running",
        "    echo   - Script must exist at path",
        "    echo   - Model file must exist at path (create Host_Model.txt)",
        "    echo   - Logs in: {}\\logs".format(object_folder_path),
        ")",
        "",
        "echo ======================================",
        "echo Completion: %date% %time%",
        "echo ======================================",
        "",
        "pause",
        "",
    ]

    bat_content = "\r\n".join(bat_lines)

    if write_txt_file(bat_path, bat_content):
        out.print_md(":white_check_mark: Bat file created: `{}`".format(bat_path))
        return bat_path
    else:
        return None


def create_scheduled_task(object_folder_path, export_time="00:00"):
    bat_path = create_bat_file(object_folder_path, export_time)
    if not bat_path:
        out.print_md(":x: Failed to create bat file")
        return False

    try:
        import subprocess

        subprocess.call('schtasks /delete /tn "{}" /f'.format(TASK_NAME), shell=True)

        cmd = 'schtasks /create /tn "{}" /tr "{}" /sc daily /st {} /f'.format(
            TASK_NAME, bat_path, export_time
        )

        result = subprocess.call(cmd, shell=True)

        if result == 0:
            out.print_md(
                ":white_check_mark: Scheduled task created at time **{}**".format(
                    export_time
                )
            )
            return True
        else:
            out.print_md(
                ":x: Failed to create scheduled task (code: {})".format(result)
            )
            return False
    except Exception as e:
        out.print_md(":x: Failed to create scheduled task: {}".format(e))
        return False


def check_scheduled_task_exists():
    try:
        import subprocess

        result = subprocess.call(
            'schtasks /query /tn "{}"'.format(TASK_NAME),
            shell=True,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
        )
        return result == 0
    except Exception:
        return False


def show_main_menu(script_root, object_folder_path, export_time, service_running=False):
    options = [
        "Export Now",
        "Start Service" if not service_running else "Stop Service",
        "Configure Object List",
        "Configure Object Folder",
        "Configure Export Time (current: {})".format(export_time),
        "Exit",
    ]

    selected = forms.SelectFromList.show(
        options, title="Daily NWC Export", button_name="Select"
    )

    return selected


def show_export_time_dialog(current_time):
    return forms.ask_for_string(
        default=current_time,
        prompt="Enter export time (format HH:MM):",
        title="Configure Export Time",
    )


def show_summary_dialog(summary):
    if not summary:
        return

    msg = """Total: {}
 Exported: {}
 Skipped (actual): {}
 Errors: {}
 Total time: {}

 Log: {}
""".format(
        summary["total"],
        summary["exported"],
        summary["skipped"],
        summary["errors"],
        summary["total_time"],
        summary["log_path"],
    )

    out.print_md("## SUMMARY")
    out.print_md(msg)

    forms.alert(msg, ok=True, exitscript=False)


def main():
    out.print_md("[DEBUG] Script started")
    out.print_md("[DEBUG] Python version: {}".format(sys.version))
    out.print_md("[DEBUG] Working dir: {}".format(os.getcwd()))

    auto_mode = AUTO_MODE_FLAG in sys.argv or "-a" in sys.argv
    out.print_md("[DEBUG] Auto mode: {}".format(auto_mode))

    export_time = "00:00"
    # Note: export_time will be loaded from config once object_folder_path is known
    for arg in sys.argv:
        if arg.startswith("--time="):
            export_time = arg.split("=")[1]
            break

    app = __revit__.Application
    revit = __revit__

    if not auto_mode:
        out.set_width(900)

    script_root = get_script_root()

    if auto_mode:
        object_folder_path = read_object_folder_path(script_root)

        out.print_md("[AUTO] Object folder: `{}`".format(object_folder_path))

        if not object_folder_path:
            out.print_md(
                ":x: [AUTO] Object folder not configured. Run script in manual mode for setup."
            )
            import sys as sys_module

            sys_module.stdin.read(1)
            return

        export_list = read_export_list(object_folder_path)

        out.print_md(
            "[AUTO] Objects for export ({}): {}".format(len(export_list), export_list)
        )

        if not export_list:
            out.print_md(
                ":x: [AUTO] Export list not configured. Run script in manual mode for setup."
            )
            return

        if check_lock_file(object_folder_path):
            out.print_md(
                ":x: [AUTO] Script is already running. Duplicate run not possible."
            )
            return

        create_lock_file(object_folder_path)

        try:
            result = export_all_objects(
                object_folder_path,
                export_list,
                auto_mode=True,
                export_enabled=EXPORT_ENABLED,
            )

            if result and result["total"] > 0:
                out.print_md(
                    "[AUTO] Export completed. Log: `{}`".format(
                        result.get("log_path", "Unknown")
                    )
                )
        finally:
            remove_lock_file(object_folder_path)
    else:
        object_folder_path = read_object_folder_path(script_root)

        if not object_folder_path:
            out.print_md("[INFO] Object folder not configured")
            object_folder_path = select_object_folder_ui()

            if object_folder_path:
                save_object_folder_path(object_folder_path, script_root)
                out.print_md(
                    ":white_check_mark: Object folder saved: `{}`".format(
                        object_folder_path
                    )
                )
            else:
                script.exit()

        export_list = read_export_list(object_folder_path)

        if not export_list:
            out.print_md("[INFO] Export list not configured")
            export_list = select_objects_ui(object_folder_path)

        if export_list:
            out.print_md(
                ":white_check_mark: Selected objects: {}".format(len(export_list))
            )
        else:
            script.exit()

        # Load export_time from config if not overridden by command line
        if export_time == "00:00":
            export_time = read_export_time(object_folder_path)
            out.print_md(
                "[INFO] Export time loaded from config: {}".format(export_time)
            )

        while True:
            service_manager = get_service_manager()
            service_running = service_manager is not None and service_manager.running

            selected = show_main_menu(
                script_root, object_folder_path, export_time, service_running
            )

            if selected is None or selected == "Exit":
                break

            if selected == "Export Now":
                if check_lock_file(object_folder_path):
                    out.print_md(":x: Script is already running.")
                    continue

                create_lock_file(object_folder_path)

                try:
                    summary = export_all_objects(
                        object_folder_path,
                        export_list,
                        auto_mode=False,
                        export_enabled=EXPORT_ENABLED,
                    )

                    if summary and summary["total"] > 0:
                        show_summary_dialog(summary)
                finally:
                    remove_lock_file(object_folder_path)

            elif selected == "Start Service":
                if service_running:
                    out.print_md(":x: Service is already running.")
                    continue

                if check_lock_file(object_folder_path):
                    out.print_md(":x: Script is already running.")
                    continue

                log_path = init_logger(object_folder_path)
                if not log_path:
                    out.print_md(":x: Failed to initialize logger.")
                    continue

                new_manager = ServiceManager(
                    object_folder_path,
                    export_list,
                    log_path,
                    app,
                    revit,
                    EXPORT_ENABLED,
                    export_time,
                )
                if new_manager.start():
                    if set_service_manager(new_manager):
                        mode_msg = (
                            "REAL EXPORT" if EXPORT_ENABLED else "DRY-RUN (no export)"
                        )
                        out.print_md(
                            ":white_check_mark: Service started (time: {}, mode: {})".format(
                                export_time, mode_msg
                            )
                        )
                        out.print_md("Log: `{}`".format(log_path))
                        out.print_md("___")
                        out.print_md(
                            ":information_source: Service is now running in background."
                        )
                        out.print_md(
                            ":information_source: To check status, run this script again."
                        )
                        out.print_md(
                            ":information_source: To stop service, select 'Stop Service'."
                        )
                        break
                    else:
                        new_manager.stop()
                        out.print_md(
                            ":x: Failed to save service manager to persistent engine."
                        )
                else:
                    out.print_md(":x: Failed to start service.")

            elif selected == "Stop Service":
                if not service_running:
                    out.print_md(":x: Service is not running.")
                    continue

                if service_manager:
                    status = service_manager.get_status()
                    service_manager.stop()
                    if status["last_error"]:
                        out.print_md(
                            ":warning: Last service error: {} ({})".format(
                                status["last_error"],
                                status["last_error_time"].strftime("%Y-%m-%d %H:%M:%S")
                                if status["last_error_time"]
                                else "Unknown",
                            )
                        )
                clear_service_manager()
                out.print_md(":white_check_mark: Service stopped.")

            elif selected == "Configure Object List":
                export_list = select_objects_ui(object_folder_path)

                if export_list:
                    out.print_md(
                        ":white_check_mark: Selected objects: {}".format(
                            len(export_list)
                        )
                    )

            elif selected == "Configure Object Folder":
                new_path = select_object_folder_ui()

                if new_path:
                    object_folder_path = new_path
                    save_object_folder_path(object_folder_path, script_root)
                    out.print_md(
                        ":white_check_mark: Object folder saved: `{}`".format(
                            object_folder_path
                        )
                    )

            elif selected.startswith("Configure Export Time"):
                new_time = show_export_time_dialog(export_time)

                if new_time:
                    export_time = new_time
                    # Save to config file
                    if save_export_time(export_time, object_folder_path):
                        out.print_md(
                            ":white_check_mark: Export time saved to config: {}".format(
                                export_time
                            )
                        )
                    else:
                        out.print_md(
                            ":warning: Export time set to {} but failed to save to config".format(
                                export_time
                            )
                        )

                    if service_running:
                        out.print_md(
                            ":warning: Service will use new time starting from next day or after restart."
                        )


if __name__ == "__main__":
    main()
