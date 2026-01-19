# -*- coding: utf-8 -*-
"""
Ежедневный экспорт NWC — автоматический экспорт RVT в NWC с проверкой актуальности.
"""

import os
import sys
import datetime
import glob
from pyrevit import script, coreutils, forms

# Локальные модули
import openbg
import closebg
from nwc_export_utils import export_rvt_to_nwc_full

# Константы
OBJECT_FOLDER_CONFIG = "Object_folder_path.txt"
DAILY_EXPORT_LIST = "Ежедневная выгрузка.txt"
LOG_FOLDER = "logs"
AUTO_MODE_FLAG = "--auto"
LOCK_FILE = ".export_lock"
TASK_NAME = "Ежедневный экспорт NWC"
BAT_FILE = "run_daily_export.bat"
LOG_RETENTION_DAYS = 7

out = script.get_output()
out.close_others(all_open_outputs=True)

# ========== helpers ==========


def _norm(p):
    """Нормализация пути."""
    return os.path.normpath(os.path.abspath(p)) if p else p


def _module_dir():
    """Получить директорию текущего модуля."""
    try:
        return os.path.dirname(os.path.abspath(__file__))
    except Exception:
        return os.getcwd()


def _find_scripts_root(start_dir):
    """Поднимаемся вверх и ищем папку 'Scripts'."""
    cur = _norm(start_dir)
    for _ in range(0, 8):
        if os.path.basename(cur).lower() == "scripts":
            return cur
        parent = os.path.dirname(cur)
        if not parent or parent == cur:
            break
        cur = parent
    # типовая структура: lib -> WWBIM.extension -> Scripts
    return _norm(os.path.join(start_dir, os.pardir, os.pardir))


# ========== чтение/запись файлов ==========


def read_txt_file(filepath):
    """Читать txt файл и вернуть содержимое (без пробелов/переносов)."""
    if not os.path.exists(filepath):
        return None
    try:
        # бинарный режим + универсальная декодировка (IronPython/CPython)
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
    return None


def write_txt_file(filepath, content):
    """Записать содержимое в txt файл (UTF-8)."""
    try:
        with open(filepath, "w", encoding="utf-8") as f:
            f.write(content)
        return True
    except Exception as e:
        out.print_md(":x: Ошибка записи в файл `{}`: {}".format(filepath, e))
        return False


def get_script_root():
    """Получить корневую папку скриптов."""
    module_dir = _module_dir()
    _env_root = os.environ.get("WW_SCRIPTS_ROOT")
    if _env_root and os.path.isdir(_env_root):
        return _norm(_env_root)
    return _find_scripts_root(module_dir)


def select_object_folder_ui():
    """UI: диалог выбора папки Object."""
    folder = forms.pick_folder(title="Выберите папку Object")
    if folder:
        return _norm(folder)
    return None


def read_object_folder_path(script_root):
    """Читать сохранённый путь к папке Object."""
    config_path = os.path.join(script_root, OBJECT_FOLDER_CONFIG)
    return read_txt_file(config_path)


def save_object_folder_path(object_folder_path, script_root):
    """Сохранить путь к папке Object."""
    config_path = os.path.join(script_root, OBJECT_FOLDER_CONFIG)
    return write_txt_file(config_path, object_folder_path)


def read_object_config(object_name, object_folder_path):
    """
    Читать конфигурацию объекта:
    - object_name.txt → путь к RVT
    - object_name_NWC.txt → папка для NWC
    """
    result = {
        "rvt_path": None,
        "nwc_folder": None,
        "rvt_exists": False,
        "nwc_folder_exists": False,
    }

    # Читаем путь к RVT
    rvt_file = os.path.join(object_folder_path, object_name + ".txt")
    rvt_path = read_txt_file(rvt_file)
    if rvt_path:
        result["rvt_path"] = rvt_path
        result["rvt_exists"] = os.path.exists(rvt_path)

    # Читаем папку для NWC
    nwc_file = os.path.join(object_folder_path, object_name + "_NWC.txt")
    nwc_folder = read_txt_file(nwc_file)
    if nwc_folder:
        result["nwc_folder"] = _norm(nwc_folder)
        result["nwc_folder_exists"] = os.path.isdir(result["nwc_folder"])

    return result


def get_available_objects(object_folder_path):
    """Получить список доступных объектов из папки Object."""
    if not os.path.isdir(object_folder_path):
        return []

    objects = []
    try:
        files = os.listdir(object_folder_path)
        for file in files:
            # берём *.txt без _NWC суффикса
            if file.lower().endswith(".txt") and not file.lower().endswith("_nwc.txt"):
                # Исключаем системные файлы
                if file not in [
                    OBJECT_FOLDER_CONFIG,
                    DAILY_EXPORT_LIST,
                    BAT_FILE,
                    LOCK_FILE,
                ]:
                    objects.append(file[:-4])  # убираем .txt
    except Exception as e:
        out.print_md(":x: Ошибка чтения папки `{}`: {}".format(object_folder_path, e))

    return sorted(objects)


# ========== управление списком экспорта ==========


def read_export_list(object_folder_path):
    """Читать список объектов для ежедневной выгрузки."""
    list_path = os.path.join(object_folder_path, DAILY_EXPORT_LIST)
    if not os.path.exists(list_path):
        return []

    objects = []
    try:
        with open(list_path, "r", encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if line:
                    # Проверяем, может быть список через запятую
                    if "," in line:
                        parts = line.split(",")
                        objects.extend([p.strip() for p in parts if p.strip()])
                    else:
                        objects.append(line)
    except Exception as e:
        out.print_md(":x: Ошибка чтения списка `{}`: {}".format(list_path, e))

    return objects


def save_export_list(objects, object_folder_path):
    """Сохранить список объектов в Ежедневная выгрузка.txt."""
    list_path = os.path.join(object_folder_path, DAILY_EXPORT_LIST)
    content = "\n".join(objects)
    return write_txt_file(list_path, content)


def select_objects_ui(object_folder_path):
    """UI: диалог выбора объектов для ежедневной выгрузки."""
    available = get_available_objects(object_folder_path)
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


# ========== проверка актуальности ==========


def get_file_modification_date(filepath):
    """Получить дату изменения файла."""
    if not os.path.exists(filepath):
        return None
    try:
        timestamp = os.path.getmtime(filepath)
        return datetime.datetime.fromtimestamp(timestamp)
    except Exception:
        return None


def check_need_export(rvt_path, nwc_folder, object_name):
    """
    Проверить, нужен ли экспорт NWC.
    """
    result = {
        "need_export": True,
        "reason": "",
        "rvt_date": None,
        "nwc_date": None,
        "nwc_exists": False,
    }

    # Проверяем RVT
    rvt_date = get_file_modification_date(rvt_path)
    if rvt_date:
        result["rvt_date"] = rvt_date

    # Проверяем NWC
    nwc_path = os.path.join(nwc_folder, object_name + ".nwc")
    nwc_exists = os.path.exists(nwc_path)
    result["nwc_exists"] = nwc_exists

    if nwc_exists:
        nwc_date = get_file_modification_date(nwc_path)
        if nwc_date:
            result["nwc_date"] = nwc_date
            # Сравниваем даты
            if rvt_date and rvt_date <= nwc_date:
                result["need_export"] = False
                result["reason"] = "NWC актуален"
            elif rvt_date and rvt_date > nwc_date:
                result["need_export"] = True
                result["reason"] = "RVT обновлён"
            else:
                result["need_export"] = True
                result["reason"] = "Не удалось определить даты"
        else:
            result["need_export"] = True
            result["reason"] = "Не удалось определить дату NWC"
    else:
        result["need_export"] = True
        result["reason"] = "NWC не существует"

    return result


# ========== логирование ==========


def init_logger(object_folder_path):
    """Инициализировать логгер для текущей сессии."""
    log_dir = os.path.join(object_folder_path, LOG_FOLDER)
    if not os.path.exists(log_dir):
        try:
            os.makedirs(log_dir)
        except Exception as e:
            out.print_md(":x: Ошибка создания папки логов `{}`: {}".format(log_dir, e))
            return None

    # Удаляем старые логи
    cleanup_old_logs(log_dir)

    # Создаём файл лога с датой
    today = datetime.datetime.now().strftime("%Y-%m-%d")
    log_path = os.path.join(log_dir, "export_log_{}.txt".format(today))

    try:
        with open(log_path, "a", encoding="utf-8") as f:
            timestamp = datetime.datetime.now().strftime("%H:%M:%S")
            f.write("=" * 40 + "\n")
            f.write(
                "Ежедневный экспорт NWC - {} {}\n".format(
                    datetime.datetime.now().strftime("%Y-%m-%d"), timestamp
                )
            )
            f.write("=" * 40 + "\n\n")
        return log_path
    except Exception as e:
        out.print_md(":x: Ошибка создания файла лога `{}`: {}".format(log_path, e))
        return None


def cleanup_old_logs(log_dir):
    """Удалить логи старше LOG_RETENTION_DAYS дней."""
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
    """Записать сообщение в лог файл."""
    if not log_path:
        return
    try:
        with open(log_path, "a", encoding="utf-8") as f:
            timestamp = datetime.datetime.now().strftime("[%H:%M:%S]")
            f.write("{} {}\n".format(timestamp, message))
    except Exception:
        pass


def log_export_start(log_path, object_name, rvt_path):
    """Логировать начало экспорта объекта."""
    msg = "Объект: {}".format(object_name)
    log_message(log_path, msg)
    msg = "  - RVT: {}".format(rvt_path)
    log_message(log_path, msg)


def log_export_success(log_path, object_name, elapsed_time, file_size_mb):
    """Логировать успешный экспорт."""
    msg = "  - Экспорт: УСПЕХ (время: {}, размер: {:.1f} MB)".format(
        elapsed_time, file_size_mb
    )
    log_message(log_path, msg)


def log_export_error(log_path, object_name, error_message):
    """Логировать ошибку экспорта."""
    msg = "  - Ошибка: {}".format(error_message)
    log_message(log_path, msg)


def log_export_skipped(log_path, object_name, reason, rvt_date=None, nwc_date=None):
    """Логировать пропуск экспорта."""
    msg = "  - NWC: {}".format(reason)
    log_message(log_path, msg)
    if rvt_date:
        msg = "  - RVT дата: {}".format(rvt_date.strftime("%Y-%m-%d %H:%M:%S"))
        log_message(log_path, msg)
    if nwc_date:
        msg = "  - NWC дата: {}".format(nwc_date.strftime("%Y-%m-%d %H:%M:%S"))
        log_message(log_path, msg)


def log_summary(log_path, total, exported, skipped, errors, total_time):
    """Логировать итоговую сводку."""
    log_message(log_path, "")
    log_message(log_path, "=" * 40)
    msg = "ИТОГО: Всего: {}, Экспортировано: {}, Пропущено: {}, Ошибок: {}".format(
        total, exported, skipped, errors
    )
    log_message(log_path, msg)
    msg = "Время выполнения: {}".format(total_time)
    log_message(log_path, msg)
    log_message(log_path, "=" * 40)


# ========== экспорт ==========


def export_single_object(
    object_name, rvt_path, nwc_folder, app, revit, log_path, auto_mode=False
):
    """Экспортировать один объект."""
    log_export_start(log_path, object_name, rvt_path)

    # Проверяем актуальность
    check_result = check_need_export(rvt_path, nwc_folder, object_name)

    if not check_result["need_export"]:
        log_export_skipped(
            log_path,
            object_name,
            check_result["reason"],
            check_result.get("rvt_date"),
            check_result.get("nwc_date"),
        )
        if not auto_mode:
            out.print_md(
                ":white_check_mark: {}: **{}**".format(
                    object_name, check_result["reason"]
                )
            )
        return True, False, None

    # Экспортируем
    if not auto_mode:
        out.print_md(":hourglass: Экспортирую: **{}**".format(object_name))

    export_result = export_rvt_to_nwc_full(
        rvt_path, nwc_folder, object_name, app, revit
    )

    if export_result["success"]:
        file_size = export_result.get("file_size_mb", 0)
        elapsed = export_result.get("time_export", "0s")
        log_export_success(log_path, object_name, elapsed, file_size)

        if not auto_mode:
            msg = ":white_check_mark: {}: УСПЕХ ({} MB)".format(
                object_name, round(file_size, 2)
            )
            out.print_md(msg)

        # Логируем предупреждения/ошибки открытия
        if export_result.get("warnings_count", 0) > 0:
            msg = "  :warning: Предупреждений при открытии: {}".format(
                export_result["warnings_count"]
            )
            log_message(log_path, msg)
            if export_result.get("warnings"):
                for w in export_result["warnings"][:3]:
                    log_message(log_path, "    - {}".format(w))

        if export_result.get("errors_count", 0) > 0:
            msg = "  :x: Ошибок при открытии: {}".format(export_result["errors_count"])
            log_message(log_path, msg)
            if export_result.get("errors"):
                for e in export_result["errors"][:3]:
                    log_message(log_path, "    - {}".format(e))

        return True, True, None
    else:
        error = export_result.get("error", "Неизвестная ошибка")
        log_export_error(log_path, object_name, error)

        if not auto_mode:
            out.print_md(":x: {}: **{}**".format(object_name, error))

        # Продолжаем при ошибке (как указано в требованиях)
        return False, False, error


def export_all_objects(object_folder_path, export_list, app, revit, auto_mode=False):
    """
    Экспортировать все объекты из списка.
    """
    log_path = init_logger(object_folder_path)
    if not log_path:
        return None

    t_all = coreutils.Timer()

    total = 0
    exported = 0
    skipped = 0
    errors = 0

    if not auto_mode:
        out.print_md("## ЭКСПОРТ NWC ({})".format(len(export_list)))
        out.print_md("Папка экспорта: **{}**".format(object_folder_path))
        out.print_md("___")
        out.update_progress(0, len(export_list))

    for i, object_name in enumerate(export_list):
        total += 1

        # Читаем конфигурацию объекта
        config = read_object_config(object_name, object_folder_path)

        if not config["rvt_exists"]:
            error = "Файл RVT не найден: {}".format(config["rvt_path"])
            log_export_error(log_path, object_name, error)
            if not auto_mode:
                out.print_md(":x: {}: **{}**".format(object_name, error))
            errors += 1
            if not auto_mode:
                out.update_progress(i + 1, len(export_list))
            continue

        if not config["nwc_folder_exists"]:
            error = "Папка NWC не существует: {}".format(config["nwc_folder"])
            log_export_error(log_path, object_name, error)
            if not auto_mode:
                out.print_md(":x: {}: **{}**".format(object_name, error))
            errors += 1
            if not auto_mode:
                out.update_progress(i + 1, len(export_list))
            continue

        # Экспортируем
        success, was_exported, error = export_single_object(
            object_name,
            config["rvt_path"],
            config["nwc_folder"],
            app,
            revit,
            log_path,
            auto_mode,
        )

        if success:
            if was_exported:
                exported += 1
            else:
                skipped += 1
        else:
            errors += 1

        if not auto_mode:
            out.update_progress(i + 1, len(export_list))

    all_s = str(datetime.timedelta(seconds=int(t_all.get_time())))

    # Логируем сводку
    log_summary(log_path, total, exported, skipped, errors, all_s)

    if not auto_mode:
        out.print_md("___")
        out.print_md(
            "**Готово. Всего: {}, Экспортировано: {}, Пропущено: {}, Ошибок: {}**".format(
                total, exported, skipped, errors
            )
        )
        out.print_md("**Время выполнения: {}**".format(all_s))
        out.print_md("**Лог: `{}`**".format(log_path))

    return {
        "total": total,
        "exported": exported,
        "skipped": skipped,
        "errors": errors,
        "total_time": all_s,
        "log_path": log_path,
    }


# ========== scheduled task ==========


def get_pyrevit_cli_path():
    """Получить путь к pyrevit.exe CLI."""
    import os as os_module

    username = os_module.environ.get("USERNAME", "")
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
    ]

    for path in possible_paths:
        if os.path.exists(path):
            return path

    return None


def create_bat_file(object_folder_path, export_time):
    """Создать bat-файл для запуска pyRevit."""
    pyrevit_path = get_pyrevit_cli_path()
    if not pyrevit_path:
        out.print_md(":x: Не найден pyrevit.exe")
        return None

    bat_path = os.path.join(object_folder_path, BAT_FILE)

    # Создаём bat-файл
    content = """@echo off
echo Starting daily NWC export at %date% %time%
"{}" run "WW.BIM.tab BIM.panel Экспорт.stack Ежедневный экспорт NWC.pushbutton" --auto --time "{}"
echo Export completed at %date% %time%
""".format(pyrevit_path, export_time)

    if write_txt_file(bat_path, content):
        return bat_path
    else:
        return None


def create_scheduled_task(object_folder_path, export_time="00:00"):
    """Создать Windows Scheduled Task для ежедневного запуска."""
    # Создаём bat-файл
    bat_path = create_bat_file(object_folder_path, export_time)
    if not bat_path:
        out.print_md(":x: Не удалось создать bat-файл")
        return False

    # Создаём scheduled task
    try:
        import subprocess

        # Удаляем старую задачу, если есть
        subprocess.call('schtasks /delete /tn "{}" /f'.format(TASK_NAME), shell=True)

        # Создаём новую задачу
        cmd = 'schtasks /create /tn "{}" /tr "{}" /sc daily /st {} /f'.format(
            TASK_NAME, bat_path, export_time
        )

        result = subprocess.call(cmd, shell=True)

        if result == 0:
            out.print_md(
                ":white_check_mark: Scheduled task создан на время **{}**".format(
                    export_time
                )
            )
            out.print_md("Bat-файл: `{}`".format(bat_path))
            return True
        else:
            out.print_md(":x: Ошибка создания scheduled task (код: {})".format(result))
            return False
    except Exception as e:
        out.print_md(":x: Ошибка создания scheduled task: {}".format(e))
        return False


def delete_scheduled_task():
    """Удалить Windows Scheduled Task."""
    try:
        import subprocess

        result = subprocess.call(
            'schtasks /delete /tn "{}" /f'.format(TASK_NAME), shell=True
        )
        if result == 0:
            out.print_md(":white_check_mark: Scheduled task удалён")
            return True
        else:
            out.print_md(":x: Ошибка удаления scheduled task (код: {})".format(result))
            return False
    except Exception as e:
        out.print_md(":x: Ошибка удаления scheduled task: {}".format(e))
        return False


def check_scheduled_task_exists():
    """Проверить существование scheduled task."""
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


# ========== UI ==========


def show_main_menu(script_root, object_folder_path, export_time):
    """Показать главное меню."""
    task_exists = check_scheduled_task_exists()

    options = [
        "Экспортировать сейчас",
        "Настроить список объектов",
        "Настроить папку Object",
        "Настроить автозапуск (время: {})".format(export_time),
        "Удалить автозапуск" if task_exists else "Удалить автозапуск (не настроен)",
        "Выход",
    ]

    selected = forms.SelectFromList.show(
        options, title="Ежедневный экспорт NWC", button_name="Выбрать"
    )

    return selected


def show_export_time_dialog(current_time):
    """Показать диалог настройки времени экспорта."""
    return forms.ask_for_string(
        default=current_time,
        prompt="Укажите время экспорта (формат ЧЧ:ММ):",
        title="Настройка времени экспорта",
    )


def show_summary_dialog(summary):
    """Показать диалоговое окно с итогами экспорта."""
    if not summary:
        return

    msg = """Всего объектов: {}
Экспортировано: {}
Пропущено (актуально): {}
Ошибок: {}
Время выполнения: {}

Лог: {}
""".format(
        summary["total"],
        summary["exported"],
        summary["skipped"],
        summary["errors"],
        summary["total_time"],
        summary["log_path"],
    )

    out.print_md("## ИТОГИ")
    out.print_md(msg)

    forms.alert(msg, ok=True, exitscript=False)


# ========== проверка lock-файла ==========


def check_lock_file(object_folder_path):
    """Проверить наличие lock-файла."""
    lock_path = os.path.join(object_folder_path, LOCK_FILE)
    if os.path.exists(lock_path):
        # Проверяем, не старый ли lock (более 24 часов)
        try:
            lock_time = datetime.datetime.fromtimestamp(os.path.getmtime(lock_path))
            if datetime.datetime.now() - lock_time > datetime.timedelta(hours=24):
                # Lock старый, удаляем
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
    """Создать lock-файл."""
    lock_path = os.path.join(object_folder_path, LOCK_FILE)
    try:
        with open(lock_path, "w") as f:
            f.write(str(datetime.datetime.now()))
        return True
    except Exception:
        return False


def remove_lock_file(object_folder_path):
    """Удалить lock-файл."""
    lock_path = os.path.join(object_folder_path, LOCK_FILE)
    try:
        if os.path.exists(lock_path):
            os.remove(lock_path)
    except Exception:
        pass


# ========== main ==========


def main():
    """Точка входа скрипта."""
    # Проверяем режим (автоматический/ручной)
    auto_mode = AUTO_MODE_FLAG in sys.argv or "-a" in sys.argv

    # Получаем время экспорта из аргументов
    export_time = "00:00"
    for arg in sys.argv:
        if arg.startswith("--time="):
            export_time = arg.split("=")[1]
            break

    app = __revit__.Application
    revit = __revit__

    # Получаем скрипт root
    script_root = get_script_root()

    if auto_mode:
        # Автоматический режим
        object_folder_path = read_object_folder_path(script_root)

        if not object_folder_path:
            out.print_md(
                ":x: Не настроена папка Object. Запустите скрипт в ручном режиме для настройки."
            )
            return

        export_list = read_export_list(object_folder_path)

        if not export_list:
            out.print_md(
                ":x: Не настроен список объектов. Запустите скрипт в ручном режиме для настройки."
            )
            return

        # Проверяем lock-файл
        if check_lock_file(object_folder_path):
            out.print_md(":x: Скрипт уже выполняется. Повторный запуск невозможен.")
            return

        # Создаём lock-файл
        create_lock_file(object_folder_path)

        try:
            # Выполняем экспорт
            export_all_objects(
                object_folder_path, export_list, app, revit, auto_mode=True
            )
        finally:
            # Удаляем lock-файл
            remove_lock_file(object_folder_path)
    else:
        # Ручной режим
        out.set_width(900)

        # Проверяем папку Object
        object_folder_path = read_object_folder_path(script_root)

        if not object_folder_path:
            out.print_md(":information_source: Папка Object не настроена")
            object_folder_path = select_object_folder_ui()

            if object_folder_path:
                save_object_folder_path(object_folder_path, script_root)
                out.print_md(
                    ":white_check_mark: Папка Object сохранена: `{}`".format(
                        object_folder_path
                    )
                )
            else:
                script.exit()

        # Проверяем список объектов
        export_list = read_export_list(object_folder_path)

        if not export_list:
            out.print_md(":information_source: Список объектов не настроен")
            export_list = select_objects_ui(object_folder_path)

            if export_list:
                out.print_md(
                    ":white_check_mark: Выбрано объектов: {}".format(len(export_list))
                )
            else:
                script.exit()

        # Главное меню
        while True:
            selected = show_main_menu(script_root, object_folder_path, export_time)

            if selected is None or selected == "Выход":
                break

            if selected == "Экспортировать сейчас":
                # Проверяем lock-файл
                if check_lock_file(object_folder_path):
                    out.print_md(
                        ":x: Скрипт уже выполняется. Повторный запуск невозможен."
                    )
                    continue

                # Создаём lock-файл
                create_lock_file(object_folder_path)

                try:
                    summary = export_all_objects(
                        object_folder_path, export_list, app, revit, auto_mode=False
                    )

                    if summary and summary["total"] > 0:
                        # Показываем итоги
                        show_summary_dialog(summary)
                finally:
                    # Удаляем lock-файл
                    remove_lock_file(object_folder_path)

            elif selected == "Настроить список объектов":
                export_list = select_objects_ui(object_folder_path)

                if export_list:
                    out.print_md(
                        ":white_check_mark: Выбрано объектов: {}".format(
                            len(export_list)
                        )
                    )

            elif selected == "Настроить папку Object":
                new_path = select_object_folder_ui()

                if new_path:
                    object_folder_path = new_path
                    save_object_folder_path(object_folder_path, script_root)
                    out.print_md(
                        ":white_check_mark: Папка Object сохранена: `{}`".format(
                            object_folder_path
                        )
                    )

            elif selected.startswith("Настроить автозапуск"):
                # Показываем диалог настройки времени
                new_time = show_export_time_dialog(export_time)

                if new_time:
                    export_time = new_time
                    # Создаём scheduled task
                    if create_scheduled_task(object_folder_path, export_time):
                        out.print_md(
                            ":white_check_mark: Автозапуск настроен на время {}".format(
                                export_time
                            )
                        )

            elif selected.startswith("Удалить автозапуск"):
                if delete_scheduled_task():
                    out.print_md(":white_check_mark: Автозапуск удалён")


if __name__ == "__main__":
    main()
