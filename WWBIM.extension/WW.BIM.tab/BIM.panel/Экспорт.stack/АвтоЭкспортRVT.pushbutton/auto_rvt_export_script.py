# -*- coding: utf-8 -*-
__title__ = "Экспорт RVT"
__author__ = "vlad / you"
__doc__ = "Экспорт .rvt без интерфейса. Читает объекты из Y:\\BIM\\Scripts\\Objects\\AUTO_RVT\\AUTO_RVT_Objects.txt, пути из <object>.txt, папку экспорта из <object>_RVT.txt. Открытие РН через openbg: все, кроме начинающихся с '00_' и содержащих 'Link'/'Связь'. Режим detach: preserve. Совместимость Revit 2022/2023."

import os
import sys
import datetime
import codecs
import re
from pyrevit import script, coreutils

lib_path = r"D:\Share\OneDrive - SIYA Project\03_YandexDisk_Vlad\BIM\Scripts\WWBIM.extension\lib"
if lib_path not in sys.path:
    sys.path.insert(0, lib_path)

from Autodesk.Revit.DB import (
    ModelPathUtils,
    SaveAsOptions,
    WorksharingSaveAsOptions,
    OpenOptions,
    DetachFromCentralOption,
)

import openbg
import closebg

DETACH_MODE = "preserve"
COMPACT_ON_SAVE = True
OVERWRITE_SAME = True

OBJECTS_FILE = r"Y:\BIM\Scripts\Objects\AUTO_RVT\AUTO_RVT_Objects.txt"
OBJECTS_BASE_DIR = r"Y:\BIM\Scripts\Objects"

LOG_DIR = r"Y:\BIM\Scripts\Objects\AUTO_RVT\logs"

if not os.path.exists(LOG_DIR):
    try:
        os.makedirs(LOG_DIR)
    except:
        pass

today = datetime.datetime.now().strftime("%Y-%m-%d")
log_file = os.path.join(LOG_DIR, "export_log_{}.txt".format(today))
service_log_file = os.path.join(LOG_DIR, "service_{}.txt".format(today))


def log(msg):
    with codecs.open(log_file, "a", encoding="utf-8") as f:
        ts = datetime.datetime.now().strftime("[%H:%M:%S]")
        f.write("{} {}\n".format(ts, msg))
    try:
        print(msg)
    except:
        pass


def service_log(msg):
    try:
        ts = datetime.datetime.now().strftime("[%Y-%m-%d %H:%M:%S]")
        log_line = "[auto_rvt_export] {} {}".format(ts, msg)
        with codecs.open(service_log_file, "a", encoding="utf-8") as f:
            f.write("{}\n".format(log_line))
        print(log_line)
    except:
        pass


separator = "=" * 40
with codecs.open(log_file, "a", encoding="utf-8") as f:
    f.write("{}\n".format(separator))
    f.write(
        "RVT Export - {} {}\n".format(
            today, datetime.datetime.now().strftime("%H:%M:%S")
        )
    )
    f.write("{}\n\n".format(separator))

log("File: {}".format(OBJECTS_FILE))
log("Log: {}".format(log_file))
service_log("Export process started")


class DialogSuppressor:
    def __init__(self, log_func):
        self.log = log_func
        self.total_dialogs = 0
        self.dialogs_list = []

        self.suppress_messages = [
            "не найдена подходящая геометрия",
            "элементов для геометрии не найдено",
            "не найдена геометрия для экспорта",
            "no suitable geometry",
            "no appropriate geometry",
            "no elements for geometry",
            "no geometry for export",
            "требуется просмотр координации",
            "для экземпляра связи требуется просмотр координации",
            "coordination review",
            "coordination review required",
            "coordination review is required",
            "марка помещение вне элемента помещение",
            "room tag is outside of its room",
            "опорных элементов размеров",
            "dimension reference",
        ]

        self.suppress_dialog_ids = [
            "docwarndialog",
            "coordination",
            "review",
        ]

    def should_suppress(self, message, dialog_id):
        if not message and not dialog_id:
            return False

        message_lower = message.lower() if message else ""
        dialog_id_lower = dialog_id.lower() if dialog_id else ""

        for msg in self.suppress_messages:
            if msg.lower() in message_lower:
                return True

        for did in self.suppress_dialog_ids:
            if did.lower() in dialog_id_lower:
                return True

        return False

    def record_dialog(self, message, dialog_id):
        self.total_dialogs += 1
        dialog_info = {"message": message, "dialog_id": dialog_id}
        self.dialogs_list.append(dialog_info)

        if self.should_suppress(message, dialog_id):
            self.log(
                "  Dialog suppressed: ID={}, Message={}".format(
                    dialog_id or "None", message[:100] if message else "None"
                )
            )
        else:
            self.log(
                "  Dialog shown: ID={}, Message={}".format(
                    dialog_id or "None", message[:100] if message else "None"
                )
            )

    def get_summary(self):
        return {"total_dialogs": self.total_dialogs, "dialogs": self.dialogs_list}


_dialog_suppressor = None


def on_dialog_box_showing(sender, args):
    if _dialog_suppressor is None:
        return

    try:
        dialog_id = str(args.DialogId) if args.DialogId else ""

        try:
            message = getattr(args, "Message", "")
            _dialog_suppressor.record_dialog(message, dialog_id)

            if _dialog_suppressor.should_suppress(message, dialog_id):
                from Autodesk.Revit.UI import TaskDialogResult

                args.OverrideResult(int(TaskDialogResult.Ok))
        except:
            pass

        try:
            message = getattr(args, "Message", "")
            _dialog_suppressor.record_dialog(message, dialog_id)

            if _dialog_suppressor.should_suppress(message, dialog_id):
                args.OverrideResult(1)
        except:
            pass
    except:
        pass


try:
    dialog_suppressor_instance = DialogSuppressor(service_log)
    _dialog_suppressor = dialog_suppressor_instance

    try:
        __revit__.Application.DialogBoxShowing += on_dialog_box_showing
    except:
        pass
except:
    pass


def read_object_names():
    if not os.path.exists(OBJECTS_FILE):
        log("ERROR: File not found: {}".format(OBJECTS_FILE))
        return []
    try:
        with codecs.open(OBJECTS_FILE, "r", encoding="utf-8") as f:
            names = [line.strip() for line in f if line.strip()]
        if not names:
            log("WARNING: File is empty: {}".format(OBJECTS_FILE))
        return names
    except Exception as e:
        log("ERROR: Cannot read file: {}".format(e))
        return []


def read_model_paths(object_name):
    log("  DEBUG: OBJECTS_BASE_DIR = {}".format(OBJECTS_BASE_DIR))
    log("  DEBUG: object_name = {}".format(object_name))
    paths_file = os.path.join(OBJECTS_BASE_DIR, object_name + ".txt")
    log("  DEBUG: paths_file = {}".format(paths_file))
    if not os.path.exists(paths_file):
        log("WARNING: File not found: {}".format(paths_file))
        return []
    try:
        with codecs.open(paths_file, "r", encoding="utf-8") as f:
            paths = [line.strip() for line in f if line.strip()]
        if not paths:
            log("WARNING: File is empty: {}".format(paths_file))
        return paths
    except Exception as e:
        log("ERROR: Cannot read file {}: {}".format(paths_file, e))
        return []


def read_export_folder(object_name):
    folder_file = os.path.join(OBJECTS_BASE_DIR, object_name + "_RVT.txt")
    if not os.path.exists(folder_file):
        log("WARNING: Folder file not found: {}".format(folder_file))
        return default_export_root()
    try:
        with codecs.open(folder_file, "r", encoding="utf-8") as f:
            folder = f.read().strip()
        if not folder:
            log("WARNING: Folder file is empty: {}".format(folder_file))
            return default_export_root()
        if not os.path.exists(folder):
            try:
                os.makedirs(folder)
            except Exception as e:
                log("ERROR: Cannot create folder {}: {}".format(folder, e))
                return default_export_root()
        return os.path.normpath(folder)
    except Exception as e:
        log("ERROR: Cannot read folder file {}: {}".format(folder_file, e))
        return default_export_root()


out = script.get_output()
out.close_others(all_open_outputs=True)


def to_model_path(user_visible_path):
    if not user_visible_path:
        return None
    try:
        return ModelPathUtils.ConvertUserVisiblePathToModelPath(user_visible_path)
    except Exception:
        return None


def default_export_root():
    docs = os.path.join(os.path.expanduser("~"), "Documents")
    root = os.path.join(docs, "RVT_Export")
    if not os.path.exists(root):
        try:
            os.makedirs(root)
        except Exception:
            pass
    return root


def save_document(doc, full_path):
    dst_dir = os.path.dirname(full_path)
    if not os.path.exists(dst_dir):
        try:
            os.makedirs(dst_dir)
        except Exception:
            pass

    mp_out = ModelPathUtils.ConvertUserVisiblePathToModelPath(full_path)

    sao = SaveAsOptions()
    sao.Compact = bool(COMPACT_ON_SAVE)
    sao.OverwriteExistingFile = bool(OVERWRITE_SAME)

    is_workshared = doc.IsWorkshared

    if DETACH_MODE == "preserve" and is_workshared:
        wsa = WorksharingSaveAsOptions()
        wsa.SaveAsCentral = True
        sao.SetWorksharingOptions(wsa)
        doc.SaveAs(mp_out, sao)
    else:
        doc.SaveAs(mp_out, sao)


def is_workshared_file(mp):
    try:
        from Autodesk.Revit.DB import BasicFileInfo

        file_info = BasicFileInfo.Extract(
            ModelPathUtils.ConvertModelPathToUserVisiblePath(mp)
        )
        return file_info.IsWorkshared
    except Exception:
        return True


def main():
    log("=" * 60)
    object_names = read_object_names()
    if not object_names:
        log("ERROR: No objects found")
        log("=" * 60)
        return

    all_models = []
    for obj_name in object_names:
        models = read_model_paths(obj_name)
        if models:
            for model_path in models:
                all_models.append((obj_name, model_path))

    if not all_models:
        log("ERROR: No models found for export")
        log("=" * 60)
        return

    log("Total models: {}".format(len(all_models)))
    log("=" * 60)

    exported_count = 0
    skipped_count = 0
    error_count = 0

    t_all = coreutils.Timer()
    for i, (obj_name, user_path) in enumerate(all_models):
        model_name = os.path.basename(user_path)
        file_wo_ext = os.path.splitext(model_name)[0]
        dest_folder = read_export_folder(obj_name)

        out_file_expected = os.path.join(dest_folder, model_name)
        log(
            "[{}/{}] Object: {}, Model: {}".format(
                i + 1, len(all_models), obj_name, model_name
            )
        )
        log("  -> {}".format(out_file_expected))

        mp = to_model_path(user_path)
        if mp is None:
            log("  ERROR: Cannot convert path to ModelPath")
            error_count += 1
            continue

        t_open = coreutils.Timer()

        def workset_filter(ws_name):
            name = (ws_name or "").strip()
            if name.startswith("00_"):
                return False
            name_lower = name.lower()
            if "link" in name_lower or "связь" in name_lower:
                return False
            return True

        try:
            doc, failure_handler, dialog_suppressor = openbg.open_in_background(
                __revit__.Application,
                __revit__,
                mp,
                audit=False,
                worksets=("predicate", workset_filter),
                detach=True,
                suppress_dialogs=True,
            )
        except Exception as e:
            log("  ERROR: Cannot open model: {}".format(e))
            error_count += 1
            continue
        open_s = str(datetime.timedelta(seconds=int(t_open.get_time())))

        if failure_handler is not None:
            try:
                summary = failure_handler.get_summary()
                total_w = summary.get("total_warnings", 0)
                total_e = summary.get("total_errors", 0)
                if total_w > 0 or total_e > 0:
                    log("  WARNING: {} warnings, {} errors".format(total_w, total_e))
                    if total_w > 0:
                        warnings = summary.get("warnings", [])
                        for idx, w in enumerate(warnings[:3], 1):
                            log("    {}. {}".format(idx, w))
                        if total_w > 3:
                            log("    ... and {} more warnings".format(total_w - 3))
                    if total_e > 0:
                        errors = summary.get("errors", [])
                        for idx, err in enumerate(errors[:2], 1):
                            log("    Error {}: {}".format(idx, err))
                        if total_e > 2:
                            log("    ... and {} more errors".format(total_e - 2))
            except Exception:
                pass

        t_save = coreutils.Timer()
        ok, err = True, None
        try:
            save_document(doc, out_file_expected)
        except Exception as e:
            ok, err = False, str(e)
        save_s = str(datetime.timedelta(seconds=int(t_save.get_time())))

        try:
            closebg.close_with_policy(doc, do_sync=False, save_if_not_ws=False)
        except Exception:
            pass

        if dialog_suppressor is not None:
            try:
                dialog_summary = dialog_suppressor.get_summary()
                total_dialogs = dialog_summary.get("total_dialogs", 0)
                if total_dialogs > 0:
                    log("  Dialogs suppressed: {}".format(total_dialogs))
            except Exception:
                pass
            try:
                dialog_suppressor.detach()
            except Exception:
                pass

        if ok and os.path.exists(out_file_expected):
            exported_count += 1
            file_size = 0
            try:
                if os.path.exists(out_file_expected):
                    file_size = os.path.getsize(out_file_expected) / (1024.0 * 1024.0)
                    log("  SUCCESS: Open: {}, Save: {}".format(open_s, save_s))
                    log("  File: {} ({:.2f} MB)".format(out_file_expected, file_size))
                else:
                    log("  SUCCESS: Open: {}, Save: {}".format(open_s, save_s))
                    log("  File: {}".format(out_file_expected))
            except:
                log("  SUCCESS: Open: {}, Save: {}".format(open_s, save_s))
                log("  File: {}".format(out_file_expected))
        else:
            error_count += 1
            log("  ERROR: {}".format(err if err else "Unknown error"))

    all_s = str(datetime.timedelta(seconds=int(t_all.get_time())))
    log("")
    log("=== SUMMARY ===")
    log(
        "Total: {}, Exported: {}, Skipped: {}, Errors: {}".format(
            len(all_models), exported_count, skipped_count, error_count
        )
    )
    log("=" * 60)
    log("DONE. Total time: {}".format(all_s))
    log("=" * 60)
    service_log(
        "Export completed: {} exported, {} skipped, {} errors".format(
            exported_count, skipped_count, error_count
        )
    )


if __name__ == "__main__":
    main()
