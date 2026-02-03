# -*- coding: utf-8 -*-
__title__ = "Экспорт NWC"
__author__ = "vlad / you"
__doc__ = "Экспорт .nwc без интерфейса. Читает объекты из Y:\\BIM\\Scripts\\Objects\\Ежедневная выгрузка.txt, пути из <object>.txt, папку экспорта из <object>_NWC.txt. Открытие РН через openbg: все, кроме начинающихся с '00_' и содержащих 'Link'/'Связь'. Совместимость Revit 2022/2023."

import os
import sys
import datetime
import codecs
import re
from pyrevit import script, coreutils

# Добавление пути к lib для импорта openbg и closebg
lib_path = r"D:\Share\OneDrive - SIYA Project\03_YandexDisk_Vlad\BIM\Scripts\WWBIM.extension\lib"
if lib_path not in sys.path:
    sys.path.insert(0, lib_path)


# Revit API
from Autodesk.Revit.DB import (
    ModelPathUtils,
    View3D,
    ViewFamilyType,
    ViewFamily,
    FilteredElementCollector,
    CategoryType,
    BuiltInCategory,
    BuiltInParameter,
    Transaction,
    NavisworksExportOptions,
    NavisworksExportScope,
    Category,
    ElementId,
    ImportInstance,
    Options,
    GeometryElement,
    Solid,
    Mesh,
    GeometryInstance,
)
from System import Enum
from System.Collections.Generic import List

# ваши либы
import openbg
import closebg

SAVE_CREATED_VIEW = False

OBJECTS_FILE = r"Y:\BIM\Scripts\Objects\AUTO_NWC\AUTO_NWC_Objects.txt"
OBJECTS_BASE_DIR = r"Y:\BIM\Scripts\Objects"

# Логирование
LOG_DIR = r"Y:\BIM\Scripts\Objects\AUTO_NWC\logs"

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
        log_line = "[auto_nwis_export] {} {}".format(ts, msg)
        with codecs.open(service_log_file, "a", encoding="utf-8") as f:
            f.write("{}\n".format(log_line))
        print(log_line)
    except:
        pass


separator = "=" * 40
with codecs.open(log_file, "a", encoding="utf-8") as f:
    f.write("{}\n".format(separator))
    f.write(
        "NWC Export - {} {}\n".format(
            today, datetime.datetime.now().strftime("%H:%M:%S")
        )
    )
    f.write("{}\n\n".format(separator))

log("File: {}".format(OBJECTS_FILE))
log("Log: {}".format(log_file))
service_log("Export process started")


# Пункт 4: Подавление диалогов (из C# DailyNwcExport)
class DialogSuppressor:
    """Обработчик диалогов для подавления предупреждений"""

    def __init__(self, log_func):
        self.log = log_func
        self.total_dialogs = 0
        self.dialogs_list = []

        # Сообщения для подавления (из C#)
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

        # ID диалогов для подавления
        self.suppress_dialog_ids = [
            "navisworks",
            "docwarndialog",
            "coordination",
            "review",
        ]

    def should_suppress(self, message, dialog_id):
        """Проверяет, нужно ли подавить диалог"""
        if not message and not dialog_id:
            return False

        message_lower = message.lower() if message else ""
        dialog_id_lower = dialog_id.lower() if dialog_id else ""

        # Проверка сообщения
        for msg in self.suppress_messages:
            if msg.lower() in message_lower:
                return True

        # Проверка ID диалога
        for did in self.suppress_dialog_ids:
            if did.lower() in dialog_id_lower:
                return True

        return False

    def record_dialog(self, message, dialog_id):
        """Логирование диалога"""
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
        """Возвращает сводку подавленных диалогов"""
        return {"total_dialogs": self.total_dialogs, "dialogs": self.dialogs_list}


# Глобальные переменные для обработчика диалогов
_dialog_suppressor = None


def on_dialog_box_showing(sender, args):
    """Обработчик события DialogBoxShowing"""
    if _dialog_suppressor is None:
        return

    try:
        dialog_id = str(args.DialogId) if args.DialogId else ""

        # TaskDialog
        try:
            message = getattr(args, "Message", "")
            _dialog_suppressor.record_dialog(message, dialog_id)

            if _dialog_suppressor.should_suppress(message, dialog_id):
                from Autodesk.Revit.UI import TaskDialogResult

                args.OverrideResult(int(TaskDialogResult.Ok))
        except:
            pass

        # MessageBox
        try:
            message = getattr(args, "Message", "")
            _dialog_suppressor.record_dialog(message, dialog_id)

            if _dialog_suppressor.should_suppress(message, dialog_id):
                args.OverrideResult(1)
        except:
            pass
    except:
        pass


# Инициализация обработчика диалогов
try:
    dialog_suppressor_instance = DialogSuppressor(service_log)
    _dialog_suppressor = dialog_suppressor_instance

    # Регистрация обработчика
    try:
        __revit__.Application.DialogBoxShowing += on_dialog_box_showing
    except:
        pass
except:
    pass


# Сообщения об ошибках геометрии (из C# DailyNwcExport)
GEOMETRY_ERROR_MESSAGES = [
    "не найдена подходящая геометрия",
    "элементов для геометрии не найдено",
    "не найдена геометрия для экспорта",
    "no suitable geometry",
    "no appropriate geometry",
    "no elements for geometry",
    "no geometry for export",
]


def is_geometry_error(message):
    """Проверяет, является ли ошибка ошибкой геометрии (should be SKIP, not ERROR)"""
    if not message:
        return False
    msg_lower = message.lower()
    for geo_msg in GEOMETRY_ERROR_MESSAGES:
        if geo_msg.lower() in msg_lower:
            return True
    return False


def get_exception_info(ex):
    """Детальная информация об исключении"""
    return {
        "type": type(ex).__name__,
        "message": str(ex),
        "args": str(ex.args) if hasattr(ex, "args") else "",
    }


# Пункт 5: Проверка необходимости экспорта (из C# DailyNwcExport)
def get_path_type(path):
    """Определяет тип пути: revit_server или local"""
    if not path:
        return "unknown"
    upper_path = path.upper()
    if (
        upper_path.startswith("RSN://")
        or upper_path.startswith("RSN:/")
        or upper_path.startswith("RSN:\\")
        or upper_path.startswith("RSN:\\\\")
    ):
        return "revit_server"
    return "local"


def get_file_modification_date(path, path_type):
    """Получает дату изменения файла"""
    try:
        if path_type == "revit_server":
            return None
        else:
            if os.path.exists(path):
                return datetime.datetime.fromtimestamp(os.path.getmtime(path))
            return None
    except:
        return None


def normalize_revit_server_path(path):
    """Нормализует путь Revit Server (из C#)"""
    if not path:
        return path

    upper_path = path.upper()
    if not upper_path.startswith("RSN:"):
        return path

    path = path.replace("rsn:", "RSN:")
    import re

    path = re.sub(r"\\+", r"\\", path)
    path = path.replace("\\", "/")
    path = path.replace("RSN:/", "RSN://")
    path = path.replace("RSN:///", "RSN://")

    return path


def check_need_export(rvt_path, nwc_folder, object_name):
    """Проверяет, нужен ли экспорт (аналог C# CheckNeedExport)"""
    result = {
        "need_export": False,
        "reason": "",
        "rvt_date": None,
        "nwc_date": None,
        "nwc_used_path": None,
        "target_nwc_path": None,
    }

    try:
        result["path_type"] = get_path_type(rvt_path)
        result["rvt_date"] = get_file_modification_date(rvt_path, result["path_type"])

        rvt_filename = os.path.splitext(os.path.basename(rvt_path))[0]
        nwc_path1 = os.path.join(nwc_folder, rvt_filename + ".nwc")
        nwc_date1 = get_file_modification_date(nwc_path1, "local")

        nwc_path2 = None
        nwc_date2 = None

        # Проверка суффикса _RXX
        import re

        match = re.search(r"_R(\d+)$", rvt_filename)
        if match:
            suffix = match.group(0)
            version_num = int(match.group(1))
            new_suffix = "_N{}".format(version_num + 1)
            nwc_filename2 = rvt_filename.replace(suffix, new_suffix)
            nwc_path2 = os.path.join(nwc_folder, nwc_filename2 + ".nwc")
            nwc_date2 = get_file_modification_date(nwc_path2, "local")

        # Определение используемой даты NWC
        if nwc_date1 and nwc_date2:
            if nwc_date1 > nwc_date2:
                result["nwc_date"] = nwc_date1
                result["nwc_used_path"] = nwc_path1
            else:
                result["nwc_date"] = nwc_date2
                result["nwc_used_path"] = nwc_path2
        elif nwc_date1:
            result["nwc_date"] = nwc_date1
            result["nwc_used_path"] = nwc_path1
        elif nwc_date2:
            result["nwc_date"] = nwc_date2
            result["nwc_used_path"] = nwc_path2

        # Целевой путь для NWC
        if nwc_path2:
            result["target_nwc_path"] = nwc_path2
        else:
            result["target_nwc_path"] = nwc_path1

        # Проверка необходимости экспорта
        if result["path_type"] == "revit_server":
            result["need_export"] = True
            result["reason"] = "Revit Server (always export)"
        elif not result["nwc_date"]:
            result["need_export"] = True
            result["reason"] = "NWC does not exist"
        elif not result["rvt_date"]:
            result["need_export"] = True
            result["reason"] = "Cannot determine RVT date"
        elif result["rvt_date"] > result["nwc_date"]:
            result["need_export"] = True
            result["reason"] = "RVT updated"
        else:
            result["need_export"] = False
            result["reason"] = "NWC is up to date"
    except Exception as e:
        result["need_export"] = True
        result["reason"] = "Error checking: {}".format(e)

    return result


def has_exportable_geometry(doc, view):
    """Проверяет наличие экспортируемой геометрии в виде (из C# DailyNwcExport)"""
    try:
        log("  Checking geometry in view...")
        options = Options()
        options.View = view
        options.IncludeNonVisibleObjects = False

        elements = FilteredElementCollector(doc, view.Id).WhereElementIsNotElementType()

        for elem in elements:
            try:
                geo = elem.get_Geometry(options)
                if geo is None:
                    continue

                if _contains_exportable_geometry(geo):
                    log("  Geometry found: True")
                    return True
            except:
                continue

        log("  Geometry found: False")
        return False
    except Exception as e:
        log("  WARNING: Could not check geometry: {}".format(e))
        return True


def _contains_exportable_geometry(geometry_element):
    """Рекурсивно проверяет наличие экспортируемой геометрии"""
    try:
        for obj in geometry_element:
            if isinstance(obj, Solid):
                if obj.Volume > 1e-6:
                    return True
            elif isinstance(obj, Mesh):
                if obj.NumTriangles > 0:
                    return True
            elif isinstance(obj, GeometryInstance):
                inst_geo = obj.GetInstanceGeometry()
                if inst_geo is not None:
                    if _contains_exportable_geometry(inst_geo):
                        return True
        return False
    except:
        return False


def read_object_names():
    """Читает имена объектов из файла Ежедневная выгрузка.txt"""
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
    r"""Читает пути к RVT файлам из <object_name>.txt в папке Y:\BIM\Scripts\Objects"""
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
    r"""Читает папку для экспорта NWC из <object_name>_NWC.txt в папке Y:\BIM\Scripts\Objects"""
    folder_file = os.path.join(OBJECTS_BASE_DIR, object_name + "_NWC.txt")
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

# ---------- helpers ----------


def to_model_path(user_visible_path):
    if not user_visible_path:
        return None
    try:
        return ModelPathUtils.ConvertUserVisiblePathToModelPath(user_visible_path)
    except Exception:
        return None


def default_export_root():
    docs = os.path.join(os.path.expanduser("~"), "Documents")
    root = os.path.join(docs, "NWC_Export")
    if not os.path.exists(root):
        try:
            os.makedirs(root)
        except Exception:
            pass
    return root


# ---- SAFE BIC helpers (без getattr к BuiltInCategory) ----


def _resolve_bic(name):
    """Вернуть BuiltInCategory по строке или None, если такого имени нет в текущей версии Revit."""
    if not name:
        return None
    try:
        if Enum.IsDefined(BuiltInCategory, name):
            return Enum.Parse(BuiltInCategory, name)
    except Exception:
        pass
    return None


def _resolve_bip(name):
    if not name:
        return None
    try:
        if Enum.IsDefined(BuiltInParameter, name):
            return Enum.Parse(BuiltInParameter, name)
    except Exception:
        pass
    return None


def _try_set_bip_int(element, bip_name, value):
    bip = _resolve_bip(bip_name)
    if bip is None:
        return False
    try:
        p = element.get_Parameter(bip)
        if p and (not p.IsReadOnly):
            p.Set(int(value))
            return True
    except Exception:
        pass
    return False


def _cat_id(doc, bic):
    if bic is None:
        return None
    try:
        cat = Category.GetCategory(doc, bic)
        if cat:
            return cat.Id
    except Exception:
        return None
    return None


def _hide_categories_by_names(doc, view, names):
    ids = List[ElementId]()
    for nm in names or []:
        bic = _resolve_bic(nm)
        eid = _cat_id(doc, bic)
        if eid:
            ids.Add(eid)

    # скрыть все аннотации через тип категории (это стабильно во всех версиях)
    try:
        for c in doc.Settings.Categories:
            try:
                if c.CategoryType == CategoryType.Annotation:
                    ids.Add(c.Id)
            except Exception:
                pass
    except Exception:
        pass

    if ids.Count == 0:
        return 0

    hidden = 0
    try:
        view.HideCategories(ids)
        hidden = ids.Count
        return hidden
    except Exception:
        pass

    for eid in ids:
        ok = False
        try:
            view.SetCategoryHidden(eid, True)
            ok = True
        except Exception:
            try:
                view.SetCategoryHidden(eid.IntegerValue, True)
                ok = True
            except Exception:
                ok = False
        if ok:
            hidden += 1
    return hidden


def hide_annos_and_links_safe(view):
    doc = view.Document

    # View template can lock Visibility/Graphics. Detach for this export session (doc is opened detached and not saved).
    try:
        vtid = getattr(view, "ViewTemplateId", None)
        if vtid and (vtid.IntegerValue != -1):
            view.ViewTemplateId = ElementId.InvalidElementId
    except Exception:
        pass
    names = [
        "OST_RvtLinks",
        "OST_LinkInstances",
        # Все варианты импорта (DWG, DXF и др.)
        "OST_ExportLayer",
        "OST_ImportInstance",
        "OST_ImportsInFamilies",
        "OST_ImportObjectStyles",  # Импорт в семействах (стили объектов)
        "OST_Cameras",
        "OST_Views",
        "OST_Lines",
        "OST_PointClouds",
        "OST_PointCloudsHardware",
        "OST_Levels",
        "OST_Grids",
        "OST_Annotations",
        "OST_TitleBlocks",
        "OST_Viewports",
        "OST_TextNotes",
        "OST_Dimensions",
    ]
    hidden = _hide_categories_by_names(doc, view, names)
    # ВАЖНО: Отключаем чекбоксы "Показывать импортированные/аннотации на этом виде"
    try:
        # Скрыть все импортированные категории (вкладка "Импортированные категории")
        view.AreImportCategoriesHidden = True
    except Exception:
        pass
    _try_set_bip_int(view, "VIEW_SHOW_IMPORT_CATEGORIES", 0)
    _try_set_bip_int(view, "VIEW_SHOW_IMPORT_CATEGORIES_IN_VIEW", 0)

    try:
        # Скрыть все категории аннотаций (вкладка "Категории аннотаций")
        view.AreAnnotationCategoriesHidden = True
    except Exception:
        pass
    _try_set_bip_int(view, "VIEW_SHOW_ANNOTATION_CATEGORIES", 0)
    _try_set_bip_int(view, "VIEW_SHOW_ANNOTATION_CATEGORIES_IN_VIEW", 0)

    # Extra safety: explicitly hide ImportInstance elements so they won't be exported even if category flags are blocked.
    try:
        ids = List[ElementId]()
        for ii in FilteredElementCollector(doc, view.Id).OfClass(ImportInstance):
            try:
                if view.CanElementBeHidden(ii.Id):
                    ids.Add(ii.Id)
            except Exception:
                ids.Add(ii.Id)
        if ids.Count > 0:
            view.HideElements(ids)
    except Exception:
        pass

    return hidden


def find_or_create_navis_view(doc):
    for v in FilteredElementCollector(doc).OfClass(View3D):
        try:
            if (not v.IsTemplate) and v.Name == "Navisworks":
                # на всякий — привести вид к нужному набору скрытий
                with Transaction(doc, "Configure Navisworks view") as t:
                    t.Start()
                    hide_annos_and_links_safe(v)
                    t.Commit()
                # отключить 3D подрезку вида (Границы 3D вида)
                with Transaction(doc, "Отключить 3D подрезку") as t:
                    t.Start()
                    try:
                        v.IsSectionBoxActive = False
                    except Exception:
                        pass
                    t.Commit()
                return v, False
        except Exception:
            pass
    vft = None
    for t in FilteredElementCollector(doc).OfClass(ViewFamilyType):
        if t.ViewFamily == ViewFamily.ThreeDimensional:
            vft = t
            break
    if vft is None:
        raise Exception("Не найден тип 3D-вида для создания 'Navisworks'.")
    with Transaction(doc, "Создать вид Navisworks") as t:
        t.Start()
        view = View3D.CreateIsometric(doc, vft.Id)
        view.Name = "Navisworks"
        hide_annos_and_links_safe(view)
        # отключить 3D подрезку вида (Границы 3D вида)
        try:
            view.IsSectionBoxActive = False
        except Exception:
            pass
        t.Commit()
    return view, True


def count_visible_elements(doc, view):
    return (
        FilteredElementCollector(doc, view.Id)
        .WhereElementIsNotElementType()
        .GetElementCount()
    )


def export_view_to_nwc(doc, view, target_folder, file_wo_ext):
    r"""Экспорт указанного вида в .nwc. Возвращает (api_ok, out_path)."""
    log("    Export function started...")
    log("    target_folder: {}".format(target_folder))
    log("    file_wo_ext: {}".format(file_wo_ext))

    if not target_folder or not file_wo_ext:
        log("    ERROR: Missing target_folder or file_wo_ext")
        return False, None

    # Проверка папки перед экспортом
    if not os.path.exists(target_folder):
        log("    WARNING: Target folder does not exist, attempting to create...")
        try:
            os.makedirs(target_folder)
            log("    Target folder created successfully")
        except Exception as e:
            log("    ERROR: Cannot create target folder: {}".format(e))
            return False, None
    else:
        log("    Target folder exists")

    # Создание опций экспорта
    log("    Creating export options...")
    opts = NavisworksExportOptions()
    opts.ExportScope = NavisworksExportScope.View
    opts.ViewId = view.Id
    log("    ExportScope: View")
    log("    ViewId: {}".format(view.Id))

    # Расчёт ожидаемого пути
    out_path = os.path.join(target_folder, file_wo_ext + ".nwc")
    log("    Expected output path: {}".format(out_path))

    # Проверка существования файла перед экспортом
    if os.path.exists(out_path):
        try:
            old_size = os.path.getsize(out_path)
            log("    Existing file found, size: {} bytes".format(old_size))
        except Exception as e:
            log("    WARNING: Cannot get existing file size: {}".format(e))

    # Вызов экспорта
    log("    Calling doc.Export()...")
    api_ok = False
    try:
        api_ok = doc.Export(target_folder, file_wo_ext, opts)
        log("    doc.Export() returned: {}".format(api_ok))
    except Exception as e:
        log("    EXPORT EXCEPTION:")
        log("      Type: {}".format(type(e).__name__))
        log("      Message: {}".format(str(e)))
        try:
            import traceback

            log("      Traceback: {}".format(traceback.format_exc()))
        except:
            pass
        api_ok = False

    # Проверка результата
    file_exists = os.path.exists(out_path)
    if file_exists:
        try:
            file_size = os.path.getsize(out_path)
            log("    Output file exists: True, size: {} bytes".format(file_size))
        except Exception as e:
            log("    WARNING: Cannot get file size: {}".format(e))
    else:
        log("    Output file exists: False")

    log("    Export function completed")
    return api_ok, out_path


# ---------- main ----------


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

    # Счётчики статистики
    exported_count = 0
    skipped_count = 0
    error_count = 0

    t_all = coreutils.Timer()
    for i, (obj_name, user_path) in enumerate(all_models):
        model_name = os.path.basename(user_path)
        file_wo_ext = os.path.splitext(model_name)[0]
        dest_folder = read_export_folder(obj_name)

        out_file_expected = os.path.join(dest_folder, file_wo_ext + ".nwc")
        log(
            "[{}/{}] Object: {}, Model: {}".format(
                i + 1, len(all_models), obj_name, model_name
            )
        )
        log("  -> {}".format(out_file_expected))

        mp = to_model_path(user_path)
        if mp is None:
            log("  ERROR: Cannot convert path to ModelPath")
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
            continue
        open_s = str(datetime.timedelta(seconds=int(t_open.get_time())))

        if failure_handler is not None:
            try:
                summary = failure_handler.get_summary()
                total_w = summary.get("total_warnings", 0)
                total_e = summary.get("total_errors", 0)
                if total_w > 0 or total_e > 0:
                    log("  WARNING: {} warnings, {} errors".format(total_w, total_e))
            except Exception:
                pass

        try:
            view, created = find_or_create_navis_view(doc)
        except Exception as e:
            log("  ERROR: Cannot create Navisworks view: {}".format(e))
            if dialog_suppressor is not None:
                try:
                    dialog_suppressor.detach()
                except Exception:
                    pass
            try:
                closebg.close_with_policy(doc, do_sync=False, save_if_not_ws=False)
            except Exception:
                pass
            continue

        try:
            doc.Regenerate()
        except Exception:
            pass

        vis_count = count_visible_elements(doc, view)
        log("  View: {}, Elements: {}".format(view.Name, vis_count))

        # Пункт 5: Проверка необходимости экспорта
        export_check = check_need_export(user_path, dest_folder, obj_name)
        log(
            "  Export check: need={}, reason={}".format(
                export_check["need_export"], export_check["reason"]
            )
        )

        if not export_check["need_export"]:
            log("  SKIP: {}".format(export_check["reason"]))
            log("  Open: {}, Export: 0:00:00".format(open_s))
            skipped_count += 1
            try:
                closebg.close_with_policy(doc, do_sync=False, save_if_not_ws=False)
            except:
                pass
            continue

        t_exp = coreutils.Timer()
        api_ok, out_path = False, out_file_expected
        err_text = None
        skip_reason = None

        # Целевой путь для NWC
        target_nwc = export_check["target_nwc_path"]
        if target_nwc:
            out_file_expected = target_nwc
        else:
            out_file_expected = os.path.join(dest_folder, file_wo_ext + ".nwc")

        log("  Target NWC: {}".format(out_file_expected))

        # Удаление существующего NWC файла (как в C#)
        if os.path.exists(out_file_expected):
            try:
                os.remove(out_file_expected)
                log("  Deleted existing file")
            except Exception as e:
                log("  WARNING: Cannot delete existing file: {}".format(e))

        # Пункт 3: Проверка геометрии перед экспортом
        has_geo = has_exportable_geometry(doc, view)

        try:
            if has_geo:
                log("  Starting export to: {}".format(dest_folder))
                nwc_file_wo_ext = os.path.splitext(os.path.basename(out_file_expected))[
                    0
                ]
                api_ok, out_path = export_view_to_nwc(
                    doc, view, dest_folder, nwc_file_wo_ext
                )
            else:
                skip_reason = "View has no exportable geometry"
                log("  SKIP: {}".format(skip_reason))
                skipped_count += 1
        except Exception as e:
            err_msg = str(e)
            ex_info = get_exception_info(e)
            log("  EXPORT ERROR:")
            log("    Type: {}".format(ex_info["type"]))
            log("    Message: {}".format(err_msg))
            if ex_info["args"]:
                log("    Args: {}".format(ex_info["args"]))

            if is_geometry_error(err_msg):
                skip_reason = "Model contains no exportable geometry"
                log("  SKIP: {}".format(skip_reason))
                skipped_count += 1
            else:
                err_text = err_msg
                log("  ERROR: {}".format(err_text))

        file_ok = os.path.exists(out_path) and (os.path.getsize(out_path) > 0)
        ok = (api_ok or file_ok) and (err_text is None)
        exp_s = str(datetime.timedelta(seconds=int(t_exp.get_time())))

        if file_ok and not api_ok and err_text is None:
            log("  WARNING: API returned False but file exists")

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

        if skip_reason:
            log("  SKIP: {}".format(skip_reason))
            log("  Open: {}, Export: {}".format(open_s, exp_s))
        elif ok:
            exported_count += 1
            # Размер файла
            file_size = 0
            try:
                if out_path and os.path.exists(out_path):
                    file_size = os.path.getsize(out_path) / (1024.0 * 1024.0)
                    log("  SUCCESS: Open: {}, Export: {}".format(open_s, exp_s))
                    log("  File: {} ({:.2f} MB)".format(out_path, file_size))
                else:
                    log("  SUCCESS: Open: {}, Export: {}".format(open_s, exp_s))
                    log("  File: {}".format(out_path))
            except:
                log("  SUCCESS: Open: {}, Export: {}".format(open_s, exp_s))
                log("  File: {}".format(out_path))
        else:
            error_count += 1
            # Детальная диагностика "Unknown error"
            if err_text is None:
                log("  ERROR: Export failed without exception")
                if out_path:
                    file_exists = os.path.exists(out_path)
                    if file_exists:
                        try:
                            file_size = os.path.getsize(out_path)
                            log(
                                "  ERROR: File exists but export failed: {} bytes".format(
                                    file_size
                                )
                            )
                        except Exception as e:
                            log(
                                "  ERROR: File exists but cannot get size: {}".format(e)
                            )
                    else:
                        log("  ERROR: File not created: {}".format(out_path))
                else:
                    log("  ERROR: No output path specified")
            else:
                log("  ERROR: {}".format(err_text))

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
