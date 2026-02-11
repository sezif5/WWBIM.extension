# -*- coding: utf-8 -*-
# -*- coding: utf-8 -*-
import os
import sys
import imp

from pyrevit import script, forms

out = script.get_output()
from Autodesk.Revit.DB import (
    FilteredElementCollector,
    WorksetConfiguration,
    WorksetConfigurationOption,
    WorksetId,
)
from System.Collections.Generic import List

# Импорт модулей пакетных операций
import openbg
import closebg

# ---------- Константы ----------

SCRIPT_DIR = os.path.dirname(__file__)
PANEL_DIR = os.path.dirname(SCRIPT_DIR)
TAB_DIR = os.path.dirname(PANEL_DIR)
EXTENSION_ROOT = os.path.dirname(TAB_DIR)
SCRIPTS_ROOT = os.path.dirname(EXTENSION_ROOT)
LIB_DIR = os.path.join(EXTENSION_ROOT, "lib")
OBJECTS_DIR = os.path.join(SCRIPTS_ROOT, "Objects")
PYTHON_SCRIPTS_DIR = os.path.join(LIB_DIR, "Batch Operations")


# ---------- Вспомогательные функции ----------


def to_model_path(model_path_str):
    """Преобразовать путь к модели в ModelPath для Revit API."""
    if not model_path_str:
        return None

    if os.path.isfile(model_path_str):
        try:
            from Autodesk.Revit.DB import ModelPathUtils

            return ModelPathUtils.ConvertUserVisiblePathToModelPath(model_path_str)
        except Exception:
            return None
    else:
        return None


def list_txt_files(folder_path):
    """Получить список txt файлов без расширения."""
    lst = []
    try:
        for f in os.listdir(folder_path):
            if f.lower().endswith(".txt"):
                lst.append(f[:-4])
    except OSError:
        pass
    return lst


def get_open_documents():
    """Получить список открытых документов Revit (кроме семей)."""
    docs = []
    app = __revit__
    for doc in app.Documents:
        try:
            if not doc.IsFamilyDocument:
                docs.append(doc)
        except Exception:
            pass
    return docs


def select_open_documents():
    """Выбрать открытые документы через UI."""
    docs = get_open_documents()

    if not docs:
        forms.alert("Нет открытых документов для выбора.", warn_icon=True)
        return None

    class DocItem:
        def __init__(self, doc):
            self.doc = doc
            self.name = doc.Title

        def __str__(self):
            return self.name

    doc_items = [DocItem(d) for d in docs]
    selected = forms.SelectFromList.show(
        doc_items,
        title="Выберите документы",
        multiselect=True,
        width=600,
        height=400,
        button_name="Выбрать",
    )

    if not selected:
        return None

    if not isinstance(selected, list):
        selected = [selected]

    return [item.doc for item in selected]


def _clear_transmission_flag(model_path_str):
    """Сбрасывает флаг переданного файла (IsTransmitted -> False)."""
    try:
        from Autodesk.Revit.DB import ModelPathUtils, TransmissionData

        mp = ModelPathUtils.ConvertUserVisiblePathToModelPath(model_path_str)
        td = TransmissionData.ReadTransmissionData(mp)
        if td is None:
            # ReadTransmissionData вернул None - логируем причину
            out.print_md(
                ":warning: [_clear_transmission_flag] ReadTransmissionData returned None for: {}".format(
                    model_path_str
                )
            )
            return True  # Считаем, что флага нет
        if td is not None and hasattr(td, "IsTransmitted") and td.IsTransmitted:
            td.IsTransmitted = False
            TransmissionData.WriteTransmissionData(mp, td)
            return True
        return True  # Флаг уже снят или не был установлен
    except Exception as e:
        # Логируем ошибку, а не глушим молча
        out.print_md(
            ":warning: [_clear_transmission_flag] Error: {} (path: {})".format(
                str(e), model_path_str
            )
        )
        return False


def read_models_from_txt(file_path):
    """Прочитать список моделей из txt файла."""
    models = []
    try:
        with open(file_path, "rb") as f:
            for raw in f:
                try:
                    line = raw.decode("utf-8").strip()
                except Exception:
                    try:
                        line = raw.decode("cp1251").strip()
                    except Exception:
                        line = raw.strip()
                if line:
                    models.append(line)
    except Exception:
        pass
    return models


def get_open_families():
    """Получить список открытых семейств (документов-семейств)."""
    app = __revit__.Application
    families = []
    for doc in app.Documents:
        try:
            if doc.IsFamilyDocument:
                families.append(doc)
        except Exception:
            pass
    return families


# ---------- Действия ----------


def should_close_workset(ws_name):
    """Проверить, нужно ли закрыть рабочий набор."""
    name = (ws_name or "").strip()
    name_lower = name.lower()
    # Закрываем если начинается с 00_ или содержит Связь/Links
    if name.startswith("00_"):
        return True
    if "связь" in name_lower:
        return True
    if "link" in name_lower:
        return True
    return False


def get_worksets_to_open(uiapp, mp):
    """Получить список ID рабочих наборов, которые нужно открыть."""
    from Autodesk.Revit.DB import WorksharingUtils

    try:
        previews = WorksharingUtils.GetUserWorksetInfo(mp)
    except Exception:
        try:
            previews = WorksharingUtils.GetUserWorksetInfoForOpen(uiapp, mp)
        except Exception:
            return None

    if not previews:
        return None

    ids_to_open = List[WorksetId]()
    for p in previews:
        try:
            if not should_close_workset(p.Name):
                ids_to_open.Add(p.Id)
        except Exception:
            pass

    return ids_to_open


def action_open_models(models):
    """Открыть выбранные модели с созданием локальной копии."""
    from Autodesk.Revit.DB import OpenOptions

    out.print_md("## ОТКРЫТИЕ МОДЕЛЕЙ ({})".format(len(models)))
    out.print_md("Создание локальных копий, закрытие РН: 00_*, Связь, Links")
    out.print_md("---")

    uiapp = __revit__
    app = uiapp.Application

    for i, model_path in enumerate(models):
        model_name = os.path.basename(model_path)
        out.print_md(":open_file_folder: **{}**".format(model_name))

        mp = to_model_path(model_path)
        if mp is None:
            out.print_md(":x: Не удалось преобразовать путь.")
            out.update_progress(i + 1, len(models))
            continue

        dialog_suppressor = None

        try:
            result = openbg.open_in_background(
                __revit__,
                None,
                mp,
                audit=False,
                worksets=("all_except_prefixes", ["00_", "Связь", "Links"]),
                detach=False,
                suppress_warnings=True,
                suppress_dialogs=True,
            )
            if len(result) >= 3:
                doc = result[0]
                dialog_suppressor = result[2]
            else:
                raise Exception("openbg не вернул документ")

            dialog_summary = dialog_suppressor.get_summary()
            out.print_md(
                "  dialogs: total={}, suppressed={}, unknown={}, errors={}".format(
                    dialog_summary.get("total", 0),
                    dialog_summary.get("suppressed", 0),
                    dialog_summary.get("unknown", 0),
                    dialog_summary.get("errors", 0),
                )
            )
            if dialog_summary.get("transmitted_dialog_handled"):
                out.print_md("  :white_check_mark: transmitted_dialog_handled")
            if dialog_summary.get("unknown", 0) > 0:
                unknown = dialog_summary.get("unknown_details", [])[:3]
                for uid, snippet in unknown:
                    out.print_md("  unknown_dialog: [{}] {}".format(uid, snippet[:50]))

            out.print_md(
                "  doc: workshared={}, readonly={}, path={}".format(
                    doc.IsWorkshared,
                    doc.IsReadOnly,
                    os.path.basename(doc.PathName) if doc.PathName else "None",
                )
            )

            # Загрузить все выбранные семейства
            loaded_count = 0
            from Autodesk.Revit.DB import Transaction

            t = Transaction(doc, "Загрузка семейств")
            t.Start()
            try:
                for family_doc in family_docs:
                    try:
                        loaded_family = doc.LoadFamily(family_doc)
                        if loaded_family:
                            loaded_count += 1
                    except Exception:
                        pass
                t.Commit()

                if loaded_count == len(family_docs):
                    out.print_md(
                        ":white_check_mark: Загружено семейств: {}/{}".format(
                            loaded_count, len(family_docs)
                        )
                    )
                    success_models += 1
                elif loaded_count > 0:
                    out.print_md(
                        ":warning: Загружено семейств: {}/{} (остальные уже существуют)".format(
                            loaded_count, len(family_docs)
                        )
                    )
                    success_models += 1
                else:
                    out.print_md(":warning: Семейства уже существуют или не загружены")

            except Exception as e:
                t.RollBack()
                out.print_md(":x: Ошибка загрузки: `{}`".format(e))

            out.print_md(
                "  before_close: readonly={}, path={}".format(
                    doc.IsReadOnly,
                    os.path.basename(doc.PathName) if doc.PathName else "None",
                )
            )

            # Синхронизация и закрытие
            res = closebg.close_with_policy(
                doc,
                do_sync=True,
                comment="Загрузка семейств",
                dialog_suppressor=dialog_suppressor,
                source_path=model_path,
            )

            if not res.get("success"):
                out.print_md(
                    ":x: Ошибка закрытия: {}".format(
                        res.get(
                            "save_error", res.get("close_error", "Неизвестная ошибка")
                        )
                    )
                )
                if doc and doc.IsValidObject:
                    out.print_md(
                        "  doc_state: readonly={}, path={}".format(
                            doc.IsReadOnly,
                            os.path.basename(doc.PathName) if doc.PathName else "None",
                        )
                    )
                if dialog_summary.get("transmitted_dialog_handled"):
                    out.print_md("  :white_check_mark: transmitted_dialog_handled")

        except Exception as e:
            out.print_md(":x: Ошибка открытия: `{}`".format(e))
        finally:
            if dialog_suppressor:
                dialog_suppressor.detach()

        out.update_progress(i + 1, len(models))

    out.print_md("---")
    out.print_md(
        "**Готово. Моделей обработано: {}/{}**".format(success_models, len(models))
    )


def action_add_link(models):
    """Добавить связь."""
    links_to_add = select_links_from_txt()

    if not links_to_add:
        return

    out.print_md("## ДОБАВЛЕНИЕ СВЯЗЕЙ")
    out.print_md("**Связей для добавления:** {}".format(len(links_to_add)))
    out.print_md("**Моделей:** {}".format(len(models)))
    out.print_md("---")

    success_models = 0

    for i, model_path in enumerate(models):
        model_name = os.path.basename(model_path)
        out.print_md(":open_file_folder: **{}**".format(model_name))

        mp = to_model_path(model_path)
        if mp is None:
            out.print_md(":x: Не удалось преобразовать путь.")
            out.update_progress(i + 1, len(models))
            continue

        dialog_suppressor = None

        try:
            result = openbg.open_in_background(
                __revit__,
                None,
                mp,
                audit=False,
                worksets=("all_except_prefixes", ["00_", "Связь", "Links"]),
                detach=False,
                suppress_warnings=True,
                suppress_dialogs=True,
            )
            if len(result) >= 3:
                doc = result[0]
                dialog_suppressor = result[2]
            else:
                raise Exception("openbg не вернул документ")

            dialog_summary = dialog_suppressor.get_summary()
            out.print_md(
                "  dialogs: total={}, suppressed={}, unknown={}, errors={}".format(
                    dialog_summary.get("total", 0),
                    dialog_summary.get("suppressed", 0),
                    dialog_summary.get("unknown", 0),
                    dialog_summary.get("errors", 0),
                )
            )
            if dialog_summary.get("transmitted_dialog_handled"):
                out.print_md("  :white_check_mark: transmitted_dialog_handled")
            if dialog_summary.get("unknown", 0) > 0:
                unknown = dialog_summary.get("unknown_details", [])[:3]
                for uid, snippet in unknown:
                    out.print_md("  unknown_dialog: [{}] {}".format(uid, snippet[:50]))

            out.print_md(
                "  doc: workshared={}, readonly={}, path={}".format(
                    doc.IsWorkshared,
                    doc.IsReadOnly,
                    os.path.basename(doc.PathName) if doc.PathName else "None",
                )
            )

            # Получить существующие связи
            existing_links = get_existing_links(doc)

            from Autodesk.Revit.DB import Transaction, RevitLinkType

            t = Transaction(doc, "Добавление связей")
            t.Start()
            try:
                added_count = 0
                skipped_count = 0
                for link_path in links_to_add:
                    link_mp = to_model_path(link_path)
                    if link_mp:
                        try:
                            link_type = RevitLinkType.Load
                            link = doc.LoadLink(link_mp, link_type, None)
                            added_count += 1
                            out.print_md(
                                "  :white_check_mark: Добавлена связь: {}".format(
                                    os.path.basename(link_path)
                                )
                            )
                        except Exception as e:
                            out.print_md(
                                "  :warning: Не удалось добавить связь: {}".format(e)
                            )
                            skipped_count += 1
                    else:
                        skipped_count += 1

                t.Commit()

                if added_count > 0:
                    out.print_md(
                        ":white_check_mark: Связей добавлено: {}".format(added_count)
                    )
                    success_models += 1
                else:
                    out.print_md(":warning: Связи не добавлены")

            except Exception as e:
                t.RollBack()
                out.print_md(":x: Ошибка транзакции: `{}`".format(e))

            out.print_md(
                "  before_close: readonly={}, path={}".format(
                    doc.IsReadOnly,
                    os.path.basename(doc.PathName) if doc.PathName else "None",
                )
            )

            # Синхронизация и закрытие
            res = closebg.close_with_policy(
                doc,
                do_sync=True,
                comment="Добавление связей",
                dialog_suppressor=dialog_suppressor,
                source_path=model_path,
            )

            if not res.get("success"):
                out.print_md(
                    ":x: Ошибка закрытия: {}".format(
                        res.get(
                            "save_error", res.get("close_error", "Неизвестная ошибка")
                        )
                    )
                )
                if doc and doc.IsValidObject:
                    out.print_md(
                        "  doc_state: readonly={}, path={}".format(
                            doc.IsReadOnly,
                            os.path.basename(doc.PathName) if doc.PathName else "None",
                        )
                    )
                if dialog_summary.get("transmitted_dialog_handled"):
                    out.print_md("  :white_check_mark: transmitted_dialog_handled")

        except Exception as e:
            out.print_md(":x: Ошибка открытия: `{}`".format(e))
        finally:
            if dialog_suppressor:
                dialog_suppressor.detach()

        out.update_progress(i + 1, len(models))

    out.print_md("---")
    out.print_md(
        "**Готово. Моделей обработано: {}/{}**".format(success_models, len(models))
    )


def action_run_python_script_on_opened_docs(docs, scripts):
    """Выполнить python скрипты из библиотеки на выбранных открытых документах.

    ВАЖНО: Документы, выбранные пользователем, НЕ ЗАКРЫВАТЬСЯ.
    """
    script_names = [os.path.basename(s) for s in scripts]

    out.print_md("## ВЫПОЛНЕНИЕ PYTHON СКРИПТОВ НА ОТКРЫТЫХ ДОКУМЕНТАХ")
    out.print_md("**Скрипты ({}):** {}".format(len(scripts), ", ".join(script_names)))
    out.print_md("**Документов:** {}".format(len(docs)))
    out.print_md("---")

    success_docs = 0

    for i, doc in enumerate(docs):
        doc_name = doc.Title
        out.print_md(":open_file_folder: **{}**".format(doc_name))

        scripts_succeeded = 0
        scripts_failed = 0

        try:
            for script_rel_path in scripts:
                script_path = os.path.join(PYTHON_SCRIPTS_DIR, script_rel_path)
                script_name = os.path.basename(script_rel_path)

                module_name = (
                    os.path.splitext(script_rel_path)[0]
                    .replace("\\", "_")
                    .replace("/", "_")
                )

                try:
                    sys.modules.pop(module_name, None)

                    module = imp.load_source(module_name, script_path)

                    if hasattr(module, "Execute") and callable(module.Execute):
                        result = module.Execute(doc)

                        if not isinstance(result, dict):
                            out.print_md("  :white_check_mark: {}".format(script_name))
                            scripts_succeeded += 1
                            continue

                        success = result.get("success", False)
                        message = result.get("message", "")
                        parameters = result.get("parameters", {})
                        fill = result.get("fill", {})
                        diagnostics = result.get("diagnostics", {})
                        debug_error = (result.get("info") or {}).get("debug_error")

                        if success:
                            out.print_md("  :white_check_mark: {}".format(script_name))

                            if message:
                                out.print_md("  &nbsp;&nbsp;{}".format(message))

                            if parameters:
                                added = len(parameters.get("added", []))
                                existing = len(parameters.get("existing", []))
                                failed = len(parameters.get("failed", []))
                                if added > 0 or existing > 0 or failed > 0:
                                    out.print_md(
                                        "  &nbsp;&nbsp;Параметры: добавлено {}, существует {}, ошибок {}".format(
                                            added, existing, failed
                                        )
                                    )

                            if fill:
                                target_param = fill.get("target_param", "")
                                filled = fill.get("filled", False)
                                planned_value = fill.get("planned_value", None)
                                total = fill.get("total", 0)
                                updated = fill.get("updated_count", 0)
                                skipped = fill.get("skipped_count", 0)
                                skip_reasons = fill.get("skip_reasons", {})
                                values = fill.get("values", [])
                                fill_message = fill.get("message", "")

                                if target_param:
                                    out.print_md(
                                        "  &nbsp;&nbsp;Целевой параметр: {}".format(
                                            target_param
                                        )
                                    )

                                if total > 0:
                                    out.print_md(
                                        "  &nbsp;&nbsp;Элементы: всего {}, обновлено {}, пропущено {}".format(
                                            total, updated, skipped
                                        )
                                    )

                                if planned_value:
                                    out.print_md(
                                        "  &nbsp;&nbsp;Запланировано: {}".format(
                                            planned_value
                                        )
                                    )

                                if skip_reasons:
                                    reasons_str = ", ".join(
                                        [
                                            "{}={}".format(k, v)
                                            for k, v in skip_reasons.items()
                                            if v > 0
                                        ]
                                    )
                                    if reasons_str:
                                        out.print_md(
                                            "  &nbsp;&nbsp;Причины пропуска: {}".format(
                                                reasons_str
                                            )
                                        )

                                if values:
                                    values_str = ", ".join(str(v) for v in values)
                                    if len(values_str) > 100:
                                        values_str = values_str[:97] + "..."
                                    out.print_md(
                                        "  &nbsp;&nbsp;Значения: {}".format(values_str)
                                    )

                                if fill_message:
                                    out.print_md(
                                        "  &nbsp;&nbsp;{}".format(fill_message)
                                    )

                            scripts_succeeded += 1
                        else:
                            out.print_md("  :x: {}".format(script_name))
                            if message:
                                out.print_md("  &nbsp;&nbsp;{}".format(message))
                            if debug_error:
                                out.print_md(
                                    "  &nbsp;&nbsp;DEBUG: {}".format(debug_error)
                                )
                            scripts_failed += 1
                    else:
                        out.print_md("  :white_check_mark: {}".format(script_name))
                        scripts_succeeded += 1

                except Exception as e:
                    out.print_md("  :x: {} - {}".format(script_name, e))
                    scripts_failed += 1
                finally:
                    sys.modules.pop(module_name, None)

            if scripts_failed == 0 and scripts_succeeded > 0:
                success_docs += 1

        except Exception as e:
            out.print_md(":x: Ошибка при выполнении скриптов: {}".format(e))
            scripts_failed += len(scripts)

        out.update_progress(i + 1, len(docs))

    out.print_md("---")
    out.print_md(
        "**Готово. Документов обработано: {}/{}**".format(success_docs, len(docs))
    )


def action_run_python_script(models, scripts):
    """Выполнить python скрипты из библиотеки на выбранных моделях."""
    script_names = [os.path.basename(s) for s in scripts]

    out.print_md("## ВЫПОЛНЕНИЕ PYTHON СКРИПТОВ")
    out.print_md("**Скрипты ({}):** {}".format(len(scripts), ", ".join(script_names)))
    out.print_md("**Моделей:** {}".format(len(models)))
    out.print_md("---")

    success_models = 0

    for i, model_path in enumerate(models):
        model_name = os.path.basename(model_path)
        out.print_md(":open_file_folder: **{}**".format(model_name))

        mp = to_model_path(model_path)
        if mp is None:
            out.print_md(":x: Не удалось преобразовать путь.")
            out.update_progress(i + 1, len(models))
            continue

        dialog_suppressor = None

        try:
            result = openbg.open_in_background(
                __revit__,
                None,
                mp,
                audit=False,
                worksets=("all_except_prefixes", ["00_", "Связь", "Links"]),
                detach=False,
                suppress_warnings=True,
                suppress_dialogs=True,
            )
            if len(result) >= 3:
                doc = result[0]
                dialog_suppressor = result[2]
            else:
                raise Exception("openbg не вернул документ")

            # Проверяем валидность документа
            if not doc or not doc.IsValidObject:
                out.print_md(":warning: Документ невалиден, пропускаем эту модель")
                out.update_progress(i + 1, len(models))
                continue

            dialog_summary = dialog_suppressor.get_summary()
            out.print_md(
                "  dialogs: total={}, suppressed={}, unknown={}, errors={}".format(
                    dialog_summary.get("total", 0),
                    dialog_summary.get("suppressed", 0),
                    dialog_summary.get("unknown", 0),
                    dialog_summary.get("errors", 0),
                )
            )
            if dialog_summary.get("transmitted_dialog_handled"):
                out.print_md("  :white_check_mark: transmitted_dialog_handled")
            if dialog_summary.get("unknown", 0) > 0:
                unknown = dialog_summary.get("unknown_details", [])[:3]
                for uid, snippet in unknown:
                    out.print_md("  unknown_dialog: [{}] {}".format(uid, snippet[:50]))

            out.print_md(
                "  doc: workshared={}, readonly={}, path={}".format(
                    doc.IsWorkshared,
                    doc.IsReadOnly,
                    os.path.basename(doc.PathName) if doc.PathName else "None",
                )
            )

            scripts_succeeded = 0
            scripts_failed = 0

            for script_rel_path in scripts:
                script_path = os.path.join(PYTHON_SCRIPTS_DIR, script_rel_path)
                script_name = os.path.basename(script_rel_path)

                module_name = (
                    os.path.splitext(script_rel_path)[0]
                    .replace("\\", "_")
                    .replace("/", "_")
                )

                try:
                    sys.modules.pop(module_name, None)

                    module = imp.load_source(module_name, script_path)

                    if hasattr(module, "Execute") and callable(module.Execute):
                        result = module.Execute(doc)

                        if not isinstance(result, dict):
                            out.print_md("  :white_check_mark: {}".format(script_name))
                            scripts_succeeded += 1
                            continue

                        success = result.get("success", False)
                        message = result.get("message", "")
                        parameters = result.get("parameters", {})
                        fill = result.get("fill", {})
                        diagnostics = result.get("diagnostics", {})
                        debug_error = (result.get("info") or {}).get("debug_error")

                        if success:
                            confirmed = diagnostics.get("bound_after_operation", None)
                            show_warning = (
                                "bound_after_operation" in diagnostics and not confirmed
                            )

                            if show_warning:
                                out.print_md(
                                    "  :white_check_mark: {} - :warning: Операция не подтверждена".format(
                                        script_name
                                    )
                                )
                            else:
                                out.print_md(
                                    "  :white_check_mark: {}".format(script_name)
                                )

                            if message:
                                out.print_md("  &nbsp;&nbsp;{}".format(message))

                            if parameters:
                                added = len(parameters.get("added", []))
                                existing = len(parameters.get("existing", []))
                                failed = len(parameters.get("failed", []))
                                if added > 0 or existing > 0 or failed > 0:
                                    out.print_md(
                                        "  &nbsp;&nbsp;Параметры: добавлено {}, существует {}, ошибок {}".format(
                                            added, existing, failed
                                        )
                                    )

                            if fill:
                                target_param = fill.get("target_param", "")
                                filled = fill.get("filled", False)
                                planned_value = fill.get("planned_value", None)
                                total = fill.get("total", 0)
                                updated = fill.get("updated_count", 0)
                                skipped = fill.get("skipped_count", 0)
                                skip_reasons = fill.get("skip_reasons", {})
                                values = fill.get("values", [])
                                fill_message = fill.get("message", "")

                                if target_param:
                                    out.print_md(
                                        "  &nbsp;&nbsp;Целевой параметр: {}".format(
                                            target_param
                                        )
                                    )

                                if total > 0:
                                    out.print_md(
                                        "  &nbsp;&nbsp;Элементы: всего {}, обновлено {}, пропущено {}".format(
                                            total, updated, skipped
                                        )
                                    )

                                if planned_value:
                                    out.print_md(
                                        "  &nbsp;&nbsp;Запланировано: {}".format(
                                            planned_value
                                        )
                                    )

                                if skip_reasons:
                                    reasons_str = ", ".join(
                                        [
                                            "{}={}".format(k, v)
                                            for k, v in skip_reasons.items()
                                            if v > 0
                                        ]
                                    )
                                    if reasons_str:
                                        out.print_md(
                                            "  &nbsp;&nbsp;Причины пропуска: {}".format(
                                                reasons_str
                                            )
                                        )

                                if values:
                                    values_str = ", ".join(str(v) for v in values)
                                    if len(values_str) > 100:
                                        values_str = values_str[:97] + "..."
                                    out.print_md(
                                        "  &nbsp;&nbsp;Значения: {}".format(values_str)
                                    )

                                if fill_message:
                                    out.print_md(
                                        "  &nbsp;&nbsp;{}".format(fill_message)
                                    )

                            scripts_succeeded += 1
                        else:
                            out.print_md("  :x: {}".format(script_name))
                            if message:
                                out.print_md("  &nbsp;&nbsp;{}".format(message))
                            if debug_error:
                                out.print_md(
                                    "  &nbsp;&nbsp;DEBUG: {}".format(debug_error)
                                )
                            scripts_failed += 1
                    else:
                        out.print_md("  :white_check_mark: {}".format(script_name))
                        scripts_succeeded += 1

                except Exception as e:
                    out.print_md("  :x: {} - {}".format(script_name, e))
                    scripts_failed += 1
                finally:
                    sys.modules.pop(module_name, None)

            out.print_md(
                "  before_close: readonly={}, path={}".format(
                    doc.IsReadOnly,
                    os.path.basename(doc.PathName) if doc.PathName else "None",
                )
            )

            # Закрываем документ ОДИН РАЗ после всех скриптов
            res = closebg.close_with_policy(
                doc,
                do_sync=True,
                comment="Выполнение python скриптов",
                dialog_suppressor=dialog_suppressor,
                source_path=model_path,
            )

            if not res.get("success"):
                out.print_md(
                    ":x: Ошибка закрытия: {}".format(
                        res.get(
                            "save_error", res.get("close_error", "Неизвестная ошибка")
                        )
                    )
                )
                if doc and doc.IsValidObject:
                    out.print_md(
                        "  doc_state: readonly={}, path={}".format(
                            doc.IsReadOnly,
                            os.path.basename(doc.PathName) if doc.PathName else "None",
                        )
                    )
                if dialog_summary.get("transmitted_dialog_handled"):
                    out.print_md("  :white_check_mark: transmitted_dialog_handled")

            if scripts_failed == 0:
                success_models += 1

        except Exception as e:
            out.print_md(":x: Ошибка открытия: `{}`".format(e))
        finally:
            if dialog_suppressor:
                dialog_suppressor.detach()

        out.update_progress(i + 1, len(models))

    out.print_md("---")
    out.print_md(
        "**Готово. Моделей обработано: {}/{}**".format(success_models, len(models))
    )


# ---------- Действия в разработке ----------


def action_create_worksets(models):
    """Создать рабочие наборы (в разработке)."""
    forms.alert(
        "Функция 'Создать рабочие наборы' находится в разработке.",
        title="В разработке",
        warn_icon=True,
    )


def action_add_shared_params(models):
    """Добавить общие параметры (в разработке)."""
    forms.alert(
        "Функция 'Добавить общие параметры' находится в разработке.",
        title="В разработке",
        warn_icon=True,
    )


def list_python_scripts():
    """Получить список python скриптов из папки lib/Пакетные операции."""
    scripts = []
    if not PYTHON_SCRIPTS_DIR:
        print("PYTHON_SCRIPTS_DIR не задан")
        return scripts

    if not os.path.isdir(PYTHON_SCRIPTS_DIR):
        print("Папка не существует: {}".format(PYTHON_SCRIPTS_DIR))
        return scripts

    print("Поиск скриптов в: {}".format(PYTHON_SCRIPTS_DIR))

    for root, dirs, files in os.walk(PYTHON_SCRIPTS_DIR):
        print("Папка: {}, файлов: {}".format(root, len(files)))
        for f in files:
            if f.lower().endswith(".py"):
                rel_path = os.path.relpath(os.path.join(root, f), PYTHON_SCRIPTS_DIR)
                scripts.append(rel_path)
                print("Найден: {}".format(rel_path))

    print("Всего найдено скриптов: {}".format(len(scripts)))
    return sorted(scripts)


# ---------- Main ----------


def main():
    # 1. Выбор действия
    actions = [
        "Открыть модели",
        "Загрузить семейство (из открытых)",
        "Добавить связь",
        "Выполнение python скриптов из библиотеки",
        "Выполнить python скрипты в открытых документах",
        "Создать рабочие наборы (в разработке)",
        "Добавить общие параметры (в разработке)",
    ]

    selected_action = forms.SelectFromList.show(
        actions, title="Выберите действие", width=400, button_name="Далее"
    )

    if not selected_action:
        script.exit()

    selected_scripts = None

    # Если выбрано выполнение python скриптов - сначала выбрать скрипты
    if selected_action in [
        "Выполнение python скриптов из библиотеки",
        "Выполнить python скрипты в открытых документах",
    ]:
        scripts = list_python_scripts()

        if not scripts:
            forms.alert(
                "Python скрипты не найдены в папке:\n{}".format(PYTHON_SCRIPTS_DIR),
                warn_icon=True,
            )
            script.exit()

        selected_scripts = forms.SelectFromList.show(
            scripts,
            title="Выберите python скрипты",
            multiselect=True,
            width=600,
            height=400,
            button_name="Далее",
        )

        if not selected_scripts:
            script.exit()

        if not isinstance(selected_scripts, list):
            selected_scripts = [selected_scripts]

    # Для выполнения скриптов на открытых документах - пропускаем выбор объекта и моделей
    if selected_action == "Выполнить python скрипты в открытых документах":
        selected_docs = select_open_documents()

        if not selected_docs:
            script.exit()

        # 4. Выполнение действия
        out.print_md("# Пакетные операции")
        out.print_md("**Действие:** {}".format(selected_action))
        out.print_md("**Документов:** {}".format(len(selected_docs)))
        out.print_md("**Скриптов:** {}".format(len(selected_scripts)))
        out.print_md("")

        out.update_progress(0, len(selected_docs))

        action_run_python_script_on_opened_docs(selected_docs, selected_scripts)
        return

    # 2. Выбор объекта (txt файла)
    txt_files = list_txt_files(OBJECTS_DIR)

    if not txt_files:
        forms.alert(
            "Папка с объектами пуста или недоступна:\n{}".format(OBJECTS_DIR),
            warn_icon=True,
        )
        script.exit()

    selected_object = forms.SelectFromList.show(
        txt_files, title="Выберите объект", width=400, button_name="Далее"
    )

    if not selected_object:
        script.exit()

    # 3. Выбор моделей из списка
    txt_path = os.path.join(OBJECTS_DIR, selected_object + ".txt")
    models_list = read_models_from_txt(txt_path)

    if not models_list:
        forms.alert("Файл пуст или не найден: {}".format(txt_path), warn_icon=True)
        script.exit()

    selected_models = forms.SelectFromList.show(
        models_list,
        title="Выберите модели",
        multiselect=True,
        width=800,
        height=600,
        button_name="Выполнить",
    )

    if not selected_models:
        script.exit()

    if not isinstance(selected_models, list):
        selected_models = [selected_models]

    # 4. Выполнение действия
    out.print_md("# Пакетные операции")
    out.print_md("**Действие:** {}".format(selected_action))
    out.print_md("**Объект:** {}".format(selected_object))
    if selected_scripts:
        out.print_md("**Скриптов:** {}".format(len(selected_scripts)))
    out.print_md("")

    out.update_progress(0, len(selected_models))

    if selected_action == "Открыть модели":
        action_open_models(selected_models)
    elif selected_action == "Загрузить семейство (из открытых)":
        action_load_family(selected_models)
    elif selected_action == "Добавить связь":
        action_add_link(selected_models)
    elif selected_action == "Выполнение python скриптов из библиотеки":
        action_run_python_script(selected_models, selected_scripts)
    elif "Создать рабочие наборы" in selected_action:
        action_create_worksets(selected_models)
    elif "Добавить общие параметры" in selected_action:
        action_add_shared_params(selected_models)


if __name__ == "__main__":
    main()
