# -*- coding: utf-8 -*-
"""
openbg.py — фон. открытие RVT + безопасная подготовка 3D‑вида для NWC.
Важное изменение: полностью исключён прямой доступ к BuiltInCategory через getattr().
Теперь используется только Enum.IsDefined/Enum.Parse — это устраняет AttributeError
в сборках, где, например, отсутствует OST_ImportsInFamilies (Revit 2022).
"""

from Autodesk.Revit.DB import (
    ModelPathUtils,
    WorksharingUtils,
    WorksetId,
    OpenOptions,
    WorksetConfiguration,
    WorksetConfigurationOption,
    DetachFromCentralOption,
    BuiltInCategory,
    Category,
    ElementId,
    View3D,
    ViewFamilyType,
    ViewFamily,
    Transaction,
    FilteredElementCollector,
    IFailuresPreprocessor,
    FailureProcessingResult,
    FailureSeverity,
)
from Autodesk.Revit.UI.Events import (
    DialogBoxShowingEventArgs,
    TaskDialogShowingEventArgs,
    MessageBoxShowingEventArgs,
)
from System.Collections.Generic import List
from System import Enum

# ---------------- Failures Processor ----------------


class SuppressWarningsPreprocessor(IFailuresPreprocessor):
    """
    Автоматический обработчик предупреждений и ошибок Revit.
    Удаляет все предупреждения, чтобы диалоги не блокировали выполнение скрипта.
    Собирает информацию об обработанных предупреждениях/ошибках для отчета.
    """

    def __init__(self):
        self.warnings = []
        self.errors = []

    def PreprocessFailures(self, failuresAccessor):
        try:
            failures = failuresAccessor.GetFailureMessages()
            for failure in failures:
                try:
                    severity = failure.GetSeverity()
                    desc = ""
                    try:
                        desc = failure.GetDescriptionText() or ""
                    except Exception:
                        pass

                    # Удаляем предупреждения (Warning)
                    if severity == FailureSeverity.Warning:
                        self.warnings.append(desc)
                        failuresAccessor.DeleteWarning(failure)
                    # Для ошибок пытаемся использовать дефолтное решение
                    elif severity == FailureSeverity.Error:
                        self.errors.append(desc)
                        try:
                            failuresAccessor.ResolveFailure(failure)
                        except Exception:
                            pass
                except Exception:
                    pass
            return FailureProcessingResult.Continue
        except Exception:
            return FailureProcessingResult.Continue

    def get_summary(self):
        """Возвращает сводку об обработанных предупреждениях и ошибках."""
        return {
            "warnings": list(self.warnings),
            "errors": list(self.errors),
            "total_warnings": len(self.warnings),
            "total_errors": len(self.errors),
        }


# ---------------- Dialog Suppressor ----------------


class DialogSuppressor(object):
    """
    Контроллер подавления диалогов Revit с режимом логирования и whitelist.

    Реализовано корректно с использованием TaskDialogResult enum.

    Режимы работы:
    - log_only: только логировать диалоги, не подавлять
    - suppress: активно подавлять диалоги согласно whitelist

    Критично: подписка на событие включается только на время операции.
    """

    def __init__(self, log_only=False):
        """
        Args:
            log_only: Если True - только логировать, не подавлять
        """
        self.suppressed_dialogs = []
        self._uiapp = None
        self._log_only = log_only
        self.transmitted_dialog_handled = 0

    def _on_task_dialog_showing(self, sender, args):
        """
        Обработчик для TaskDialogShowingEventArgs.

        Реализовано по whitelist-подходу из C# примера Application.cs.

        НЕ использует "link/reference" напрямую - слишком широкая проверка.
        Использует DialogId для точного определения (как в C#).
        """
        try:
            from Autodesk.Revit.UI import TaskDialogResult

            dialog_id = args.DialogId or ""
            message = args.Message

            # Для диагностики: если DialogId содержит "transmitted" или текст соответствует, логируем детально
            main_instruction = ""
            try:
                main_instruction = args.MainInstruction or ""
            except Exception:
                pass

            # Логируем информацию о диалоге
            dialog_info = {
                "dialog_id": dialog_id,
                "type": "TaskDialogShowingEventArgs",
                "message": message,
                "main_instruction": main_instruction,
                "action": "logged" if self._log_only else "unknown",
            }
            self.suppressed_dialogs.append(dialog_info)

            # Если режим только логирования - не подавляем
            if self._log_only:
                return

            message_lower = (message or "").lower()
            main_instruction_lower = (main_instruction or "").lower()
            identifier_lower = dialog_id.lower()

            # 1. Проверка по DialogId (как в C# IsSuppressedDialog)
            if any(
                kw in identifier_lower
                for kw in ["navisworks", "docwarndialog", "coordination", "review"]
            ):
                dialog_info["action"] = "suppressed (OK)"
                args.OverrideResult(TaskDialogResult.Ok)
                return

            # 2. Coordination Review → OK
            if (
                "coordination review" in message_lower
                or "требуется просмотр координаций" in message_lower
            ):
                dialog_info["action"] = "suppressed (OK)"
                args.OverrideResult(TaskDialogResult.Ok)
                return

            # 3. Missing/Unresolved links → No (ТОЛЬКО точные фразы)
            # НЕ используем общие "link"/"reference" - слишком широко
            # Проверяем контекст - только если явно про связи
            if any(
                kw in message_lower
                for kw in [
                    "не удается найти",
                    "cannot find",
                    "не найден",
                    "unresolved",
                    "missing file",
                    "отсутствующий файл",
                ]
            ):
                if any(
                    kw in message_lower
                    for kw in ["связ", "ссылка", "link", "reference", "file"]
                ):
                    dialog_info["action"] = "suppressed (No - open without search)"
                    args.OverrideResult(TaskDialogResult.No)
                    return

            # 4. Geometry errors / No suitable geometry → OK
            if any(
                kw in message_lower
                for kw in [
                    "не найдена подходящая геометрия",
                    "элементов для геометрии не найдено",
                    "не найдена геометрия для экспорта",
                    "no suitable geometry",
                    "no appropriate geometry",
                    "no elements for geometry",
                    "no geometry for export",
                ]
            ):
                dialog_info["action"] = "suppressed (OK)"
                args.OverrideResult(TaskDialogResult.Ok)
                return

            # 4. Room tag outside room → OK
            if any(
                kw in message_lower
                for kw in [
                    "марка помещение вне элемента помещение",
                    "room tag is outside of its room",
                ]
            ):
                dialog_info["action"] = "suppressed (OK)"
                args.OverrideResult(TaskDialogResult.Ok)
                return

            # 5. Dimension reference / Wrong dimension → OK
            if any(
                kw in message_lower
                for kw in ["опорных элементов размеров", "dimension reference"]
            ):
                dialog_info["action"] = "suppressed (OK)"
                args.OverrideResult(TaskDialogResult.Ok)
                return

            # 6. Transmitted model dialog → CommandLink1 (Save this model)
            # Проверяем по DialogId и тексту диалога (рус/англ) - в message И main_instruction
            full_text = message_lower + " " + main_instruction_lower
            if "transmitted" in identifier_lower or any(
                kw in full_text
                for kw in [
                    "переданная модель",
                    "данная модель была получена из другой папки",
                    "transmitted model",
                    "received from another folder",
                ]
            ):
                dialog_info["action"] = "suppressed (CommandLink1 - Save this model)"
                args.OverrideResult(TaskDialogResult.CommandLink1)
                self.transmitted_dialog_handled += 1
                return

            # Неизвестный диалог - не трогаем, пусть Revit покажет
            dialog_info["action"] = "unknown (not suppressed)"

        except Exception as e:
            self.suppressed_dialogs.append(
                {
                    "dialog_id": None,
                    "type": "TaskDialogShowingEventArgs-Error",
                    "message": str(e),
                    "action": "handler_error",
                }
            )

    def _on_message_box_showing(self, sender, args):
        """
        Обработчик для MessageBoxShowingEventArgs.

        Используем целочисленные ID кнопок как в C#:
        - IDOK = 1
        - IDYES = 6
        - IDNO = 7
        - IDCANCEL = 2
        """
        try:
            dialog_id = None
            try:
                dialog_id = args.DialogId
            except Exception:
                pass

            message = args.Message

            # Логируем информацию о диалоге
            dialog_info = {
                "dialog_id": dialog_id,
                "type": "MessageBoxShowingEventArgs",
                "message": message,
                "action": "logged" if self._log_only else "unknown",
            }
            self.suppressed_dialogs.append(dialog_info)

            # Если режим только логирования - не подавляем
            if self._log_only:
                return

            # Для MessageBox используем целочисленные ID (как в C# примере)
            message_lower = (message or "").lower()

            # Coordination Review → OK (1)
            if "coordination review" in message_lower:
                dialog_info["action"] = "suppressed (OK)"
                args.OverrideResult(1)
                return

            # Missing/Unresolved links → No (ТОЛЬКО точные фразы)
            if any(
                kw in message_lower
                for kw in [
                    "не удается найти",
                    "cannot find",
                    "не найден",
                    "unresolved",
                    "missing file",
                    "отсутствующий файл",
                ]
            ):
                if any(
                    kw in message_lower
                    for kw in ["связ", "ссылка", "link", "reference", "file"]
                ):
                    dialog_info["action"] = "suppressed (No - open without search)"
                    args.OverrideResult(7)
                    return

            # Geometry errors / Minor warnings → OK (1)
            if any(
                kw in message_lower
                for kw in [
                    "не найдена подходящая геометрия",
                    "элементов для геометрии не найдено",
                    "не найдена геометрия для экспорта",
                    "no suitable geometry",
                    "no appropriate geometry",
                    "no elements for geometry",
                    "no geometry for export",
                ]
            ):
                dialog_info["action"] = "suppressed (OK)"
                args.OverrideResult(1)
                return

            # Room tag outside room → OK (1)
            if any(
                kw in message_lower
                for kw in [
                    "марка помещение вне элемента помещение",
                    "room tag is outside of its room",
                ]
            ):
                dialog_info["action"] = "suppressed (OK)"
                args.OverrideResult(1)
                return

            # Dimension reference → OK (1)
            if any(
                kw in message_lower
                for kw in ["опорных элементов размеров", "dimension reference"]
            ):
                dialog_info["action"] = "suppressed (OK)"
                args.OverrideResult(1)
                return

            # Неизвестный диалог - не трогаем, пусть Revit покажет
            dialog_info["action"] = "unknown (not suppressed)"

        except Exception as e:
            # Логируем ошибку обработки
            self.suppressed_dialogs.append(
                {
                    "dialog_id": None,
                    "type": "MessageBoxShowingEventArgs-Error",
                    "message": str(e),
                    "action": "handler_error",
                }
            )

    def _on_dialog_showing(self, sender, args):
        """
        Главный обработчик события DialogBoxShowing.
        Разделяет типы диалогов и перенаправляет на соответствующие обработчики.
        """
        try:
            if isinstance(args, TaskDialogShowingEventArgs):
                self._on_task_dialog_showing(sender, args)
            elif isinstance(args, MessageBoxShowingEventArgs):
                self._on_message_box_showing(sender, args)
            else:
                # Неизвестный тип аргумента - логируем
                dialog_info = {
                    "dialog_id": getattr(args, "DialogId", None),
                    "type": type(args).__name__,
                    "message": getattr(args, "Message", None),
                    "action": "unknown type",
                }
                self.suppressed_dialogs.append(dialog_info)
        except Exception as e:
            self.suppressed_dialogs.append(
                {
                    "dialog_id": None,
                    "type": "DialogBoxShowing-Error",
                    "message": str(e),
                    "action": "handler_error",
                }
            )

    def attach(self, uiapp):
        """Подключить обработчик к UIApplication."""
        self._uiapp = uiapp
        if uiapp is not None:
            try:
                uiapp.DialogBoxShowing += self._on_dialog_showing
            except Exception:
                pass

    def detach(self):
        """Отключить обработчик от UIApplication."""
        if self._uiapp is not None:
            try:
                self._uiapp.DialogBoxShowing -= self._on_dialog_showing
            except Exception:
                pass
            self._uiapp = None

    def get_summary(self):
        """Возвращает сводку о диалогах."""
        suppressed = [
            d for d in self.suppressed_dialogs if "suppressed" in d.get("action", "")
        ]
        logged = [d for d in self.suppressed_dialogs if d.get("action") == "logged"]
        unknown = [
            d for d in self.suppressed_dialogs if "unknown" in d.get("action", "")
        ]
        errors = [
            d for d in self.suppressed_dialogs if d.get("action") == "handler_error"
        ]

        # Формируем краткое описание unknown диалогов (первые 5)
        unknown_dialogs_brief = []
        max_brief = 5
        for d in unknown[:max_brief]:
            dialog_id = d.get("dialog_id", "")
            message = d.get("message", "") or ""
            main_instruction = d.get("main_instruction", "") or ""
            # Объединяем message и main_instruction, обрезаем до 120 символов
            combined_text = (main_instruction + " " + message).strip()[:120]
            unknown_dialogs_brief.append(
                {
                    "dialog_id": dialog_id,
                    "text_preview": combined_text,
                }
            )

        return {
            "total": len(self.suppressed_dialogs),
            "suppressed": len(suppressed),
            "logged": len(logged),
            "unknown": len(unknown),
            "errors": len(errors),
            "transmitted_dialog_handled": self.transmitted_dialog_handled,
            "dialogs": list(self.suppressed_dialogs),
            "suppressed_dialogs": suppressed,
            "logged_dialogs": logged,
            "unknown_dialogs": unknown,
            "error_dialogs": errors,
            "unknown_dialogs_brief": unknown_dialogs_brief,
        }

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.detach()
        return False


# ---------------- helpers ----------------


def _enum(name):
    try:
        return Enum.Parse(WorksetConfigurationOption, name)
    except Exception:
        return Enum.Parse(WorksetConfigurationOption, "OpenAllWorksets")


def _cfg_from_optname(name):
    opt = _enum(name)
    return WorksetConfiguration(opt)


def _is_string(x):
    try:
        return isinstance(x, basestring)
    except NameError:
        try:
            return isinstance(x, str) or isinstance(x, unicode)
        except NameError:
            return isinstance(x, str)


def _coerce_app_uiapp(app_or_uiapp, maybe_uiapp=None):
    if maybe_uiapp is not None:
        return app_or_uiapp, maybe_uiapp
    try:
        _ = app_or_uiapp.Application
        uiapp = app_or_uiapp
        app = uiapp.Application
    except Exception:
        app = app_or_uiapp
        uiapp = None
    return app, uiapp


def _to_model_path(path_or_mp):
    try:
        if hasattr(path_or_mp, "ServerPath") or hasattr(
            path_or_mp, "CentralServerPath"
        ):
            return path_or_mp
    except Exception:
        pass
    return ModelPathUtils.ConvertUserVisiblePathToModelPath(path_or_mp)


def _get_workset_previews(uiapp, mp):
    previews = None
    if uiapp is not None:
        try:
            previews = WorksharingUtils.GetUserWorksetInfoForOpen(uiapp, mp)
        except Exception:
            previews = None
    if previews is None:
        try:
            previews = WorksharingUtils.GetUserWorksetInfo(mp)
        except Exception:
            previews = []
    return list(previews or [])


def _ids_all_except_prefixes(previews, prefixes):
    ids = List[WorksetId]()
    for p in previews:
        try:
            nm = p.Name or ""
            # Проверяем префиксы (startswith) и подстроки (in)
            if not any(
                nm.startswith(pr) or pr.lower() in nm.lower() for pr in prefixes
            ):
                ids.Add(p.Id)
        except Exception:
            pass
    return ids


def _ids_only_prefixes(previews, prefixes):
    ids = List[WorksetId]()
    for p in previews:
        try:
            nm = p.Name or ""
            # Проверяем префиксы (startswith) и подстроки (in)
            if any(nm.startswith(pr) or pr.lower() in nm.lower() for pr in prefixes):
                ids.Add(p.Id)
        except Exception:
            pass
    return ids


def _ids_only_names(previews, names, case_sensitive=False):
    ids = List[WorksetId]()
    if not names:
        return ids
    if case_sensitive:
        names_set = set(names)
        for p in previews:
            try:
                if (p.Name or "") in names_set:
                    ids.Add(p.Id)
            except Exception:
                pass
    else:
        names_l = set([(n or "").lower() for n in names])
        for p in previews:
            try:
                if (p.Name or "").lower() in names_l:
                    ids.Add(p.Id)
            except Exception:
                pass
    return ids


def _ids_by_predicate(previews, pred):
    ids = List[WorksetId]()
    if not callable(pred):
        return ids
    for p in previews:
        try:
            if pred(p.Name or ""):
                ids.Add(p.Id)
        except Exception:
            pass
    return ids


def _build_ws_config(uiapp, mp, worksets_rule):
    if _is_string(worksets_rule):
        key = (worksets_rule or "").strip().lower()
        if key in ("all", "open_all"):
            return _cfg_from_optname("OpenAllWorksets")
        if key in ("close", "close_all"):
            return _cfg_from_optname("CloseAllWorksets")
        if key in ("lastviewed", "last_viewed", "last"):
            return _cfg_from_optname("LastViewed")

    previews = _get_workset_previews(uiapp, mp)
    if not previews:
        return _cfg_from_optname("LastViewed")

    cfg = _cfg_from_optname("CloseAllWorksets")

    if (
        _is_string(worksets_rule)
        and (worksets_rule or "").strip().lower() == "all_except_00"
    ):
        ids = _ids_all_except_prefixes(previews, ("00_",))
        if ids.Count > 0:
            cfg.Open(ids)
        return cfg

    if isinstance(worksets_rule, tuple) and len(worksets_rule) > 0:
        mode = (
            ((worksets_rule[0] or "") if len(worksets_rule) > 0 else "").strip().lower()
        )
        if mode == "all_except_prefixes":
            prefixes = tuple(worksets_rule[1]) if len(worksets_rule) > 1 else ("00_",)
            ids = _ids_all_except_prefixes(previews, prefixes)
            if ids.Count > 0:
                cfg.Open(ids)
            return cfg
        if mode == "only_prefixes":
            prefixes = tuple(worksets_rule[1]) if len(worksets_rule) > 1 else tuple()
            ids = _ids_only_prefixes(previews, prefixes)
            if ids.Count > 0:
                cfg.Open(ids)
            return cfg
        if mode == "only_names":
            names = tuple(worksets_rule[1]) if len(worksets_rule) > 1 else tuple()
            ids = _ids_only_names(previews, names, case_sensitive=False)
            if ids.Count > 0:
                cfg.Open(ids)
            return cfg
        if mode == "predicate":
            pred = worksets_rule[1] if len(worksets_rule) > 1 else None
            ids = _ids_by_predicate(previews, pred)
            if ids.Count > 0:
                cfg.Open(ids)
            return cfg

    if isinstance(worksets_rule, dict):
        mode = (worksets_rule.get("mode") or "").strip().lower()
        if mode == "all_except_prefixes":
            prefixes = tuple(worksets_rule.get("prefixes") or ("00_",))
            ids = _ids_all_except_prefixes(previews, prefixes)
            if ids.Count > 0:
                cfg.Open(ids)
            return cfg
        if mode == "only_prefixes":
            prefixes = tuple(worksets_rule.get("prefixes") or tuple())
            ids = _ids_only_prefixes(previews, prefixes)
            if ids.Count > 0:
                cfg.Open(ids)
            return cfg
        if mode == "only_names":
            names = tuple(worksets_rule.get("names") or tuple())
            case = bool(worksets_rule.get("case_sensitive", False))
            ids = _ids_only_names(previews, names, case_sensitive=case)
            if ids.Count > 0:
                cfg.Open(ids)
            return cfg

    return _cfg_from_optname("LastViewed")


# ----------- BIC safe -----------


def _resolve_bic(name):
    """Вернуть BuiltInCategory по имени или None, если такой константы нет в этой версии."""
    if not name:
        return None
    try:
        if Enum.IsDefined(BuiltInCategory, name):
            return Enum.Parse(BuiltInCategory, name)
    except Exception:
        pass
    return None


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

    if ids.Count == 0:
        return 0

    hidden = 0
    try:
        view.HideCategories(ids)
        return ids.Count
    except Exception:
        pass

    for eid in ids:
        try:
            view.SetCategoryHidden(eid, True)
            hidden += 1
        except Exception:
            try:
                view.SetCategoryHidden(eid.IntegerValue, True)
                hidden += 1
            except Exception:
                pass
    return hidden


# ----------- public API -----------


def open_in_background(
    app_or_uiapp,
    maybe_uiapp,
    model_path_or_str,
    audit=False,
    worksets="lastviewed",
    detach=False,
    suppress_warnings=True,
    suppress_dialogs=True,
    log_only=False,
):
    """
    Открыть документ в фоне.

    Args:
        app_or_uiapp: Application или UIApplication. Для работы подавления диалогов
                     (suppress_dialogs=True) необходимо передать UIApplication (в pyRevit это __revit__)
        maybe_uiapp: UIApplication (опционально, если app_or_uiapp — Application)
        model_path_or_str: путь к модели (строка или ModelPath)
        detach: если True — открыть с опцией "Отсоединить с сохранением рабочих наборов"
                (DetachAndPreserveWorksets)
        suppress_warnings: если True — автоматически подавлять предупреждения и ошибки при открытии
                          (через IFailuresPreprocessor)
        suppress_dialogs: если True — автоматически закрывать диалоговые окна Revit
                         (через DialogBoxShowing event). ВАЖНО: для работы требует UIApplication!
        log_only: если True — только логировать диалоги, не подавлять (для отладки)

    Returns:
        tuple: (doc, failure_handler, dialog_suppressor) - документ, обработчик предупреждений и подавитель диалогов
               Используйте failure_handler.get_summary() для информации об обработанных ошибках/предупреждениях
               Используйте dialog_suppressor.get_summary() для информации о диалогах
               ВАЖНО: dialog_suppressor остаётся активным после открытия! Отключите его вызовом detach()
                      когда закончите работу с документом
    """
    app, uiapp = _coerce_app_uiapp(app_or_uiapp, maybe_uiapp)
    mp = _to_model_path(model_path_or_str)

    cfg = _build_ws_config(uiapp, mp, worksets)

    opts = OpenOptions()
    try:
        opts.Audit = bool(audit)
    except Exception:
        pass
    try:
        opts.SetOpenWorksetsConfiguration(cfg)
    except Exception:
        pass

    # Отсоединить с сохранением рабочих наборов
    if detach:
        try:
            opts.DetachFromCentralOption = (
                DetachFromCentralOption.DetachAndPreserveWorksets
            )
        except Exception:
            pass

    # Обработчик предупреждений транзакций (Failures API)
    failure_handler = None
    if suppress_warnings:
        failure_handler = SuppressWarningsPreprocessor()
        try:
            app.FailuresProcessing += failure_handler.PreprocessFailures
        except Exception:
            pass

    # Контроллер диалоговых окон (TaskDialog и др.)
    dialog_suppressor = None
    if suppress_dialogs and uiapp is not None:
        dialog_suppressor = DialogSuppressor(log_only=log_only)
        dialog_suppressor.attach(uiapp)

    try:
        try:
            doc = app.OpenDocumentFile(mp, opts)
            return (doc, failure_handler, dialog_suppressor)
        except Exception as ex1:
            msg = "{}".format(ex1)
            need_retry_lv = False
            key = (worksets or "").strip().lower() if _is_string(worksets) else ""
            if key in ("lastviewed", "last_viewed", "last"):
                need_retry_lv = True
            if "LastViewed" in msg or ("attribute" in msg and "LastViewed" in msg):
                need_retry_lv = True

            if need_retry_lv:
                try:
                    cfg2 = _build_ws_config(uiapp, mp, "all_except_00")
                    opts2 = OpenOptions()
                    try:
                        opts2.Audit = bool(audit)
                    except Exception:
                        pass
                    try:
                        opts2.SetOpenWorksetsConfiguration(cfg2)
                    except Exception:
                        pass
                    doc = app.OpenDocumentFile(mp, opts2)
                    return (doc, failure_handler, dialog_suppressor)
                except Exception:
                    cfg3 = _cfg_from_optname("OpenAllWorksets")
                    opts3 = OpenOptions()
                    try:
                        opts3.Audit = bool(audit)
                    except Exception:
                        pass
                    try:
                        opts3.SetOpenWorksetsConfiguration(cfg3)
                    except Exception:
                        pass
                    doc = app.OpenDocumentFile(mp, opts3)
                    return (doc, failure_handler, dialog_suppressor)
            raise
    finally:
        # Отписываемся от события Failures (всегда)
        if failure_handler is not None:
            try:
                app.FailuresProcessing -= failure_handler.PreprocessFailures
            except Exception:
                pass
        # НЕ отключаем dialog_suppressor здесь - он нужен для подавления диалогов
        # при последующих операциях (создание вида, экспорт и т.д.)
        # Вызывающий код должен вызвать dialog_suppressor.detach() когда закончит


def prepare_navisworks_view(doc, view):
    """Скрыть служебные/аннотационные категории. Отсутствующие — пропускаются. Возвращает число скрытых."""
    names = [
        # ссылки и импорты
        "OST_RvtLinks",
        "OST_LinkInstances",
        "OST_ExportLayer",
        "OST_ImportInstance",
        "OST_ImportsInFamilies",
        # служебное
        "OST_Cameras",
        "OST_Views",
        "OST_Lines",
        "OST_PointClouds",
        "OST_PointCloudsHardware",
        "OST_Levels",
        "OST_Grids",
        # аннотации
        "OST_Annotations",
        "OST_TitleBlocks",
        "OST_Viewports",
        "OST_TextNotes",
        "OST_Dimensions",
    ]
    return _hide_categories_by_names(doc, view, names)


def get_or_create_navisworks_view(doc, name="Navisworks"):
    """Найдёт или создаст изометрический 3D‑вид с именем name. Удобно для батч‑экспорта."""
    # поиск
    try:
        for v in FilteredElementCollector(doc).OfClass(View3D):
            try:
                if not v.IsTemplate and v.Name == name:
                    return v
            except Exception:
                pass
    except Exception:
        pass

    # создание
    vft = None
    try:
        for t in FilteredElementCollector(doc).OfClass(ViewFamilyType):
            if t.ViewFamily == ViewFamily.ThreeDimensional:
                vft = t
                break
    except Exception:
        pass
    if vft is None:
        return None

    created = None
    t = Transaction(doc, "Create 3D View for Navisworks")
    t.Start()
    try:
        created = View3D.CreateIsometric(doc, vft.Id)
        try:
            created.Name = name
        except Exception:
            pass
        t.Commit()
    except Exception:
        try:
            t.RollBack()
        except Exception:
            pass
        created = None
    return created
