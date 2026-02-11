# -*- coding: utf-8 -*-
from Autodesk.Revit.DB import (
    TransactWithCentralOptions,
    SynchronizeWithCentralOptions,
    RelinquishOptions,
    SaveOptions,
    TransmissionData,
    ModelPathUtils,
    SaveAsOptions,
    WorksharingSaveAsOptions,
)
from System.IO import File, FileAttributes
import os
import sys


def _is_transmitted(doc):
    """
    Проверяет, является ли документ переданной моделью.

    Returns:
        bool: True если документ является переданной моделью, False в противном случае
    """
    try:
        path = doc.PathName
        if not path:
            return False

        modelPath = ModelPathUtils.ConvertUserVisiblePathToModelPath(path)
        td = TransmissionData.ReadTransmissionData(modelPath)

        return bool(td and td.IsTransmitted)
    except Exception:
        return False


def _clear_readonly_attribute(file_path):
    """
    Снимает атрибут read-only с файла на диске.

    Args:
        file_path: путь к файлу

    Returns:
        bool: True если атрибут был успешно снят или файл не был read-only, False при ошибке
    """
    try:
        if File.Exists(file_path):
            attrs = File.GetAttributes(file_path)
            if attrs & FileAttributes.ReadOnly:
                File.SetAttributes(file_path, attrs & ~FileAttributes.ReadOnly)
            return True
        return True
    except Exception:
        return False


def _looks_like_detached_error(msg):
    """
    Проверяет, похоже ли сообщение об ошибке на detached/temporary/readonly проблему.

    Args:
        msg: строка с сообщением об ошибке

    Returns:
        bool: True если сообщение указывает на detached/temporary/readonly проблему
    """
    m = (msg or "").lower()
    return (
        "read-only" in m
        or "readonly" in m
        or "detached" in m
        or "central" in m
        or "transmitted" in m
    )


def _resolve_orig_path(doc):
    """
    Нормализует путь к документу до абсолютного.

    Args:
        doc: документ Revit

    Returns:
        str: абсолютный путь к документу

    Raises:
        Exception: если путь не абсолютный
    """
    p = doc.PathName or ""
    if not os.path.isabs(p):
        raise Exception("Document PathName is not absolute: '{}'".format(p))
    return p


def _get_orig_path(doc, source_path=None):
    """
    Возвращает оригинальный путь для сохранения.

    Если source_path предоставлен — использует его (для detached-копий).
    Иначе использует doc.PathName.

    Args:
        doc: документ Revit
        source_path: оригинальный путь (строка), если доступен

    Returns:
        str: путь к файлу для сохранения
    """
    if source_path:
        p = source_path
        if not os.path.isabs(p):
            p = os.path.abspath(p)
        return p
    return _resolve_orig_path(doc)


def _clear_transmission_flag(doc):
    """
    Снимает флаг Transmitted Model с документа перед сохранением/синхронизацией.

    Returns:
        bool: True если флаг был успешно снят или его не было, False при ошибке
    """
    try:
        path = doc.PathName
        if not path:
            return True

        modelPath = ModelPathUtils.ConvertUserVisiblePathToModelPath(path)
        td = TransmissionData.ReadTransmissionData(modelPath)

        if td is None:
            return True

        if hasattr(td, "IsTransmitted") and td.IsTransmitted:
            td.IsTransmitted = False
            TransmissionData.WriteTransmissionData(modelPath, td)
            return True

        return True
    except Exception:
        return False


def close_with_policy(
    doc,
    do_sync=False,
    comment="",
    compact=True,
    relinquish=True,
    save_if_not_ws=True,
    dialog_suppressor=None,
    source_path=None,
):
    """
    Если do_sync=True и документ совместный — SWC с параметрами.
    Иначе: для не-совместных при желании просто Save, и закрыть без сохранения.

    Args:
        doc: документ для закрытия
        do_sync: выполнять ли синхронизацию
        comment: комментарий к синхронизации
        compact: компактная синхронизация
        relinquish: освобождение элементов
        save_if_not_ws: сохранять не-workshared документы
        dialog_suppressor: объект DialogSuppressor для отключения (опционально)
        source_path: оригинальный путь к файлу (строка), используется вместо doc.PathName для detached-копий

    Returns:
        dict: словарь с результатами операции:
            - saved_or_synced (bool): успешно сохранено или синхронизировано
            - closed (bool): успешно закрыт
            - success (bool): операция в целом успешна
            - swc_error (str): ошибка при синхронизации с центральной моделью
            - save_error (str): ошибка при сохранении
            - close_error (str): ошибка при закрытии
            - transmission_flag_cleared (bool): флаг передан
            - fallback_saved (bool): альтернативное сохранение выполнено успешно
            - was_transmitted (bool): документ был переданной моделью
            - is_readonly (bool): документ был в состоянии IsReadOnly
            - save_operation (str): описание выполненной операции
            - fallback_reason (str): причина запуска fallback (если применимо)
            - doc_pathname (str): оригинальный путь документа (doc.PathName или source_path)
            - is_path_absolute (bool): является ли путь абсолютным
    """
    # Вычисляем was_transmitted ДО сброса флага
    was_transmitted = _is_transmitted(doc)
    transmission_flag_cleared = _clear_transmission_flag(doc)

    saved_or_synced = False
    closed = False
    swc_error = None
    save_error = None
    close_error = None
    fallback_saved = False
    save_operation = None
    fallback_reason = None
    is_readonly = doc.IsReadOnly
    doc_pathname = doc.PathName or ""
    is_path_absolute = os.path.isabs(doc_pathname) if doc_pathname else False

    try:
        if do_sync and doc.IsWorkshared:
            try:
                twc = TransactWithCentralOptions()
                swc = SynchronizeWithCentralOptions()
                swc.Comment = comment or ""
                swc.Compact = bool(compact)
                swc.SetRelinquishOptions(RelinquishOptions(bool(relinquish)))
                doc.SynchronizeWithCentral(twc, swc)
                saved_or_synced = True
                save_operation = "SynchronizeWithCentral"
            except Exception as e:
                swc_error = str(e)
                if (
                    was_transmitted
                    or doc.IsReadOnly
                    or _looks_like_detached_error(swc_error)
                ):
                    save_operation = (
                        "SaveAs via temporary file (fallback after SWC error)"
                    )
                    if was_transmitted:
                        fallback_reason = "was_transmitted"
                    elif doc.IsReadOnly:
                        fallback_reason = "doc_is_readonly"
                    elif _looks_like_detached_error(swc_error):
                        fallback_reason = "swc_error_detached"
                    try:
                        orig_path = _get_orig_path(doc, source_path)
                        sys.stderr.write(
                            "[closebg] doc.PathName: '{}'\n".format(doc_pathname)
                        )
                        sys.stderr.write(
                            "[closebg] is_path_absolute: {}\n".format(is_path_absolute)
                        )
                        sys.stderr.write(
                            "[closebg] orig_path: '{}'\n".format(orig_path)
                        )
                        _clear_readonly_attribute(orig_path)
                        sao = SaveAsOptions()
                        sao.OverwriteExistingFile = True
                        wsao = WorksharingSaveAsOptions()
                        wsao.SaveAsCentral = True
                        sao.SetWorksharingOptions(wsao)

                        # Сохраняем во временный файл
                        dirname = os.path.dirname(orig_path)
                        basename = os.path.basename(orig_path)
                        tmp_path = os.path.join(dirname, ".__tmp__." + basename)
                        doc.SaveAs(tmp_path, sao)

                        # Закрываем документ
                        doc.Close(False)
                        closed = True

                        # Безопасная замена файла: Copy + Delete вместо Delete + Move
                        File.Copy(tmp_path, orig_path, True)
                        File.Delete(tmp_path)

                        saved_or_synced = True
                        fallback_saved = True
                    except Exception as save_err:
                        save_error = str(save_err)
            else:
                save_operation = "Save (fallback after SWC error)"
                try:
                    doc.Save(SaveOptions())
                    saved_or_synced = True
                    fallback_saved = True
                except Exception as save_err:
                    save_error = str(save_err)
                    if _looks_like_detached_error(save_error):
                        save_operation = (
                            "SaveAs via temporary file (detached fix after Save error)"
                        )
                        fallback_reason = "save_error_detached"
                        try:
                            orig_path = _get_orig_path(doc, source_path)
                            sys.stderr.write(
                                "[closebg] doc.PathName: '{}'\n".format(doc_pathname)
                            )
                            sys.stderr.write(
                                "[closebg] is_path_absolute: {}\n".format(
                                    is_path_absolute
                                )
                            )
                            sys.stderr.write(
                                "[closebg] orig_path: '{}'\n".format(orig_path)
                            )
                            _clear_readonly_attribute(orig_path)
                            sao = SaveAsOptions()
                            sao.OverwriteExistingFile = True
                            wsao = WorksharingSaveAsOptions()
                            wsao.SaveAsCentral = True
                            sao.SetWorksharingOptions(wsao)

                            # Сохраняем во временный файл
                            dirname = os.path.dirname(orig_path)
                            basename = os.path.basename(orig_path)
                            tmp_path = os.path.join(dirname, ".__tmp__." + basename)
                            doc.SaveAs(tmp_path, sao)

                            # Закрываем документ
                            doc.Close(False)
                            closed = True

                            # Безопасная замена файла: Copy + Delete вместо Delete + Move
                            File.Copy(tmp_path, orig_path, True)
                            File.Delete(tmp_path)

                            saved_or_synced = True
                            fallback_saved = True
                        except Exception as saveas_err:
                            save_error = str(saveas_err)
        elif save_if_not_ws and not doc.IsWorkshared:
            if was_transmitted or doc.IsReadOnly:
                save_operation = "SaveAs via temporary file to clear transmitted state"
                if was_transmitted:
                    fallback_reason = "was_transmitted"
                elif doc.IsReadOnly:
                    fallback_reason = "doc_is_readonly"
                try:
                    orig_path = _get_orig_path(doc, source_path)
                    sys.stderr.write(
                        "[closebg] doc.PathName: '{}'\n".format(doc_pathname)
                    )
                    sys.stderr.write(
                        "[closebg] is_path_absolute: {}\n".format(is_path_absolute)
                    )
                    sys.stderr.write("[closebg] orig_path: '{}'\n".format(orig_path))
                    _clear_readonly_attribute(orig_path)
                    sao = SaveAsOptions()
                    sao.OverwriteExistingFile = True

                    # Сохраняем во временный файл
                    dirname = os.path.dirname(orig_path)
                    basename = os.path.basename(orig_path)
                    tmp_path = os.path.join(dirname, ".__tmp__." + basename)
                    doc.SaveAs(tmp_path, sao)

                    # Закрываем документ
                    doc.Close(False)
                    closed = True

                    # Безопасная замена файла: Copy + Delete вместо Delete + Move
                    File.Copy(tmp_path, orig_path, True)
                    File.Delete(tmp_path)

                    saved_or_synced = True
                except Exception as e:
                    save_error = str(e)
            else:
                save_operation = "normal save"
                try:
                    doc.Save(SaveOptions())
                    saved_or_synced = True
                except Exception as e:
                    save_error = str(e)
                    if _looks_like_detached_error(save_error):
                        save_operation = "SaveAs via temporary file (detached fix)"
                        fallback_reason = "save_error_detached"
                        try:
                            orig_path = _get_orig_path(doc, source_path)
                            sys.stderr.write(
                                "[closebg] doc.PathName: '{}'\n".format(doc_pathname)
                            )
                            sys.stderr.write(
                                "[closebg] is_path_absolute: {}\n".format(
                                    is_path_absolute
                                )
                            )
                            sys.stderr.write(
                                "[closebg] orig_path: '{}'\n".format(orig_path)
                            )
                            _clear_readonly_attribute(orig_path)
                            sao = SaveAsOptions()
                            sao.OverwriteExistingFile = True

                            # Сохраняем во временный файл
                            dirname = os.path.dirname(orig_path)
                            basename = os.path.basename(orig_path)
                            tmp_path = os.path.join(dirname, ".__tmp__." + basename)
                            doc.SaveAs(tmp_path, sao)

                            # Закрываем документ
                            doc.Close(False)
                            closed = True

                            # Безопасная замена файла: Copy + Delete вместо Delete + Move
                            File.Copy(tmp_path, orig_path, True)
                            File.Delete(tmp_path)

                            saved_or_synced = True
                        except Exception as saveas_err:
                            save_error = str(saveas_err)

        if not closed:
            try:
                doc.Close(False)
                closed = True
            except Exception as e:
                close_error = str(e)
    finally:
        if dialog_suppressor is not None:
            dialog_suppressor.detach()

    success = saved_or_synced and closed

    return {
        "saved_or_synced": saved_or_synced,
        "closed": closed,
        "success": success,
        "swc_error": swc_error,
        "save_error": save_error,
        "close_error": close_error,
        "transmission_flag_cleared": transmission_flag_cleared,
        "fallback_saved": fallback_saved,
        "was_transmitted": was_transmitted,
        "is_readonly": is_readonly,
        "save_operation": save_operation,
        "fallback_reason": fallback_reason,
        "doc_pathname": doc_pathname,
        "is_path_absolute": is_path_absolute,
    }
