# -*- coding: utf-8 -*-
from Autodesk.Revit.DB import (
    TransactWithCentralOptions,
    SynchronizeWithCentralOptions,
    RelinquishOptions,
    SaveOptions,
)


def close_with_policy(
    doc,
    do_sync=False,
    comment="",
    compact=True,
    relinquish=True,
    save_if_not_ws=True,
    dialog_suppressor=None,
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
    """
    try:
        if do_sync and doc.IsWorkshared:
            try:
                twc = TransactWithCentralOptions()
                swc = SynchronizeWithCentralOptions()
                swc.Comment = comment or ""
                swc.Compact = bool(compact)
                swc.SetRelinquishOptions(RelinquishOptions(bool(relinquish)))
                doc.SynchronizeWithCentral(twc, swc)
            except Exception:
                pass
        else:
            if save_if_not_ws and not doc.IsWorkshared:
                try:
                    doc.Save(SaveOptions())
                except Exception:
                    pass
        try:
            doc.Close(False)
        except Exception:
            pass
    except Exception:
        try:
            doc.Close(False)
        except Exception:
            pass
    finally:
        # КРИТИЧНО: отключаем dialog_suppressor после закрытия документа
        if dialog_suppressor is not None:
            try:
                dialog_suppressor.detach()
            except Exception:
                pass
