# -*- coding: utf-8 -*-
# pyRevit pushbutton: копируем "ADSK_Номер секции" и "O_Комплект" с трубы/соед.детали на изоляцию этой трубы
# Работает в IronPython/CPython. Берём из instance->type у хоста; пишем в instance, иначе type у изоляции.
from __future__ import print_function
__title__  = u"Перенести параметры\nс трубы на изоляцию"
__author__ = "Влад"
__doc__    = u"""Переносит параметры с трубы на изоляцию"""


import clr

# Параметры для переноса: (параметр источника, параметр назначения, коэффициент конверсии)
# Коэффициент: None = без конверсии, число = умножить значение на это число
# Revit хранит длину в футах, 1 фут = 0.3048 м, поэтому для перевода футов в метры: * 0.3048
FEET_TO_METERS = 0.3048

PARAM_MAPPING = [
    (u"ADSK_Номер секции", u"ADSK_Номер секции", None),
    (u"О_Комплект", u"О_Комплект", None),
    (u"ADSK_Размер_Длина", u"ADSK_Количество", FEET_TO_METERS),  # Длина трубы (футы -> метры) -> Количество изоляции
]

# 1) Пробуем doc через pyRevit
use_pyrevit = False
try:
    from pyrevit import revit, DB
    doc = revit.doc
    use_pyrevit = True
except Exception:
    # 2) Фолбэк через RevitServices
    clr.AddReference('RevitServices')
    from RevitServices.Persistence import DocumentManager
    doc = DocumentManager.Instance.CurrentDBDocument

if doc is None:
    raise Exception(u"Откройте проект Revit и запустите скрипт из него (doc is None).")

# Унификация API-типов
if use_pyrevit:
    BuiltInCategory = DB.BuiltInCategory
    StorageType = DB.StorageType
    ElementId = DB.ElementId
    FilteredElementCollector = DB.FilteredElementCollector
    TransactionCls = None  # не нужен, в pyRevit используем контекст-менеджер
else:
    clr.AddReference('RevitAPI')
    from Autodesk.Revit.DB import (
        BuiltInCategory, StorageType, ElementId, FilteredElementCollector, Transaction as TransactionCls
    )

# ------- helpers -------

def lookup_param_on(el, name):
    """Возвращает (Parameter, 'instance'|'type'|None). Сначала ищем в экземпляре, потом в типе."""
    if el is None:
        return (None, None)
    p = el.LookupParameter(name)
    if p is not None:
        return (p, 'instance')
    # пробуем в типе
    typ = getattr(el, "Symbol", None)
    if typ is None:
        # у некоторых элементов тип берётся иначе
        typ = getattr(el, "GetTypeId", None)
        if callable(typ):
            tid = el.GetTypeId()
            if tid and tid.IntegerValue > 0:
                typ = doc.GetElement(tid)
            else:
                typ = None
    if typ is not None:
        p2 = typ.LookupParameter(name)
        if p2 is not None:
            return (p2, 'type')
    return (None, None)

def value_of(p):
    """Безопасно читает значение параметра. Пустые строки -> None."""
    if p is None or p.IsReadOnly and not p.HasValue:
        return None
    if not p.HasValue:
        return None
    st = p.StorageType
    if st == StorageType.String:
        s = p.AsString()
        return s if (s is not None and s.strip() != u"") else None
    if st == StorageType.Integer:
        return p.AsInteger()
    if st == StorageType.Double:
        return p.AsDouble()
    if st == StorageType.ElementId:
        eid = p.AsElementId()
        if eid and eid.IntegerValue > 0:
            return eid
        return None
    return None

def set_param(p, v):
    """Пишет значение в параметр p с приведением типов. Возвращает True/False и текст причины при неудаче."""
    if p is None:
        return False, u"target param is None"
    if p.IsReadOnly:
        return False, u"target param is ReadOnly"
    if v is None:
        return False, u"source value is None"

    try:
        st = p.StorageType
        if st == StorageType.String:
            # ElementId -> строка имени/Id
            if isinstance(v, ElementId):
                el = doc.GetElement(v)
                text = el.Name if el is not None and hasattr(el, "Name") else str(v.IntegerValue)
                p.Set(text)
                return True, u""
            p.Set(str(v))
            return True, u""
        if st == StorageType.Integer:
            # true/false/строки чисел -> int
            try:
                p.Set(int(v))
                return True, u""
            except:
                try:
                    p.Set(int(str(v).strip()))
                    return True, u""
                except:
                    return False, u"int cast failed"
        if st == StorageType.Double:
            try:
                p.Set(float(v))
                return True, u""
            except:
                try:
                    p.Set(float(str(v).replace(',', '.').strip()))
                    return True, u""
                except:
                    return False, u"float cast failed"
        if st == StorageType.ElementId:
            # Не будем пытаться угадывать целевой ElementId из строки — небезопасно
            if isinstance(v, ElementId):
                p.Set(v); return True, u""
            return False, u"source is not ElementId"
        return False, u"unsupported storage type"
    except Exception as ex:
        return False, u"exception: {}".format(ex)

# ------- main -------

def process():
    # статистика (ключ — имя целевого параметра)
    wrote = {tgt: 0 for src, tgt, conv in PARAM_MAPPING}
    src_empty = {tgt: 0 for src, tgt, conv in PARAM_MAPPING}
    no_target = {tgt: 0 for src, tgt, conv in PARAM_MAPPING}
    ro_target = {tgt: 0 for src, tgt, conv in PARAM_MAPPING}
    fails = []

    insulations = list(
        FilteredElementCollector(doc)
        .OfCategory(BuiltInCategory.OST_PipeInsulations)
        .WhereElementIsNotElementType()
    )
    if not insulations:
        print(u"Изоляции труб (OST_PipeInsulations) не найдены.")
        return

    for ins in insulations:
        host_id = getattr(ins, "HostElementId", None)
        if not isinstance(host_id, ElementId) or host_id.IntegerValue <= 0:
            fails.append(u"[{}] Нет валидного HostElementId".format(ins.Id.IntegerValue))
            continue
        host = doc.GetElement(host_id)
        if host is None:
            fails.append(u"[{}] Хост (Id={}) не найден".format(ins.Id.IntegerValue, host_id.IntegerValue))
            continue

        for src_pname, tgt_pname, conversion in PARAM_MAPPING:
            src_p, src_where = lookup_param_on(host, src_pname)
            val = value_of(src_p)
            if val is None:
                src_empty[tgt_pname] += 1
                continue

            # Применяем коэффициент конверсии, если задан
            if conversion is not None and isinstance(val, (int, float)):
                val = val * conversion

            # Куда писать: сначала экземпляр изоляции, иначе тип изоляции
            tgt_p, tgt_where = lookup_param_on(ins, tgt_pname)
            if tgt_p is None:
                no_target[tgt_pname] += 1
                fails.append(u"[{}] У изоляции нет параметра '{}' ни в экземпляре, ни в типе".format(
                    ins.Id.IntegerValue, tgt_pname))
                continue
            if tgt_p.IsReadOnly:
                ro_target[tgt_pname] += 1
                fails.append(u"[{}] Параметр '{}' ({}) только для чтения".format(
                    ins.Id.IntegerValue, tgt_pname, tgt_where))
                continue

            ok, reason = set_param(tgt_p, val)
            if ok:
                wrote[tgt_pname] += 1
            else:
                fails.append(u"[{}] Не удалось записать '{}' (из '{}'): {} (src from {})".format(
                    ins.Id.IntegerValue, tgt_pname, src_pname, reason, src_where))

    # отчёт
    total_wrote = sum(wrote.values())
    print(u"Готово. Записано значений: {}".format(total_wrote))
    for src_pname, tgt_pname, conversion in PARAM_MAPPING:
        label = u"{} -> {}".format(src_pname, tgt_pname) if src_pname != tgt_pname else tgt_pname
        print(u"  - {}: записано {}, источник пуст: {}, нет целевого: {}, ReadOnly: {}".format(
            label, wrote[tgt_pname], src_empty[tgt_pname], no_target[tgt_pname], ro_target[tgt_pname]
        ))
    if fails:
        print(u"\nПроблемы:")
        for f in fails:
            print(u" - " + f)

# Транзакция
if use_pyrevit:
    with revit.Transaction(u"Копировать параметры (труба → изоляция)"):
        process()
else:
    tx = TransactionCls(doc, u"Копировать параметры (труба → изоляция)")
    tx.Start()
    try:
        process()
    finally:
        tx.Commit()
