# -*- coding: utf-8 -*-
from __future__ import print_function

from pyrevit import revit, DB, script, forms
from System.IO import FileInfo, Path, File
from System import Guid
from collections import defaultdict

doc = revit.doc
output = script.get_output()
app = doc.Application

# флаг: нужно ли пропускать проверки спецификации/библиотеки
# для моделей разделов АР/AR, КР/KR, КМ/KM (по имени файла/проекта)
try:
    model_name = doc.Title or u""
except Exception:
    model_name = u""

try:
    if (not model_name) and doc.PathName:
        model_name = Path.GetFileNameWithoutExtension(doc.PathName)
except Exception:
    pass

model_name_upper = (model_name or u"").upper()

def _has_spec_lib_marker(name):
    markers = [u"АР", u"AR", u"КР", u"KR", u"КМ", u"KM"]
    # считаем маркером отдельную часть имени, отделённую подчёркиванием или дефисом
    for tag in markers:
        if name.startswith(tag + u"_") or name.endswith(u"_" + tag) or (u"_" + tag + u"_") in name:
            return True
        if name.startswith(tag + u"-") or name.endswith(u"-" + tag) or (u"-" + tag + u"-") in name:
            return True
    return False

skip_spec_and_lib_checks = _has_spec_lib_marker(model_name_upper)




# -----------------------
# ВСТУПИТЕЛЬНЫЙ ВОПРОС
# -----------------------
resp = forms.alert(
    u"Проверять размер загружаемых семейств?\n"
    u"Это может занять дополнительное время.",
    yes=True, no=True, cancel=True
)

if resp is None:
    # нажали Cancel или закрыли окно
    script.exit()

# поддерживаем и bool, и строковые ответы
if resp in (True, 'yes', 'Yes', 'YES'):
    check_family_sizes = True
elif resp in (False, 'no', 'No', 'NO'):
    check_family_sizes = False
else:
    # что-то непонятное — на всякий случай выходим
    script.exit()


# -----------------------
# ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ
# -----------------------
def format_size(bytes_count):
    size = float(bytes_count)
    for unit in [u'Б', u'КБ', u'МБ', u'ГБ']:
        if size < 1024.0 or unit == u'ГБ':
            return u"{0:.1f} {1}".format(size, unit)
        size /= 1024.0
    return u"{0:.1f} {1}".format(size, u'ГБ')


def print_section_header(title, emoji=u"📌"):
    output.print_html(
        u"<h3>{emoji} {title}</h3>".format(
            emoji=emoji,
            title=title
        )
    )


def print_ok(msg):
    output.print_html(u'<p style="color:#4CAF50;">✅ {}</p>'.format(msg))


def print_warn(msg):
    output.print_html(u'<p style="color:#FF5722;">⚠️ {}</p>'.format(msg))


def print_info(msg):
    output.print_html(u'<p style="color:#2196F3;">ℹ️ {}</p>'.format(msg))


def is_param_zero_or_empty(param):
    """True, если параметр нет, 0 или '0'/'0,0'/'0.0'/пусто."""
    if param is None:
        return True

    st = param.StorageType
    if st == DB.StorageType.Double:
        val = abs(param.AsDouble())
        return val < 1e-9
    elif st == DB.StorageType.Integer:
        return param.AsInteger() == 0
    else:
        val = param.AsString() or u""
        val = val.strip()
        if not val:
            return True
        low = val.replace(" ", "").replace(",", ".")
        if low in [u"0", u"0.0", u"0.00"]:
            return True
        try:
            f = float(low)
            return abs(f) < 1e-9
        except Exception:
            # любая ненулевая строка — считаем ненулевой
            return False


def is_param_positive(param):
    return not is_param_zero_or_empty(param)


# кэш имён рабочих наборов
_workset_name_cache = {}
_workset_table = doc.GetWorksetTable()


def get_workset_name(el):
    try:
        wsid = el.WorksetId
    except Exception:
        return None

    if wsid is None:
        return None

    key = wsid.IntegerValue
    if key in _workset_name_cache:
        return _workset_name_cache[key]

    try:
        ws = _workset_table.GetWorkset(wsid)
    except Exception:
        ws = None

    name = ws.Name if ws else None
    _workset_name_cache[key] = name
    return name


def print_element_line(el, prefix=u"-"):
    try:
        cat_name = el.Category.Name if el.Category else u"(нет категории)"
    except Exception:
        cat_name = u"(нет категории)"

    ws_name = get_workset_name(el) or u"(нет рабочего набора)"
    link = output.linkify(el.Id)
    output.print_md(
        u"{prefix} {link} | {cat} | WS: `{ws}`".format(
            prefix=prefix,
            link=link,
            cat=cat_name,
            ws=ws_name
        )
    )


# -----------------------
# GUID ИЗ ФАЙЛА ОБЩИХ ПАРАМЕТРОВ
# -----------------------
shared_param_guids_in_file = None
try:
    sp_file = app.OpenSharedParameterFile()
except Exception:
    sp_file = None

if sp_file:
    shared_param_guids_in_file = set()
    try:
        for group in sp_file.Groups:
            for defn in group.Definitions:
                guid = None
                try:
                    guid = defn.GUID
                except Exception:
                    guid = None
                if guid and guid != Guid.Empty:
                    shared_param_guids_in_file.add(guid)
    except Exception:
        shared_param_guids_in_file = None


# -----------------------
# 1. РАЗМЕР ФАЙЛА И КОЛ-ВО ЭЛЕМЕНТОВ
# -----------------------
print_section_header(u"Общая информация по модели", emoji=u"📊")

path = doc.PathName
filename = Path.GetFileName(path) if path else u"(файл ещё не сохранён)"

output.print_md(u"**Файл:** `{}`".format(filename))

size_bytes = None

if not path:
    print_warn(u"Файл ещё не сохранён. Невозможно определить размер файла на диске.")
else:
    try:
        fi = FileInfo(path)
        if not fi.Exists:
            print_warn(
                u"Файл по пути `{}` не найден. Размер файла определить невозможно.".format(path)
            )
        else:
            size_bytes = fi.Length
            size_str = format_size(size_bytes)
            output.print_md(u"**Путь:** `{}`".format(path))
            output.print_md(u"**Размер файла:** `{}`".format(size_str))
    except Exception as ex:
        print_warn(
            u"Не удалось получить размер файла ({}). Размер будет пропущен.".format(ex)
        )

elem_count = DB.FilteredElementCollector(doc) \
    .WhereElementIsNotElementType() \
    .GetElementCount()
output.print_md(
    u"**Количество элементов (без типов):** `{}`".format(elem_count)
)

if size_bytes is not None:
    limit_bytes = 500 * 1024 * 1024
    size_str = format_size(size_bytes)
    if size_bytes > limit_bytes:
        print_warn(
            u"Размер файла больше 500 МБ ({}). Есть риск проблем с производительностью 🚨".format(size_str)
        )
    else:
        print_ok(u"Размер файла меньше 500 МБ 👍")


# -----------------------
# СБОР ЭЛЕМЕНТОВ ОДНИМ ПРОХОДОМ
# -----------------------
allelems = list(
    DB.FilteredElementCollector(doc)
    .WhereElementIsNotElementType()
    .ToElements()
)

# для проверки рабочих наборов
ws00_invalid_elements = defaultdict(list)   # workset_name -> [elements]
dwg_wrong_workset = []                      # DWG, лежащие не в WS с "DWG"

# для связей
links_by_name = defaultdict(list)           # link type name -> [RevitLinkInstance]

# семейства, у которых хотя бы один экземпляр имеет положительное ADSK_Количество
good_families_qty = set()

# любое размещённое в модели семейство -> любой его экземпляр
family_any_instance = {}

for el in allelems:
    # --- экземпляры семейств для ADSK_Количество и карта "семейство -> любой экземпляр" ---
    if isinstance(el, DB.FamilyInstance):
        try:
            sym = el.Symbol
            fam = sym.Family if sym else None
        except Exception:
            fam = None

        if fam is not None:
            try:
                if fam.IsInPlace:
                    fam = None
            except Exception:
                pass

        if fam is not None:
            # аннотационные семейства в проверке объёма не нужны
            try:
                fcat = fam.FamilyCategory
                if fcat and fcat.CategoryType == DB.CategoryType.Annotation:
                    fam = None
            except Exception:
                pass

        if fam is not None:
            fam_id_int = fam.Id.IntegerValue

            # запоминаем любой размещённый экземпляр семейства
            if fam_id_int not in family_any_instance:
                family_any_instance[fam_id_int] = el

            # ADSK_Количество на экземпляре
            param_inst = el.LookupParameter("ADSK_Количество")
            if is_param_positive(param_inst):
                good_families_qty.add(fam_id_int)

    # --- рабочие наборы, начинающиеся с "00_" ---
    ws_name = get_workset_name(el)
    if ws_name and ws_name.lower().startswith(u"00_"):
        ws_lower = ws_name.lower()
        # Если в названии рабочего набора есть "DWG", полностью игнорируем его
        # в этой проверке (там допустимы любые элементы, в т.ч. DWG).
        if u"dwg" in ws_lower:
            pass
        else:
            # Во всех прочих 00_* рабочих наборах допускаются только Revit-связи.
            if isinstance(el, DB.RevitLinkInstance):
                pass
            else:
                ws00_invalid_elements[ws_name].append(el)

    # --- DWG в неверном рабочем наборе ---
    if isinstance(el, DB.ImportInstance):
        try:
            symbol = el.Symbol
        except Exception:
            symbol = None

        name_candidate = None
        if symbol:
            try:
                name_candidate = symbol.Name
            except Exception:
                pass

        if not name_candidate:
            try:
                name_candidate = el.Name
            except Exception:
                name_candidate = None

        if name_candidate and u".dwg" in name_candidate.lower():
            ws_name = get_workset_name(el) or u""
            # DWG считаются корректными только в рабочих наборах, где в названии есть "DWG"
            if u"dwg" not in ws_name.lower():
                dwg_wrong_workset.append(el)

    # --- связи Revit для поиска дублей ---
    if isinstance(el, DB.RevitLinkInstance):
        try:
            linktype = doc.GetElement(el.GetTypeId())
            lname = linktype.Name if linktype else el.Name
        except Exception:
            lname = el.Name
        links_by_name[lname].append(el)


# общий список семейств
families = list(
    DB.FilteredElementCollector(doc)
    .OfClass(DB.Family)
    .ToElements()
)


# -----------------------
# 2. ТОП-5 ТЯЖЁЛЫХ ЗАГРУЖАЕМЫХ СЕМЕЙСТВ ПО РАЗМЕРУ ФАЙЛА
# -----------------------
# 2. ТОП-5 ТЯЖЁЛЫХ ЗАГРУЖАЕМЫХ СЕМЕЙСТВ ПО РАЗМЕРУ ФАЙЛА
# -----------------------
print_section_header(
    u"Первые 5 самых «тяжёлых» загружаемых семейств (по размеру файла)",
    emoji=u"🐘"
)

family_file_sizes = {}   # Family -> size_bytes

if check_family_sizes:
    for fam in families:
        # только загружаемые, не in-place
        try:
            if fam.IsInPlace:
                continue
        except Exception:
            pass

        fam_doc = None
        try:
            fam_doc = doc.EditFamily(fam)
        except Exception:
            fam_doc = None

        if fam_doc is None:
            continue

        try:
            fpath = fam_doc.PathName
            if fpath and File.Exists(fpath):
                fi = FileInfo(fpath)
                family_file_sizes[fam] = fi.Length
        except Exception:
            pass
        finally:
            try:
                fam_doc.Close(False)
            except Exception:
                pass

    if not family_file_sizes:
        print_info(
            u"Не удалось определить размер ни одного загружаемого семейства "
            u"(возможно, семейства не имеют сохранённого файла RFA)."
        )
    else:
        # учитываем только семейства, у которых есть размещённые экземпляры
        family_file_sizes_in_use = {
            fam: size for fam, size in family_file_sizes.items()
            if fam.Id.IntegerValue in family_any_instance
        }

        if not family_file_sizes_in_use:
            print_info(
                u"Не найдено загружаемых семейств с размещёнными экземплярами "
                u"для оценки по размеру файла."
            )
        else:
            top5 = sorted(
                family_file_sizes_in_use.items(),
                key=lambda x: x[1],
                reverse=True
            )[:5]

            for fam, size_b in top5:
                size_str = format_size(size_b)
                inst = family_any_instance.get(fam.Id.IntegerValue)
                if inst is not None:
                    link = output.linkify(inst.Id)
                else:
                    link = output.linkify(fam.Id)

                output.print_md(
                    u"- {link} `{name}` — размер файла: **{size}**".format(
                        link=link,
                        name=fam.Name,
                        size=size_str
                    )
                )
else:
    print_info(u"Проверка размера файлов семейств отключена пользователем.")


# -----------------------
print_section_header(
    u"Проверка рабочих наборов '00_' и DWG",
    emoji=u"🧱"
)

if ws00_invalid_elements:
    print_warn(
        u"Найдены элементы в рабочих наборах, имя которых начинается с '00_', "
        u"кроме связанных файлов (Revit Links):"
    )
    for ws_name, elems in sorted(ws00_invalid_elements.items(), key=lambda x: x[0]):
        output.print_md(u"**Рабочий набор:** `{}`".format(ws_name))
        for el in elems:
            print_element_line(el)
else:
    print_ok(
        u"Во всех рабочих наборах, начинающихся с '00_', лежат только Revit-связи ✅"
    )

if dwg_wrong_workset:
    print_warn(
        u"Найдены DWG связи/импорты, которые находятся не в рабочем наборе, "
        u"название которого содержит 'DWG':"
    )
    for el in dwg_wrong_workset:
        print_element_line(el)
else:
    print_ok(
        u"Все DWG связи/импорты находятся в рабочих наборах, содержащих 'DWG' в названии 👍"
    )


# -----------------------
# 4. ДУБЛИРОВАННЫЕ СВЯЗИ (ОДНО И ТО ЖЕ ИМЯ)
# -----------------------
print_section_header(
    u"Проверка дублей Revit-связей по имени",
    emoji=u"🔗"
)

duplicates = {name: insts for name, insts in links_by_name.items() if len(insts) > 1}

if duplicates:
    print_warn(
        u"Найдены связи, у которых одно и то же имя (тип связи) "
        u"размещено больше одного раза:"
    )
    for lname, insts in sorted(duplicates.items(), key=lambda x: x[0]):
        output.print_md(
            u"**Связь:** `{}` — экземпляров: {}".format(lname, len(insts))
        )
        for inst in insts:
            print_element_line(inst)
else:
    print_ok(u"Нет дублей Revit-связей с одинаковым именем (типом) ✅")


# -----------------------
# 5. СЕМЕЙСТВА С ADSK_Количество = 0 ИЛИ ПУСТО (БЕЗ АННОТАЦИЙ)
# -----------------------
# 5. ВОЗМОЖНО НЕ УЧИТЫВАЮТСЯ В СПЕЦИФИКАЦИИ (ADSK_Количество)
# -----------------------

# -----------------------
# 4б. НЕЗАКРЕПЛЁННЫЕ ЭЛЕМЕНТЫ (ОСИ И СВЯЗИ)
# -----------------------

print_section_header(
    u"Незакреплённые элементы",
    emoji=u"📌"
)

unpinned_grids = []
unpinned_links = []
unpinned_levels = []

# незакреплённые оси
try:
    for g in DB.FilteredElementCollector(doc).OfClass(DB.Grid):
        try:
            if not g.Pinned:
                unpinned_grids.append(g)
        except Exception:
            pass
except Exception:
    pass

# незакреплённые связи Revit
try:
    for l in DB.FilteredElementCollector(doc).OfClass(DB.RevitLinkInstance):
        try:
            if not l.Pinned:
                unpinned_links.append(l)
        except Exception:
            pass
except Exception:
    pass

# незакреплённые уровни
try:
    for lvl in DB.FilteredElementCollector(doc).OfClass(DB.Level):
        try:
            if not lvl.Pinned:
                unpinned_levels.append(lvl)
        except Exception:
            pass
except Exception:
    pass

if not unpinned_grids and not unpinned_links and not unpinned_levels:
    print_ok(u"Все оси, уровни и связи Revit закреплены (Pinned) ✅")
else:
    if unpinned_grids:
        output.print_md(u"**Незакреплённые оси:**")
        for g in unpinned_grids:
            print_element_line(g)
    else:
        print_info(u"Незакреплённых осей не найдено.")

    if unpinned_levels:
        output.print_md(u"**Незакреплённые уровни:**")
        for lvl in unpinned_levels:
            print_element_line(lvl)
    else:
        print_info(u"Незакреплённых уровней не найдено.")

    if unpinned_links:
        output.print_md(u"**Незакреплённые связи Revit:**")
        for l in unpinned_links:
            print_element_line(l)
    else:
        print_info(u"Незакреплённых связей Revit не найдено.")


if not skip_spec_and_lib_checks:
    print_section_header(
        u"Возможно не учитываются в спецификации (ADSK_Количество пусто или равно 0)",
        emoji=u"📉"
    )

    problem_families_qty = {}

    family_symbols = DB.FilteredElementCollector(doc)     .OfClass(DB.FamilySymbol)     .ToElements()

    for sym in family_symbols:
        try:
            fam = sym.Family
        except Exception:
            fam = None

        if fam is None:
            continue

        # не учитываем in-place семейства
        try:
            if fam.IsInPlace:
                continue
        except Exception:
            pass

        # исключаем аннотационные семейства
        try:
            fcat = fam.FamilyCategory
            if fcat and fcat.CategoryType == DB.CategoryType.Annotation:
                continue
        except Exception:
            pass

        fam_id = fam.Id.IntegerValue

        # учитываем только семейства, у которых есть размещённые экземпляры
        if fam_id not in family_any_instance:
            continue

        # если хотя бы один экземпляр этого семейства имеет положительное ADSK_Количество — всё ок
        if fam_id in good_families_qty:
            continue

        # уже добавили в список
        if fam_id in problem_families_qty:
            continue

        # проверяем тип/семейство
        param = sym.LookupParameter("ADSK_Количество")
        if param is None and fam is not None:
            param = fam.LookupParameter("ADSK_Количество")

        if is_param_zero_or_empty(param):
            problem_families_qty[fam_id] = fam

    if problem_families_qty:
        for fam_id in sorted(problem_families_qty.keys()):
            fam = problem_families_qty[fam_id]
            inst = family_any_instance.get(fam_id)
            if inst is not None:
                link = output.linkify(inst.Id)
            else:
                link = output.linkify(fam.Id)

            output.print_md(u"- {link} `{name}`".format(link=link, name=fam.Name))
    else:
        print_ok(
            u"Все загружаемые (не аннотационные) семейства учитываются по ADSK_Количество ✅"
        )


    # -----------------------
    # 6. СЕМЕЙСТВА БЕЗ НИЖНЕГО ПОДЧЁРКИВАНИЯ В ИМЕНИ
    # -----------------------
if not skip_spec_and_lib_checks:
    print_section_header(
        u"Возможно семейства не из библиотеки (нет '_' в имени семейства)",
        emoji=u"📁"
    )

    families_no_underscore = []

    for fam in families:
        try:
            name = fam.Name or u""
        except Exception:
            continue

        # не учитываем in-place
        try:
            if fam.IsInPlace:
                continue
        except Exception:
            pass

        fam_id_int = fam.Id.IntegerValue

        # учитываем только семейства, у которых есть размещённые экземпляры
        if fam_id_int not in family_any_instance:
            continue

        if u"_" not in name:
            families_no_underscore.append(fam)

    if families_no_underscore:
        for fam in sorted(families_no_underscore, key=lambda f: f.Name):
            fam_id_int = fam.Id.IntegerValue
            inst = family_any_instance.get(fam_id_int)
            if inst is not None:
                link = output.linkify(inst.Id)
            else:
                link = output.linkify(fam.Id)

            output.print_md(u"- {link} `{name}`".format(link=link, name=fam.Name))
    else:
        print_ok(
            u"Все загружаемые семейства содержат '_' в имени (по признаку библиотеки) ✅"
        )


    # -----------------------
    # 7. ДУБЛИРОВАННЫЕ ОБЩИЕ (SHARED) ПАРАМЕТРЫ ПО НАЗВАНИЮ
    # -----------------------
print_section_header(
    u"Дублированные общие параметры (по названию)",
    emoji=u"⚙️"
)

param_elems = DB.FilteredElementCollector(doc) \
    .OfClass(DB.ParameterElement) \
    .ToElements()

params_by_name = defaultdict(list)

for pe in param_elems:
    try:
        defn = pe.GetDefinition()
    except Exception:
        defn = None

    if defn is None:
        continue

    if not isinstance(pe, DB.SharedParameterElement):
        continue

    name = defn.Name
    try:
        guid = pe.GuidValue
    except Exception:
        guid = Guid.Empty

    if guid is None or guid == Guid.Empty:
        continue

    params_by_name[name].append((pe, guid))

duplicated_shared_params = {
    name: items for name, items in params_by_name.items() if len(items) > 1
}

if duplicated_shared_params:
    if shared_param_guids_in_file is None:
        print_warn(
            u"Найдены имена общих параметров, которые повторяются в проекте. "
            u"(Файл общих параметров не найден или не прочитан — без проверки GUID.) "
            u"Ниже имя → GUID → ID (кликабельно):"
        )
    else:
        print_warn(
            u"Найдены имена общих параметров, которые повторяются в проекте. "
            u"GUID, присутствующие в файле общих параметров, отмечены ✅. "
            u"Ниже имя → статус → GUID → ID:"
        )

    for name in sorted(duplicated_shared_params.keys()):
        output.print_md(u"**Параметр:** `{}`".format(name))
        for pe, guid in duplicated_shared_params[name]:
            link = output.linkify(pe.Id)

            status = u""
            if shared_param_guids_in_file is not None:
                if guid in shared_param_guids_in_file:
                    status = u'<span style="color:#4CAF50;">✅</span> '
                else:
                    status = u'<span style="color:#FF5722;">⚠️</span> '

            output.print_html(
                u"{status}- GUID: <code>{guid}</code> | ID: {link}".format(
                    status=status,
                    guid=str(guid),
                    link=link
                )
            )
else:
    print_ok(u"Не найдено дублированных по имени общих параметров ✅")


print_info(u"Проверка BIM-модели завершена 🎉")
