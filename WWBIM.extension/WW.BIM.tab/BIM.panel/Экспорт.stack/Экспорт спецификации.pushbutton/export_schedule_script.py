# -*- coding: utf-8 -*-
from __future__ import print_function

import os
import re
import zipfile
from xml.sax.saxutils import escape, quoteattr

from pyrevit import revit, DB, forms


NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def text_value(value):
    if value is None:
        return u""
    try:
        return unicode(value)
    except NameError:
        return str(value)


def xml_text(value):
    return escape(text_value(value), {"\"": "&quot;"})


def excel_column(index):
    result = u""
    index += 1
    while index:
        index, remainder = divmod(index - 1, 26)
        result = chr(65 + remainder) + result
    return result


def safe_sheet_name(name):
    name = re.sub(r"[\\/*?:\[\]]", u"_", text_value(name)).strip()
    if not name:
        name = u"Schedule"
    return name[:31]


def safe_file_name(name):
    name = re.sub(r"[\\/:*?\"<>|]", u"_", text_value(name)).strip()
    return (name or u"schedule") + u".xlsx"


def section_text(schedule, section_type, row, column):
    try:
        return text_value(schedule.GetCellText(section_type, row, column))
    except Exception:
        return u""


def section_dimensions(section):
    try:
        return int(section.NumberOfRows), int(section.NumberOfColumns)
    except Exception:
        return 0, 0


def section_size(section, row, column, default):
    try:
        value = section.GetRowHeight(row) if column is None else section.GetColumnWidth(column)
        # Revit table dimensions are internal feet. Excel uses points/character widths.
        value = float(value)
        if value > 0:
            return value * 304.8
    except Exception:
        pass
    return default


def merged_bounds(section, row, column):
    try:
        merged = section.GetMergedCell(row, column)
        top = int(merged.Top)
        bottom = int(merged.Bottom)
        left = int(merged.Left)
        right = int(merged.Right)
        if top <= bottom and left <= right:
            return top, bottom, left, right
    except Exception:
        pass
    return row, row, column, column


def read_schedule(schedule):
    sections = [
        (DB.SectionType.Header, u"header"),
        (DB.SectionType.Body, u"body"),
        (DB.SectionType.Footer, u"footer"),
    ]
    rows = []
    merges = []
    widths_mm = []
    row_offset = 0
    max_columns = 0

    for section_type, section_name in sections:
        section = schedule.GetTableData().GetSectionData(section_type)
        row_count, column_count = section_dimensions(section)
        if row_count <= 0 or column_count <= 0:
            continue
        max_columns = max(max_columns, column_count)
        while len(widths_mm) < column_count:
            widths_mm.append(0.0)
        for column in range(column_count):
            widths_mm[column] = max(widths_mm[column], section_size(section, 0, column, 25.0))

        for row in range(row_count):
            row_height = section_size(section, row, None, 15.0)
            values = []
            for column in range(column_count):
                values.append(section_text(schedule, section_type, row, column))
                top, bottom, left, right = merged_bounds(section, row, column)
                if top == row and left == column and (bottom > top or right > left):
                    merges.append((row_offset + top, row_offset + bottom, left, right))
            rows.append({
                "values": values,
                "height_mm": row_height,
                "header": section_type == DB.SectionType.Header,
                "section": section_name,
            })
        row_offset += row_count

    if not rows or max_columns == 0:
        raise Exception(u"Активная спецификация не содержит экспортируемых строк.")
    while len(widths_mm) < max_columns:
        widths_mm.append(25.0)
    for row in rows:
        row["values"] += [u""] * (max_columns - len(row["values"]))
    return rows, widths_mm, unique_merges(merges)


def unique_merges(merges):
    result = []
    seen = set()
    for merge in merges:
        if merge not in seen:
            seen.add(merge)
            result.append(merge)
    return result


def excel_width(width_mm):
    # Approximation for Excel's default 96 dpi / 7 px character width.
    pixels = max(12.0, width_mm / 25.4 * 96.0)
    return max(0.5, min(255.0, (pixels - 5.0) / 7.0))


def cell_xml(row, column, value, header):
    reference = excel_column(column) + str(row + 1)
    style = " s=\"1\"" if header else ""
    if not value:
        return u"<c r=\"{0}\"{1}/ >".format(reference, style).replace(u"/ >", u"/>")
    return u"<c r=\"{0}\" t=\"inlineStr\"{1}><is><t xml:space=\"preserve\">{2}</t></is></c>".format(
        reference, style, xml_text(value))


def build_xlsx(path, sheet_name, rows, widths_mm, merges):
    sheet_rows = []
    for row_index, row in enumerate(rows):
        cells = [cell_xml(row_index, column, row["values"][column], row["header"])
                 for column in range(len(widths_mm))]
        height_points = max(1.0, row["height_mm"] / 25.4 * 72.0)
        sheet_rows.append(u"<row r=\"{0}\" ht=\"{1:.2f}\" customHeight=\"1\">{2}</row>".format(
            row_index + 1, height_points, u"".join(cells)))

    columns = []
    for index, width in enumerate(widths_mm):
        columns.append(u"<col min=\"{0}\" max=\"{0}\" width=\"{1:.2f}\" customWidth=\"1\"/>".format(
            index + 1, excel_width(width)))
    merge_xml = u""
    if merges:
        refs = [u"{0}{1}:{2}{3}".format(excel_column(left), top + 1,
                                           excel_column(right), bottom + 1)
                for top, bottom, left, right in merges
                if bottom > top or right > left]
        if refs:
            merge_xml = u"<mergeCells count=\"{0}\">{1}</mergeCells>".format(
                len(refs), u"".join(u"<mergeCell ref=\"{0}\"/>".format(ref) for ref in refs))

    worksheet = u'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="{main}" xmlns:r="{rel}"><dimension ref="A1:{last_col}{last_row}"/>
<sheetViews><sheetView workbookViewId="0"/></sheetViews><sheetFormatPr defaultRowHeight="15"/>
<cols>{columns}</cols><sheetData>{rows}</sheetData>{merges}</worksheet>'''.format(
        main=NS_MAIN, rel=NS_REL, last_col=excel_column(len(widths_mm) - 1),
        last_row=len(rows), columns=u"".join(columns), rows=u"".join(sheet_rows), merges=merge_xml)
    workbook = u'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="{main}" xmlns:r="{rel}"><sheets><sheet name={name} sheetId="1" r:id="rId1"/></sheets></workbook>'''.format(
        main=NS_MAIN, rel=NS_REL, name=quoteattr(sheet_name))
    styles = u'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="{main}"><fonts count="2"><font><sz val="11"/><name val="Arial"/></font><font><b/><sz val="11"/><name val="Arial"/></font></fonts>
<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/><xf numFmtId="0" fontId="1" fillId="0" borderId="0"/></cellXfs></styleSheet>'''.format(main=NS_MAIN)
    content_types = u'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>'''
    root_rels = u'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>'''
    workbook_rels = u'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>'''
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as archive:
        def write_xml(name, value):
            archive.writestr(name, value.encode("utf-8"))
        write_xml("[Content_Types].xml", content_types)
        write_xml("_rels/.rels", root_rels)
        write_xml("xl/workbook.xml", workbook)
        write_xml("xl/_rels/workbook.xml.rels", workbook_rels)
        write_xml("xl/worksheets/sheet1.xml", worksheet)
        write_xml("xl/styles.xml", styles)


def main():
    uidoc = revit.uidoc
    schedule = uidoc.ActiveView
    if not isinstance(schedule, DB.ViewSchedule) or schedule.IsTemplate:
        forms.alert(u"Активный вид не является спецификацией Revit. Откройте спецификацию и повторите экспорт.",
                    title=u"Экспорт спецификации")
        return
    try:
        rows, widths, merges = read_schedule(schedule)
        default_name = safe_file_name(schedule.Name)
        path = forms.save_file(
            files_filter=u"Excel workbook (*.xlsx)|*.xlsx",
            default_name=default_name)
        if not path:
            return
        if not path.lower().endswith(u".xlsx"):
            path += u".xlsx"
        build_xlsx(path, safe_sheet_name(schedule.Name), rows, widths, merges)
        forms.alert(u"Спецификация экспортирована:\n{0}".format(path), title=u"Экспорт спецификации")
    except Exception as error:
        forms.alert(u"Не удалось экспортировать спецификацию:\n{0}".format(error),
                    title=u"Экспорт спецификации")


if __name__ == "__main__":
    main()
