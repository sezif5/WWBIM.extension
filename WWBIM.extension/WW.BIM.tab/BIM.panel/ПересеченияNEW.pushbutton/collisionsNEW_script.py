# -*- coding: utf-8 -*-
"""
Скрипт для чтения и анализа XML-отчётов Navisworks о пересечениях.

Показывает все проверки (включая пустые), фильтрует по текущей модели Revit,
печатает таблицы с подсветкой категорий и выводит сводку.
"""

__title__ = "Пересечения"
__author__ = "vlad / you"
__doc__ = "Читает XML-отчёт Navisworks, показывает проверки, фильтрует по текущей модели Revit, выводит таблицы с подсветкой категории и сводку."
__persistentengine__ = True

import os
import re
import json
import clr
from collections import namedtuple
from datetime import datetime

clr.AddReference("System.Data")
clr.AddReference("RevitAPIUI")

from pyrevit import forms, script
from Autodesk.Revit.DB import ElementId, View3D, BoundingBoxXYZ, XYZ, Transaction
from Autodesk.Revit.UI import IExternalEventHandler, ExternalEvent
from System.Data import DataTable
from System.Collections.Generic import List
from System.Windows import Thickness, FrameworkElementFactory
from System.Windows.Controls import (
    Button,
    DataGridComboBoxColumn,
    DataGridTemplateColumn,
    Image,
    Orientation,
    StackPanel,
    TextBlock,
)
from System.Windows import DataTemplate
from System.Windows.Data import Binding
from System.Windows.Media import Stretch
from System.Windows.Media import Brushes
from System.Windows.Media.Imaging import BitmapImage, BitmapCacheOption
from System import Uri, UriKind


# =============================================================================
# КОНСТАНТЫ И НАСТРОЙКИ
# =============================================================================

# Фильтрация результатов
FILTER_BY_FILE = True  # Сверка файла из pathlink с текущей моделью
FILTER_BY_CATEGORY = False  # Дополнительная сверка категории в тексте pathlink

# Имена атрибутов для поиска ID объекта
OBJECT_ID_NAMES = frozenset(
    ["объект id", "object id", "object 1 id", "object 2 id", "id объекта"]
)

# Атрибуты даты в XML
DATE_ATTRIBUTES = (
    "date",
    "created",
    "generated",
    "exported",
    "timestamp",
    "time",
    "start",
    "lastsaved",
    "creationtime",
    "modificationtime",
)

ALLOWED_STATUSES = ("Новая", "В работе", "Закрыта")

# CSS-стиль для подсветки категорий
BADGE_STYLE = "background:#2f3b4a; color:#fff; padding:1px 6px; border-radius:3px; font-weight:600;"
MODEL_HEADER_LABEL_STYLE = "background:#0d6efd; color:#ffffff; padding:3px 10px; border-radius:999px; font-weight:700;"
MODEL_HEADER_NAME_STYLE = "background:#e7f1ff; color:#0b3d91; padding:3px 10px; border-radius:6px; font-weight:700; border:1px solid #b6d4fe;"


# =============================================================================
# РЕГУЛЯРНЫЕ ВЫРАЖЕНИЯ
# =============================================================================


class Patterns:
    """Скомпилированные регулярные выражения для парсинга."""

    # Экранирование некорректных амперсандов в XML
    BAD_AMPERSAND = re.compile(
        "&(?!amp;|lt;|gt;|apos;|quot;|#\\d+;|#x[0-9A-Fa-f]+;)", re.UNICODE
    )

    # Имя файла модели
    MODEL_FILE = re.compile(".+\\.(nwc|nwd|nwf|rvt)$", re.IGNORECASE)

    # Атрибут даты в XML-тексте
    DATE_ATTR = re.compile(
        r"\b(?:date|created|generated|export(?:date|ed)?|timestamp|time|"
        r'start|lastsaved|creationtime|modificationtime)\s*=\s*"([^"]+)"',
        re.IGNORECASE,
    )

    # Fallback-парсер: блоки и атрибуты
    CLASHTEST_BLOCK = re.compile(
        r"<clashtest\b[^>]*>.*?</clashtest>", re.DOTALL | re.IGNORECASE
    )
    TEST_BLOCK = re.compile(r"<test\b[^>]*>.*?</test>", re.DOTALL | re.IGNORECASE)
    CLASHRESULT_BLOCK = re.compile(
        r"<clashresult\b[^>]*>.*?</clashresult>", re.DOTALL | re.IGNORECASE
    )
    CLASHOBJECT_BLOCK = re.compile(
        r"<clashobject\b[^>]*>.*?</clashobject>", re.DOTALL | re.IGNORECASE
    )

    ATTR_NAME = re.compile(r"\bname=\"([^\"]+)\"", re.IGNORECASE)
    ATTR_TESTNAME = re.compile(r"\btestname=\"([^\"]+)\"", re.IGNORECASE)
    ATTR_HREF = re.compile(r"\bhref=\"([^\"]*)\"", re.IGNORECASE)
    ATTR_WW_STATUS = re.compile(r"\bww_status=\"([^\"]*)\"", re.IGNORECASE)

    SMARTTAG_PAIR = re.compile(
        r"<smarttag>.*?<name>(.*?)</name>.*?<value>(.*?)</value>.*?</smarttag>",
        re.DOTALL | re.IGNORECASE,
    )
    OBJECTATTR_PAIR = re.compile(
        r"<objectattribute>.*?<name>(.*?)</name>.*?<value>(.*?)</value>.*?</objectattribute>",
        re.DOTALL | re.IGNORECASE,
    )
    PATHLINK_BLOCK = re.compile(r"<pathlink>.*?</pathlink>", re.DOTALL | re.IGNORECASE)
    NODE_TEXT = re.compile(r"<node>(.*?)</node>", re.DOTALL | re.IGNORECASE)
    HTML_TAGS = re.compile(r"<[^>]+>")


# =============================================================================
# СТРУКТУРЫ ДАННЫХ
# =============================================================================

ClashObject = namedtuple("ClashObject", ["element_id", "path"])
ClashRow = namedtuple(
    "ClashRow",
    [
        "name",
        "image_path",
        "element_id",
        "other_element_id",
        "path",
        "other_path",
        "category",
        "other_category",
    ],
)


# =============================================================================
# ИНИЦИАЛИЗАЦИЯ REVIT
# =============================================================================

out = script.get_output()
out.close_others(all_open_outputs=True)
out.set_width(1200)

SCRIPT_DIR = os.path.dirname(__file__)
NAVIGATOR_XAML_PATH = os.path.join(SCRIPT_DIR, "Navigator.xaml")
NAVIGATOR_WINDOW = None

uidoc = __revit__.ActiveUIDocument
doc = uidoc.Document


# =============================================================================
# КЭШ ЭЛЕМЕНТОВ И КАТЕГОРИЙ
# =============================================================================


class ElementCache:
    """Кэш для элементов Revit и их категорий."""

    def __init__(self, document):
        self._doc = document
        self._cache = {}
        self._category_keys = self._build_category_keys()
        self._doc_title_key = self._normalize_key(document.Title)

    def _normalize_key(self, text):
        """Нормализует строку для сравнения."""
        if not text:
            return ""
        return re.sub("[^0-9a-zA-Zа-яА-Я]+", "", text, flags=re.UNICODE).lower()

    def _build_category_keys(self):
        """Собирает нормализованные имена категорий документа."""
        keys = set()
        try:
            for cat in self._doc.Settings.Categories:
                try:
                    if cat.Name:
                        keys.add(self._normalize_key(cat.Name))
                except Exception:
                    pass
        except Exception:
            pass
        return keys

    def get_element_and_category(self, element_id):
        """Возвращает (element, category_name) для заданного ID."""
        if element_id in self._cache:
            return self._cache[element_id]

        try:
            element = self._doc.GetElement(ElementId(element_id))
            category = ""
            if element and element.Category:
                category = element.Category.Name or ""
        except Exception:
            element, category = None, ""

        self._cache[element_id] = (element, category)
        return element, category

    def matches_current_model(self, path_text):
        """Проверяет, относится ли путь к текущей модели."""
        filename = self._extract_filename(path_text)
        if not filename:
            return True

        filename_key = self._normalize_key(filename)
        return (
            filename_key in self._doc_title_key or self._doc_title_key in filename_key
        )

    def _extract_filename(self, path_text):
        """Извлекает имя файла модели из пути."""
        if not path_text:
            return ""

        for token in path_text.split("/"):
            token = token.strip()
            if Patterns.MODEL_FILE.match(token):
                return os.path.splitext(token)[0]
        return ""

    def find_category_in_path(self, segment_normalized):
        """Ищет категорию документа в сегменте пути."""
        if not segment_normalized or len(segment_normalized) < 4:
            return None

        for cat_key in self._category_keys:
            if not cat_key or len(cat_key) < 4:
                continue
            if cat_key in segment_normalized or segment_normalized in cat_key:
                return cat_key
        return None

    @property
    def doc_title_key(self):
        return self._doc_title_key

    def normalize_key(self, text):
        return self._normalize_key(text)


# Глобальный экземпляр кэша
element_cache = ElementCache(doc)


# =============================================================================
# ЧТЕНИЕ И ОЧИСТКА XML
# =============================================================================


class XmlReader:
    """Класс для безопасного чтения XML-файлов."""

    ENCODINGS = ("utf-8-sig", "utf-16", "utf-8", "cp1251")

    @classmethod
    def read_file(cls, path):
        """Читает файл с автоопределением кодировки."""
        with open(path, "rb") as f:
            data = f.read()

        for encoding in cls.ENCODINGS:
            try:
                return data.decode(encoding)
            except Exception:
                pass
        return data.decode("utf-8", "ignore")

    @classmethod
    def sanitize(cls, text):
        """Очищает текст от некорректных XML-символов."""
        # Нормализация переносов строк
        text = text.replace("\r\n", "\n").replace("\r", "\n")
        text = text.replace("\ufeff", " ")

        # Удаление недопустимых символов
        text = cls._strip_invalid_chars(text)

        # Экранирование амперсандов
        text = Patterns.BAD_AMPERSAND.sub("&amp;", text)

        # Обрезка до начала XML
        start = text.find("<")
        if start > 0:
            text = text[start:]

        return text

    @staticmethod
    def _strip_invalid_chars(text):
        """Удаляет недопустимые XML-символы."""
        result = []
        for char in text:
            code = ord(char)
            if (
                code == 0x9
                or code == 0xA
                or code == 0xD
                or (0x20 <= code <= 0xD7FF)
                or (0xE000 <= code <= 0xFFFD)
                or (0x10000 <= code <= 0x10FFFF)
            ):
                result.append(char)
            else:
                result.append(" ")
        return "".join(result)

    @classmethod
    def parse(cls, xml_path):
        """Парсит XML-файл и возвращает корневой элемент."""
        try:
            from lxml import etree as ET
        except ImportError:
            import xml.etree.ElementTree as ET

        raw_text = cls.read_file(xml_path)
        cleaned_text = cls.sanitize(raw_text)
        return ET.fromstring(cleaned_text)


def build_clash_key(test_name, clash_name, id1, id2):
    """Строит стабильный ключ коллизии для статуса/комментария."""
    try:
        a = int(id1) if id1 is not None else 0
    except Exception:
        a = 0
    try:
        b = int(id2) if id2 is not None else 0
    except Exception:
        b = 0

    if a and b:
        lo, hi = (a, b) if a < b else (b, a)
        pair = "{}:{}".format(lo, hi)
    else:
        pair = "{}:{}".format(a, b)

    return "{}|{}|{}".format(
        (test_name or "").strip(), (clash_name or "").strip(), pair
    )


def normalize_status(raw_status):
    """Нормализует текст статуса к одному из допустимых значений."""
    status = (raw_status or "").strip().lower()
    if not status:
        return ALLOWED_STATUSES[0]
    if status in ("новая", "новый", "new"):
        return "Новая"
    if status in ("в работе", "work", "in progress", "in_work"):
        return "В работе"
    if status in ("закрыта", "закрыто", "closed", "done"):
        return "Закрыта"
    return ALLOWED_STATUSES[0]


def comments_json_path(xml_path):
    """Путь к json с комментариями рядом с XML."""
    base, _ = os.path.splitext(xml_path)
    return base + ".comments.json"


def load_comments_map(json_path):
    """Загружает комментарии из JSON: clash_key -> comment."""
    if not json_path or not os.path.exists(json_path):
        return {}
    try:
        with open(json_path, "r", encoding="utf-8") as f:
            data = json.load(f)
    except Exception:
        return {}

    if not isinstance(data, dict):
        return {}

    result = {}
    for key, value in data.items():
        if isinstance(value, dict):
            comment = str(value.get("comment") or "")
        else:
            comment = str(value or "")
        result[str(key)] = comment
    return result


def save_comments_map(json_path, comments_map):
    """Сохраняет комментарии в JSON рядом с XML."""
    payload = {}
    for key, comment in comments_map.items():
        text = (comment or "").strip()
        if text:
            payload[key] = {"comment": text}

    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)


# =============================================================================
# ПАРСИНГ XML-ОТЧЁТА
# =============================================================================


class NavisworksReportParser:
    """Парсер XML-отчётов Navisworks."""

    def __init__(self, xml_path):
        self.xml_path = xml_path
        self.base_dir = os.path.dirname(xml_path)

    def parse(self):
        """Парсит отчёт и возвращает dict: test_name -> [ClashRow, ...]."""
        try:
            return self._parse_with_xml_parser()
        except Exception:
            return self._parse_with_regex()

    def _parse_with_xml_parser(self):
        """Основной парсер на базе XML."""
        root = XmlReader.parse(self.xml_path)

        # Собираем имена всех тестов
        test_names = self._collect_test_names(root)
        groups = {name: [] for name in test_names}

        # Строим карту родительских элементов
        parent_map = self._build_parent_map(root)

        # Парсим результаты
        for clash_result in root.findall(".//clashresult"):
            self._process_clash_result(clash_result, groups, parent_map)

        return groups

    def _collect_test_names(self, root):
        """Собирает имена всех тестов."""
        names = set()
        for tag in ("clashtest", "test"):
            for element in root.findall(".//{}".format(tag)):
                name = element.get("name") or element.get("displayname")
                if name:
                    names.add(name)
        return names

    def _build_parent_map(self, root):
        """Строит карту ребёнок -> родитель."""
        parent_map = {}
        for parent in root.iter():
            for child in list(parent):
                parent_map[child] = parent
        return parent_map

    def _process_clash_result(self, clash_result, groups, parent_map):
        """Обрабатывает один clashresult."""
        test_name = self._get_test_name(clash_result, parent_map)
        clash_name = clash_result.get("name") or "Без имени"
        status = normalize_status(
            clash_result.get("ww_status") or clash_result.get("status")
        )
        image_path = self._resolve_image_path(clash_result.get("href"))

        objects = []
        for clash_object in clash_result.findall("./clashobjects/clashobject"):
            obj = self._parse_clash_object(clash_object)
            if obj:
                objects.append(obj)

        if not objects:
            return

        rows = groups.setdefault(test_name, [])
        self._add_rows(rows, test_name, clash_name, image_path, objects, status)

    def _get_test_name(self, clash_result, parent_map):
        """Определяет имя теста для clashresult."""
        # Сначала ищем в атрибутах самого элемента
        for attr in ("testname", "test", "groupname"):
            name = clash_result.get(attr)
            if name:
                return name

        # Ищем в родительских элементах
        parent = parent_map.get(clash_result)
        while parent is not None:
            name = (
                parent.get("name")
                or parent.get("displayname")
                or parent.get("testname")
            )
            if name:
                return name
            parent = parent_map.get(parent)

        return "(Без названия проверки)"

    def _parse_clash_object(self, element):
        """Парсит clashobject и возвращает ClashObject."""
        element_id = self._extract_object_id(element)
        if element_id is None:
            return None

        path = self._extract_pathlink(element)
        return ClashObject(element_id=element_id, path=path)

    def _extract_object_id(self, element):
        """Извлекает ID объекта из smarttags или objectattribute."""
        # Ищем в smarttags
        for smarttag in element.findall("./smarttags/smarttag"):
            name = (smarttag.findtext("name") or "").strip().lower()
            if name in OBJECT_ID_NAMES:
                value = (smarttag.findtext("value") or "").strip()
                if value.isdigit():
                    return int(value)

        # Ищем в objectattribute
        for obj_attr in element.findall("./objectattribute"):
            name = (obj_attr.findtext("name") or "").strip().lower()
            if name in OBJECT_ID_NAMES:
                value = (obj_attr.findtext("value") or "").strip()
                if value.isdigit():
                    return int(value)

        return None

    def _extract_pathlink(self, element):
        """Извлекает путь из pathlink."""
        parts = []

        # Пробуем node
        for node in element.findall("./pathlink/node"):
            text = (node.text or "").strip()
            if text:
                parts.append(text)

        # Если не нашли, пробуем path
        if not parts:
            for path in element.findall("./pathlink/path"):
                text = (path.text or "").strip()
                if text:
                    parts.append(text)

        return " / ".join(parts)

    def _resolve_image_path(self, href):
        """Преобразует относительный путь к изображению в абсолютный."""
        if not href:
            return ""
        href = href.replace("\\", "/").lstrip("./")
        return os.path.normpath(os.path.join(self.base_dir, href))

    def _add_rows(self, rows, test_name, clash_name, image_path, objects, status):
        """Добавляет строки для пары объектов."""
        if len(objects) >= 2:
            obj_a, obj_b = objects[0], objects[1]
            clash_key = build_clash_key(
                test_name, clash_name, obj_a.element_id, obj_b.element_id
            )
            # Добавляем обе перестановки
            rows.append(
                {
                    "name": clash_name,
                    "img": image_path,
                    "id": obj_a.element_id,
                    "id_other": obj_b.element_id,
                    "path": obj_a.path,
                    "path_other": obj_b.path,
                    "status": status,
                    "clash_key": clash_key,
                }
            )
            rows.append(
                {
                    "name": clash_name,
                    "img": image_path,
                    "id": obj_b.element_id,
                    "id_other": obj_a.element_id,
                    "path": obj_b.path,
                    "path_other": obj_a.path,
                    "status": status,
                    "clash_key": clash_key,
                }
            )
        else:
            obj = objects[0]
            clash_key = build_clash_key(test_name, clash_name, obj.element_id, None)
            rows.append(
                {
                    "name": clash_name,
                    "img": image_path,
                    "id": obj.element_id,
                    "id_other": None,
                    "path": obj.path,
                    "path_other": "",
                    "status": status,
                    "clash_key": clash_key,
                }
            )

    def _parse_with_regex(self):
        """Fallback-парсер на базе регулярных выражений."""
        text = XmlReader.sanitize(XmlReader.read_file(self.xml_path))
        groups = {}

        # Ищем блоки тестов
        blocks = list(Patterns.CLASHTEST_BLOCK.finditer(text))
        if not blocks:
            blocks = list(Patterns.TEST_BLOCK.finditer(text))
        if not blocks:
            # Обрабатываем весь текст как один блок
            blocks = [type("FakeMatch", (), {"group": lambda self, n=0: text})()]

        for block_match in blocks:
            block_text = (
                block_match.group(0)
                if hasattr(block_match, "group")
                else str(block_match)
            )

            # Определяем имя теста
            name_match = Patterns.ATTR_NAME.search(block_text)
            default_test_name = (
                name_match.group(1) if name_match else "(Без названия проверки)"
            )
            groups.setdefault(default_test_name, [])

            # Парсим результаты
            for result_match in Patterns.CLASHRESULT_BLOCK.finditer(block_text):
                self._process_regex_result(
                    result_match.group(0), groups, default_test_name
                )

        return groups

    def _process_regex_result(self, result_text, groups, default_test_name):
        """Обрабатывает clashresult через regex."""
        # Имя теста
        testname_match = Patterns.ATTR_TESTNAME.search(result_text)
        test_name = testname_match.group(1) if testname_match else default_test_name

        # Путь к изображению
        href_match = Patterns.ATTR_HREF.search(result_text)
        image_path = self._resolve_image_path(href_match.group(1) if href_match else "")

        # Имя коллизии
        name_match = Patterns.ATTR_NAME.search(result_text)
        clash_name = name_match.group(1) if name_match else "Без имени"
        status_match = Patterns.ATTR_WW_STATUS.search(result_text)
        status = normalize_status(status_match.group(1) if status_match else "")

        # Парсим объекты
        objects = []
        for obj_match in Patterns.CLASHOBJECT_BLOCK.finditer(result_text):
            obj = self._parse_regex_object(obj_match.group(0))
            if obj:
                objects.append(obj)

        if not objects:
            return

        rows = groups.setdefault(test_name, [])
        self._add_rows(rows, test_name, clash_name, image_path, objects, status)

    def _parse_regex_object(self, obj_text):
        """Парсит clashobject через regex."""
        element_id = None

        # Ищем в smarttags
        for name, value in Patterns.SMARTTAG_PAIR.findall(obj_text):
            if (name or "").strip().lower() in OBJECT_ID_NAMES:
                v = (value or "").strip()
                if v.isdigit():
                    element_id = int(v)
                    break

        # Ищем в objectattribute
        if element_id is None:
            for name, value in Patterns.OBJECTATTR_PAIR.findall(obj_text):
                if (name or "").strip().lower() in OBJECT_ID_NAMES:
                    v = (value or "").strip()
                    if v.isdigit():
                        element_id = int(v)
                        break

        if element_id is None:
            return None

        # Извлекаем путь
        path = ""
        pathlink_match = Patterns.PATHLINK_BLOCK.search(obj_text)
        if pathlink_match:
            nodes = Patterns.NODE_TEXT.findall(pathlink_match.group(0))
            parts = [Patterns.HTML_TAGS.sub("", n).strip() for n in nodes if n.strip()]
            path = " / ".join(parts)

        return ClashObject(element_id=element_id, path=path)


def _extract_object_id_from_xml_object(clash_object):
    """Извлекает ID из XML-узла clashobject."""
    for smarttag in clash_object.findall("./smarttags/smarttag"):
        name = (smarttag.findtext("name") or "").strip().lower()
        if name in OBJECT_ID_NAMES:
            value = (smarttag.findtext("value") or "").strip()
            if value.isdigit():
                return int(value)

    for obj_attr in clash_object.findall("./objectattribute"):
        name = (obj_attr.findtext("name") or "").strip().lower()
        if name in OBJECT_ID_NAMES:
            value = (obj_attr.findtext("value") or "").strip()
            if value.isdigit():
                return int(value)

    return None


def save_statuses_to_xml(xml_path, statuses_map):
    """Сохраняет статусы коллизий в XML как атрибут ww_status у clashresult."""
    if not statuses_map:
        return

    try:
        try:
            from lxml import etree as ET
        except ImportError:
            import xml.etree.ElementTree as ET

        root = XmlReader.parse(xml_path)

        parent_map = {}
        for parent in root.iter():
            for child in list(parent):
                parent_map[child] = parent

        parser_stub = NavisworksReportParser(xml_path)

        modified = 0
        for clash_result in root.findall(".//clashresult"):
            test_name = parser_stub._get_test_name(clash_result, parent_map)
            clash_name = clash_result.get("name") or "Без имени"

            ids = []
            for clash_object in clash_result.findall("./clashobjects/clashobject"):
                obj_id = _extract_object_id_from_xml_object(clash_object)
                if obj_id is not None:
                    ids.append(obj_id)

            id1 = ids[0] if len(ids) > 0 else None
            id2 = ids[1] if len(ids) > 1 else None
            key = build_clash_key(test_name, clash_name, id1, id2)

            if key in statuses_map:
                clash_result.set("ww_status", normalize_status(statuses_map[key]))
                modified += 1

        if modified == 0:
            return

        xml_bytes = ET.tostring(root, encoding="utf-8")
        with open(xml_path, "wb") as f:
            f.write(xml_bytes)
    except Exception as exc:
        out.print_md(":warning: Не удалось сохранить статусы в XML: `{}`".format(exc))


def deduplicate_groups_by_clash_key(groups):
    """Убирает дубли пар (A-B и B-A), оставляя одну строку на clash_key."""
    result = {}
    for test_name, rows in groups.items():
        seen = set()
        deduped = []
        for item in rows:
            clash_key = item.get("clash_key") or build_clash_key(
                test_name,
                item.get("name"),
                item.get("id"),
                item.get("id_other"),
            )
            if clash_key in seen:
                continue
            seen.add(clash_key)

            row = dict(item)
            row["clash_key"] = clash_key
            deduped.append(row)
        result[test_name] = deduped
    return result


# =============================================================================
# ИЗВЛЕЧЕНИЕ ДАТЫ ОТЧЁТА
# =============================================================================


def extract_report_datetime(xml_path):
    """Извлекает дату отчёта из XML или метаданных файла."""
    # Пробуем из XML-атрибутов
    try:
        root = XmlReader.parse(xml_path)

        # Ищем в корневом элементе
        for attr in DATE_ATTRIBUTES:
            value = root.get(attr)
            if value:
                return value, "из XML"

        # Ищем в дочерних элементах
        for tag in ("report", "batchtest", "tests"):
            for element in root.findall(".//{}".format(tag)):
                for attr in DATE_ATTRIBUTES:
                    value = element.get(attr)
                    if value:
                        return value, "из XML"
    except Exception:
        pass

    # Пробуем regex-поиск
    try:
        text = XmlReader.sanitize(XmlReader.read_file(xml_path))
        match = Patterns.DATE_ATTR.search(text)
        if match:
            return match.group(1).strip(), "из XML"
    except Exception:
        pass

    # Берём дату изменения файла
    try:
        timestamp = os.path.getmtime(xml_path)
        dt = datetime.fromtimestamp(timestamp)
        return dt.strftime("%d.%m.%Y %H:%M"), "по дате изменения файла"
    except Exception:
        return "", ""


# =============================================================================
# ФИЛЬТРАЦИЯ И АННОТАЦИЯ
# =============================================================================


class ResultFilter:
    """Фильтрация и аннотация результатов."""

    def __init__(self, cache):
        self.cache = cache

    def filter_and_annotate(self, items):
        """Фильтрует строки и добавляет информацию о категориях."""
        result = []

        for item in items:
            annotated = self._process_item(item)
            if annotated:
                result.append(annotated)

        return result

    def _process_item(self, item):
        """Обрабатывает одну строку."""
        try:
            element_id = int(item["id"])
            element, category = self.cache.get_element_and_category(element_id)

            if not element:
                return None

            # Фильтр по файлу
            if FILTER_BY_FILE:
                path_matches = self.cache.matches_current_model(
                    item.get("path", "")
                ) or self.cache.matches_current_model(item.get("path_other", ""))
                if not path_matches:
                    return None

            # Фильтр по категории
            if FILTER_BY_CATEGORY and category:
                cat_lower = category.strip().lower()
                path1 = (item.get("path", "") or "").lower()
                path2 = (item.get("path_other", "") or "").lower()
                if cat_lower not in path1 and cat_lower not in path2:
                    return None

            # Создаём аннотированную копию
            result = dict(item)
            result["cat"] = category

            # Категория второго элемента
            other_id = item.get("id_other")
            if other_id:
                other_element, other_category = self.cache.get_element_and_category(
                    int(other_id)
                )
                result["cat_other"] = other_category if other_element else ""
            else:
                result["cat_other"] = ""

            return result

        except Exception:
            return None


# =============================================================================
# ПОДСВЕТКА КАТЕГОРИЙ В ПУТИ
# =============================================================================


class PathHighlighter:
    """Подсветка категорий в путях элементов."""

    def __init__(self, cache):
        self.cache = cache

    def highlight(self, path_text, category_name):
        """Подсвечивает категорию в пути."""
        if not path_text:
            return "—"

        segments = [s.strip() for s in path_text.split("/")]
        highlighted = False

        # 1. Пробуем подсветить по имени категории
        if category_name:
            category_key = self.cache.normalize_key(category_name.strip())
            if category_key:
                for i, segment in enumerate(segments):
                    if category_key in self.cache.normalize_key(segment):
                        segments[i] = self._wrap_in_badge(segment)
                        highlighted = True
                        break

        # 2. Если не нашли - ищем по списку категорий документа
        if not highlighted:
            start_index = self._find_start_after_filename(segments)
            for i in range(start_index, len(segments)):
                if "<span" in segments[i]:
                    continue

                segment_key = self.cache.normalize_key(segments[i])
                if self.cache.find_category_in_path(segment_key):
                    segments[i] = self._wrap_in_badge(segments[i])
                    break

        return " / ".join(segments)

    def _find_start_after_filename(self, segments):
        """Находит индекс сегмента после имени файла."""
        for i, segment in enumerate(segments):
            if Patterns.MODEL_FILE.match(segment):
                return i + 1
        return 0

    def _wrap_in_badge(self, text):
        """Оборачивает текст в span с подсветкой."""
        return '<span style="{style}">{text}</span>'.format(
            style=BADGE_STYLE, text=text
        )


# =============================================================================
# СТАТИСТИКА
# =============================================================================


class StatisticsBuilder:
    """Построение статистики по коллизиям."""

    def __init__(self, highlighter):
        self.highlighter = highlighter

    def build(self, filtered_groups):
        """Строит статистику по группам."""
        stats = {}

        for test_name, rows in filtered_groups.items():
            seen_pairs = set()
            category_pairs = set()

            for item in rows:
                pair = self._extract_pair(item)
                if not pair or pair in seen_pairs:
                    continue

                seen_pairs.add(pair)
                cat_pair = self._extract_category_pair(item)
                if cat_pair:
                    category_pairs.add(cat_pair)

            stats[test_name] = {"pairs": len(seen_pairs), "catpairs": category_pairs}

        return stats

    def _extract_pair(self, item):
        """Извлекает пару ID элементов."""
        id1 = item.get("id")
        id2 = item.get("id_other")

        if not id1 or not id2:
            return None

        try:
            a, b = int(id1), int(id2)
            return (a, b) if a < b else (b, a)
        except Exception:
            return None

    def _extract_category_pair(self, item):
        """Извлекает пару категорий."""
        cat1 = (item.get("cat") or "").strip()
        cat2 = (item.get("cat_other") or "").strip()

        # Если вторая категория не определена - пробуем извлечь из пути
        if not cat2:
            highlighted = self.highlighter.highlight(item.get("path_other") or "", "")
            # Извлекаем текст из span
            match = re.search("<span[^>]*>([^<]+)</span>", highlighted)
            if match:
                cat2 = match.group(1)
            else:
                cat2 = re.sub("<[^>]*>", "", highlighted)

        if cat1 or cat2:
            return tuple(sorted([cat1 or "—", cat2 or "—"]))
        return None


class ModelGroupingBuilder:
    """Группировка строк по NWC-моделям, затем по проверкам."""

    UNKNOWN_MODEL = "(Без NWC модели)"

    def __init__(self, cache):
        self.cache = cache

    def build(self, filtered_groups):
        """Возвращает dict: model_name -> {test_name: [rows]}"""
        grouped = {}

        for test_name, rows in filtered_groups.items():
            for item in rows:
                model_name = self._pick_model_name(item)
                model_bucket = grouped.setdefault(model_name, {})
                model_bucket.setdefault(test_name, []).append(item)

        return grouped

    def _pick_model_name(self, item):
        """Выбирает одну модель для строки.

        Приоритет:
        1) Модель, отличная от текущей открытой Revit-модели.
        2) Если обе стороны текущие/неопределимы — fallback на path, затем path_other.
        """
        path = item.get("path") or ""
        path_other = item.get("path_other") or ""

        path_model = self._extract_primary_nwc_from_path(path)
        path_other_model = self._extract_primary_nwc_from_path(path_other)

        path_is_current = bool(path_model) and self.cache.matches_current_model(path)
        path_other_is_current = bool(
            path_other_model
        ) and self.cache.matches_current_model(path_other)

        # Главное правило: группировать по внешней (не текущей) модели
        if path_other_model and not path_other_is_current:
            return path_other_model
        if path_model and not path_is_current:
            return path_model

        # Fallback: чтобы строка всегда попадала ровно в одну группу
        if path_model:
            return path_model
        if path_other_model:
            return path_other_model

        return self.UNKNOWN_MODEL

    def _extract_primary_nwc_from_path(self, path_text):
        """Извлекает первое имя файла .nwc (без расширения) из pathlink."""
        if not path_text:
            return None

        for token in path_text.split("/"):
            token = token.strip()
            if token.lower().endswith(".nwc"):
                return os.path.splitext(token)[0]
        return None


class SelectElementExternalEventHandler(IExternalEventHandler):
    """ExternalEvent для безопасного выделения элементов из немодального окна."""

    def __init__(self):
        self._element_id = None
        self._action = "select"
        self._last_error = ""

    def set_target(self, element_id, action):
        self._element_id = element_id
        self._action = action or "select"

    def Execute(self, uiapp):
        self._last_error = ""
        if self._element_id is None:
            self._last_error = "Не задан ID элемента для действия."
            return

        try:
            active_uidoc = uiapp.ActiveUIDocument
            if not active_uidoc:
                self._last_error = "Нет активного документа Revit."
                return

            _perform_revit_action(active_uidoc, int(self._element_id), self._action)
        except Exception as exc:
            self._last_error = str(exc)

    def GetName(self):
        return "WWBIM Select Element External Event"


def _perform_revit_action(uidoc, element_id, action):
    """Выполняет выделение/подрезку в контексте Revit API."""
    ids = List[ElementId]()
    ids.Add(ElementId(int(element_id)))
    uidoc.Selection.SetElementIds(ids)
    uidoc.ShowElements(ElementId(int(element_id)))

    if action != "crop":
        return

    doc = uidoc.Document
    view = doc.ActiveView
    if not isinstance(view, View3D) or view.IsTemplate:
        raise Exception("Подрезка работает только в активном 3D-виде.")

    element = doc.GetElement(ElementId(int(element_id)))
    if not element:
        raise Exception("Элемент с ID {} не найден в документе.".format(element_id))

    bbox = element.get_BoundingBox(view) or element.get_BoundingBox(None)
    if not bbox:
        raise Exception("Не удалось получить bounding box элемента.")

    pad = 1.0
    section = BoundingBoxXYZ()
    section.Min = XYZ(bbox.Min.X - pad, bbox.Min.Y - pad, bbox.Min.Z - pad)
    section.Max = XYZ(bbox.Max.X + pad, bbox.Max.Y + pad, bbox.Max.Z + pad)

    tx = Transaction(doc, "WWBIM Подрезка коллизии")
    tx.Start()
    try:
        if not view.IsSectionBoxActive:
            view.IsSectionBoxActive = True
        view.SetSectionBox(section)
        tx.Commit()
    except Exception:
        tx.RollBack()
        raise


class ClashNavigatorWindow(forms.WPFWindow):
    """Интерактивный навигатор коллизий (WPF: сортировка + фильтр)."""

    HEADER_FILTER_COLUMNS = (
        "Модель NWC",
        "Проверка",
        "Пересечение",
        "ID",
        "ID_2",
        "Статус",
        "Категория",
        "Категория_2",
    )

    def __init__(self, rows, uidoc, xaml_path, xml_path, comments_path):
        self._uidoc = uidoc
        self._rows = rows
        self._xml_path = xml_path
        self._comments_path = comments_path
        self._source_table = None
        self._view = None
        self._suspend_filter_events = False
        self._column_value_filters = {}
        self._header_filter_buttons = {}
        self._select_handler = SelectElementExternalEventHandler()
        self._select_external_event = None

        forms.WPFWindow.__init__(self, xaml_path)

        try:
            self._select_external_event = ExternalEvent.Create(self._select_handler)
        except Exception as exc:
            forms.alert(
                "Не удалось инициализировать ExternalEvent: {}".format(exc),
                title="Пересечения",
            )

        self.FilterTextBox.TextChanged += self._on_filter_changed
        self.ResetFilterButton.Click += self._on_reset_clicked
        self.SaveChangesButton.Click += self._on_save_changes_clicked
        self.SelectElementButton.Click += self._on_select_element_clicked
        self.CropToElementButton.Click += self._on_crop_to_element_clicked
        self.ClashesGrid.AutoGeneratingColumn += self._on_grid_auto_generating_column
        self.ClashesGrid.AutoGeneratedColumns += self._on_grid_auto_generated_columns
        self.ClashesGrid.MouseDoubleClick += self._on_grid_double_click

        self._bind_data()

    def _bind_data(self):
        table = DataTable("Clashes")
        columns = [
            "Ключ",
            "Модель NWC",
            "Проверка",
            "Пересечение",
            "ID",
            "ID_2",
            "Статус",
            "Комментарий",
            "Категория",
            "Категория_2",
            "Путь элемента",
            "Путь второго",
            "Снимок",
        ]
        for col_name in columns:
            table.Columns.Add(col_name)

        for row in self._rows:
            test_name = (row.get("test") or "").strip()
            table.Rows.Add(
                row.get("clash_key", ""),
                row.get("model", ""),
                test_name,
                row.get("name", ""),
                str(row.get("id", "") or ""),
                str(row.get("id_other", "") or ""),
                normalize_status(row.get("status", "")),
                row.get("comment", "") or "",
                row.get("cat", "") or "",
                row.get("cat_other", "") or "",
                row.get("path", "") or "",
                row.get("path_other", "") or "",
                self._build_preview_image(row.get("img", "") or ""),
            )

        self._source_table = table
        self._view = table.DefaultView
        self.ClashesGrid.ItemsSource = self._view

    def _on_grid_auto_generating_column(self, sender, args):
        header = str(args.Column.Header)
        if header == "Ключ":
            args.Cancel = True
            return

        if header == "Снимок":
            template_column = DataGridTemplateColumn()
            template_column.Header = "Снимок"
            image_factory = FrameworkElementFactory(Image)
            image_factory.SetBinding(Image.SourceProperty, Binding("[Снимок]"))
            image_factory.SetValue(Image.WidthProperty, 96.0)
            image_factory.SetValue(Image.HeightProperty, 72.0)
            image_factory.SetValue(Image.StretchProperty, Stretch.Uniform)
            cell_template = DataTemplate()
            cell_template.VisualTree = image_factory
            template_column.CellTemplate = cell_template
            template_column.IsReadOnly = True
            template_column.Header = "Снимок"
            args.Column = template_column
            return

        if header == "Статус":
            combo = DataGridComboBoxColumn()
            combo.Header = "Статус"
            combo.ItemsSource = list(ALLOWED_STATUSES)
            combo.SelectedItemBinding = args.Column.Binding
            combo.IsReadOnly = False
            args.Column = combo

        if header in self.HEADER_FILTER_COLUMNS:
            self._set_filter_header(args.Column, header)

        if header in ("Комментарий", "Статус"):
            args.Column.IsReadOnly = False
        else:
            args.Column.IsReadOnly = True

    def _set_filter_header(self, column, column_name):
        header_panel = StackPanel()
        header_panel.Orientation = Orientation.Horizontal

        title_row = StackPanel()
        title_row.Orientation = Orientation.Horizontal

        title = TextBlock()
        title.Text = column_name
        title.Margin = Thickness(0, 0, 4, 2)

        filter_button = Button()
        filter_button.Content = "▼"
        filter_button.Width = 18
        filter_button.Height = 18
        filter_button.Tag = column_name
        filter_button.Padding = Thickness(0)
        filter_button.ToolTip = "Фильтр значений"
        filter_button.Click += self._on_header_values_filter_click

        title_row.Children.Add(title)
        title_row.Children.Add(filter_button)
        header_panel.Children.Add(title_row)

        column.Header = header_panel
        self._header_filter_buttons[column_name] = filter_button
        self._update_header_filter_indicator(column_name)

    def _on_header_values_filter_click(self, sender, args):
        column_name = str(sender.Tag or "").strip()
        if not column_name or self._source_table is None:
            return

        unique_values = self._collect_unique_values(column_name)
        if not unique_values:
            forms.alert(
                "В колонке '{}' нет значений для фильтрации.".format(column_name),
                title="Пересечения",
            )
            return

        selected_prev = set(self._column_value_filters.get(column_name, []))
        options = [
            forms.TemplateListItem(v, checked=(v in selected_prev))
            for v in unique_values
        ]

        selected = forms.SelectFromList.show(
            options,
            title="Фильтр значений: {}".format(column_name),
            multiselect=True,
            button_name="Применить",
        )

        if selected is None:
            return

        selected_values = []
        for v in selected:
            value = v.item if hasattr(v, "item") else v
            value = str(value).strip()
            if value:
                selected_values.append(value)
        all_selected = len(selected_values) == len(unique_values)

        if not selected_values or all_selected:
            if column_name in self._column_value_filters:
                del self._column_value_filters[column_name]
        else:
            self._column_value_filters[column_name] = selected_values

        self._update_header_filter_indicator(column_name)
        self._apply_filters()

    def _collect_unique_values(self, column_name):
        if self._source_table is None:
            return []

        values = set()
        for data_row in self._source_table.Rows:
            value = str(data_row[column_name] or "").strip()
            if value:
                values.add(value)
        return sorted(values, key=lambda s: s.lower())

    def _on_grid_auto_generated_columns(self, sender, args):
        try:
            image_col = None
            for col in self.ClashesGrid.Columns:
                if str(col.Header) == "Снимок":
                    image_col = col
                    break

            if image_col is None:
                return

            image_col.DisplayIndex = 0

            next_index = 1
            for col in self.ClashesGrid.Columns:
                if col is image_col:
                    continue
                col.DisplayIndex = next_index
                next_index += 1
        except Exception:
            pass

    def _build_preview_image(self, image_path):
        if not image_path or not os.path.exists(image_path):
            return None

        try:
            bmp = BitmapImage()
            bmp.BeginInit()
            bmp.UriSource = Uri(image_path, UriKind.Absolute)
            bmp.CacheOption = BitmapCacheOption.OnLoad
            bmp.EndInit()
            bmp.Freeze()
            return bmp
        except Exception:
            return None

    def _on_save_changes_clicked(self, sender, args):
        statuses_map = {}
        comments_map = {}

        if self._source_table is None:
            return

        try:
            for data_row in self._source_table.Rows:
                key = str(data_row["Ключ"] or "").strip()
                if not key:
                    continue

                status = normalize_status(str(data_row["Статус"] or ""))
                comment = str(data_row["Комментарий"] or "").strip()

                data_row["Статус"] = status
                statuses_map[key] = status
                if comment:
                    comments_map[key] = comment
        except Exception as exc:
            forms.alert(
                "Не удалось собрать изменения: {}".format(exc),
                title="Пересечения",
            )
            return

        save_statuses_to_xml(self._xml_path, statuses_map)
        try:
            save_comments_map(self._comments_path, comments_map)
        except Exception as exc:
            forms.alert(
                "Не удалось сохранить комментарии JSON: {}".format(exc),
                title="Пересечения",
            )
            return

        out.print_md(
            "✅ Сохранено: статусы в XML, комментарии в `{}`".format(
                self._comments_path
            )
        )

    def _on_reset_clicked(self, sender, args):
        self._suspend_filter_events = True
        try:
            self.FilterTextBox.Text = ""
            self._column_value_filters = {}
        finally:
            self._suspend_filter_events = False

        self._update_all_header_filter_indicators()
        self._apply_filters()

    def _update_all_header_filter_indicators(self):
        for col_name in self._header_filter_buttons.keys():
            self._update_header_filter_indicator(col_name)

    def _update_header_filter_indicator(self, column_name):
        button = self._header_filter_buttons.get(column_name)
        if button is None:
            return

        selected_values = self._column_value_filters.get(column_name)
        if selected_values:
            button.Content = "●▼"
            button.Foreground = Brushes.OrangeRed
            button.ToolTip = "Фильтр активен ({})".format(len(selected_values))
        else:
            button.Content = "▼"
            button.Foreground = Brushes.DimGray
            button.ToolTip = "Фильтр значений"

    def _on_filter_changed(self, sender, args):
        if self._suspend_filter_events:
            return
        self._apply_filters()

    def _apply_filters(self):
        if self._view is None:
            return

        term = (self.FilterTextBox.Text or "").strip()
        parts = []

        if term:
            escaped = self._escape_row_filter_value(term)
            parts.append(
                "("
                "[Модель NWC] LIKE '%{0}%' OR "
                "[Проверка] LIKE '%{0}%' OR "
                "[Пересечение] LIKE '%{0}%' OR "
                "[ID] LIKE '%{0}%' OR "
                "[ID_2] LIKE '%{0}%' OR "
                "[Категория] LIKE '%{0}%' OR "
                "[Категория_2] LIKE '%{0}%' OR "
                "[Путь элемента] LIKE '%{0}%' OR "
                "[Путь второго] LIKE '%{0}%'"
                ")".format(escaped)
            )

        for col_name, selected_values in self._column_value_filters.items():
            escaped_col = self._escape_row_filter_column(col_name)
            exprs = []
            for value in selected_values:
                escaped_val = self._escape_row_filter_literal(value)
                exprs.append("[{}] = '{}'".format(escaped_col, escaped_val))
            if exprs:
                parts.append("({})".format(" OR ".join(exprs)))

        if not parts:
            self._view.RowFilter = ""
            return

        try:
            self._view.RowFilter = " AND ".join(parts)
        except Exception:
            self._view.RowFilter = ""

    def _escape_row_filter_value(self, value):
        value = value.replace("'", "''")
        value = value.replace("[", "[[]")
        value = value.replace("]", "[]]")
        value = value.replace("%", "[%]")
        value = value.replace("_", "[_]")
        value = value.replace("*", "[*]")
        return value

    def _escape_row_filter_literal(self, value):
        return value.replace("'", "''")

    def _escape_row_filter_column(self, value):
        return value.replace("]", "]]")

    def _on_grid_double_click(self, sender, args):
        self._run_action_on_selected("select")

    def _on_select_element_clicked(self, sender, args):
        self._run_action_on_selected("select")

    def _on_crop_to_element_clicked(self, sender, args):
        self._run_action_on_selected("crop")

    def _run_action_on_selected(self, action):
        element_id = self._get_selected_element_id()
        if element_id is None:
            forms.alert("Выберите строку с валидным ID.", title="Пересечения")
            return

        # В modeless окне работаем только через ExternalEvent (без прямых API-вызовов).
        if self._select_external_event is None:
            forms.alert(
                "ExternalEvent не инициализирован. Перезапустите команду.",
                title="Пересечения",
            )
            return

        try:
            self._select_handler.set_target(element_id, action)
            result = self._select_external_event.Raise()
            if str(result) not in ("Accepted", "Pending"):
                forms.alert(
                    "ExternalEvent отклонен: {}".format(result),
                    title="Пересечения",
                )
        except Exception as exc:
            forms.alert(
                "Не удалось выполнить действие '{}': {}".format(action, exc),
                title="Пересечения",
            )

    def _get_selected_element_id(self):
        selected = self.ClashesGrid.SelectedItem
        if selected is None:
            return None

        try:
            id_value = selected["ID"]
        except Exception:
            try:
                id_value = selected.Row["ID"]
            except Exception:
                return None

        try:
            return int(str(id_value))
        except Exception:
            return None


def build_navigator_rows(grouped_by_model, comments_map):
    """Преобразует grouped_by_model в плоский набор строк для интерактивного грида."""
    rows = []

    for model_name in sorted(grouped_by_model.keys(), key=lambda s: s.lower()):
        test_groups = grouped_by_model[model_name]
        for test_name in sorted(test_groups.keys(), key=lambda s: s.lower()):
            for item in test_groups[test_name]:
                row = dict(item)
                row["model"] = model_name
                row["test"] = test_name
                clash_key = row.get("clash_key") or build_clash_key(
                    test_name,
                    row.get("name"),
                    row.get("id"),
                    row.get("id_other"),
                )
                row["clash_key"] = clash_key
                row["comment"] = comments_map.get(clash_key, "")
                rows.append(row)

    return rows


# =============================================================================
# ВЫВОД РЕЗУЛЬТАТОВ
# =============================================================================


class ResultPrinter:
    """Вывод результатов в UI."""

    def __init__(self, output, highlighter):
        self.out = output
        self.highlighter = highlighter

    def print_report_date(self, date_value, date_source):
        """Печатает дату отчёта."""
        if date_value:
            self.out.print_md(
                "**Дата отчёта Navisworks:** {} _({})_".format(date_value, date_source)
            )

    def print_summary(self, filtered_groups, stats):
        """Печатает краткую сводку."""
        self.out.print_md("### Краткая сводка по выбранным проверкам")

        lines = []
        for test_name in sorted(filtered_groups.keys(), key=lambda s: s.lower()):
            st = stats.get(test_name, {"pairs": 0, "catpairs": set()})
            cat_pairs_str = "; ".join(
                "{} × {}".format(a, b) for (a, b) in sorted(st["catpairs"])
            )
            lines.append(
                "- **{}** — {} коллизий; пары категорий: {}".format(
                    test_name, st["pairs"], cat_pairs_str or "—"
                )
            )

        self.out.print_md("\n".join(lines))

    def print_total(self, filtered_groups):
        """Печатает общее количество."""
        total_rows = sum(len(rows) for rows in filtered_groups.values())
        self.out.print_md(
            "## Пересечения: {} строк ({} проверок)".format(
                total_rows, len(filtered_groups)
            )
        )

    def print_group(self, title, rows):
        """Печатает таблицу для группы проверок."""
        self.out.print_md(
            "\n---\n### Проверка: **{}**  _(строк: {})_".format(title, len(rows))
        )

        if not rows:
            self.out.print_md("—")
            return

        table_data = []
        for i, item in enumerate(rows, 1):
            path1 = self.highlighter.highlight(
                item.get("path") or "", item.get("cat") or ""
            )
            path2 = self.highlighter.highlight(
                item.get("path_other") or "", item.get("cat_other") or ""
            )

            table_data.append(
                [
                    i,
                    self._format_image(item.get("img")),
                    title,
                    item.get("name") or "",
                    self.out.linkify(ElementId(int(item["id"]))),
                    path1 or "—",
                    path2 or "—",
                ]
            )

        self.out.print_table(
            table_data=table_data,
            columns=[
                "№",
                "Снимок",
                "Название проверки",
                "Пересечение",
                "ID",
                "Путь элемента",
                "Путь второго элемента",
            ],
            title=None,
        )

    def print_footer(self):
        """Печатает подвал."""
        self.out.print_md("_Клик по **ID** выделяет элемент. Можно кликать подряд._")

    def _format_image(self, path, width=96):
        """Форматирует ячейку с изображением."""
        if not path or not os.path.exists(path):
            return "—"
        uri = "file:///" + path.replace("\\", "/")
        return '<img src="{}" width="{}" />'.format(uri, int(width))


# =============================================================================
# ГЛАВНАЯ ФУНКЦИЯ
# =============================================================================


def main():
    global NAVIGATOR_WINDOW
    # Выбор файла
    xml_path = forms.pick_file(
        files_filter="XML (*.xml)|*.xml", title="Выберите XML отчёт Navisworks"
    )
    if not xml_path:
        forms.alert("Файл не выбран.", title="Пересечения")
        return

    # Парсинг
    parser = NavisworksReportParser(xml_path)
    groups = parser.parse()

    if not groups:
        forms.alert("Не удалось извлечь данные из отчёта.", title="Пересечения")
        return

    # Инициализация компонентов
    result_filter = ResultFilter(element_cache)
    model_grouping_builder = ModelGroupingBuilder(element_cache)

    # Фильтрация
    filtered_groups = {
        test: result_filter.filter_and_annotate(rows) for test, rows in groups.items()
    }
    filtered_groups = deduplicate_groups_by_clash_key(filtered_groups)

    grouped_by_model = model_grouping_builder.build(filtered_groups)
    comments_path = comments_json_path(xml_path)
    comments_map = load_comments_map(comments_path)

    navigator_rows = build_navigator_rows(grouped_by_model, comments_map)
    if not navigator_rows:
        forms.alert(
            "После фильтрации не осталось строк для навигации.", title="Пересечения"
        )
        return

    if not os.path.exists(NAVIGATOR_XAML_PATH):
        forms.alert("Не найден Navigator.xaml рядом со скриптом.", title="Пересечения")
        return

    NAVIGATOR_WINDOW = ClashNavigatorWindow(
        navigator_rows,
        uidoc,
        NAVIGATOR_XAML_PATH,
        xml_path,
        comments_path,
    )
    NAVIGATOR_WINDOW.show(modal=False)


if __name__ == "__main__":
    main()
