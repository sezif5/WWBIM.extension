# -*- coding: utf-8 -*-
"""
Пакетные операции с моделями: Открыть, Загрузить семейство, Добавить связь.
Использует openbg/closebg из библиотеки.
"""

from __future__ import print_function, division, unicode_literals

import os
import sys
import importlib
import datetime
from pyrevit import script, forms

from Autodesk.Revit.DB import (
    ModelPathUtils,
    Transaction,
    FilteredElementCollector,
    Family,
    RevitLinkOptions,
    RevitLinkType,
    RevitLinkInstance,
    ImportPlacement,
    OpenOptions,
    WorksetConfiguration,
    WorksetConfigurationOption,
    WorksharingUtils,
    WorksetId,
)
from System.Collections.Generic import List

import openbg
import closebg

out = script.get_output()
out.close_others(all_open_outputs=True)

# ---------- Пути ----------


def _norm(p):
    return os.path.normpath(os.path.abspath(p)) if p else p


def _module_dir():
    try:
        return os.path.dirname(os.path.abspath(__file__))
    except Exception:
        return os.getcwd()


def _find_scripts_root(start_dir):
    cur = _norm(start_dir)
    for _ in range(0, 8):
        if os.path.basename(cur).lower() == "scripts":
            return cur
        parent = os.path.dirname(cur)
        if not parent or parent == cur:
            break
        cur = parent
    return _norm(os.path.join(start_dir, os.pardir, os.pardir))


_env_root = os.environ.get("WW_SCRIPTS_ROOT")
if _env_root and os.path.isdir(_env_root):
    SCRIPTS_ROOT = _norm(_env_root)
else:
    SCRIPTS_ROOT = _find_scripts_root(_module_dir())

OBJECTS_DIR = _norm(os.path.join(SCRIPTS_ROOT, "Objects"))

PYTHON_SCRIPTS_DIR = None
if SCRIPTS_ROOT:
    module_dir = _module_dir()
    PYTHON_SCRIPTS_DIR = _norm(
        os.path.join(module_dir, "..", "..", "..", "lib", "Batch Operations")
    )

# ---------- Helpers ----------


def to_model_path(user_visible_path):
    if not user_visible_path:
        return None
    try:
        return ModelPathUtils.ConvertUserVisiblePathToModelPath(user_visible_path)
    except Exception:
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

        try:
            # Настройки открытия
            opts = OpenOptions()

            # Получить рабочие наборы для открытия
            ws_ids = get_worksets_to_open(uiapp, mp)
            if ws_ids is not None and ws_ids.Count > 0:
                # Закрыть все, потом открыть нужные
                ws_config = WorksetConfiguration(
                    WorksetConfigurationOption.CloseAllWorksets
                )
                ws_config.Open(ws_ids)
                opts.SetOpenWorksetsConfiguration(ws_config)
            else:
                # Если не удалось получить РН - открыть все
                ws_config = WorksetConfiguration(
                    WorksetConfigurationOption.OpenAllWorksets
                )
                opts.SetOpenWorksetsConfiguration(ws_config)

            # Открыть с созданием локальной копии (OpenAndActivateDocument создаёт локальную)
            uidoc = uiapp.OpenAndActivateDocument(mp, opts, False)

            if uidoc and uidoc.Document:
                out.print_md(
                    ":white_check_mark: Открыто: **{}**".format(uidoc.Document.Title)
                )
            else:
                out.print_md(":warning: Документ открыт, но не активирован")

        except Exception as e:
            out.print_md(":x: Ошибка: `{}`".format(e))

        out.update_progress(i + 1, len(models))

    out.print_md("---")
    out.print_md("**Готово.**")


def action_load_family(models):
    """Загрузить семейства из открытых документов в выбранные модели."""
    # Получить открытые семейства
    open_families = get_open_families()

    if not open_families:
        forms.alert("Нет открытых семейств для загрузки.", warn_icon=True)
        return

    # Выбрать семейства (multiselect)
    class FamilyItem:
        def __init__(self, doc):
            self.doc = doc
            self.name = doc.Title

        def __str__(self):
            return self.name

    family_items = [FamilyItem(d) for d in open_families]
    selected = forms.SelectFromList.show(
        family_items,
        title="Выберите семейства для загрузки",
        multiselect=True,
        button_name="Выбрать",
    )

    if not selected:
        return

    # Преобразовать в список если выбрано одно
    if not isinstance(selected, list):
        selected = [selected]

    family_docs = [item.doc for item in selected]
    family_names = [item.name for item in selected]

    out.print_md("## ЗАГРУЗКА СЕМЕЙСТВ")
    out.print_md(
        "**Семейства ({}):** {}".format(len(family_docs), ", ".join(family_names))
    )
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

        try:
            doc, _ = openbg.open_in_background(
                __revit__.Application,
                __revit__,
                mp,
                audit=False,
                worksets="all",
                detach=False,
                suppress_warnings=True,
            )

            # Загрузить все выбранные семейства
            loaded_count = 0
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

            # Синхронизация и закрытие
            closebg.close_with_policy(doc, do_sync=True, comment="Загрузка семейств")

        except Exception as e:
            out.print_md(":x: Ошибка открытия: `{}`".format(e))

        out.update_progress(i + 1, len(models))

    out.print_md("---")
    out.print_md(
        "**Готово. Моделей обработано: {}/{}**".format(success_models, len(models))
    )


def select_links_from_txt():
    """Выбрать связи через txt файлы из папки Objects."""
    # 1. Выбор объекта (txt файла)
    txt_files = list_txt_files(OBJECTS_DIR)

    if not txt_files:
        forms.alert(
            "Папка с объектами пуста или недоступна:\n{}".format(OBJECTS_DIR),
            warn_icon=True,
        )
        return None

    selected_object = forms.SelectFromList.show(
        txt_files, title="Выберите объект (для связей)", width=400, button_name="Далее"
    )

    if not selected_object:
        return None

    # 2. Выбор связей из списка
    txt_path = os.path.join(OBJECTS_DIR, selected_object + ".txt")
    links_list = read_models_from_txt(txt_path)

    if not links_list:
        forms.alert("Файл пуст или не найден: {}".format(txt_path), warn_icon=True)
        return None

    selected_links = forms.SelectFromList.show(
        links_list,
        title="Выберите связи для добавления",
        multiselect=True,
        width=800,
        height=600,
        button_name="Выбрать",
    )

    return selected_links


def get_existing_links(doc):
    """Получить список имён существующих связей в документе."""
    links = FilteredElementCollector(doc).OfClass(RevitLinkInstance).ToElements()
    names = set()
    for link in links:
        try:
            # Имя инстанса часто "file.rvt : location"
            name = link.Name.split(" ")[0]
            names.add(name)
        except Exception:
            pass
    return names


def add_single_link(doc, path_str, existing_links):
    """Добавить одну связь. Сначала Shared, при ошибке — Origin.
    Возвращает (success, message).
    """
    name_model = os.path.basename(path_str)

    # Проверка на дубликат
    if name_model in existing_links:
        return False, "уже существует"

    mp = ModelPathUtils.ConvertUserVisiblePathToModelPath(path_str)
    rlo = RevitLinkOptions(False)

    try:
        # Создаём тип связи
        rl_type = RevitLinkType.Create(doc, mp, rlo)

        try:
            # Пробуем по общим координатам
            rl_inst = RevitLinkInstance.Create(
                doc, rl_type.ElementId, ImportPlacement.Shared
            )
            return True, "по общим координатам"
        except Exception as e:
            # Если не совпадают СК - по внутреннему началу
            if "coordinate system" in str(e).lower() or "координат" in str(e).lower():
                rl_inst = RevitLinkInstance.Create(
                    doc, rl_type.ElementId, ImportPlacement.Origin
                )
                return True, "в начало координат"
            else:
                raise

    except Exception as e:
        return False, unicode(e)


def action_add_link(models):
    """Добавить связи в выбранные модели."""
    # Выбрать связи через txt файлы
    selected_links = select_links_from_txt()

    if not selected_links:
        return

    link_names = [os.path.basename(l) for l in selected_links]

    out.print_md("## ДОБАВЛЕНИЕ СВЯЗЕЙ")
    out.print_md(
        "**Связей ({}):** {}".format(len(selected_links), ", ".join(link_names[:5]))
    )
    if len(link_names) > 5:
        out.print_md("*... и ещё {}*".format(len(link_names) - 5))
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

        try:
            doc, _ = openbg.open_in_background(
                __revit__.Application,
                __revit__,
                mp,
                audit=False,
                worksets="all",
                detach=False,
                suppress_warnings=True,
            )

            # Получить существующие связи
            existing_links = get_existing_links(doc)

            # Добавить связи
            added = 0
            skipped = 0
            errors = 0

            t = Transaction(doc, "Добавление связей")
            t.Start()
            try:
                for link_path in selected_links:
                    link_name = os.path.basename(link_path)

                    # Не добавлять саму себя
                    if link_name == model_name or link_name.startswith(
                        model_name.replace(".rvt", "")
                    ):
                        out.print_md(
                            "  :information_source: {} — пропуск (сама себя)".format(
                                link_name
                            )
                        )
                        skipped += 1
                        continue

                    success, msg = add_single_link(doc, link_path, existing_links)
                    if success:
                        out.print_md(
                            "  :white_check_mark: {} — {}".format(link_name, msg)
                        )
                        added += 1
                    elif msg == "уже существует":
                        out.print_md(
                            "  :information_source: {} — уже существует".format(
                                link_name
                            )
                        )
                        skipped += 1
                    else:
                        out.print_md("  :x: {} — {}".format(link_name, msg))
                        errors += 1

                t.Commit()

                if added > 0:
                    success_models += 1

                out.print_md(
                    "  **Итого:** добавлено {}, пропущено {}, ошибок {}".format(
                        added, skipped, errors
                    )
                )

            except Exception as e:
                t.RollBack()
                out.print_md(":x: Ошибка транзакции: `{}`".format(e))

            # Синхронизация и закрытие
            closebg.close_with_policy(doc, do_sync=True, comment="Добавление связей")

        except Exception as e:
            out.print_md(":x: Ошибка открытия: `{}`".format(e))

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

        try:
            result = openbg.open_in_background(
                __revit__.Application,
                __revit__,
                mp,
                audit=False,
                worksets="all",
                detach=False,
                suppress_warnings=True,
            )
            if len(result) >= 1:
                doc = result[0]
            else:
                raise Exception("openbg не вернул документ")

            scripts_succeeded = 0
            scripts_failed = 0

            for script_rel_path in scripts:
                script_path = os.path.join(PYTHON_SCRIPTS_DIR, script_rel_path)
                script_name = os.path.basename(script_rel_path)
                module_name = os.path.splitext(script_name)[0]

                try:
                    # Добавляем директорию скрипта в sys.path для импорта
                    script_dir = os.path.dirname(script_path)
                    if script_dir not in sys.path:
                        sys.path.insert(0, script_dir)

                    # Пытаемся импортировать модуль по имени файла
                    module = None
                    try:
                        module = importlib.import_module(module_name)
                    except ImportError as ie:
                        out.print_md(
                            "  :x: {} - Ошибка импорта: {}".format(script_name, str(ie))
                        )
                        scripts_failed += 1
                        continue
                    except Exception as e:
                        out.print_md(
                            "  :x: {} - Ошибка загрузки: {}".format(script_name, str(e))
                        )
                        scripts_failed += 1
                        continue

                    if not module:
                        out.print_md(
                            "  :x: {} - Модуль не загружен".format(script_name)
                        )
                        scripts_failed += 1
                        continue

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
                                total = fill.get("total_elements", 0)
                                updated = fill.get("updated_elements", 0)
                                skipped = fill.get("skipped_elements", 0)
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
                            scripts_failed += 1
                    else:
                        out.print_md("  :white_check_mark: {}".format(script_name))
                        scripts_succeeded += 1

                except Exception as e:
                    out.print_md("  :x: {} - {}".format(script_name, e))
                    scripts_failed += 1
            closebg.close_with_policy(
                doc, do_sync=True, comment="Выполнение python скриптов"
            )

        except Exception as e:
            out.print_md(":x: Ошибка открытия: `{}`".format(e))

        out.update_progress(i + 1, len(models))

    out.print_md("---")
    out.print_md(
        "**Готово. Моделей обработано: {}/{}**".format(success_models, len(models))
    )


# ---------- Main ----------


def main():
    # 1. Выбор действия
    actions = [
        "Открыть модели",
        "Загрузить семейство (из открытых)",
        "Добавить связь",
        "Выполнение python скриптов из библиотеки",
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
    if selected_action == "Выполнение python скриптов из библиотеки":
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

    # 2. Выбор объекта (txt файла)
    txt_files = list_txt_files(OBJECTS_DIR)

    if not txt_files:
        forms.alert(
            "Папка с объектами пуста или недоступна:\n{}\n\nSCRIPTS_ROOT: {}".format(
                OBJECTS_DIR, SCRIPTS_ROOT
            ),
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
