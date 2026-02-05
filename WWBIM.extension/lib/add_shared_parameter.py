# -*- coding: utf-8 -*-
from Autodesk.Revit.DB import (
    BuiltInParameter,
    BuiltInCategory,
    CategorySet,
    InstanceBinding,
    TypeBinding,
    BuiltInParameterGroup,
    ExternalDefinition,
    Transaction,
    FilteredElementCollector,
)
import traceback


CONFIG = {
    "PARAMETER_NAME": "ADSK_КомплектШифр",
    "BINDING_TYPE": "Instance",
    "PARAMETER_GROUP": BuiltInParameterGroup.PG_IDENTITY_DATA,
    "CATEGORIES": [
        BuiltInCategory.OST_Walls,
        BuiltInCategory.OST_Doors,
        BuiltInCategory.OST_Windows,
        BuiltInCategory.OST_Floors,
        BuiltInCategory.OST_Roofs,
        BuiltInCategory.OST_StructuralColumns,
        BuiltInCategory.OST_StructuralFraming,
        BuiltInCategory.OST_MechanicalEquipment,
        BuiltInCategory.OST_PlumbingFixtures,
    ],
}


def _add_exception(diagnostics, context, exc):
    if diagnostics is None:
        return
    try:
        trace = traceback.format_exc()
    except Exception:
        trace = str(exc)
    diagnostics.setdefault("exception_trace", []).append(
        {"context": context, "error": str(exc), "trace": trace}
    )


def FindBoundDefinitionByGuid(doc, target_guid, diagnostics=None):
    if not target_guid:
        return None
    try:
        bindings = doc.ParameterBindings
        iterator = bindings.ForwardIterator()
        iterator.Reset()
        while iterator.MoveNext():
            definition = iterator.Key
            if hasattr(definition, "GUID") and definition.GUID == target_guid:
                return definition
    except Exception as e:
        _add_exception(diagnostics, "FindBoundDefinitionByGuid", e)
    return None


def GetSharedParameterFile(doc):
    app = doc.Application
    filepath = app.SharedParametersFilename

    if not filepath:
        return None, "Файл общих параметров не настроен"

    try:
        def_file = app.OpenSharedParameterFile()
        if not def_file:
            return None, "Не удалось открыть файл общих параметров"

        return def_file, filepath
    except Exception as e:
        return None, "Ошибка открытия файла: {0}".format(str(e))


def FindExternalDefinition(def_file, param_name, diagnostics=None):
    try:
        for group in def_file.Groups:
            for definition in group.Definitions:
                if definition.Name == param_name:
                    return definition, group.Name
    except Exception as e:
        _add_exception(diagnostics, "FindExternalDefinition", e)

    return None, None


def IsParameterAlreadyBound(doc, ext_def, diagnostics=None):
    """
    Проверяет, привязан ли уже общий параметр к документу по GUID.
    """
    try:
        target_guid = ext_def.GUID if hasattr(ext_def, "GUID") else None
        if not target_guid:
            return False

        bindings = doc.ParameterBindings
        iterator = bindings.ForwardIterator()
        iterator.Reset()
        while iterator.MoveNext():
            definition = iterator.Key
            if hasattr(definition, "GUID") and definition.GUID == target_guid:
                return True
        return False
    except Exception as e:
        _add_exception(diagnostics, "IsParameterAlreadyBound", e)
        return False


def CheckForDuplicateParameters(doc, param_name, ext_def_guid=None, diagnostics=None):
    """
    Проверяет наличие конфликтующих параметров.

    ВАЖНО: Семейные параметры с таким же именем НЕ являются конфликтом!
    Конфликтом является только общий параметр с ДРУГИМ GUID, но таким же именем.
    Если имя совпадает и GUID совпадает - это НЕ конфликт (параметр уже существует).

    Возвращает список конфликтов (пустой если конфликтов нет).
    """
    conflicts = []

    if ext_def_guid is None:
        return conflicts  # Нет GUID для проверки - считаем что конфликтов нет

    try:
        # Проверяем только привязанные параметры проекта
        bindings = doc.ParameterBindings
        iterator = bindings.ForwardIterator()
        iterator.Reset()

        while iterator.MoveNext():
            definition = iterator.Key
            # Проверяем только параметры с таким же именем
            if definition.Name != param_name:
                continue

            # Если это общий параметр - проверяем GUID
            if hasattr(definition, "GUID"):
                # Если GUID совпадает - это НЕ конфликт, параметр уже существует
                if definition.GUID == ext_def_guid:
                    continue
                else:
                    conflicts.append("Shared with different GUID")
            else:
                # Это не общий параметр (например, параметр проекта) с таким же именем
                conflicts.append("Project parameter with same name")
    except Exception as e:
        _add_exception(diagnostics, "CheckForDuplicateParameters", e)

    return conflicts


def CreateCategorySet(doc, categories, diagnostics=None):
    app = doc.Application
    cat_set = app.Create.NewCategorySet()
    skipped_categories = []

    for category in categories:
        try:
            cat = doc.Settings.Categories.get_Item(category)
            if cat:
                # Проверяем разрешена ли привязка параметров к категории
                if cat.AllowsBoundParameters:
                    cat_set.Insert(cat)
                else:
                    skipped_categories.append(cat.Name)
        except Exception as e:
            _add_exception(diagnostics, "CreateCategorySet: {0}".format(category), e)

    return cat_set, skipped_categories


def CreateBinding(app, binding_type, cat_set):
    if binding_type == "Instance":
        return app.Create.NewInstanceBinding(cat_set)
    elif binding_type == "Type":
        return app.Create.NewTypeBinding(cat_set)
    else:
        return app.Create.NewInstanceBinding(cat_set)


def BindParameter(doc, ext_def, binding, param_group, diagnostics=None):
    bindings = doc.ParameterBindings
    result_info = {
        "insert_result": None,
        "reinsert_result": None,
        "exception_trace": None,
        "bound_after_operation": False,
        "bound_definition_name": None,
        "bound_definition_guid": None,
    }

    try:
        insert_result = bindings.Insert(ext_def, binding, param_group)
        result_info["insert_result"] = insert_result

        ext_def_guid = ext_def.GUID if hasattr(ext_def, "GUID") else None
        bound_def = FindBoundDefinitionByGuid(doc, ext_def_guid, diagnostics)
        result_info["bound_after_operation"] = bound_def is not None
        if bound_def:
            result_info["bound_definition_name"] = bound_def.Name
            result_info["bound_definition_guid"] = str(bound_def.GUID)

        if bound_def:
            return True, "Параметр успешно добавлен", result_info

        reinsert_result = bindings.ReInsert(ext_def, binding, param_group)
        result_info["reinsert_result"] = reinsert_result

        bound_def = FindBoundDefinitionByGuid(doc, ext_def_guid, diagnostics)
        result_info["bound_after_operation"] = bound_def is not None
        if bound_def:
            result_info["bound_definition_name"] = bound_def.Name
            result_info["bound_definition_guid"] = str(bound_def.GUID)

        if bound_def:
            return True, "Параметр успешно обновлен (ReInsert)", result_info

        return False, "Не удалось добавить параметр", result_info
    except Exception as e:
        _add_exception(diagnostics, "BindParameter", e)
        result_info["exception_trace"] = str(e)
        error_msg = str(e)
        if "not allowed" in error_msg.lower():
            return False, "Категория не разрешает привязку параметров", result_info
        else:
            return False, "Ошибка привязки: {0}".format(error_msg), result_info


def AddSharedParameterToDoc(doc, config=None):
    if config is None:
        config = CONFIG

    diagnostics = {
        "SharedParametersFilename": None,
        "group_name": None,
        "ext_def.Name": None,
        "ext_def.GUID": None,
        "cat_set.Size": None,
        "skipped_categories": [],
        "insert_result": None,
        "reinsert_result": None,
        "exception_trace": [],
        "bound_after_operation": False,
        "bound_definition_name": None,
        "bound_definition_guid": None,
    }

    result = {
        "success": True,
        "parameters": {"added": [], "existing": [], "failed": []},
        "message": "",
        "diagnostics": diagnostics,
    }

    def_file, filepath = GetSharedParameterFile(doc)
    diagnostics["SharedParametersFilename"] = filepath
    if not def_file:
        result["success"] = False
        result["message"] = filepath
        return result

    ext_def, group_name = FindExternalDefinition(
        def_file, config["PARAMETER_NAME"], diagnostics
    )
    diagnostics["group_name"] = group_name
    diagnostics["ext_def.Name"] = getattr(ext_def, "Name", None)
    diagnostics["ext_def.GUID"] = (
        str(ext_def.GUID) if hasattr(ext_def, "GUID") else None
    )
    if not ext_def:
        result["success"] = False
        result["message"] = "Параметр '{0}' не найден в ФОП".format(
            config["PARAMETER_NAME"]
        )
        result["parameters"]["failed"].append(config["PARAMETER_NAME"])
        return result

    ext_def_guid = ext_def.GUID if hasattr(ext_def, "GUID") else None
    diagnostics["ext_def.GUID"] = str(ext_def_guid) if ext_def_guid else None
    if not ext_def_guid:
        result["success"] = False
        result["message"] = "Параметр '{0}' найден в ФОП, но не имеет GUID".format(
            config["PARAMETER_NAME"]
        )
        result["parameters"]["failed"].append(config["PARAMETER_NAME"])
        return result

    conflicts = CheckForDuplicateParameters(
        doc, config["PARAMETER_NAME"], ext_def_guid, diagnostics
    )
    if conflicts:
        result["success"] = False
        result["message"] = (
            "Конфликт параметров: '{0}'. "
            "Причины: {1}. Удалите конфликтующие параметры и повторите операцию.".format(
                config["PARAMETER_NAME"], ", ".join(conflicts)
            )
        )
        result["parameters"]["failed"].append(config["PARAMETER_NAME"])
        return result

    bound_def = FindBoundDefinitionByGuid(doc, ext_def_guid, diagnostics)
    if bound_def:
        diagnostics["bound_after_operation"] = True
        diagnostics["bound_definition_name"] = bound_def.Name
        diagnostics["bound_definition_guid"] = str(bound_def.GUID)
        result["success"] = True
        result["message"] = "Параметр '{0}' уже привязан".format(
            config["PARAMETER_NAME"]
        )
        result["parameters"]["existing"].append(config["PARAMETER_NAME"])
        return result

    cat_set, skipped_categories = CreateCategorySet(
        doc, config["CATEGORIES"], diagnostics
    )
    diagnostics["cat_set.Size"] = cat_set.Size if cat_set else None
    diagnostics["skipped_categories"] = skipped_categories
    if cat_set.Size == 0:
        if skipped_categories:
            result["success"] = False
            result["message"] = (
                "Не удалось добавить ни одной категории. Пропущено: {0}".format(
                    ", ".join(skipped_categories)
                )
            )
            result["parameters"]["failed"].append(config["PARAMETER_NAME"])
        else:
            result["success"] = False
            result["message"] = "Не удалось добавить ни одной категории"
            result["parameters"]["failed"].append(config["PARAMETER_NAME"])
        return result

    app = doc.Application
    binding = CreateBinding(app, config["BINDING_TYPE"], cat_set)

    bind_success, bind_message, bind_info = BindParameter(
        doc, ext_def, binding, config["PARAMETER_GROUP"], diagnostics
    )

    diagnostics.update(bind_info)

    all_skipped = list(skipped_categories)

    bound_after_operation = bind_info["bound_after_operation"]
    insert_result = bind_info["insert_result"]
    reinsert_result = bind_info["reinsert_result"]

    if bound_after_operation:
        result["success"] = True
        if insert_result:
            result["parameters"]["added"].append(config["PARAMETER_NAME"])
            message = "Параметр успешно добавлен"
        elif reinsert_result:
            result["parameters"]["added"].append(config["PARAMETER_NAME"])
            message = "Параметр успешно обновлен (ReInsert)"
        else:
            result["parameters"]["existing"].append(config["PARAMETER_NAME"])
            message = "Параметр найден после операции"

        if all_skipped:
            result["message"] = (
                "{0}. Пропущенные категории (не разрешают привязку): {1}".format(
                    message, ", ".join(all_skipped)
                )
            )
        else:
            result["message"] = message

        return result

    result["success"] = False
    result["parameters"]["failed"].append(config["PARAMETER_NAME"])
    result["message"] = (
        "Не удалось добавить параметр. Insert={0}, ReInsert={1}, Bound={2}".format(
            insert_result, reinsert_result, bound_after_operation
        )
    )
    return result


def Execute(doc, config=None):
    t = Transaction(doc, "Добавление общего параметра из ФОП")
    t.Start()

    try:
        result = AddSharedParameterToDoc(doc, config)

        if result["success"]:
            t.Commit()
        else:
            t.RollBack()

        return result

    except Exception as e:
        t.RollBack()
        return {
            "success": False,
            "message": "Ошибка: {0}".format(str(e)),
            "parameters": {"added": [], "existing": [], "failed": []},
        }


if __name__ == "__main__":
    Execute(__revit__.ActiveUIDocument.Document)
