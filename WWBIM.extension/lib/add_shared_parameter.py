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
    SharedParameterElement,
    ParameterElement,
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


def FindSharedParameterElementByGuid(doc, target_guid, diagnostics=None):
    """Находит SharedParameterElement по GUID через GuidValue"""
    if not target_guid:
        return None
    try:
        collector = FilteredElementCollector(doc)
        for elem in collector.OfClass(SharedParameterElement).ToElements():
            if elem.GuidValue == target_guid:
                return elem
    except Exception as e:
        _add_exception(diagnostics, "FindSharedParameterElementByGuid", e)
    return None


def IsDefinitionBoundByName(doc, param_name, diagnostics=None):
    """Проверяет, привязан ли параметр по имени в ParameterBindings"""
    try:
        bindings = doc.ParameterBindings
        iterator = bindings.ForwardIterator()
        iterator.Reset()
        while iterator.MoveNext():
            definition = iterator.Key
            if definition.Name == param_name:
                return True
        return False
    except Exception as e:
        _add_exception(diagnostics, "IsDefinitionBoundByName", e)
        return False


def FindSharedParameterElementsByName(doc, param_name, diagnostics=None):
    """Находит все SharedParameterElement по имени"""
    if not param_name:
        return []
    try:
        collector = FilteredElementCollector(doc)
        same_name_spes = []
        for elem in collector.OfClass(SharedParameterElement).ToElements():
            if elem.Name == param_name:
                same_name_spes.append(elem)
        return same_name_spes
    except Exception as e:
        _add_exception(diagnostics, "FindSharedParameterElementsByName", e)
        return []


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
    Проверяет, привязан ли уже общий параметр к документу.
    Проверка выполняется через SharedParameterElement.GuidValue + ParameterBindings по имени.
    """
    try:
        target_guid = ext_def.GUID if hasattr(ext_def, "GUID") else None
        param_name = ext_def.Name if hasattr(ext_def, "Name") else None

        if not target_guid or not param_name:
            return False

        # Проверяем наличие SharedParameterElement по GUID
        spe = FindSharedParameterElementByGuid(doc, target_guid, diagnostics)
        if not spe:
            return False

        # Проверяем, что параметр привязан по имени в ParameterBindings
        return IsDefinitionBoundByName(doc, param_name, diagnostics)
    except Exception as e:
        _add_exception(diagnostics, "IsParameterAlreadyBound", e)
        return False


def CheckForDuplicateParameters(doc, param_name, ext_def_guid=None, diagnostics=None):
    """
    Проверяет наличие конфликтующих параметров.

    ВАЖНО: Семейные параметры с таким же именем НЕ являются конфликтом!
    Конфликтом является только параметр с тем же именем, но ДРУГИМ GUID.
    Если имя совпадает и GUID совпадает (есть SharedParameterElement по этому GUID) - это НЕ конфликт.

    Возвращает список конфликтов (пустой если конфликтов нет).
    """
    conflicts = []

    if ext_def_guid is None:
        return conflicts  # Нет GUID для проверки - считаем что конфликтов нет

    try:
        # Сначала проверяем, есть ли SharedParameterElement по целевому GUID
        spe_by_guid = FindSharedParameterElementByGuid(doc, ext_def_guid, diagnostics)

        # Проверяем только привязанные параметры проекта
        bindings = doc.ParameterBindings
        iterator = bindings.ForwardIterator()
        iterator.Reset()

        while iterator.MoveNext():
            definition = iterator.Key
            # Проверяем только параметры с таким же именем
            if definition.Name != param_name:
                continue

            # Если SharedParameterElement по целевому GUID найден - это НЕ конфликт
            if spe_by_guid:
                continue

            # SharedParameterElement по целевому GUID не найден, но binding по имени есть - конфликт
            conflicts.append(
                "ParameterBindings содержит параметр с таким именем, но SharedParameterElement не найден"
            )
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
        "allow_vary_between_groups": None,
    }

    try:
        ext_def_guid = ext_def.GUID if hasattr(ext_def, "GUID") else None
        param_name = ext_def.Name if hasattr(ext_def, "Name") else None

        if not ext_def_guid or not param_name:
            return False, "Параметр не имеет GUID или имени", result_info

        insert_result = bindings.Insert(ext_def, binding, param_group)
        result_info["insert_result"] = insert_result

        spe = FindSharedParameterElementByGuid(doc, ext_def_guid, diagnostics)
        if spe:
            is_bound = IsDefinitionBoundByName(doc, param_name, diagnostics)
            result_info["bound_after_operation"] = is_bound
            result_info["bound_definition_name"] = param_name
            result_info["bound_definition_guid"] = str(ext_def_guid)

            if is_bound:
                try:
                    ParameterElement.SetAllowVaryBetweenGroups(doc, spe.Id, True)
                    result_info["allow_vary_between_groups"] = True
                except Exception as e:
                    _add_exception(diagnostics, "SetAllowVaryBetweenGroups", e)
                    result_info["allow_vary_between_groups"] = False

                return True, "Параметр успешно добавлен", result_info

        reinsert_result = bindings.ReInsert(ext_def, binding, param_group)
        result_info["reinsert_result"] = reinsert_result

        spe = FindSharedParameterElementByGuid(doc, ext_def_guid, diagnostics)
        if spe:
            is_bound = IsDefinitionBoundByName(doc, param_name, diagnostics)
            result_info["bound_after_operation"] = is_bound
            result_info["bound_definition_name"] = param_name
            result_info["bound_definition_guid"] = str(ext_def_guid)

            if is_bound:
                try:
                    ParameterElement.SetAllowVaryBetweenGroups(doc, spe.Id, True)
                    result_info["allow_vary_between_groups"] = True
                except Exception as e:
                    _add_exception(diagnostics, "SetAllowVaryBetweenGroups", e)
                    result_info["allow_vary_between_groups"] = False

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
        "allow_vary_between_groups": None,
        "mode": None,
        "guid_match": None,
        "expected_guid": None,
        "found_shared_guids": None,
        "use_param_name": None,
    }

    result = {
        "success": True,
        "parameters": {"added": [], "existing": [], "failed": []},
        "message": "",
        "diagnostics": diagnostics,
        "mode": None,
        "guid_match": None,
        "expected_guid": None,
        "found_shared_guids": None,
        "use_param_name": None,
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
    if not ext_def:
        result["success"] = False
        result["message"] = "Параметр '{0}' не найден в ФОП".format(
            config["PARAMETER_NAME"]
        )
        result["parameters"]["failed"].append(config["PARAMETER_NAME"])
        return result

    diagnostics["ext_def.Name"] = ext_def.Name
    diagnostics["ext_def.GUID"] = (
        str(ext_def.GUID) if hasattr(ext_def, "GUID") else None
    )

    ext_def_guid = ext_def.GUID if hasattr(ext_def, "GUID") else None
    if not ext_def_guid:
        result["success"] = False
        result["message"] = "Параметр '{0}' найден в ФОП, но не имеет GUID".format(
            config["PARAMETER_NAME"]
        )
        result["parameters"]["failed"].append(config["PARAMETER_NAME"])
        return result

    diagnostics["expected_guid"] = str(ext_def_guid)
    result["expected_guid"] = str(ext_def_guid)

    # B) Проверка "мой shared уже в проекте"
    if IsParameterAlreadyBound(doc, ext_def, diagnostics):
        spe = FindSharedParameterElementByGuid(doc, ext_def_guid, diagnostics)
        if spe:
            diagnostics["bound_after_operation"] = True
            diagnostics["bound_definition_name"] = config["PARAMETER_NAME"]
            diagnostics["bound_definition_guid"] = str(ext_def_guid)
            diagnostics["mode"] = "existing_guid_match"
            diagnostics["guid_match"] = True
            diagnostics["found_shared_guids"] = [str(ext_def_guid)]
            diagnostics["use_param_name"] = config["PARAMETER_NAME"]
            result["mode"] = "existing_guid_match"
            result["guid_match"] = True
            result["found_shared_guids"] = [str(ext_def_guid)]
            result["use_param_name"] = config["PARAMETER_NAME"]
            result["success"] = True
            result["message"] = "Параметр '{0}' уже привязан".format(
                config["PARAMETER_NAME"]
            )
            result["parameters"]["existing"].append(config["PARAMETER_NAME"])
            return result

    # C) НОВОЕ: если моего GUID нет — проверить наличие SharedParameterElement по ИМЕНИ
    same_name_spes = FindSharedParameterElementsByName(
        doc, config["PARAMETER_NAME"], diagnostics
    )
    if same_name_spes:
        found_guids = [str(spe.GuidValue) for spe in same_name_spes]
        bound_by_name = IsDefinitionBoundByName(
            doc, config["PARAMETER_NAME"], diagnostics
        )

        if not bound_by_name:
            diagnostics["mode"] = "failed"
            diagnostics["bound_after_operation"] = False
            diagnostics["found_shared_guids"] = found_guids
            result["mode"] = "failed"
            result["success"] = False
            result["message"] = (
                "SharedParameterElement с именем '{0}' найден, но parameter binding отсутствует".format(
                    config["PARAMETER_NAME"]
                )
            )
            result["parameters"]["failed"].append(config["PARAMETER_NAME"])
            return result

        diagnostics["mode"] = "use_existing_shared_by_name"
        diagnostics["guid_match"] = False
        diagnostics["found_shared_guids"] = found_guids
        diagnostics["bound_after_operation"] = True
        diagnostics["bound_definition_name"] = config["PARAMETER_NAME"]
        diagnostics["use_param_name"] = config["PARAMETER_NAME"]
        result["mode"] = "use_existing_shared_by_name"
        result["guid_match"] = False
        result["found_shared_guids"] = found_guids
        result["use_param_name"] = config["PARAMETER_NAME"]
        result["bound_after_operation"] = True
        result["success"] = True
        result["message"] = (
            "В проекте уже есть Shared параметр с таким именем. Будет использован существующий; GUID отличается."
        )
        result["parameters"]["existing"].append(config["PARAMETER_NAME"])
        return result

    # Проверка: если параметр с именем есть, но он не shared
    if IsDefinitionBoundByName(doc, config["PARAMETER_NAME"], diagnostics):
        diagnostics["mode"] = "failed"
        result["mode"] = "failed"
        result["success"] = False
        result["message"] = (
            "ParameterBindings содержит параметр с таким именем, но SharedParameterElement не найден"
        )
        result["parameters"]["failed"].append(config["PARAMETER_NAME"])
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
    exception_trace = bind_info.get("exception_trace")

    if bound_after_operation:
        result["success"] = True
        result["mode"] = "added"
        result["guid_match"] = True
        result["use_param_name"] = config["PARAMETER_NAME"]
        diagnostics["mode"] = "added"
        diagnostics["guid_match"] = True
        diagnostics["use_param_name"] = config["PARAMETER_NAME"]
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
    result["mode"] = "failed"
    diagnostics["mode"] = "failed"
    diagnostics["use_param_name"] = config["PARAMETER_NAME"]
    result["use_param_name"] = config["PARAMETER_NAME"]
    result["parameters"]["failed"].append(config["PARAMETER_NAME"])

    error_parts = [bind_message]
    if exception_trace:
        error_parts.append("Exception: {0}".format(exception_trace))
    result["message"] = ". ".join(error_parts)
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
