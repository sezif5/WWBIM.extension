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
)


CONFIG = {
    "PARAMETER_NAME": "ADSK_КомплектШифр",
    "BINDING_TYPE": "Instance",
    "PARAMETER_GROUP": BuiltInParameterGroup.INVALID,
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


def FindExternalDefinition(def_file, param_name):
    try:
        for group in def_file.Groups:
            for definition in group.Definitions:
                if definition.Name == param_name:
                    return definition, group.Name
    except Exception as e:
        pass

    return None, None


def IsParameterAlreadyBound(doc, ext_def):
    bindings = doc.ParameterBindings
    try:
        return bindings.Contains(ext_def)
    except:
        return False


def CreateCategorySet(doc, categories):
    app = doc.Application
    cat_set = app.Create.NewCategorySet()

    for category in categories:
        try:
            cat = doc.Settings.Categories.get_Item(category)
            if cat:
                cat_set.Insert(cat)
        except:
            pass

    return cat_set


def CreateBinding(binding_type, cat_set):
    if binding_type == "Instance":
        return InstanceBinding(cat_set)
    elif binding_type == "Type":
        return TypeBinding(cat_set)
    else:
        return InstanceBinding(cat_set)


def BindParameter(doc, ext_def, binding, param_group):
    bindings = doc.ParameterBindings

    try:
        if bindings.Insert(ext_def, binding, param_group):
            return True, "Параметр успешно добавлен"
        else:
            return False, "Не удалось добавить параметр"
    except Exception as e:
        return False, "Ошибка привязки: {0}".format(str(e))


def AddSharedParameterToDoc(doc, config=None):
    if config is None:
        config = CONFIG

    result = {
        "success": True,
        "parameters": {"added": [], "existing": [], "failed": []},
        "message": "",
    }

    def_file, filepath = GetSharedParameterFile(doc)
    if not def_file:
        result["success"] = False
        result["message"] = filepath
        return result

    ext_def, group_name = FindExternalDefinition(def_file, config["PARAMETER_NAME"])
    if not ext_def:
        result["success"] = False
        result["message"] = "Параметр '{0}' не найден в ФОП".format(
            config["PARAMETER_NAME"]
        )
        result["parameters"]["failed"].append(config["PARAMETER_NAME"])
        return result

    if IsParameterAlreadyBound(doc, ext_def):
        result["success"] = True
        result["message"] = "Параметр '{0}' уже привязан".format(
            config["PARAMETER_NAME"]
        )
        result["parameters"]["existing"].append(config["PARAMETER_NAME"])
        return result

    cat_set = CreateCategorySet(doc, config["CATEGORIES"])
    if cat_set.Size == 0:
        result["success"] = False
        result["message"] = "Не удалось добавить ни одной категории"
        result["parameters"]["failed"].append(config["PARAMETER_NAME"])
        return result

    binding = CreateBinding(config["BINDING_TYPE"], cat_set)

    success, message = BindParameter(doc, ext_def, binding, config["PARAMETER_GROUP"])

    if success:
        result["parameters"]["added"].append(config["PARAMETER_NAME"])
    else:
        result["parameters"]["failed"].append(config["PARAMETER_NAME"])
        result["success"] = False

    result["message"] = message
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
