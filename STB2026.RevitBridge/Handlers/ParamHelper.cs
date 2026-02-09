using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;

namespace STB2026.RevitBridge.Handlers
{
    /// <summary>
    /// Утилиты для работы с параметрами Revit.
    /// Используются всеми обработчиками.
    /// 
    /// Совместимость: Revit 2024-2025 API.
    /// Избегаем deprecated: ParameterType, ParameterGroup.
    /// </summary>
    internal static class ParamHelper
    {
        /// <summary>Извлечь краткую сводку элемента (для списков).</summary>
        public static Dictionary<string, object> ElementSummary(Element el)
        {
            var result = new Dictionary<string, object>
            {
                ["id"] = el.Id.IntegerValue,
                ["name"] = el.Name ?? "",
                ["category"] = el.Category?.Name ?? "—"
            };

            // Тип
            var typeId = el.GetTypeId();
            if (typeId != null && typeId != ElementId.InvalidElementId)
            {
                var type = el.Document.GetElement(typeId);
                if (type != null)
                    result["type_name"] = type.Name;
            }

            // Семейство
            if (el is FamilyInstance fi)
                result["family"] = fi.Symbol?.Family?.Name ?? "";

            // Уровень
            var levelParam = el.get_Parameter(BuiltInParameter.INSTANCE_REFERENCE_LEVEL_PARAM)
                          ?? el.get_Parameter(BuiltInParameter.FAMILY_LEVEL_PARAM)
                          ?? el.get_Parameter(BuiltInParameter.SCHEDULE_LEVEL_PARAM);
            if (levelParam != null && levelParam.AsElementId() != ElementId.InvalidElementId)
            {
                var level = el.Document.GetElement(levelParam.AsElementId());
                if (level != null)
                    result["level"] = level.Name;
            }

            // Для MEP — система
            if (el is Duct duct)
            {
                var sys = duct.MEPSystem;
                if (sys != null)
                    result["system"] = sys.Name;
            }

            return result;
        }

        /// <summary>Извлечь ВСЕ параметры элемента с деталями.</summary>
        public static List<Dictionary<string, object>> AllParameters(Element el)
        {
            var list = new List<Dictionary<string, object>>();

            foreach (Parameter p in el.Parameters)
            {
                if (p?.Definition == null || !p.HasValue) continue;

                var info = new Dictionary<string, object>
                {
                    ["name"] = p.Definition.Name ?? "?",
                    ["id"] = p.Id.IntegerValue,
                    ["value"] = GetParamDisplayValue(p),
                    ["storage"] = p.StorageType.ToString(),
                    ["readonly"] = p.IsReadOnly
                };

                // Группа параметра (Revit 2024+ API)
                try
                {
                    var groupId = p.Definition.GetGroupTypeId();
                    if (groupId != null)
                        info["group"] = LabelUtils.GetLabelForGroup(groupId);
                }
                catch
                {
                    info["group"] = "Прочее";
                }

                // Единицы (для числовых)
                if (p.StorageType == StorageType.Double)
                {
                    info["raw_value"] = p.AsDouble();
                    try
                    {
                        var unitTypeId = p.GetUnitTypeId();
                        if (unitTypeId != null)
                            info["units"] = unitTypeId.TypeId ?? "";
                    }
                    catch { /* Нет единиц */ }
                }

                list.Add(info);
            }

            return list.OrderBy(x => x.ContainsKey("group") ? x["group"].ToString() : "")
                       .ThenBy(x => x["name"].ToString())
                       .ToList();
        }

        /// <summary>Значение параметра как строка.</summary>
        public static string GetParamDisplayValue(Parameter p)
        {
            if (p == null || !p.HasValue) return "";

            switch (p.StorageType)
            {
                case StorageType.String:
                    return p.AsString() ?? "";

                case StorageType.Integer:
                    // Проверяем Да/Нет через ForgeTypeId (Revit 2024+ способ)
                    try
                    {
                        var specId = p.Definition.GetDataType();
                        if (specId == SpecTypeId.Boolean.YesNo)
                            return p.AsInteger() == 1 ? "Да" : "Нет";
                    }
                    catch { /* Старый API fallback */ }
                    return p.AsInteger().ToString();

                case StorageType.Double:
                    return p.AsValueString() ?? p.AsDouble().ToString("F4");

                case StorageType.ElementId:
                    var elemId = p.AsElementId();
                    if (elemId == ElementId.InvalidElementId) return "—";
                    var elem = p.Element?.Document?.GetElement(elemId);
                    return elem?.Name ?? $"id:{elemId.IntegerValue}";

                default:
                    return "";
            }
        }

        /// <summary>Парсинг строки ID элементов.</summary>
        public static List<ElementId> ParseIds(string idsString)
        {
            var result = new List<ElementId>();
            if (string.IsNullOrWhiteSpace(idsString)) return result;

            foreach (var part in idsString.Split(','))
            {
                if (int.TryParse(part.Trim(), out int id))
                    result.Add(new ElementId(id));
            }
            return result;
        }

        /// <summary>Найти параметр по имени (экземпляра или типа).</summary>
        public static Parameter FindByName(Element el, string paramName)
        {
            var param = el.LookupParameter(paramName);
            if (param != null) return param;

            var typeId = el.GetTypeId();
            if (typeId != null && typeId != ElementId.InvalidElementId)
            {
                var type = el.Document.GetElement(typeId);
                param = type?.LookupParameter(paramName);
            }
            return param;
        }
    }
}
