using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace STB2026.RevitBridge.Handlers
{
    /// <summary>
    /// get_element_params — все параметры конкретных элементов.
    /// Один ID → полная развёртка всех параметров.
    /// Несколько ID + param_name → значение одного параметра для каждого.
    /// </summary>
    internal static class ElementParamsHandler
    {
        public static object Handle(UIApplication uiApp, Dictionary<string, object> p)
        {
            var doc = uiApp.ActiveUIDocument?.Document;
            if (doc == null)
                return new { error = "Нет открытого документа" };

            string idsStr = p.TryGetValue("element_ids", out var ids) ? ids?.ToString() ?? "" : "";
            string paramName = p.TryGetValue("param_name", out var pn) ? pn?.ToString() ?? "" : "";

            var elementIds = ParamHelper.ParseIds(idsStr);
            if (elementIds.Count == 0)
                return new { error = "Не указаны ID элементов. Передайте element_ids через запятую." };

            // Один элемент без фильтра → полная развёртка
            if (elementIds.Count == 1 && string.IsNullOrWhiteSpace(paramName))
            {
                var el = doc.GetElement(elementIds[0]);
                if (el == null)
                    return new { error = $"Элемент {elementIds[0].IntegerValue} не найден" };

                return SingleElementFull(el);
            }

            // Несколько элементов + конкретный параметр
            if (!string.IsNullOrWhiteSpace(paramName))
            {
                return MultipleElementsOneParam(doc, elementIds, paramName);
            }

            // Несколько элементов без фильтра → краткая сводка для каждого
            return MultipleElementsSummary(doc, elementIds);
        }

        private static object SingleElementFull(Element el)
        {
            var summary = ParamHelper.ElementSummary(el);
            var parameters = ParamHelper.AllParameters(el);

            // Также параметры типа
            List<Dictionary<string, object>> typeParams = null;
            var typeId = el.GetTypeId();
            if (typeId != null && typeId != ElementId.InvalidElementId)
            {
                var type = el.Document.GetElement(typeId);
                if (type != null)
                {
                    typeParams = ParamHelper.AllParameters(type);
                    summary["type_id"] = typeId.IntegerValue;
                    summary["type_name"] = type.Name;
                }
            }

            // BoundingBox
            var bb = el.get_BoundingBox(null);
            Dictionary<string, object> bbox = null;
            if (bb != null)
            {
                bbox = new Dictionary<string, object>
                {
                    ["min_x_mm"] = FeetToMm(bb.Min.X),
                    ["min_y_mm"] = FeetToMm(bb.Min.Y),
                    ["min_z_mm"] = FeetToMm(bb.Min.Z),
                    ["max_x_mm"] = FeetToMm(bb.Max.X),
                    ["max_y_mm"] = FeetToMm(bb.Max.Y),
                    ["max_z_mm"] = FeetToMm(bb.Max.Z)
                };
            }

            // Location
            Dictionary<string, object> location = null;
            if (el.Location is LocationPoint lp)
            {
                location = new Dictionary<string, object>
                {
                    ["type"] = "point",
                    ["x_mm"] = FeetToMm(lp.Point.X),
                    ["y_mm"] = FeetToMm(lp.Point.Y),
                    ["z_mm"] = FeetToMm(lp.Point.Z)
                };
            }
            else if (el.Location is LocationCurve lc)
            {
                var start = lc.Curve.GetEndPoint(0);
                var end = lc.Curve.GetEndPoint(1);
                location = new Dictionary<string, object>
                {
                    ["type"] = "curve",
                    ["start_x_mm"] = FeetToMm(start.X),
                    ["start_y_mm"] = FeetToMm(start.Y),
                    ["start_z_mm"] = FeetToMm(start.Z),
                    ["end_x_mm"] = FeetToMm(end.X),
                    ["end_y_mm"] = FeetToMm(end.Y),
                    ["end_z_mm"] = FeetToMm(end.Z),
                    ["length_mm"] = FeetToMm(lc.Curve.Length)
                };
            }

            return new
            {
                element = summary,
                instance_params = parameters,
                type_params = typeParams,
                bounding_box = bbox,
                location
            };
        }

        private static object MultipleElementsOneParam(Document doc, List<ElementId> ids, string paramName)
        {
            var results = new List<Dictionary<string, object>>();

            foreach (var id in ids)
            {
                var el = doc.GetElement(id);
                if (el == null) continue;

                var param = ParamHelper.FindByName(el, paramName);
                results.Add(new Dictionary<string, object>
                {
                    ["id"] = id.IntegerValue,
                    ["name"] = el.Name,
                    ["param_value"] = param != null ? ParamHelper.GetParamDisplayValue(param) : "—",
                    ["param_found"] = param != null
                });
            }

            return new
            {
                param_name = paramName,
                count = results.Count,
                elements = results
            };
        }

        private static object MultipleElementsSummary(Document doc, List<ElementId> ids)
        {
            var results = ids
                .Select(id => doc.GetElement(id))
                .Where(el => el != null)
                .Select(ParamHelper.ElementSummary)
                .ToList();

            return new
            {
                count = results.Count,
                elements = results
            };
        }

        private static double FeetToMm(double feet)
            => System.Math.Round(feet * 304.8, 1);
    }
}
