using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Newtonsoft.Json.Linq;

namespace STB2026.RevitBridge.Handlers
{
    /// <summary>
    /// modify_model â€” Ð²ÑÐµ Ð¾Ð¿ÐµÑ€Ð°Ñ†Ð¸Ð¸ Ð·Ð°Ð¿Ð¸ÑÐ¸ Ð² Ð¼Ð¾Ð´ÐµÐ»ÑŒ Revit.
    /// action: set_param, set_color, reset_color, select, move, rotate, delete, isolate, create_tag.
    /// </summary>
    internal static class ModifyHandler
    {
        public static object Handle(UIApplication uiApp, Dictionary<string, object> p)
        {
            var uidoc = uiApp.ActiveUIDocument;
            if (uidoc == null)
                return new { error = "ÐÐµÑ‚ Ð¾Ñ‚ÐºÑ€Ñ‹Ñ‚Ð¾Ð³Ð¾ Ð´Ð¾ÐºÑƒÐ¼ÐµÐ½Ñ‚Ð°" };

            var doc = uidoc.Document;
            string action = p.TryGetValue("action", out var a) ? a?.ToString() ?? "" : "";
            string idsStr = p.TryGetValue("element_ids", out var ids) ? ids?.ToString() ?? "" : "";
            string dataStr = p.TryGetValue("data", out var d) ? d?.ToString() ?? "{}" : "{}";

            var elementIds = ParamHelper.ParseIds(idsStr);
            JObject data;
            try { data = JObject.Parse(dataStr); }
            catch { return new { error = $"ÐÐµÐ²Ð°Ð»Ð¸Ð´Ð½Ñ‹Ð¹ JSON Ð² data: {dataStr}" }; }

            switch (action.ToLowerInvariant())
            {
                case "set_param": return SetParam(doc, elementIds, data);
                case "set_color": return SetColor(doc, uidoc, elementIds, data);
                case "reset_color": return ResetColor(doc, uidoc, elementIds);
                case "select": return SelectElements(uidoc, elementIds);
                case "move": return MoveElements(doc, elementIds, data);
                case "rotate": return RotateElements(doc, elementIds, data);
                case "delete": return DeleteElements(doc, elementIds);
                case "isolate": return IsolateElements(uidoc, elementIds);
                case "create_tag": return CreateTag(doc, uidoc, elementIds, data);
                default:
                    return new
                    {
                        error = $"ÐÐµÐ¸Ð·Ð²ÐµÑÑ‚Ð½Ð¾Ðµ Ð´ÐµÐ¹ÑÑ‚Ð²Ð¸Ðµ: '{action}'",
                        available = new[] { "set_param", "set_color", "reset_color", "select",
                                           "move", "rotate", "delete", "isolate", "create_tag" }
                    };
            }
        }

        // â•â•â• set_param â•â•â•
        private static object SetParam(Document doc, List<ElementId> ids, JObject data)
        {
            string paramName = data.Value<string>("param_name") ?? "";
            string paramValue = data.Value<string>("param_value") ?? "";

            if (string.IsNullOrWhiteSpace(paramName))
                return new { error = "Ð£ÐºÐ°Ð¶Ð¸Ñ‚Ðµ param_name" };

            int success = 0, failed = 0;
            var errors = new List<string>();

            using (var trans = new Transaction(doc, $"STB2026: set {paramName}"))
            {
                trans.Start();

                foreach (var id in ids)
                {
                    var el = doc.GetElement(id);
                    if (el == null) { failed++; continue; }

                    var param = ParamHelper.FindByName(el, paramName);
                    if (param == null || param.IsReadOnly)
                    {
                        failed++;
                        errors.Add($"id {id.IntegerValue}: Ð¿Ð°Ñ€Ð°Ð¼ÐµÑ‚Ñ€ '{paramName}' Ð½Ðµ Ð½Ð°Ð¹Ð´ÐµÐ½ Ð¸Ð»Ð¸ read-only");
                        continue;
                    }

                    try
                    {
                        SetParamValue(param, paramValue);
                        success++;
                    }
                    catch (Exception ex)
                    {
                        failed++;
                        errors.Add($"id {id.IntegerValue}: {ex.Message}");
                    }
                }

                trans.Commit();
            }

            return new { action = "set_param", param_name = paramName, success, failed, errors };
        }

        // â•â•â• set_color â•â•â•
        private static object SetColor(Document doc, UIDocument uidoc, List<ElementId> ids, JObject data)
        {
            int r = data.Value<int?>("r") ?? 255;
            int g = data.Value<int?>("g") ?? 0;
            int b = data.Value<int?>("b") ?? 0;
            int viewIdInt = data.Value<int?>("view_id") ?? -1;

            var view = viewIdInt > 0
                ? doc.GetElement(new ElementId(viewIdInt)) as View
                : uidoc.ActiveView;

            if (view == null)
                return new { error = "Ð’Ð¸Ð´ Ð½Ðµ Ð½Ð°Ð¹Ð´ÐµÐ½" };

            var color = new Color((byte)r, (byte)g, (byte)b);
            var settings = new OverrideGraphicSettings();
            settings.SetProjectionLineColor(color);
            settings.SetSurfaceForegroundPatternColor(color);

            // Ð¡Ð¿Ð»Ð¾ÑˆÐ½Ð°Ñ Ð·Ð°Ð»Ð¸Ð²ÐºÐ°
            var solidPattern = new FilteredElementCollector(doc)
                .OfClass(typeof(FillPatternElement))
                .Cast<FillPatternElement>()
                .FirstOrDefault(fp => fp.GetFillPattern().IsSolidFill);

            if (solidPattern != null)
                settings.SetSurfaceForegroundPatternId(solidPattern.Id);

            int count = 0;
            using (var trans = new Transaction(doc, $"STB2026: color ({r},{g},{b})"))
            {
                trans.Start();
                foreach (var id in ids)
                {
                    view.SetElementOverrides(id, settings);
                    count++;
                }
                trans.Commit();
            }

            return new { action = "set_color", r, g, b, colored = count, view = view.Name };
        }

        // â•â•â• reset_color â•â•â•
        private static object ResetColor(Document doc, UIDocument uidoc, List<ElementId> ids)
        {
            var view = uidoc.ActiveView;
            var clean = new OverrideGraphicSettings();
            int count = 0;

            using (var trans = new Transaction(doc, "STB2026: reset colors"))
            {
                trans.Start();
                foreach (var id in ids)
                {
                    view.SetElementOverrides(id, clean);
                    count++;
                }
                trans.Commit();
            }

            return new { action = "reset_color", reset = count, view = view.Name };
        }

        // â•â•â• select â•â•â•
        private static object SelectElements(UIDocument uidoc, List<ElementId> ids)
        {
            uidoc.Selection.SetElementIds(ids);
            return new { action = "select", selected = ids.Count };
        }

        // â•â•â• move (dx, dy, dz Ð² Ð¼Ð¼) â•â•â•
        private static object MoveElements(Document doc, List<ElementId> ids, JObject data)
        {
            double dxMm = data.Value<double?>("dx") ?? 0;
            double dyMm = data.Value<double?>("dy") ?? 0;
            double dzMm = data.Value<double?>("dz") ?? 0;

            var vector = new XYZ(MmToFeet(dxMm), MmToFeet(dyMm), MmToFeet(dzMm));
            int count = 0;

            using (var trans = new Transaction(doc, "STB2026: move"))
            {
                trans.Start();
                foreach (var id in ids)
                {
                    var el = doc.GetElement(id);
                    if (el == null) continue;
                    ElementTransformUtils.MoveElement(doc, id, vector);
                    count++;
                }
                trans.Commit();
            }

            return new { action = "move", dx_mm = dxMm, dy_mm = dyMm, dz_mm = dzMm, moved = count };
        }

        // â•â•â• rotate (angle Ð² Ð³Ñ€Ð°Ð´ÑƒÑÐ°Ñ…) â•â•â•
        private static object RotateElements(Document doc, List<ElementId> ids, JObject data)
        {
            double angleDeg = data.Value<double?>("angle") ?? 0;
            double angleRad = angleDeg * Math.PI / 180.0;
            int count = 0;

            using (var trans = new Transaction(doc, "STB2026: rotate"))
            {
                trans.Start();
                foreach (var id in ids)
                {
                    var el = doc.GetElement(id);
                    if (el == null) continue;

                    // ÐžÑÑŒ Ð²Ñ€Ð°Ñ‰ÐµÐ½Ð¸Ñ â€” Ñ‡ÐµÑ€ÐµÐ· LocationPoint Ð¿Ð¾ Z
                    XYZ center;
                    if (el.Location is LocationPoint lp) center = lp.Point;
                    else if (el.Location is LocationCurve lc) center = lc.Curve.Evaluate(0.5, true);
                    else continue;

                    var axis = Line.CreateBound(center, center + XYZ.BasisZ);
                    ElementTransformUtils.RotateElement(doc, id, axis, angleRad);
                    count++;
                }
                trans.Commit();
            }

            return new { action = "rotate", angle_deg = angleDeg, rotated = count };
        }

        // â•â•â• delete â•â•â•
        private static object DeleteElements(Document doc, List<ElementId> ids)
        {
            int count = 0;
            var deletedIds = new List<int>();

            using (var trans = new Transaction(doc, "STB2026: delete"))
            {
                trans.Start();
                foreach (var id in ids)
                {
                    try
                    {
                        doc.Delete(id);
                        deletedIds.Add(id.IntegerValue);
                        count++;
                    }
                    catch { /* skip locked/protected */ }
                }
                trans.Commit();
            }

            return new { action = "delete", deleted = count, ids = deletedIds,
                         warning = "ÐÐ•ÐžÐ‘Ð ÐÐ¢Ð˜ÐœÐž! Ð­Ð»ÐµÐ¼ÐµÐ½Ñ‚Ñ‹ ÑƒÐ´Ð°Ð»ÐµÐ½Ñ‹." };
        }

        // â•â•â• isolate â•â•â•
        private static object IsolateElements(UIDocument uidoc, List<ElementId> ids)
        {
            var view = uidoc.ActiveView;

            using (var trans = new Transaction(uidoc.Document, "STB2026: isolate"))
            {
                trans.Start();
                view.IsolateElementsTemporary(ids);
                trans.Commit();
            }

            return new { action = "isolate", isolated = ids.Count, view = view.Name };
        }

        // â•â•â• create_tag â•â•â•
        private static object CreateTag(Document doc, UIDocument uidoc, List<ElementId> ids, JObject data)
        {
            string tagFamily = data.Value<string>("tag_family") ?? "";
            int count = 0;
            var errors = new List<string>();
            var view = uidoc.ActiveView;

            using (var trans = new Transaction(doc, "STB2026: tag"))
            {
                trans.Start();
                foreach (var id in ids)
                {
                    var el = doc.GetElement(id);
                    if (el == null) continue;

                    try
                    {
                        // Ð¢Ð¾Ñ‡ÐºÐ° Ð´Ð»Ñ Ð¼Ð°Ñ€ÐºÐ¸ â€” Ñ†ÐµÐ½Ñ‚Ñ€ BoundingBox
                        var bb = el.get_BoundingBox(view);
                        if (bb == null) { errors.Add($"id {id.IntegerValue}: Ð½ÐµÑ‚ bbox"); continue; }

                        var center = (bb.Min + bb.Max) / 2;
                        var tagRef = new Reference(el);
                        IndependentTag.Create(doc, view.Id, tagRef, false,
                            TagMode.TM_ADDBY_CATEGORY, TagOrientation.Horizontal,
                            center);
                        count++;
                    }
                    catch (Exception ex)
                    {
                        errors.Add($"id {id.IntegerValue}: {ex.Message}");
                    }
                }
                trans.Commit();
            }

            return new { action = "create_tag", tagged = count, errors };
        }

        // â•â•â• Helpers â•â•â•

        private static void SetParamValue(Parameter param, string value)
        {
            switch (param.StorageType)
            {
                case StorageType.String:
                    param.Set(value);
                    break;
                case StorageType.Integer:
                    if (int.TryParse(value, out int i)) param.Set(i);
                    else throw new ArgumentException($"ÐžÐ¶Ð¸Ð´Ð°Ð»Ð¾ÑÑŒ Ñ†ÐµÐ»Ð¾Ðµ Ñ‡Ð¸ÑÐ»Ð¾, Ð¿Ð¾Ð»ÑƒÑ‡ÐµÐ½Ð¾: {value}");
                    break;
                case StorageType.Double:
                    if (double.TryParse(value.Replace(",", "."),
                        System.Globalization.NumberStyles.Float,
                        System.Globalization.CultureInfo.InvariantCulture, out double d))
                        param.SetValueString(value); // Ñ‡ÐµÑ€ÐµÐ· SetValueString â€” ÑƒÑ‡Ð¸Ñ‚Ñ‹Ð²Ð°ÐµÑ‚ ÐµÐ´Ð¸Ð½Ð¸Ñ†Ñ‹
                    else
                        throw new ArgumentException($"ÐžÐ¶Ð¸Ð´Ð°Ð»Ð¾ÑÑŒ Ñ‡Ð¸ÑÐ»Ð¾, Ð¿Ð¾Ð»ÑƒÑ‡ÐµÐ½Ð¾: {value}");
                    break;
                case StorageType.ElementId:
                    if (int.TryParse(value, out int eid))
                        param.Set(new ElementId(eid));
                    else
                        throw new ArgumentException($"ÐžÐ¶Ð¸Ð´Ð°Ð»ÑÑ ID ÑÐ»ÐµÐ¼ÐµÐ½Ñ‚Ð°, Ð¿Ð¾Ð»ÑƒÑ‡ÐµÐ½Ð¾: {value}");
                    break;
            }
        }

        private static double MmToFeet(double mm) => mm / 304.8;
    }
}
