using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.UI;

namespace STB2026.RevitBridge.Handlers
{
    /// <summary>
    /// run_hvac_check — проверки систем ОВиК.
    /// Переиспользует логику Phase 1 сервисов, но возвращает данные вместо TaskDialog.
    /// check_type: velocity, system_validation, wall_intersections, tag_ducts.
    /// </summary>
    internal static class HvacCheckHandler
    {
        public static object Handle(UIApplication uiApp, Dictionary<string, object> p)
        {
            var doc = uiApp.ActiveUIDocument?.Document;
            if (doc == null)
                return new { error = "Нет открытого документа" };

            string checkType = p.TryGetValue("check_type", out var ct) ? ct?.ToString() ?? "" : "";
            string systemName = p.TryGetValue("system_name", out var sn) ? sn?.ToString() ?? "" : "";

            var ducts = CollectDucts(doc, systemName);

            switch (checkType.ToLowerInvariant())
            {
                case "velocity":
                    return CheckVelocities(doc, uiApp.ActiveUIDocument, ducts);

                case "system_validation":
                    return ValidateSystem(doc, ducts);

                case "wall_intersections":
                    return FindWallIntersections(doc, ducts);

                case "tag_ducts":
                    return TagDucts(doc, uiApp.ActiveUIDocument, ducts);

                default:
                    return new
                    {
                        error = $"Неизвестный тип проверки: '{checkType}'",
                        available = new[] { "velocity", "system_validation", "wall_intersections", "tag_ducts" }
                    };
            }
        }

        // ═══ velocity — проверка скоростей по СП 60.13330.2020 ═══
        private static object CheckVelocities(Document doc, UIDocument uidoc, List<Duct> ducts)
        {
            var results = new List<Dictionary<string, object>>();
            int ok = 0, warn = 0, critical = 0;
            var view = uidoc.ActiveView;

            using (var trans = new Transaction(doc, "STB2026: velocity check"))
            {
                trans.Start();

                foreach (var duct in ducts)
                {
                    var flow = duct.get_Parameter(BuiltInParameter.RBS_DUCT_FLOW_PARAM)?.AsDouble() ?? 0;
                    var size = GetDuctSize(duct);
                    var area = GetCrossSectionArea(duct);

                    // Расход в м³/ч → м³/с
                    double flowM3s = UnitUtils.ConvertFromInternalUnits(flow, UnitTypeId.CubicMetersPerSecond);
                    double flowM3h = flowM3s * 3600;
                    double velocityMs = area > 0 ? flowM3s / area : 0;

                    // Нормы по СП 60.13330.2020 (таблица 7.2)
                    double maxVelocity = 12.0; // м/с для магистралей
                    string status;
                    Color color;

                    if (velocityMs <= maxVelocity * 0.8)
                    {
                        status = "OK"; ok++;
                        color = new Color(0, 180, 0); // Зелёный
                    }
                    else if (velocityMs <= maxVelocity)
                    {
                        status = "WARN"; warn++;
                        color = new Color(255, 165, 0); // Оранжевый
                    }
                    else
                    {
                        status = "CRITICAL"; critical++;
                        color = new Color(255, 0, 0); // Красный
                    }

                    // Окрашиваем
                    var overrides = new OverrideGraphicSettings();
                    overrides.SetProjectionLineColor(color);
                    overrides.SetSurfaceForegroundPatternColor(color);
                    view.SetElementOverrides(duct.Id, overrides);

                    results.Add(new Dictionary<string, object>
                    {
                        ["id"] = duct.Id.IntegerValue,
                        ["system"] = duct.MEPSystem?.Name ?? "—",
                        ["size"] = size,
                        ["flow_m3h"] = Math.Round(flowM3h, 1),
                        ["velocity_ms"] = Math.Round(velocityMs, 2),
                        ["max_velocity_ms"] = maxVelocity,
                        ["status"] = status
                    });
                }

                trans.Commit();
            }

            return new
            {
                check = "velocity",
                standard = "СП 60.13330.2020, п. 7.1.9",
                total = ducts.Count,
                ok, warn, critical,
                colored_on_view = view.Name,
                elements = results.OrderByDescending(r => (double)r["velocity_ms"]).ToList()
            };
        }

        // ═══ system_validation — нулевые расходы, без системы ═══
        private static object ValidateSystem(Document doc, List<Duct> ducts)
        {
            var issues = new List<Dictionary<string, object>>();

            foreach (var duct in ducts)
            {
                var flow = duct.get_Parameter(BuiltInParameter.RBS_DUCT_FLOW_PARAM)?.AsDouble() ?? 0;
                double flowM3h = UnitUtils.ConvertFromInternalUnits(flow, UnitTypeId.CubicMetersPerSecond) * 3600;
                var systemName = duct.MEPSystem?.Name;

                var ductIssues = new List<string>();

                if (flowM3h < 0.1)
                    ductIssues.Add("Нулевой расход");

                if (string.IsNullOrWhiteSpace(systemName))
                    ductIssues.Add("Не назначена система");

                // Проверяем подключённость коннекторов
                var connMgr = duct.ConnectorManager;
                if (connMgr != null)
                {
                    int disconnected = 0;
                    foreach (Connector c in connMgr.Connectors)
                    {
                        if (!c.IsConnected) disconnected++;
                    }
                    if (disconnected > 0)
                        ductIssues.Add($"Неподключённых коннекторов: {disconnected}");
                }

                if (ductIssues.Count > 0)
                {
                    issues.Add(new Dictionary<string, object>
                    {
                        ["id"] = duct.Id.IntegerValue,
                        ["name"] = duct.Name,
                        ["system"] = systemName ?? "—",
                        ["flow_m3h"] = Math.Round(flowM3h, 1),
                        ["issues"] = ductIssues
                    });
                }
            }

            return new
            {
                check = "system_validation",
                total_ducts = ducts.Count,
                total_issues = issues.Count,
                ducts_ok = ducts.Count - issues.Count,
                problems = issues
            };
        }

        // ═══ wall_intersections — пересечения воздуховодов со стенами ═══
        private static object FindWallIntersections(Document doc, List<Duct> ducts)
        {
            var walls = new FilteredElementCollector(doc)
                .OfClass(typeof(Wall))
                .WhereElementIsNotElementType()
                .Cast<Wall>()
                .ToList();

            var intersections = new List<Dictionary<string, object>>();

            foreach (var duct in ducts)
            {
                if (!(duct.Location is LocationCurve lc)) continue;
                var ductLine = lc.Curve as Line;
                if (ductLine == null) continue;

                foreach (var wall in walls)
                {
                    if (!(wall.Location is LocationCurve wlc)) continue;
                    var wallLine = wlc.Curve as Line;
                    if (wallLine == null) continue;

                    // Проверяем пересечение в проекции XY
                    var intersection = FindIntersectionXY(ductLine, wallLine);
                    if (intersection == null) continue;

                    intersections.Add(new Dictionary<string, object>
                    {
                        ["duct_id"] = duct.Id.IntegerValue,
                        ["duct_system"] = duct.MEPSystem?.Name ?? "—",
                        ["duct_size"] = GetDuctSize(duct),
                        ["wall_id"] = wall.Id.IntegerValue,
                        ["wall_type"] = doc.GetElement(wall.GetTypeId())?.Name ?? "—",
                        ["wall_thickness_mm"] = Math.Round(
                            UnitUtils.ConvertFromInternalUnits(wall.Width, UnitTypeId.Millimeters), 0),
                        ["point_x_mm"] = Math.Round(intersection.X * 304.8, 0),
                        ["point_y_mm"] = Math.Round(intersection.Y * 304.8, 0),
                        ["point_z_mm"] = Math.Round(intersection.Z * 304.8, 0)
                    });
                }
            }

            return new
            {
                check = "wall_intersections",
                total_ducts = ducts.Count,
                total_walls = walls.Count,
                intersections_found = intersections.Count,
                intersections
            };
        }

        // ═══ tag_ducts — автомаркировка ═══
        private static object TagDucts(Document doc, UIDocument uidoc, List<Duct> ducts)
        {
            var view = uidoc.ActiveView;
            int tagged = 0;
            var errors = new List<string>();

            using (var trans = new Transaction(doc, "STB2026: auto-tag ducts"))
            {
                trans.Start();
                foreach (var duct in ducts)
                {
                    try
                    {
                        var bb = duct.get_BoundingBox(view);
                        if (bb == null) continue;

                        var center = (bb.Min + bb.Max) / 2;
                        var tagRef = new Reference(duct);
                        IndependentTag.Create(doc, view.Id, tagRef, false,
                            TagMode.TM_ADDBY_CATEGORY, TagOrientation.Horizontal, center);
                        tagged++;
                    }
                    catch (Exception ex)
                    {
                        errors.Add($"id {duct.Id.IntegerValue}: {ex.Message}");
                    }
                }
                trans.Commit();
            }

            return new
            {
                check = "tag_ducts",
                total_ducts = ducts.Count,
                tagged,
                failed = errors.Count,
                view = view.Name,
                errors = errors.Take(10).ToList()
            };
        }

        // ═══ Утилиты ═══

        private static List<Duct> CollectDucts(Document doc, string systemFilter)
        {
            var ducts = new FilteredElementCollector(doc)
                .OfClass(typeof(Duct))
                .WhereElementIsNotElementType()
                .Cast<Duct>();

            if (!string.IsNullOrWhiteSpace(systemFilter))
            {
                ducts = ducts.Where(d =>
                    d.MEPSystem?.Name?.IndexOf(systemFilter, StringComparison.OrdinalIgnoreCase) >= 0);
            }

            return ducts.ToList();
        }

        private static string GetDuctSize(Duct duct)
        {
            var w = duct.get_Parameter(BuiltInParameter.RBS_CURVE_WIDTH_PARAM)?.AsDouble();
            var h = duct.get_Parameter(BuiltInParameter.RBS_CURVE_HEIGHT_PARAM)?.AsDouble();
            var d = duct.get_Parameter(BuiltInParameter.RBS_CURVE_DIAMETER_PARAM)?.AsDouble();

            if (d.HasValue && d.Value > 0)
                return $"Ø{Math.Round(d.Value * 304.8)}";
            if (w.HasValue && h.HasValue)
                return $"{Math.Round(w.Value * 304.8)}×{Math.Round(h.Value * 304.8)}";
            return "—";
        }

        private static double GetCrossSectionArea(Duct duct)
        {
            var w = duct.get_Parameter(BuiltInParameter.RBS_CURVE_WIDTH_PARAM)?.AsDouble() ?? 0;
            var h = duct.get_Parameter(BuiltInParameter.RBS_CURVE_HEIGHT_PARAM)?.AsDouble() ?? 0;
            var d = duct.get_Parameter(BuiltInParameter.RBS_CURVE_DIAMETER_PARAM)?.AsDouble() ?? 0;

            if (d > 0) return Math.PI * d * d / 4 * 0.09290304; // ft² → m²
            if (w > 0 && h > 0) return w * h * 0.09290304;
            return 0;
        }

        private static XYZ? FindIntersectionXY(Line line1, Line line2)
        {
            var p1 = line1.GetEndPoint(0);
            var p2 = line1.GetEndPoint(1);
            var p3 = line2.GetEndPoint(0);
            var p4 = line2.GetEndPoint(1);

            double d1x = p2.X - p1.X, d1y = p2.Y - p1.Y;
            double d2x = p4.X - p3.X, d2y = p4.Y - p3.Y;
            double denom = d1x * d2y - d1y * d2x;

            if (Math.Abs(denom) < 1e-10) return null; // Параллельны

            double t = ((p3.X - p1.X) * d2y - (p3.Y - p1.Y) * d2x) / denom;
            double u = ((p3.X - p1.X) * d1y - (p3.Y - p1.Y) * d1x) / denom;

            if (t < 0 || t > 1 || u < 0 || u > 1) return null; // Не пересекаются

            double z = p1.Z + t * (p2.Z - p1.Z);
            return new XYZ(p1.X + t * d1x, p1.Y + t * d1y, z);
        }
    }
}
