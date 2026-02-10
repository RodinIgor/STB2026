using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Newtonsoft.Json.Linq;

namespace STB2026.RevitBridge.Handlers
{
    /// <summary>
    /// query_geometry — геометрические запросы, измерения, анализ.
    ///
    /// Действия:
    ///   measure_distance       — расстояние между двумя элементами
    ///   measure_length         — длина элемента (стена, воздуховод, труба)
    ///   get_bounding_box       — BoundingBox элемента(ов)
    ///   get_location           — координаты размещения элемента
    ///   get_connectors         — коннекторы MEP-элемента (точки подключения)
    ///   get_room_at_point      — помещение в заданной точке
    ///   get_room_boundaries    — границы помещения (контур)
    ///   find_nearest           — ближайший элемент заданной категории к точке
    ///   ray_cast               — луч от точки в направлении — что пересекает
    ///   check_intersections    — пересечения между двумя наборами элементов
    ///   get_area               — площадь элемента (помещение, перекрытие, стена)
    ///   get_volume             — объём элемента
    /// </summary>
    internal static class GeometryHandler
    {
        private const double FtToMm = 304.8;

        public static object Handle(UIApplication uiApp, Dictionary<string, object> p)
        {
            var uidoc = uiApp.ActiveUIDocument;
            if (uidoc == null) return new { error = "Нет открытого документа" };
            var doc = uidoc.Document;

            string action = p.TryGetValue("action", out var a) ? a?.ToString() ?? "" : "";
            string dataStr = p.TryGetValue("data", out var d) ? d?.ToString() ?? "{}" : "{}";

            JObject data;
            try { data = JObject.Parse(dataStr); }
            catch { return new { error = $"Невалидный JSON: {dataStr}" }; }

            switch (action.ToLowerInvariant())
            {
                case "measure_distance":    return MeasureDistance(doc, data);
                case "measure_length":      return MeasureLength(doc, data);
                case "get_bounding_box":    return GetBoundingBox(doc, uidoc, data);
                case "get_location":        return GetLocation(doc, data);
                case "get_connectors":      return GetConnectors(doc, data);
                case "get_room_at_point":   return GetRoomAtPoint(doc, data);
                case "get_room_boundaries": return GetRoomBoundaries(doc, data);
                case "find_nearest":        return FindNearest(doc, data);
                case "check_intersections": return CheckIntersections(doc, data);
                case "get_area":            return GetArea(doc, data);
                case "get_volume":          return GetVolume(doc, data);
                default:
                    return new
                    {
                        error = $"Неизвестное действие: '{action}'",
                        available = new[]
                        {
                            "measure_distance", "measure_length", "get_bounding_box",
                            "get_location", "get_connectors", "get_room_at_point",
                            "get_room_boundaries", "find_nearest",
                            "check_intersections", "get_area", "get_volume"
                        }
                    };
            }
        }

        // ═══ measure_distance ═══
        private static object MeasureDistance(Document doc, JObject data)
        {
            int id1 = data.Value<int?>("element1_id") ?? -1;
            int id2 = data.Value<int?>("element2_id") ?? -1;

            if (id1 <= 0 || id2 <= 0)
                return new { error = "Укажите element1_id и element2_id" };

            var el1 = doc.GetElement(new ElementId(id1));
            var el2 = doc.GetElement(new ElementId(id2));
            if (el1 == null || el2 == null)
                return new { error = "Один из элементов не найден" };

            var pt1 = GetCenterPoint(el1);
            var pt2 = GetCenterPoint(el2);
            if (pt1 == null || pt2 == null)
                return new { error = "Не удалось определить координаты" };

            double dist = pt1.DistanceTo(pt2) * FtToMm;
            double distXY = new XYZ(pt1.X, pt1.Y, 0).DistanceTo(new XYZ(pt2.X, pt2.Y, 0)) * FtToMm;
            double distZ = Math.Abs(pt1.Z - pt2.Z) * FtToMm;

            return new
            {
                action = "measure_distance",
                distance_mm = Math.Round(dist, 1),
                distance_xy_mm = Math.Round(distXY, 1),
                distance_z_mm = Math.Round(distZ, 1),
                point1_mm = Pt(pt1),
                point2_mm = Pt(pt2)
            };
        }

        // ═══ measure_length ═══
        private static object MeasureLength(Document doc, JObject data)
        {
            int elId = data.Value<int?>("element_id") ?? -1;
            if (elId <= 0) return new { error = "Укажите element_id" };

            var el = doc.GetElement(new ElementId(elId));
            if (el == null) return new { error = $"Элемент id={elId} не найден" };

            double length = 0;
            if (el.Location is LocationCurve lc)
                length = lc.Curve.Length * FtToMm;
            else
            {
                var param = el.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH);
                if (param != null && param.HasValue)
                    length = param.AsDouble() * FtToMm;
            }

            return new { action = "measure_length", element_id = elId, length_mm = Math.Round(length, 1) };
        }

        // ═══ get_bounding_box ═══
        private static object GetBoundingBox(Document doc, UIDocument uidoc, JObject data)
        {
            string idsStr = data.Value<string>("element_ids") ?? "";
            if (string.IsNullOrWhiteSpace(idsStr)) return new { error = "Укажите element_ids" };

            var ids = ParamHelper.ParseIds(idsStr);
            var results = new List<object>();

            foreach (var id in ids)
            {
                var el = doc.GetElement(id);
                if (el == null) continue;

                var bb = el.get_BoundingBox(null);
                if (bb == null) continue;

                results.Add(new
                {
                    id = id.IntegerValue,
                    name = el.Name,
                    min = Pt(bb.Min),
                    max = Pt(bb.Max),
                    size = new
                    {
                        width_mm = Math.Round((bb.Max.X - bb.Min.X) * FtToMm, 1),
                        depth_mm = Math.Round((bb.Max.Y - bb.Min.Y) * FtToMm, 1),
                        height_mm = Math.Round((bb.Max.Z - bb.Min.Z) * FtToMm, 1)
                    }
                });
            }

            return new { action = "get_bounding_box", count = results.Count, elements = results };
        }

        // ═══ get_location ═══
        private static object GetLocation(Document doc, JObject data)
        {
            string idsStr = data.Value<string>("element_ids") ?? "";
            if (string.IsNullOrWhiteSpace(idsStr)) return new { error = "Укажите element_ids" };

            var ids = ParamHelper.ParseIds(idsStr);
            var results = new List<object>();

            foreach (var id in ids)
            {
                var el = doc.GetElement(id);
                if (el == null) continue;

                if (el.Location is LocationPoint lp)
                {
                    results.Add(new
                    {
                        id = id.IntegerValue,
                        type = "point",
                        point = Pt(lp.Point),
                        rotation_deg = Math.Round(lp.Rotation * 180 / Math.PI, 2)
                    });
                }
                else if (el.Location is LocationCurve lc)
                {
                    var start = lc.Curve.GetEndPoint(0);
                    var end = lc.Curve.GetEndPoint(1);
                    results.Add(new
                    {
                        id = id.IntegerValue,
                        type = "curve",
                        start = Pt(start),
                        end = Pt(end),
                        length_mm = Math.Round(lc.Curve.Length * FtToMm, 1)
                    });
                }
                else
                {
                    var center = GetCenterPoint(el);
                    results.Add(new
                    {
                        id = id.IntegerValue,
                        type = "bbox_center",
                        point = center != null ? Pt(center) : null
                    });
                }
            }

            return new { action = "get_location", count = results.Count, elements = results };
        }

        // ═══ get_connectors ═══
        private static object GetConnectors(Document doc, JObject data)
        {
            int elId = data.Value<int?>("element_id") ?? -1;
            if (elId <= 0) return new { error = "Укажите element_id" };

            var el = doc.GetElement(new ElementId(elId));
            if (el == null) return new { error = $"Элемент id={elId} не найден" };

            ConnectorSet connectors = null;
            if (el is FamilyInstance fi && fi.MEPModel?.ConnectorManager != null)
                connectors = fi.MEPModel.ConnectorManager.Connectors;
            else if (el is MEPCurve mc && mc.ConnectorManager != null)
                connectors = mc.ConnectorManager.Connectors;

            if (connectors == null)
                return new { error = "Элемент не имеет коннекторов" };

            var result = new List<object>();
            foreach (Connector c in connectors)
            {
                var info = new Dictionary<string, object>
                {
                    ["index"] = c.Id,
                    ["domain"] = c.Domain.ToString(),
                    ["type"] = c.ConnectorType.ToString(),
                    ["is_connected"] = c.IsConnected,
                    ["origin"] = Pt(c.Origin),
                    ["shape"] = c.Shape.ToString()
                };

                if (c.Domain == Domain.DomainHvac || c.Domain == Domain.DomainPiping)
                {
                    try { info["flow"] = Math.Round(c.Flow * 101.94, 1); } catch { } // CFM→m³/h approx
                    try
                    {
                        if (c.Shape == ConnectorProfileType.Round)
                            info["diameter_mm"] = Math.Round(c.Radius * 2 * FtToMm, 1);
                        else
                        {
                            info["width_mm"] = Math.Round(c.Width * FtToMm, 1);
                            info["height_mm"] = Math.Round(c.Height * FtToMm, 1);
                        }
                    }
                    catch { }
                }

                if (c.IsConnected)
                {
                    var connected = new List<object>();
                    foreach (Connector other in c.AllRefs)
                    {
                        if (other.Owner.Id != el.Id)
                        {
                            connected.Add(new
                            {
                                element_id = other.Owner.Id.IntegerValue,
                                name = other.Owner.Name,
                                connector_id = other.Id
                            });
                        }
                    }
                    info["connected_to"] = connected;
                }

                result.Add(info);
            }

            return new { action = "get_connectors", element_id = elId, count = result.Count, connectors = result };
        }

        // ═══ get_room_at_point ═══
        private static object GetRoomAtPoint(Document doc, JObject data)
        {
            double x = (data.Value<double?>("x_mm") ?? 0) / FtToMm;
            double y = (data.Value<double?>("y_mm") ?? 0) / FtToMm;
            double z = (data.Value<double?>("z_mm") ?? 0) / FtToMm;

            var point = new XYZ(x, y, z);
            var room = doc.GetRoomAtPoint(point);

            if (room == null)
                return new { action = "get_room_at_point", found = false, message = "Помещение не найдено" };

            return new
            {
                action = "get_room_at_point",
                found = true,
                room_id = room.Id.IntegerValue,
                name = room.get_Parameter(BuiltInParameter.ROOM_NAME)?.AsString() ?? "",
                number = room.Number,
                area_m2 = Math.Round(room.Area * 0.092903, 2),
                level = room.Level?.Name ?? ""
            };
        }

        // ═══ get_room_boundaries ═══
        private static object GetRoomBoundaries(Document doc, JObject data)
        {
            int roomId = data.Value<int?>("room_id") ?? -1;
            if (roomId <= 0) return new { error = "Укажите room_id" };

            var room = doc.GetElement(new ElementId(roomId)) as Autodesk.Revit.DB.Architecture.Room;
            if (room == null) return new { error = $"Помещение id={roomId} не найдено" };

            var options = new SpatialElementBoundaryOptions();
            var segments = room.GetBoundarySegments(options);

            var loops = new List<List<object>>();
            if (segments != null)
            {
                foreach (var loop in segments)
                {
                    var pts = new List<object>();
                    foreach (var seg in loop)
                    {
                        var curve = seg.GetCurve();
                        pts.Add(new
                        {
                            start = Pt(curve.GetEndPoint(0)),
                            end = Pt(curve.GetEndPoint(1)),
                            wall_id = seg.ElementId.IntegerValue
                        });
                    }
                    loops.Add(pts);
                }
            }

            return new
            {
                action = "get_room_boundaries",
                room_id = roomId,
                name = room.get_Parameter(BuiltInParameter.ROOM_NAME)?.AsString() ?? "",
                area_m2 = Math.Round(room.Area * 0.092903, 2),
                boundary_loops = loops.Count,
                boundaries = loops
            };
        }

        // ═══ find_nearest ═══
        private static object FindNearest(Document doc, JObject data)
        {
            double x = (data.Value<double?>("x_mm") ?? 0) / FtToMm;
            double y = (data.Value<double?>("y_mm") ?? 0) / FtToMm;
            double z = (data.Value<double?>("z_mm") ?? 0) / FtToMm;
            string category = data.Value<string>("category") ?? "";
            int count = data.Value<int?>("count") ?? 1;

            if (string.IsNullOrWhiteSpace(category))
                return new { error = "Укажите category" };

            var point = new XYZ(x, y, z);

            // Используем тот же подход для поиска BuiltInCategory, что и ElementsHandler
            var elements = new FilteredElementCollector(doc)
                .WhereElementIsNotElementType()
                .ToElements()
                .Where(e => e.Category?.Name?.Equals(category, StringComparison.OrdinalIgnoreCase) == true)
                .Select(e => new { el = e, dist = GetDistToElement(e, point) })
                .Where(x2 => x2.dist >= 0)
                .OrderBy(x2 => x2.dist)
                .Take(count)
                .Select(x2 => new
                {
                    id = x2.el.Id.IntegerValue,
                    name = x2.el.Name,
                    distance_mm = Math.Round(x2.dist * FtToMm, 1),
                    location = GetCenterPoint(x2.el) != null ? Pt(GetCenterPoint(x2.el)) : null
                })
                .ToList();

            return new { action = "find_nearest", count = elements.Count, elements };
        }

        // ═══ check_intersections ═══
        private static object CheckIntersections(Document doc, JObject data)
        {
            string ids1Str = data.Value<string>("set1_ids") ?? "";
            string ids2Str = data.Value<string>("set2_ids") ?? "";

            if (string.IsNullOrWhiteSpace(ids1Str) || string.IsNullOrWhiteSpace(ids2Str))
                return new { error = "Укажите set1_ids и set2_ids" };

            var set1 = ParamHelper.ParseIds(ids1Str);
            var set2 = ParamHelper.ParseIds(ids2Str);

            var intersections = new List<object>();

            foreach (var id1 in set1)
            {
                var el1 = doc.GetElement(id1);
                if (el1 == null) continue;
                var bb1 = el1.get_BoundingBox(null);
                if (bb1 == null) continue;

                foreach (var id2 in set2)
                {
                    var el2 = doc.GetElement(id2);
                    if (el2 == null) continue;
                    var bb2 = el2.get_BoundingBox(null);
                    if (bb2 == null) continue;

                    if (BBoxOverlap(bb1, bb2))
                    {
                        intersections.Add(new
                        {
                            element1_id = id1.IntegerValue,
                            element1_name = el1.Name,
                            element2_id = id2.IntegerValue,
                            element2_name = el2.Name
                        });
                    }
                }
            }

            return new { action = "check_intersections", count = intersections.Count, intersections };
        }

        // ═══ get_area ═══
        private static object GetArea(Document doc, JObject data)
        {
            int elId = data.Value<int?>("element_id") ?? -1;
            if (elId <= 0) return new { error = "Укажите element_id" };

            var el = doc.GetElement(new ElementId(elId));
            if (el == null) return new { error = $"Элемент id={elId} не найден" };

            double area = 0;
            var param = el.get_Parameter(BuiltInParameter.ROOM_AREA)
                     ?? el.get_Parameter(BuiltInParameter.HOST_AREA_COMPUTED);
            if (param != null && param.HasValue)
                area = param.AsDouble() * 0.092903; // ft² → m²

            return new { action = "get_area", element_id = elId, area_m2 = Math.Round(area, 3) };
        }

        // ═══ get_volume ═══
        private static object GetVolume(Document doc, JObject data)
        {
            int elId = data.Value<int?>("element_id") ?? -1;
            if (elId <= 0) return new { error = "Укажите element_id" };

            var el = doc.GetElement(new ElementId(elId));
            if (el == null) return new { error = $"Элемент id={elId} не найден" };

            double volume = 0;
            var param = el.get_Parameter(BuiltInParameter.ROOM_VOLUME)
                     ?? el.get_Parameter(BuiltInParameter.HOST_VOLUME_COMPUTED);
            if (param != null && param.HasValue)
                volume = param.AsDouble() * 0.0283168; // ft³ → m³

            return new { action = "get_volume", element_id = elId, volume_m3 = Math.Round(volume, 4) };
        }

        // ═══ Helpers ═══

        private static XYZ GetCenterPoint(Element el)
        {
            if (el.Location is LocationPoint lp) return lp.Point;
            if (el.Location is LocationCurve lc) return lc.Curve.Evaluate(0.5, true);
            var bb = el.get_BoundingBox(null);
            return bb != null ? (bb.Min + bb.Max) / 2 : null;
        }

        private static double GetDistToElement(Element el, XYZ point)
        {
            var center = GetCenterPoint(el);
            return center != null ? center.DistanceTo(point) : -1;
        }

        private static bool BBoxOverlap(BoundingBoxXYZ a, BoundingBoxXYZ b)
        {
            return a.Min.X <= b.Max.X && a.Max.X >= b.Min.X &&
                   a.Min.Y <= b.Max.Y && a.Max.Y >= b.Min.Y &&
                   a.Min.Z <= b.Max.Z && a.Max.Z >= b.Min.Z;
        }

        private static object Pt(XYZ p) => p == null ? null : new
        {
            x_mm = Math.Round(p.X * FtToMm, 1),
            y_mm = Math.Round(p.Y * FtToMm, 1),
            z_mm = Math.Round(p.Z * FtToMm, 1)
        };
    }
}
