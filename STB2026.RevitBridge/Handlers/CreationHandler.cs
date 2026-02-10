using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;
using Newtonsoft.Json.Linq;

namespace STB2026.RevitBridge.Handlers
{
    /// <summary>
    /// create_elements — создание элементов в модели Revit.
    ///
    /// Действия:
    ///   place_instance       — разместить экземпляр семейства (family, type, x/y/z, level)
    ///   create_wall          — создать стену по двум точкам
    ///   create_floor         — создать перекрытие по контуру точек
    ///   create_ceiling       — создать потолок по контуру
    ///   create_room          — создать помещение в заданной точке
    ///   create_space         — создать пространство
    ///   create_duct          — создать воздуховод (start, end, type, system)
    ///   create_pipe          — создать трубу
    ///   create_flex_duct     — создать гибкий воздуховод
    ///   create_duct_fitting  — вставить фитинг (elbow/tee/transition) между воздуховодами
    ///   create_insulation    — добавить изоляцию на воздуховод/трубу
    ///   create_opening       — создать отверстие в стене/перекрытии
    ///   create_text_note     — создать текстовую аннотацию
    ///   create_dimension     — создать размер между элементами
    ///   create_filled_region — создать штриховку
    /// </summary>
    internal static class CreationHandler
    {
        private const double MmToFt = 1.0 / 304.8;

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
                case "place_instance":      return PlaceInstance(doc, data);
                case "create_wall":         return CreateWall(doc, data);
                case "create_floor":        return CreateFloor(doc, data);
                case "create_ceiling":      return CreateCeiling(doc, data);
                case "create_room":         return CreateRoom(doc, data);
                case "create_space":        return CreateSpace(doc, data);
                case "create_duct":         return CreateDuct(doc, data);
                case "create_pipe":         return CreatePipe(doc, data);
                case "create_flex_duct":    return CreateFlexDuct(doc, data);
                case "create_insulation":   return CreateInsulation(doc, data);
                case "create_opening":      return CreateOpening(doc, uidoc, data);
                case "create_text_note":    return CreateTextNote(doc, uidoc, data);
                case "create_dimension":    return CreateDimensionLine(doc, uidoc, data);
                default:
                    return new
                    {
                        error = $"Неизвестное действие: '{action}'",
                        available = new[]
                        {
                            "place_instance", "create_wall", "create_floor", "create_ceiling",
                            "create_room", "create_space", "create_duct", "create_pipe",
                            "create_flex_duct", "create_insulation", "create_opening",
                            "create_text_note", "create_dimension"
                        }
                    };
            }
        }

        // ═══ place_instance ═══
        private static object PlaceInstance(Document doc, JObject data)
        {
            string familyName = data.Value<string>("family") ?? "";
            string typeName = data.Value<string>("type") ?? "";
            double x = (data.Value<double?>("x_mm") ?? 0) * MmToFt;
            double y = (data.Value<double?>("y_mm") ?? 0) * MmToFt;
            double z = (data.Value<double?>("z_mm") ?? 0) * MmToFt;
            int levelId = data.Value<int?>("level_id") ?? -1;

            var symbol = FindFamilySymbol(doc, familyName, typeName);
            if (symbol == null) return new { error = $"Семейство '{familyName}:{typeName}' не найдено" };

            Level level = levelId > 0
                ? doc.GetElement(new ElementId(levelId)) as Level
                : new FilteredElementCollector(doc).OfClass(typeof(Level)).Cast<Level>()
                    .OrderBy(l => Math.Abs(l.Elevation - z)).FirstOrDefault();

            if (level == null) return new { error = "Уровень не найден" };

            FamilyInstance inst;
            using (var trans = new Transaction(doc, "STB2026: разместить экземпляр"))
            {
                trans.Start();
                if (!symbol.IsActive) symbol.Activate();
                inst = doc.Create.NewFamilyInstance(
                    new XYZ(x, y, z), symbol, level, StructuralType.NonStructural);
                trans.Commit();
            }

            return new
            {
                action = "place_instance",
                element_id = inst.Id.IntegerValue,
                family = symbol.Family.Name,
                type = symbol.Name,
                level = level.Name
            };
        }

        // ═══ create_wall ═══
        private static object CreateWall(Document doc, JObject data)
        {
            double x1 = (data.Value<double?>("x1_mm") ?? 0) * MmToFt;
            double y1 = (data.Value<double?>("y1_mm") ?? 0) * MmToFt;
            double x2 = (data.Value<double?>("x2_mm") ?? 0) * MmToFt;
            double y2 = (data.Value<double?>("y2_mm") ?? 0) * MmToFt;
            double height = (data.Value<double?>("height_mm") ?? 3000) * MmToFt;
            int levelId = data.Value<int?>("level_id") ?? -1;
            string wallTypeName = data.Value<string>("wall_type") ?? "";

            var level = levelId > 0
                ? doc.GetElement(new ElementId(levelId)) as Level
                : new FilteredElementCollector(doc).OfClass(typeof(Level)).Cast<Level>().FirstOrDefault();
            if (level == null) return new { error = "Уровень не найден" };

            var wallType = !string.IsNullOrWhiteSpace(wallTypeName)
                ? new FilteredElementCollector(doc).OfClass(typeof(WallType)).Cast<WallType>()
                    .FirstOrDefault(wt => wt.Name.Equals(wallTypeName, StringComparison.OrdinalIgnoreCase))
                : null;

            var line = Line.CreateBound(new XYZ(x1, y1, 0), new XYZ(x2, y2, 0));

            Wall wall;
            using (var trans = new Transaction(doc, "STB2026: создание стены"))
            {
                trans.Start();
                wall = wallType != null
                    ? Wall.Create(doc, line, wallType.Id, level.Id, height, 0, false, false)
                    : Wall.Create(doc, line, level.Id, false);
                trans.Commit();
            }

            return new { action = "create_wall", element_id = wall.Id.IntegerValue, level = level.Name };
        }

        // ═══ create_floor ═══
        private static object CreateFloor(Document doc, JObject data)
        {
            var points = data["points"] as JArray;
            if (points == null || points.Count < 3)
                return new { error = "Укажите points — массив [{\"x_mm\",\"y_mm\"},...] (минимум 3)" };

            int levelId = data.Value<int?>("level_id") ?? -1;
            string floorTypeName = data.Value<string>("floor_type") ?? "";

            var level = levelId > 0
                ? doc.GetElement(new ElementId(levelId)) as Level
                : new FilteredElementCollector(doc).OfClass(typeof(Level)).Cast<Level>().FirstOrDefault();
            if (level == null) return new { error = "Уровень не найден" };

            var floorType = !string.IsNullOrWhiteSpace(floorTypeName)
                ? new FilteredElementCollector(doc).OfClass(typeof(FloorType)).Cast<FloorType>()
                    .FirstOrDefault(ft => ft.Name.Equals(floorTypeName, StringComparison.OrdinalIgnoreCase))
                : new FilteredElementCollector(doc).OfClass(typeof(FloorType)).Cast<FloorType>().FirstOrDefault();
            if (floorType == null) return new { error = "Тип пола не найден" };

            var curveLoop = new CurveLoop();
            var pts = new List<XYZ>();
            foreach (var pt in points)
            {
                double px = (pt.Value<double?>("x_mm") ?? 0) * MmToFt;
                double py = (pt.Value<double?>("y_mm") ?? 0) * MmToFt;
                pts.Add(new XYZ(px, py, level.Elevation));
            }
            for (int i = 0; i < pts.Count; i++)
                curveLoop.Append(Line.CreateBound(pts[i], pts[(i + 1) % pts.Count]));

            Floor floor;
            using (var trans = new Transaction(doc, "STB2026: создание перекрытия"))
            {
                trans.Start();
                floor = Floor.Create(doc, new List<CurveLoop> { curveLoop }, floorType.Id, level.Id);
                trans.Commit();
            }

            return new { action = "create_floor", element_id = floor.Id.IntegerValue };
        }

        // ═══ create_ceiling ═══
        private static object CreateCeiling(Document doc, JObject data)
        {
            var points = data["points"] as JArray;
            if (points == null || points.Count < 3)
                return new { error = "Укажите points (минимум 3)" };

            int levelId = data.Value<int?>("level_id") ?? -1;
            double offsetMm = data.Value<double?>("offset_mm") ?? 2700;

            var level = levelId > 0
                ? doc.GetElement(new ElementId(levelId)) as Level
                : new FilteredElementCollector(doc).OfClass(typeof(Level)).Cast<Level>().FirstOrDefault();
            if (level == null) return new { error = "Уровень не найден" };

            var ceilingType = new FilteredElementCollector(doc)
                .OfClass(typeof(CeilingType)).Cast<CeilingType>().FirstOrDefault();
            if (ceilingType == null) return new { error = "Тип потолка не найден" };

            var curveLoop = new CurveLoop();
            var pts = new List<XYZ>();
            foreach (var pt in points)
            {
                double px = (pt.Value<double?>("x_mm") ?? 0) * MmToFt;
                double py = (pt.Value<double?>("y_mm") ?? 0) * MmToFt;
                pts.Add(new XYZ(px, py, level.Elevation + offsetMm * MmToFt));
            }
            for (int i = 0; i < pts.Count; i++)
                curveLoop.Append(Line.CreateBound(pts[i], pts[(i + 1) % pts.Count]));

            Element ceiling;
            using (var trans = new Transaction(doc, "STB2026: создание потолка"))
            {
                trans.Start();
                ceiling = Ceiling.Create(doc, new List<CurveLoop> { curveLoop }, ceilingType.Id, level.Id);
                trans.Commit();
            }

            return new { action = "create_ceiling", element_id = ceiling.Id.IntegerValue };
        }

        // ═══ create_room ═══
        private static object CreateRoom(Document doc, JObject data)
        {
            double x = (data.Value<double?>("x_mm") ?? 0) * MmToFt;
            double y = (data.Value<double?>("y_mm") ?? 0) * MmToFt;
            int levelId = data.Value<int?>("level_id") ?? -1;
            string name = data.Value<string>("name") ?? "";

            var level = levelId > 0
                ? doc.GetElement(new ElementId(levelId)) as Level
                : new FilteredElementCollector(doc).OfClass(typeof(Level)).Cast<Level>().FirstOrDefault();
            if (level == null) return new { error = "Уровень не найден" };

            Autodesk.Revit.DB.Architecture.Room room;
            using (var trans = new Transaction(doc, "STB2026: создание помещения"))
            {
                trans.Start();
                var uv = new UV(x, y);
                room = doc.Create.NewRoom(level, uv);
                if (!string.IsNullOrWhiteSpace(name))
                    room.Name = name;
                trans.Commit();
            }

            return new { action = "create_room", element_id = room.Id.IntegerValue, name = room.Name };
        }

        // ═══ create_space ═══
        private static object CreateSpace(Document doc, JObject data)
        {
            double x = (data.Value<double?>("x_mm") ?? 0) * MmToFt;
            double y = (data.Value<double?>("y_mm") ?? 0) * MmToFt;
            int levelId = data.Value<int?>("level_id") ?? -1;
            string name = data.Value<string>("name") ?? "";

            var level = levelId > 0
                ? doc.GetElement(new ElementId(levelId)) as Level
                : new FilteredElementCollector(doc).OfClass(typeof(Level)).Cast<Level>().FirstOrDefault();
            if (level == null) return new { error = "Уровень не найден" };

            Space space;
            using (var trans = new Transaction(doc, "STB2026: создание пространства"))
            {
                trans.Start();
                var uv = new UV(x, y);
                space = doc.Create.NewSpace(level, uv);
                if (!string.IsNullOrWhiteSpace(name))
                    space.get_Parameter(BuiltInParameter.ROOM_NAME)?.Set(name);
                trans.Commit();
            }

            return new { action = "create_space", element_id = space.Id.IntegerValue, name = space.Name };
        }

        // ═══ create_duct ═══
        private static object CreateDuct(Document doc, JObject data)
        {
            double x1 = (data.Value<double?>("x1_mm") ?? 0) * MmToFt;
            double y1 = (data.Value<double?>("y1_mm") ?? 0) * MmToFt;
            double z1 = (data.Value<double?>("z1_mm") ?? 3000) * MmToFt;
            double x2 = (data.Value<double?>("x2_mm") ?? 1000) * MmToFt;
            double y2 = (data.Value<double?>("y2_mm") ?? 0) * MmToFt;
            double z2 = (data.Value<double?>("z2_mm") ?? 3000) * MmToFt;
            double widthMm = data.Value<double?>("width_mm") ?? 0;
            double heightMm = data.Value<double?>("height_mm") ?? 0;
            double diameterMm = data.Value<double?>("diameter_mm") ?? 0;
            string systemName = data.Value<string>("system_name") ?? "";
            string ductTypeName = data.Value<string>("duct_type") ?? "";
            int levelId = data.Value<int?>("level_id") ?? -1;

            var level = levelId > 0
                ? doc.GetElement(new ElementId(levelId)) as Level
                : new FilteredElementCollector(doc).OfClass(typeof(Level)).Cast<Level>()
                    .OrderBy(l => l.Elevation).FirstOrDefault();
            if (level == null) return new { error = "Уровень не найден" };

            // Тип воздуховода
            var ductType = !string.IsNullOrWhiteSpace(ductTypeName)
                ? new FilteredElementCollector(doc).OfClass(typeof(DuctType)).Cast<DuctType>()
                    .FirstOrDefault(dt => dt.Name.Contains(ductTypeName, StringComparison.OrdinalIgnoreCase))
                : new FilteredElementCollector(doc).OfClass(typeof(DuctType)).Cast<DuctType>().FirstOrDefault();
            if (ductType == null) return new { error = "Тип воздуховода не найден" };

            // Система
            var mepSystem = !string.IsNullOrWhiteSpace(systemName)
                ? new FilteredElementCollector(doc).OfClass(typeof(MechanicalSystemType)).Cast<MechanicalSystemType>()
                    .FirstOrDefault(s => s.Name.Contains(systemName, StringComparison.OrdinalIgnoreCase))
                : new FilteredElementCollector(doc).OfClass(typeof(MechanicalSystemType)).Cast<MechanicalSystemType>()
                    .FirstOrDefault();

            Duct duct;
            using (var trans = new Transaction(doc, "STB2026: создание воздуховода"))
            {
                trans.Start();
                duct = Duct.Create(doc, mepSystem?.Id ?? ElementId.InvalidElementId,
                    ductType.Id, level.Id,
                    new XYZ(x1, y1, z1), new XYZ(x2, y2, z2));

                // Размеры
                if (diameterMm > 0)
                    duct.get_Parameter(BuiltInParameter.RBS_CURVE_DIAMETER_PARAM)?.Set(diameterMm * MmToFt);
                else if (widthMm > 0 && heightMm > 0)
                {
                    duct.get_Parameter(BuiltInParameter.RBS_CURVE_WIDTH_PARAM)?.Set(widthMm * MmToFt);
                    duct.get_Parameter(BuiltInParameter.RBS_CURVE_HEIGHT_PARAM)?.Set(heightMm * MmToFt);
                }

                trans.Commit();
            }

            return new
            {
                action = "create_duct",
                element_id = duct.Id.IntegerValue,
                duct_type = ductType.Name
            };
        }

        // ═══ create_pipe ═══
        private static object CreatePipe(Document doc, JObject data)
        {
            double x1 = (data.Value<double?>("x1_mm") ?? 0) * MmToFt;
            double y1 = (data.Value<double?>("y1_mm") ?? 0) * MmToFt;
            double z1 = (data.Value<double?>("z1_mm") ?? 3000) * MmToFt;
            double x2 = (data.Value<double?>("x2_mm") ?? 1000) * MmToFt;
            double y2 = (data.Value<double?>("y2_mm") ?? 0) * MmToFt;
            double z2 = (data.Value<double?>("z2_mm") ?? 3000) * MmToFt;
            double diamMm = data.Value<double?>("diameter_mm") ?? 25;
            int levelId = data.Value<int?>("level_id") ?? -1;

            var level = levelId > 0
                ? doc.GetElement(new ElementId(levelId)) as Level
                : new FilteredElementCollector(doc).OfClass(typeof(Level)).Cast<Level>().OrderBy(l => l.Elevation).FirstOrDefault();
            if (level == null) return new { error = "Уровень не найден" };

            var pipeType = new FilteredElementCollector(doc).OfClass(typeof(PipeType)).Cast<PipeType>().FirstOrDefault();
            if (pipeType == null) return new { error = "Тип трубы не найден" };

            var pipingSystem = new FilteredElementCollector(doc)
                .OfClass(typeof(PipingSystemType)).Cast<PipingSystemType>().FirstOrDefault();

            Pipe pipe;
            using (var trans = new Transaction(doc, "STB2026: создание трубы"))
            {
                trans.Start();
                pipe = Pipe.Create(doc, pipingSystem?.Id ?? ElementId.InvalidElementId,
                    pipeType.Id, level.Id,
                    new XYZ(x1, y1, z1), new XYZ(x2, y2, z2));
                pipe.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM)?.Set(diamMm * MmToFt);
                trans.Commit();
            }

            return new { action = "create_pipe", element_id = pipe.Id.IntegerValue };
        }

        // ═══ create_flex_duct ═══
        private static object CreateFlexDuct(Document doc, JObject data)
        {
            var pointsArr = data["points"] as JArray;
            if (pointsArr == null || pointsArr.Count < 2)
                return new { error = "Укажите points — массив [{\"x_mm\",\"y_mm\",\"z_mm\"},...] (минимум 2)" };

            double diamMm = data.Value<double?>("diameter_mm") ?? 200;
            int levelId = data.Value<int?>("level_id") ?? -1;

            var level = levelId > 0
                ? doc.GetElement(new ElementId(levelId)) as Level
                : new FilteredElementCollector(doc).OfClass(typeof(Level)).Cast<Level>().OrderBy(l => l.Elevation).FirstOrDefault();
            if (level == null) return new { error = "Уровень не найден" };

            var flexType = new FilteredElementCollector(doc)
                .OfClass(typeof(FlexDuctType)).Cast<FlexDuctType>().FirstOrDefault();
            if (flexType == null) return new { error = "Тип гибкого воздуховода не найден" };

            var points = new List<XYZ>();
            foreach (var pt in pointsArr)
            {
                double px = (pt.Value<double?>("x_mm") ?? 0) * MmToFt;
                double py = (pt.Value<double?>("y_mm") ?? 0) * MmToFt;
                double pz = (pt.Value<double?>("z_mm") ?? 3000) * MmToFt;
                points.Add(new XYZ(px, py, pz));
            }

            var systemType = new FilteredElementCollector(doc)
                .OfClass(typeof(MechanicalSystemType)).Cast<MechanicalSystemType>().FirstOrDefault();

            FlexDuct flexDuct;
            using (var trans = new Transaction(doc, "STB2026: создание гибкого воздуховода"))
            {
                trans.Start();
                var startPt = points[0];
                var endPt = points[points.Count - 1];
                var midPoints = points.Count > 2 ? points.GetRange(1, points.Count - 2) : new List<XYZ> { (startPt + endPt) / 2 };
                flexDuct = FlexDuct.Create(doc, systemType?.Id ?? ElementId.InvalidElementId,
                    flexType.Id, level.Id, startPt, endPt, midPoints);
                flexDuct.get_Parameter(BuiltInParameter.RBS_CURVE_DIAMETER_PARAM)?.Set(diamMm * MmToFt);
                trans.Commit();
            }

            return new { action = "create_flex_duct", element_id = flexDuct.Id.IntegerValue };
        }

        // ═══ create_insulation ═══
        private static object CreateInsulation(Document doc, JObject data)
        {
            int elementId = data.Value<int?>("element_id") ?? -1;
            double thickness = (data.Value<double?>("thickness_mm") ?? 30) * MmToFt;
            string insulTypeName = data.Value<string>("insulation_type") ?? "";

            if (elementId <= 0) return new { error = "Укажите element_id (воздуховод или трубу)" };

            var el = doc.GetElement(new ElementId(elementId));
            if (el == null) return new { error = $"Элемент id={elementId} не найден" };

            // Ищем тип изоляции
            var insulType = !string.IsNullOrWhiteSpace(insulTypeName)
                ? new FilteredElementCollector(doc)
                    .OfCategory(BuiltInCategory.OST_DuctInsulations)
                    .WhereElementIsElementType()
                    .FirstOrDefault(t => t.Name.Contains(insulTypeName, StringComparison.OrdinalIgnoreCase))
                : new FilteredElementCollector(doc)
                    .OfCategory(BuiltInCategory.OST_DuctInsulations)
                    .WhereElementIsElementType()
                    .FirstOrDefault();

            if (insulType == null) return new { error = "Тип изоляции не найден" };

            DuctInsulation insul;
            using (var trans = new Transaction(doc, "STB2026: изоляция"))
            {
                trans.Start();
                insul = DuctInsulation.Create(doc, el.Id, insulType.Id, thickness);
                trans.Commit();
            }

            return new { action = "create_insulation", insulation_id = insul.Id.IntegerValue, element_id = elementId };
        }

        // ═══ create_opening ═══
        private static object CreateOpening(Document doc, UIDocument uidoc, JObject data)
        {
            int hostId = data.Value<int?>("host_id") ?? -1;
            double x = (data.Value<double?>("x_mm") ?? 0) * MmToFt;
            double y = (data.Value<double?>("y_mm") ?? 0) * MmToFt;
            double z = (data.Value<double?>("z_mm") ?? 1500) * MmToFt;
            double width = (data.Value<double?>("width_mm") ?? 500) * MmToFt;
            double height = (data.Value<double?>("height_mm") ?? 500) * MmToFt;

            if (hostId <= 0) return new { error = "Укажите host_id (стена или перекрытие)" };

            var host = doc.GetElement(new ElementId(hostId));
            if (host == null) return new { error = $"Элемент id={hostId} не найден" };

            Opening opening;
            using (var trans = new Transaction(doc, "STB2026: отверстие"))
            {
                trans.Start();

                if (host is Wall wall)
                {
                    var point = new XYZ(x, y, z);
                    opening = doc.Create.NewOpening(wall,
                        new XYZ(x - width / 2, y, z - height / 2),
                        new XYZ(x + width / 2, y, z + height / 2));
                }
                else
                {
                    return new { error = "Отверстия поддерживаются только в стенах через MCP" };
                }

                trans.Commit();
            }

            return new { action = "create_opening", opening_id = opening.Id.IntegerValue, host_id = hostId };
        }

        // ═══ create_text_note ═══
        private static object CreateTextNote(Document doc, UIDocument uidoc, JObject data)
        {
            double x = (data.Value<double?>("x_mm") ?? 0) * MmToFt;
            double y = (data.Value<double?>("y_mm") ?? 0) * MmToFt;
            string text = data.Value<string>("text") ?? "Текст";
            int viewId = data.Value<int?>("view_id") ?? -1;

            var view = viewId > 0
                ? doc.GetElement(new ElementId(viewId)) as View
                : uidoc.ActiveView;
            if (view == null) return new { error = "Вид не найден" };

            var textType = new FilteredElementCollector(doc)
                .OfClass(typeof(TextNoteType)).Cast<TextNoteType>().FirstOrDefault();
            if (textType == null) return new { error = "Тип текстовой заметки не найден" };

            TextNote note;
            using (var trans = new Transaction(doc, "STB2026: текст"))
            {
                trans.Start();
                note = TextNote.Create(doc, view.Id, new XYZ(x, y, 0), text, textType.Id);
                trans.Commit();
            }

            return new { action = "create_text_note", note_id = note.Id.IntegerValue, view = view.Name };
        }

        // ═══ create_dimension ═══
        private static object CreateDimensionLine(Document doc, UIDocument uidoc, JObject data)
        {
            int elem1Id = data.Value<int?>("element1_id") ?? -1;
            int elem2Id = data.Value<int?>("element2_id") ?? -1;
            int viewId = data.Value<int?>("view_id") ?? -1;

            if (elem1Id <= 0 || elem2Id <= 0)
                return new { error = "Укажите element1_id и element2_id" };

            var view = viewId > 0
                ? doc.GetElement(new ElementId(viewId)) as View
                : uidoc.ActiveView;
            if (view == null) return new { error = "Вид не найден" };

            var el1 = doc.GetElement(new ElementId(elem1Id));
            var el2 = doc.GetElement(new ElementId(elem2Id));
            if (el1 == null || el2 == null)
                return new { error = "Один из элементов не найден" };

            // Берём грани элементов для привязки размера
            var refs = new ReferenceArray();
            var face1 = GetBestReference(el1);
            var face2 = GetBestReference(el2);
            if (face1 == null || face2 == null)
                return new { error = "Не удалось получить ссылки для размера" };

            refs.Append(face1);
            refs.Append(face2);

            var bb1 = el1.get_BoundingBox(view);
            var bb2 = el2.get_BoundingBox(view);
            if (bb1 == null || bb2 == null)
                return new { error = "Не удалось получить BoundingBox" };

            var mid1 = (bb1.Min + bb1.Max) / 2;
            var mid2 = (bb2.Min + bb2.Max) / 2;
            var dimLine = Line.CreateBound(mid1, mid2);

            Dimension dim;
            using (var trans = new Transaction(doc, "STB2026: размер"))
            {
                trans.Start();
                dim = doc.Create.NewDimension(view, dimLine, refs);
                trans.Commit();
            }

            return new { action = "create_dimension", dimension_id = dim.Id.IntegerValue };
        }

        // ═══ Helpers ═══

        private static FamilySymbol FindFamilySymbol(Document doc, string familyName, string typeName)
        {
            var symbols = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilySymbol))
                .Cast<FamilySymbol>();

            if (!string.IsNullOrWhiteSpace(familyName) && !string.IsNullOrWhiteSpace(typeName))
                return symbols.FirstOrDefault(s =>
                    s.Family.Name.Equals(familyName, StringComparison.OrdinalIgnoreCase) &&
                    s.Name.Equals(typeName, StringComparison.OrdinalIgnoreCase));

            if (!string.IsNullOrWhiteSpace(familyName))
                return symbols.FirstOrDefault(s =>
                    s.Family.Name.Equals(familyName, StringComparison.OrdinalIgnoreCase));

            return null;
        }

        private static Reference GetBestReference(Element el)
        {
            try
            {
                if (el.Location is LocationCurve lc)
                    return lc.Curve.Reference;
                if (el.Location is LocationPoint lp)
                    return new Reference(el);
            }
            catch { }
            try { return new Reference(el); } catch { }
            return null;
        }
    }
}
