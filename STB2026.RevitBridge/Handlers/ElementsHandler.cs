using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.UI;

namespace STB2026.RevitBridge.Handlers
{
    /// <summary>
    /// get_elements — запрос элементов из модели.
    /// Режимы: by_category, by_view, by_selection, by_ids, by_system, by_linked.
    /// 
    /// by_linked — элементы из связанного файла.
    ///   value: "имя_связи:категория" (например "ТестАР2026:Rooms")
    ///   Если имя_связи пустое — ищет во всех связях.
    /// </summary>
    internal static class ElementsHandler
    {
        public static object Handle(UIApplication uiApp, Dictionary<string, object> p)
        {
            var doc = uiApp.ActiveUIDocument?.Document;
            if (doc == null)
                return new { error = "Нет открытого документа" };

            string mode = p.TryGetValue("mode", out var m) ? m?.ToString() ?? "" : "";
            string value = p.TryGetValue("value", out var v) ? v?.ToString() ?? "" : "";
            int limit = p.TryGetValue("limit", out var lim) && int.TryParse(lim?.ToString(), out int l) ? l : 500;

            switch (mode.ToLowerInvariant())
            {
                case "by_category":
                    return BuildResult(mode, value, GetByCategory(doc, value), limit);

                case "by_view":
                    return BuildResult(mode, value, GetByView(uiApp, value), limit);

                case "by_selection":
                    return BuildResult(mode, value, GetBySelection(uiApp), limit);

                case "by_ids":
                    return BuildResult(mode, value, GetByIds(doc, value), limit);

                case "by_system":
                    return BuildResult(mode, value, GetBySystem(doc, value), limit);

                case "by_linked":
                    return HandleLinked(doc, value, limit);

                default:
                    return new
                    {
                        error = $"Неизвестный режим: '{mode}'",
                        available_modes = new[] { "by_category", "by_view", "by_selection", "by_ids", "by_system", "by_linked" }
                    };
            }
        }

        private static object BuildResult(string mode, string value, List<Element> elements, int limit)
        {
            int total = elements.Count;
            var limited = elements.Take(limit).ToList();

            return new
            {
                mode,
                value,
                total_found = total,
                returned = limited.Count,
                truncated = total > limit,
                elements = limited.Select(ParamHelper.ElementSummary).ToList()
            };
        }

        // ═══════════════════════════════════════════════════════════
        // by_linked — элементы из связанных файлов
        // ═══════════════════════════════════════════════════════════

        /// <summary>
        /// Получить элементы из связанного файла.
        /// value формат: "имя_связи:категория" или просто "категория" (поиск во всех связях).
        /// Примеры: "ТестАР2026:Rooms", "Rooms", ":Walls"
        /// </summary>
        private static object HandleLinked(Document hostDoc, string value, int limit)
        {
            // Разбираем value
            string linkFilter = "";
            string categoryFilter = "";

            if (value.Contains(":"))
            {
                var parts = value.Split(new[] { ':' }, 2);
                linkFilter = parts[0].Trim();
                categoryFilter = parts.Length > 1 ? parts[1].Trim() : "";
            }
            else
            {
                categoryFilter = value.Trim();
            }

            // Найти все RevitLinkInstance
            var linkInstances = new FilteredElementCollector(hostDoc)
                .OfClass(typeof(RevitLinkInstance))
                .Cast<RevitLinkInstance>()
                .ToList();

            if (linkInstances.Count == 0)
                return new { error = "Нет связанных файлов в проекте" };

            // Фильтрация по имени связи
            if (!string.IsNullOrWhiteSpace(linkFilter))
            {
                linkInstances = linkInstances
                    .Where(li => li.Name.IndexOf(linkFilter, StringComparison.OrdinalIgnoreCase) >= 0)
                    .ToList();

                if (linkInstances.Count == 0)
                    return new
                    {
                        error = $"Связанный файл '{linkFilter}' не найден",
                        available_links = GetAllLinkNames(hostDoc)
                    };
            }

            // Собираем элементы из всех подходящих связей
            var allResults = new List<Dictionary<string, object>>();

            foreach (var linkInstance in linkInstances)
            {
                var linkDoc = linkInstance.GetLinkDocument();
                if (linkDoc == null)
                    continue; // Связь не загружена

                string linkName = linkInstance.Name;
                var transform = linkInstance.GetTotalTransform();

                List<Element> elements;
                if (string.IsNullOrWhiteSpace(categoryFilter))
                {
                    // Без категории — вернуть сводку связи
                    return GetLinkedDocSummary(linkDoc, linkName);
                }
                else
                {
                    elements = GetByCategory(linkDoc, categoryFilter);
                }

                foreach (var el in elements)
                {
                    var summary = LinkedElementSummary(el, linkName, transform);
                    allResults.Add(summary);
                }
            }

            int total = allResults.Count;
            var limited = allResults.Take(limit).ToList();

            return new
            {
                mode = "by_linked",
                value,
                total_found = total,
                returned = limited.Count,
                truncated = total > limit,
                elements = limited
            };
        }

        /// <summary>Сводка связанного документа (когда не указана категория).</summary>
        private static object GetLinkedDocSummary(Document linkDoc, string linkName)
        {
            var categories = new List<object>();
            foreach (Category c in linkDoc.Settings.Categories)
            {
                if (c == null) continue;
                try
                {
                    var count = new FilteredElementCollector(linkDoc)
                        .OfCategoryId(c.Id)
                        .WhereElementIsNotElementType()
                        .GetElementCount();

                    if (count > 0)
                        categories.Add(new { name = c.Name, id = c.Id.IntegerValue, count });
                }
                catch { /* skip */ }
            }

            return new
            {
                mode = "by_linked",
                link_name = linkName,
                link_doc_title = linkDoc.Title,
                categories = categories
                    .Cast<dynamic>()
                    .OrderByDescending(c => (int)c.count)
                    .Take(30)
                    .ToList()
            };
        }

        /// <summary>
        /// Сводка элемента из связанного файла.
        /// Включает трансформированные координаты (в системе координат хоста).
        /// </summary>
        private static Dictionary<string, object> LinkedElementSummary(Element el, string linkName, Transform transform)
        {
            var result = new Dictionary<string, object>
            {
                ["id"] = el.Id.IntegerValue,
                ["name"] = el.Name ?? "",
                ["category"] = el.Category?.Name ?? "—",
                ["link_name"] = linkName
            };

            // Тип
            var typeId = el.GetTypeId();
            if (typeId != null && typeId != ElementId.InvalidElementId)
            {
                var type = el.Document.GetElement(typeId);
                if (type != null)
                    result["type_name"] = type.Name;
            }

            // Для Room — дополнительные данные
            if (el is Room room)
            {
                result["number"] = room.Number ?? "";
                result["room_name"] = room.get_Parameter(BuiltInParameter.ROOM_NAME)?.AsString() ?? "";

                // Площадь (м²)
                var areaParam = room.get_Parameter(BuiltInParameter.ROOM_AREA);
                if (areaParam != null && areaParam.HasValue)
                {
                    double areaFt2 = areaParam.AsDouble();
                    double areaM2 = areaFt2 * 0.092903;
                    result["area_m2"] = Math.Round(areaM2, 2);
                }

                // Объём (м³)
                var volParam = room.get_Parameter(BuiltInParameter.ROOM_VOLUME);
                if (volParam != null && volParam.HasValue)
                {
                    double volFt3 = volParam.AsDouble();
                    double volM3 = volFt3 * 0.0283168;
                    result["volume_m3"] = Math.Round(volM3, 2);
                }

                // Высота
                var heightParam = room.get_Parameter(BuiltInParameter.ROOM_HEIGHT);
                if (heightParam != null && heightParam.HasValue)
                {
                    double heightFt = heightParam.AsDouble();
                    double heightMm = heightFt * 304.8;
                    result["height_mm"] = Math.Round(heightMm, 0);
                }

                // Уровень
                if (room.Level != null)
                    result["level"] = room.Level.Name;

                // Центр помещения в координатах хоста
                var location = room.Location as LocationPoint;
                if (location != null)
                {
                    var pt = transform.OfPoint(location.Point);
                    result["center_x_mm"] = Math.Round(pt.X * 304.8, 1);
                    result["center_y_mm"] = Math.Round(pt.Y * 304.8, 1);
                    result["center_z_mm"] = Math.Round(pt.Z * 304.8, 1);
                }

                // Периметр
                var perimParam = room.get_Parameter(BuiltInParameter.ROOM_PERIMETER);
                if (perimParam != null && perimParam.HasValue)
                {
                    double perimFt = perimParam.AsDouble();
                    result["perimeter_mm"] = Math.Round(perimFt * 304.8, 0);
                }
            }
            // Для Space (пространства MEP)
            else if (el is Space space)
            {
                result["number"] = space.Number ?? "";
                result["space_name"] = space.get_Parameter(BuiltInParameter.ROOM_NAME)?.AsString() ?? "";

                var areaParam = space.get_Parameter(BuiltInParameter.ROOM_AREA);
                if (areaParam != null && areaParam.HasValue)
                    result["area_m2"] = Math.Round(areaParam.AsDouble() * 0.092903, 2);

                var volParam = space.get_Parameter(BuiltInParameter.ROOM_VOLUME);
                if (volParam != null && volParam.HasValue)
                    result["volume_m3"] = Math.Round(volParam.AsDouble() * 0.0283168, 2);

                var heightParam = space.get_Parameter(BuiltInParameter.ROOM_HEIGHT);
                if (heightParam != null && heightParam.HasValue)
                    result["height_mm"] = Math.Round(heightParam.AsDouble() * 304.8, 0);
            }
            // Для стен — габариты
            else if (el is Wall wall)
            {
                var lengthParam = wall.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH);
                if (lengthParam != null && lengthParam.HasValue)
                    result["length_mm"] = Math.Round(lengthParam.AsDouble() * 304.8, 0);

                var wallHeightParam = wall.get_Parameter(BuiltInParameter.WALL_USER_HEIGHT_PARAM);
                if (wallHeightParam != null && wallHeightParam.HasValue)
                    result["height_mm"] = Math.Round(wallHeightParam.AsDouble() * 304.8, 0);

                result["width_mm"] = Math.Round(wall.Width * 304.8, 0);
            }
            else
            {
                // Общий случай — семейство
                if (el is FamilyInstance fi)
                    result["family"] = fi.Symbol?.Family?.Name ?? "";

                // BoundingBox для координат
                var bb = el.get_BoundingBox(null);
                if (bb != null)
                {
                    var minPt = transform.OfPoint(bb.Min);
                    var maxPt = transform.OfPoint(bb.Max);
                    result["bbox_min_mm"] = new
                    {
                        x = Math.Round(minPt.X * 304.8, 1),
                        y = Math.Round(minPt.Y * 304.8, 1),
                        z = Math.Round(minPt.Z * 304.8, 1)
                    };
                    result["bbox_max_mm"] = new
                    {
                        x = Math.Round(maxPt.X * 304.8, 1),
                        y = Math.Round(maxPt.Y * 304.8, 1),
                        z = Math.Round(maxPt.Z * 304.8, 1)
                    };
                }
            }

            return result;
        }

        private static string[] GetAllLinkNames(Document doc)
        {
            return new FilteredElementCollector(doc)
                .OfClass(typeof(RevitLinkInstance))
                .Cast<RevitLinkInstance>()
                .Select(li => li.Name)
                .ToArray();
        }

        // ═══════════════════════════════════════════════════════════
        // Существующие режимы
        // ═══════════════════════════════════════════════════════════

        private static List<Element> GetByCategory(Document doc, string categoryName)
        {
            if (string.IsNullOrWhiteSpace(categoryName))
                return new List<Element>();

            // Сначала пробуем BuiltInCategory по ключевым словам
            var bic = GuessBuiltInCategory(categoryName);
            if (bic != BuiltInCategory.INVALID)
            {
                return new FilteredElementCollector(doc)
                    .OfCategory(bic)
                    .WhereElementIsNotElementType()
                    .ToList();
            }

            // Поиск категории по имени
            Category cat = null;
            foreach (Category c in doc.Settings.Categories)
            {
                if (c != null && c.Name.IndexOf(categoryName, StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    cat = c;
                    break;
                }
            }

            if (cat == null)
                return new List<Element>();

            return new FilteredElementCollector(doc)
                .OfCategoryId(cat.Id)
                .WhereElementIsNotElementType()
                .ToList();
        }

        private static List<Element> GetByView(UIApplication uiApp, string viewIdStr)
        {
            var doc = uiApp.ActiveUIDocument.Document;
            ElementId viewId;

            if (string.IsNullOrWhiteSpace(viewIdStr))
                viewId = uiApp.ActiveUIDocument.ActiveView.Id;
            else if (int.TryParse(viewIdStr, out int id))
                viewId = new ElementId(id);
            else
                return new List<Element>();

            return new FilteredElementCollector(doc, viewId)
                .WhereElementIsNotElementType()
                .ToList();
        }

        private static List<Element> GetBySelection(UIApplication uiApp)
        {
            var sel = uiApp.ActiveUIDocument?.Selection?.GetElementIds();
            if (sel == null || sel.Count == 0)
                return new List<Element>();

            var doc = uiApp.ActiveUIDocument.Document;
            return sel.Select(id => doc.GetElement(id)).Where(e => e != null).ToList();
        }

        private static List<Element> GetByIds(Document doc, string idsStr)
        {
            return ParamHelper.ParseIds(idsStr)
                .Select(id => doc.GetElement(id))
                .Where(e => e != null)
                .ToList();
        }

        private static List<Element> GetBySystem(Document doc, string systemName)
        {
            if (string.IsNullOrWhiteSpace(systemName))
                return new List<Element>();

            var system = new FilteredElementCollector(doc)
                .OfClass(typeof(MEPSystem))
                .Cast<MEPSystem>()
                .FirstOrDefault(s => s.Name.IndexOf(systemName, StringComparison.OrdinalIgnoreCase) >= 0);

            if (system?.Elements == null)
                return new List<Element>();

            return system.Elements.Cast<Element>().ToList();
        }

        /// <summary>Маппинг ключевых слов (RU/EN) на BuiltInCategory.</summary>
        private static BuiltInCategory GuessBuiltInCategory(string keyword)
        {
            keyword = keyword.ToLowerInvariant().Trim();

            var map = new Dictionary<string, BuiltInCategory>
            {
                // Воздуховоды
                ["ducts"] = BuiltInCategory.OST_DuctCurves,
                ["duct curves"] = BuiltInCategory.OST_DuctCurves,
                ["duct"] = BuiltInCategory.OST_DuctCurves,
                ["воздуховод"] = BuiltInCategory.OST_DuctCurves,

                // Фитинги воздуховодов
                ["duct fittings"] = BuiltInCategory.OST_DuctFitting,
                ["duct fitting"] = BuiltInCategory.OST_DuctFitting,
                ["фитинг"] = BuiltInCategory.OST_DuctFitting,
                ["соединительн"] = BuiltInCategory.OST_DuctFitting,

                // Арматура воздуховодов
                ["duct accessories"] = BuiltInCategory.OST_DuctAccessory,
                ["duct accessory"] = BuiltInCategory.OST_DuctAccessory,
                ["арматура воздуховод"] = BuiltInCategory.OST_DuctAccessory,

                // Воздухораспределители
                ["air terminals"] = BuiltInCategory.OST_DuctTerminal,
                ["air terminal"] = BuiltInCategory.OST_DuctTerminal,
                ["terminal"] = BuiltInCategory.OST_DuctTerminal,
                ["решётка"] = BuiltInCategory.OST_DuctTerminal,
                ["решетка"] = BuiltInCategory.OST_DuctTerminal,
                ["диффузор"] = BuiltInCategory.OST_DuctTerminal,
                ["воздухораспред"] = BuiltInCategory.OST_DuctTerminal,

                // Системы воздуховодов
                ["duct systems"] = BuiltInCategory.OST_DuctSystem,
                ["системы воздух"] = BuiltInCategory.OST_DuctSystem,

                // Оборудование
                ["mechanical equipment"] = BuiltInCategory.OST_MechanicalEquipment,
                ["equipment"] = BuiltInCategory.OST_MechanicalEquipment,
                ["оборудование"] = BuiltInCategory.OST_MechanicalEquipment,

                // Гибкие воздуховоды
                ["flex duct"] = BuiltInCategory.OST_FlexDuctCurves,
                ["flex"] = BuiltInCategory.OST_FlexDuctCurves,
                ["гибк"] = BuiltInCategory.OST_FlexDuctCurves,

                // Стены
                ["walls"] = BuiltInCategory.OST_Walls,
                ["wall"] = BuiltInCategory.OST_Walls,
                ["стен"] = BuiltInCategory.OST_Walls,

                // Помещения
                ["rooms"] = BuiltInCategory.OST_Rooms,
                ["room"] = BuiltInCategory.OST_Rooms,
                ["помещен"] = BuiltInCategory.OST_Rooms,
                ["комнат"] = BuiltInCategory.OST_Rooms,

                // Пространства MEP
                ["spaces"] = BuiltInCategory.OST_MEPSpaces,
                ["space"] = BuiltInCategory.OST_MEPSpaces,
                ["пространств"] = BuiltInCategory.OST_MEPSpaces,

                // Трубы
                ["pipes"] = BuiltInCategory.OST_PipeCurves,
                ["pipe"] = BuiltInCategory.OST_PipeCurves,
                ["труб"] = BuiltInCategory.OST_PipeCurves,

                // Фитинги труб
                ["pipe fittings"] = BuiltInCategory.OST_PipeFitting,
                ["pipe fitting"] = BuiltInCategory.OST_PipeFitting,

                // Двери
                ["doors"] = BuiltInCategory.OST_Doors,
                ["door"] = BuiltInCategory.OST_Doors,
                ["двер"] = BuiltInCategory.OST_Doors,

                // Окна
                ["windows"] = BuiltInCategory.OST_Windows,
                ["window"] = BuiltInCategory.OST_Windows,
                ["окн"] = BuiltInCategory.OST_Windows,
                ["окон"] = BuiltInCategory.OST_Windows,

                // Перекрытия
                ["floors"] = BuiltInCategory.OST_Floors,
                ["floor"] = BuiltInCategory.OST_Floors,
                ["перекрыт"] = BuiltInCategory.OST_Floors,

                // Потолки
                ["ceilings"] = BuiltInCategory.OST_Ceilings,
                ["ceiling"] = BuiltInCategory.OST_Ceilings,
                ["потолк"] = BuiltInCategory.OST_Ceilings,
                ["потолок"] = BuiltInCategory.OST_Ceilings,

                // Колонны
                ["columns"] = BuiltInCategory.OST_Columns,
                ["column"] = BuiltInCategory.OST_Columns,
                ["колонн"] = BuiltInCategory.OST_Columns,

                // Уровни
                ["levels"] = BuiltInCategory.OST_Levels,
                ["level"] = BuiltInCategory.OST_Levels,
                ["уровн"] = BuiltInCategory.OST_Levels,

                // Связанные файлы
                ["revit links"] = BuiltInCategory.OST_RvtLinks,
                ["links"] = BuiltInCategory.OST_RvtLinks,
                ["связ"] = BuiltInCategory.OST_RvtLinks,
            };

            // Точное совпадение
            if (map.TryGetValue(keyword, out var exact))
                return exact;

            // Частичное совпадение
            foreach (var kvp in map)
            {
                if (keyword.Contains(kvp.Key) || kvp.Key.Contains(keyword))
                    return kvp.Value;
            }

            return BuiltInCategory.INVALID;
        }
    }
}
