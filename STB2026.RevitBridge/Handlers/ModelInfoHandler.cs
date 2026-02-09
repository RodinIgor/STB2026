using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.UI;

namespace STB2026.RevitBridge.Handlers
{
    /// <summary>
    /// get_model_info — обзор проекта Revit.
    /// Claude вызывает первым для ориентации.
    /// </summary>
    internal static class ModelInfoHandler
    {
        public static object Handle(UIApplication uiApp, Dictionary<string, object> p)
        {
            var doc = uiApp.ActiveUIDocument?.Document;
            if (doc == null)
                return new { error = "Нет открытого документа в Revit" };

            var activeView = uiApp.ActiveUIDocument.ActiveView;

            // Категории с количеством элементов
            var categories = new List<Dictionary<string, object>>();
            foreach (Category cat in doc.Settings.Categories)
            {
                if (cat == null || !cat.AllowsBoundParameters) continue;
                var count = new FilteredElementCollector(doc)
                    .OfCategoryId(cat.Id)
                    .WhereElementIsNotElementType()
                    .GetElementCount();

                if (count > 0)
                {
                    categories.Add(new Dictionary<string, object>
                    {
                        ["name"] = cat.Name,
                        ["id"] = cat.Id.IntegerValue,
                        ["count"] = count
                    });
                }
            }

            // MEP системы
            var systems = new FilteredElementCollector(doc)
                .OfClass(typeof(MEPSystem))
                .Cast<MEPSystem>()
                .Select(s => new Dictionary<string, object>
                {
                    ["name"] = s.Name,
                    ["id"] = s.Id.IntegerValue,
                    ["type"] = GetMepSystemType(s),
                    ["elements_count"] = s.Elements?.Size ?? 0
                })
                .OrderBy(s => s["name"])
                .ToList();

            // Уровни
            var levels = new FilteredElementCollector(doc)
                .OfClass(typeof(Level))
                .Cast<Level>()
                .Select(l => new Dictionary<string, object>
                {
                    ["name"] = l.Name,
                    ["id"] = l.Id.IntegerValue,
                    ["elevation_m"] = System.Math.Round(
                        UnitUtils.ConvertFromInternalUnits(l.Elevation, UnitTypeId.Meters), 3)
                })
                .OrderBy(l => (double)l["elevation_m"])
                .ToList();

            // Виды
            var views = new FilteredElementCollector(doc)
                .OfClass(typeof(View))
                .Cast<View>()
                .Where(v => !v.IsTemplate)
                .Select(v => new Dictionary<string, object>
                {
                    ["name"] = v.Name,
                    ["id"] = v.Id.IntegerValue,
                    ["type"] = v.ViewType.ToString(),
                    ["is_active"] = v.Id == activeView.Id
                })
                .OrderBy(v => v["type"].ToString())
                .ThenBy(v => v["name"].ToString())
                .ToList();

            // Связанные файлы
            var links = new FilteredElementCollector(doc)
                .OfClass(typeof(RevitLinkInstance))
                .Select(l => new Dictionary<string, object>
                {
                    ["name"] = l.Name,
                    ["id"] = l.Id.IntegerValue
                })
                .ToList();

            return new
            {
                project_name = doc.Title,
                file_path = doc.PathName,
                active_view = new
                {
                    name = activeView.Name,
                    id = activeView.Id.IntegerValue,
                    type = activeView.ViewType.ToString()
                },
                units = doc.GetUnits().GetFormatOptions(SpecTypeId.Length).GetUnitTypeId().TypeId,
                categories = categories.OrderByDescending(c => (int)c["count"]).Take(30).ToList(),
                mep_systems = systems,
                levels,
                views = views.Take(50).ToList(),
                linked_files = links,
                summary = new
                {
                    total_categories = categories.Count,
                    total_systems = systems.Count,
                    total_levels = levels.Count,
                    total_views = views.Count,
                    total_links = links.Count
                }
            };
        }

        /// <summary>Получить тип MEP системы (Supply/Return/Exhaust/Other).</summary>
        private static string GetMepSystemType(MEPSystem sys)
        {
            if (sys is MechanicalSystem ms)
                return ms.SystemType.ToString();

            // Для остальных (Piping, Electrical) — по категории
            return sys.Category?.Name ?? sys.GetType().Name;
        }
    }
}
