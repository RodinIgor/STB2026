using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Newtonsoft.Json.Linq;

namespace STB2026.RevitBridge.Handlers
{
    /// <summary>
    /// manage_project — управление параметрами проекта.
    ///
    /// Действия:
    ///   get_project_info     — общая информация о проекте
    ///   set_project_info     — изменить информацию о проекте
    ///   list_phases          — список стадий
    ///   list_worksets         — список рабочих наборов
    ///   list_design_options  — список вариантов конструкции
    ///   list_links           — связанные файлы (с трансформацией)
    ///   reload_link          — перезагрузить связь
    ///   list_warnings        — предупреждения модели
    ///   get_units            — текущие единицы проекта
    ///   list_materials       — список материалов
    ///   list_line_styles     — список стилей линий
    ///   list_fill_patterns   — список штриховок
    ///   purge_unused         — анализ неиспользуемых элементов (без удаления)
    ///   sync_to_central      — синхронизация с центральной моделью
    /// </summary>
    internal static class ProjectHandler
    {
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
                case "get_project_info":    return GetProjectInfo(doc);
                case "set_project_info":    return SetProjectInfo(doc, data);
                case "list_phases":         return ListPhases(doc);
                case "list_worksets":        return ListWorksets(doc);
                case "list_design_options":  return ListDesignOptions(doc);
                case "list_links":          return ListLinks(doc);
                case "reload_link":         return ReloadLink(doc, data);
                case "list_warnings":       return ListWarnings(doc);
                case "get_units":           return GetUnits(doc);
                case "list_materials":      return ListMaterials(doc);
                case "list_line_styles":    return ListLineStyles(doc);
                case "list_fill_patterns":  return ListFillPatterns(doc);
                case "purge_unused":        return PurgeAnalysis(doc);
                default:
                    return new
                    {
                        error = $"Неизвестное действие: '{action}'",
                        available = new[]
                        {
                            "get_project_info", "set_project_info",
                            "list_phases", "list_worksets", "list_design_options",
                            "list_links", "reload_link", "list_warnings",
                            "get_units", "list_materials", "list_line_styles",
                            "list_fill_patterns", "purge_unused"
                        }
                    };
            }
        }

        // ═══ get_project_info ═══
        private static object GetProjectInfo(Document doc)
        {
            var pi = doc.ProjectInformation;
            return new
            {
                action = "get_project_info",
                name = pi.Name,
                number = pi.Number,
                address = pi.Address,
                author = pi.Author,
                building_name = pi.BuildingName,
                client_name = pi.ClientName,
                status = pi.Status,
                issue_date = pi.IssueDate,
                organization_name = pi.OrganizationName,
                organization_description = pi.OrganizationDescription,
                file_path = doc.PathName,
                is_workshared = doc.IsWorkshared,
                is_central = doc.IsWorkshared && !doc.IsDetached
            };
        }

        // ═══ set_project_info ═══
        private static object SetProjectInfo(Document doc, JObject data)
        {
            var pi = doc.ProjectInformation;
            var changed = new List<string>();

            using (var trans = new Transaction(doc, "STB2026: проектная информация"))
            {
                trans.Start();

                void TrySet(string key, Action<string> setter)
                {
                    var val = data.Value<string>(key);
                    if (val != null) { setter(val); changed.Add(key); }
                }

                TrySet("name", v => pi.Name = v);
                TrySet("number", v => pi.Number = v);
                TrySet("address", v => pi.Address = v);
                TrySet("author", v => pi.Author = v);
                TrySet("building_name", v => pi.BuildingName = v);
                TrySet("client_name", v => pi.ClientName = v);
                TrySet("status", v => pi.Status = v);
                TrySet("issue_date", v => pi.IssueDate = v);
                TrySet("organization_name", v => pi.OrganizationName = v);
                TrySet("organization_description", v => pi.OrganizationDescription = v);

                trans.Commit();
            }

            return new { action = "set_project_info", changed };
        }

        // ═══ list_phases ═══
        private static object ListPhases(Document doc)
        {
            var phases = doc.Phases.Cast<Phase>()
                .Select(ph => new
                {
                    id = ph.Id.IntegerValue,
                    name = ph.Name
                }).ToList();

            return new { action = "list_phases", count = phases.Count, phases };
        }

        // ═══ list_worksets ═══
        private static object ListWorksets(Document doc)
        {
            if (!doc.IsWorkshared)
                return new { action = "list_worksets", is_workshared = false, message = "Проект не совместный" };

            var worksets = new FilteredWorksetCollector(doc)
                .OfKind(WorksetKind.UserWorkset)
                .Select(ws => new
                {
                    id = ws.Id.IntegerValue,
                    name = ws.Name,
                    is_open = ws.IsOpen,
                    is_default = ws.IsDefaultWorkset,
                    is_visible = ws.IsVisibleByDefault,
                    owner = ws.Owner
                }).ToList();

            return new { action = "list_worksets", is_workshared = true, count = worksets.Count, worksets };
        }

        // ═══ list_design_options ═══
        private static object ListDesignOptions(Document doc)
        {
            var options = new FilteredElementCollector(doc)
                .OfClass(typeof(DesignOption))
                .Cast<DesignOption>()
                .Select(opt => new
                {
                    id = opt.Id.IntegerValue,
                    name = opt.Name,
                    is_primary = opt.IsPrimary
                }).ToList();

            return new { action = "list_design_options", count = options.Count, options };
        }

        // ═══ list_links ═══
        private static object ListLinks(Document doc)
        {
            var links = new FilteredElementCollector(doc)
                .OfClass(typeof(RevitLinkInstance))
                .Cast<RevitLinkInstance>()
                .Select(li =>
                {
                    var transform = li.GetTotalTransform();
                    var linkDoc = li.GetLinkDocument();
                    return new
                    {
                        id = li.Id.IntegerValue,
                        name = li.Name,
                        is_loaded = linkDoc != null,
                        file_path = linkDoc?.PathName ?? "",
                        origin_x_mm = Math.Round(transform.Origin.X * 304.8, 1),
                        origin_y_mm = Math.Round(transform.Origin.Y * 304.8, 1),
                        origin_z_mm = Math.Round(transform.Origin.Z * 304.8, 1),
                        link_type_id = li.GetTypeId().IntegerValue
                    };
                }).ToList();

            return new { action = "list_links", count = links.Count, links };
        }

        // ═══ reload_link ═══
        private static object ReloadLink(Document doc, JObject data)
        {
            int linkTypeId = data.Value<int?>("link_type_id") ?? -1;
            if (linkTypeId <= 0) return new { error = "Укажите link_type_id" };

            var linkType = doc.GetElement(new ElementId(linkTypeId)) as RevitLinkType;
            if (linkType == null) return new { error = $"RevitLinkType id={linkTypeId} не найден" };

            linkType.Reload();
            return new { action = "reload_link", link_type_id = linkTypeId, name = linkType.Name };
        }

        // ═══ list_warnings ═══
        private static object ListWarnings(Document doc)
        {
            var warnings = doc.GetWarnings()
                .Select(w =>
                {
                    var ids = w.GetFailingElements().Select(id => id.IntegerValue).ToList();
                    var addIds = w.GetAdditionalElements().Select(id => id.IntegerValue).ToList();
                    return new
                    {
                        severity = w.GetSeverity().ToString(),
                        description = w.GetDescriptionText(),
                        failing_elements = ids,
                        additional_elements = addIds
                    };
                }).ToList();

            // Группировка по типу предупреждения
            var grouped = warnings.GroupBy(w => w.description)
                .Select(g => new
                {
                    message = g.Key,
                    count = g.Count(),
                    first_elements = g.First().failing_elements
                })
                .OrderByDescending(g => g.count)
                .ToList();

            return new
            {
                action = "list_warnings",
                total = warnings.Count,
                unique_types = grouped.Count,
                by_type = grouped.Take(30).ToList()
            };
        }

        // ═══ get_units ═══
        private static object GetUnits(Document doc)
        {
            var units = doc.GetUnits();
            var specs = new[]
            {
                (name: "length", spec: SpecTypeId.Length),
                (name: "area", spec: SpecTypeId.Area),
                (name: "volume", spec: SpecTypeId.Volume),
                (name: "angle", spec: SpecTypeId.Angle),
                (name: "hvac_velocity", spec: SpecTypeId.HvacVelocity),
                (name: "hvac_airflow", spec: SpecTypeId.AirFlow),
                (name: "hvac_pressure", spec: SpecTypeId.HvacPressure),
                (name: "pipe_size", spec: SpecTypeId.PipeSize),
                (name: "duct_size", spec: SpecTypeId.DuctSize)
            };

            var result = new Dictionary<string, object>();
            foreach (var (name, spec) in specs)
            {
                try
                {
                    var fo = units.GetFormatOptions(spec);
                    result[name] = new
                    {
                        unit = fo.GetUnitTypeId().TypeId,
                        accuracy = fo.Accuracy
                    };
                }
                catch { result[name] = "N/A"; }
            }

            return new { action = "get_units", units = result };
        }

        // ═══ list_materials ═══
        private static object ListMaterials(Document doc)
        {
            var materials = new FilteredElementCollector(doc)
                .OfClass(typeof(Material))
                .Cast<Material>()
                .Select(m => new
                {
                    id = m.Id.IntegerValue,
                    name = m.Name,
                    color = m.Color.IsValid ? $"#{m.Color.Red:X2}{m.Color.Green:X2}{m.Color.Blue:X2}" : "",
                    transparency = m.Transparency
                })
                .OrderBy(m => m.name)
                .ToList();

            return new { action = "list_materials", count = materials.Count, materials = materials.Take(100).ToList() };
        }

        // ═══ list_line_styles ═══
        private static object ListLineStyles(Document doc)
        {
            var lineStylesCat = doc.Settings.Categories.get_Item(BuiltInCategory.OST_Lines);
            var styles = new List<object>();

            if (lineStylesCat?.SubCategories != null)
            {
                foreach (Category sub in lineStylesCat.SubCategories)
                {
                    styles.Add(new
                    {
                        id = sub.Id.IntegerValue,
                        name = sub.Name,
                        color = sub.LineColor.IsValid
                            ? $"#{sub.LineColor.Red:X2}{sub.LineColor.Green:X2}{sub.LineColor.Blue:X2}"
                            : ""
                    });
                }
            }

            return new { action = "list_line_styles", count = styles.Count, styles = styles.OrderBy(s => ((dynamic)s).name).ToList() };
        }

        // ═══ list_fill_patterns ═══
        private static object ListFillPatterns(Document doc)
        {
            var patterns = new FilteredElementCollector(doc)
                .OfClass(typeof(FillPatternElement))
                .Cast<FillPatternElement>()
                .Select(fp =>
                {
                    var pat = fp.GetFillPattern();
                    return new
                    {
                        id = fp.Id.IntegerValue,
                        name = fp.Name,
                        target = pat.Target.ToString(),
                        is_solid = pat.IsSolidFill
                    };
                })
                .OrderBy(p => p.name)
                .ToList();

            return new { action = "list_fill_patterns", count = patterns.Count, patterns };
        }

        // ═══ purge_unused ═══
        private static object PurgeAnalysis(Document doc)
        {
            // Анализ неиспользуемых типов (без удаления!)
            var unusedFamilies = new List<object>();

            var allSymbols = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilySymbol))
                .Cast<FamilySymbol>()
                .ToList();

            foreach (var symbol in allSymbols)
            {
                var instances = new FilteredElementCollector(doc)
                    .OfClass(typeof(FamilyInstance))
                    .Cast<FamilyInstance>()
                    .Count(fi => fi.Symbol.Id == symbol.Id);

                if (instances == 0)
                {
                    unusedFamilies.Add(new
                    {
                        id = symbol.Id.IntegerValue,
                        family = symbol.Family.Name,
                        type = symbol.Name,
                        category = symbol.Category?.Name ?? ""
                    });
                }
            }

            return new
            {
                action = "purge_unused",
                note = "Только анализ, ничего не удалено",
                unused_family_types = unusedFamilies.Count,
                types = unusedFamilies.Take(50).ToList()
            };
        }
    }
}
