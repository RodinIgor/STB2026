using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Newtonsoft.Json.Linq;

namespace STB2026.RevitBridge.Handlers
{
    /// <summary>
    /// manage_views — управление видами Revit.
    /// 
    /// Действия:
    ///   list_views          — список видов (фильтр: type, prefix)
    ///   create_plan          — создать план этажа (level_id, view_family_type_id)
    ///   create_section       — создать разрез (min/max BoundingBox)
    ///   create_3d            — создать 3D вид
    ///   create_legend        — создать легенду
    ///   duplicate_view       — дублировать вид (Duplicate/WithDetailing/AsDependent)
    ///   delete_view          — удалить вид
    ///   rename_view          — переименовать вид
    ///   set_active           — сделать вид активным
    ///   set_scale            — установить масштаб
    ///   set_template         — назначить шаблон вида
    ///   remove_template      — снять шаблон вида
    ///   list_templates       — список шаблонов видов
    ///   set_crop_box         — установить границы обрезки
    ///   toggle_crop          — вкл/выкл обрезку
    ///   set_detail_level     — уровень детализации (coarse/medium/fine)
    ///   set_display_style    — стиль отображения (wireframe/hidden/shaded/realistic)
    ///   get_view_filters     — список фильтров на виде
    ///   add_view_filter      — добавить фильтр
    ///   remove_view_filter   — удалить фильтр
    ///   set_category_visibility — скрыть/показать категорию
    /// </summary>
    internal static class ViewHandler
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
                case "list_views":        return ListViews(doc, data);
                case "create_plan":       return CreatePlan(doc, data);
                case "create_section":    return CreateSection(doc, data);
                case "create_3d":         return Create3D(doc, data);
                case "duplicate_view":    return DuplicateView(doc, data);
                case "delete_view":       return DeleteView(doc, data);
                case "rename_view":       return RenameView(doc, data);
                case "set_active":        return SetActiveView(uidoc, data);
                case "set_scale":         return SetScale(doc, data);
                case "set_template":      return SetTemplate(doc, data);
                case "remove_template":   return RemoveTemplate(doc, data);
                case "list_templates":    return ListTemplates(doc, data);
                case "set_crop_box":      return SetCropBox(doc, data);
                case "toggle_crop":       return ToggleCrop(doc, data);
                case "set_detail_level":  return SetDetailLevel(doc, data);
                case "set_display_style": return SetDisplayStyle(doc, data);
                case "get_view_filters":  return GetViewFilters(doc, data);
                case "add_view_filter":   return AddViewFilter(doc, data);
                case "remove_view_filter": return RemoveViewFilter(doc, data);
                case "set_category_visibility": return SetCategoryVisibility(doc, data);
                default:
                    return new
                    {
                        error = $"Неизвестное действие: '{action}'",
                        available = new[]
                        {
                            "list_views", "create_plan", "create_section", "create_3d",
                            "duplicate_view", "delete_view", "rename_view", "set_active",
                            "set_scale", "set_template", "remove_template", "list_templates",
                            "set_crop_box", "toggle_crop", "set_detail_level", "set_display_style",
                            "get_view_filters", "add_view_filter", "remove_view_filter",
                            "set_category_visibility"
                        }
                    };
            }
        }

        // ═══ list_views ═══
        private static object ListViews(Document doc, JObject data)
        {
            string typeFilter = data.Value<string>("type") ?? "";
            string prefix = data.Value<string>("prefix") ?? "";

            var views = new FilteredElementCollector(doc)
                .OfClass(typeof(View))
                .Cast<View>()
                .Where(v => !v.IsTemplate);

            if (!string.IsNullOrWhiteSpace(typeFilter))
                views = views.Where(v => v.ViewType.ToString().Equals(typeFilter, StringComparison.OrdinalIgnoreCase));

            if (!string.IsNullOrWhiteSpace(prefix))
                views = views.Where(v => v.Name.StartsWith(prefix, StringComparison.OrdinalIgnoreCase));

            var result = views.Select(v => new
            {
                id = v.Id.IntegerValue,
                name = v.Name,
                type = v.ViewType.ToString(),
                scale = v.Scale,
                template = v.ViewTemplateId != ElementId.InvalidElementId
                    ? doc.GetElement(v.ViewTemplateId)?.Name ?? ""
                    : "",
                is_on_sheet = v is View vw && IsOnSheet(doc, vw.Id),
                detail_level = v.DetailLevel.ToString()
            }).OrderBy(v => v.type).ThenBy(v => v.name).ToList();

            return new { action = "list_views", count = result.Count, views = result };
        }

        // ═══ create_plan ═══
        private static object CreatePlan(Document doc, JObject data)
        {
            int levelId = data.Value<int?>("level_id") ?? -1;
            string name = data.Value<string>("name") ?? "";
            int vftId = data.Value<int?>("view_family_type_id") ?? -1;

            if (levelId <= 0) return new { error = "Укажите level_id" };

            var level = doc.GetElement(new ElementId(levelId)) as Level;
            if (level == null) return new { error = $"Уровень id={levelId} не найден" };

            // Если не указан тип — берём первый FloorPlan
            ElementId viewFamilyTypeId;
            if (vftId > 0)
                viewFamilyTypeId = new ElementId(vftId);
            else
            {
                viewFamilyTypeId = new FilteredElementCollector(doc)
                    .OfClass(typeof(ViewFamilyType))
                    .Cast<ViewFamilyType>()
                    .FirstOrDefault(vft => vft.ViewFamily == ViewFamily.FloorPlan)?.Id
                    ?? ElementId.InvalidElementId;
            }

            if (viewFamilyTypeId == ElementId.InvalidElementId)
                return new { error = "Не найден тип семейства вида для плана" };

            ViewPlan view;
            using (var trans = new Transaction(doc, "STB2026: создание плана"))
            {
                trans.Start();
                view = ViewPlan.Create(doc, viewFamilyTypeId, level.Id);
                if (!string.IsNullOrWhiteSpace(name))
                    view.Name = name;
                trans.Commit();
            }

            return new { action = "create_plan", view_id = view.Id.IntegerValue, name = view.Name };
        }

        // ═══ create_section ═══
        private static object CreateSection(Document doc, JObject data)
        {
            double minX = data.Value<double?>("min_x_mm") ?? 0;
            double minY = data.Value<double?>("min_y_mm") ?? 0;
            double minZ = data.Value<double?>("min_z_mm") ?? 0;
            double maxX = data.Value<double?>("max_x_mm") ?? 1000;
            double maxY = data.Value<double?>("max_y_mm") ?? 1000;
            double maxZ = data.Value<double?>("max_z_mm") ?? 3000;
            string name = data.Value<string>("name") ?? "";

            var vft = new FilteredElementCollector(doc)
                .OfClass(typeof(ViewFamilyType))
                .Cast<ViewFamilyType>()
                .FirstOrDefault(v => v.ViewFamily == ViewFamily.Section);

            if (vft == null) return new { error = "Не найден тип для разреза" };

            var min = new XYZ(minX / 304.8, minY / 304.8, minZ / 304.8);
            var max = new XYZ(maxX / 304.8, maxY / 304.8, maxZ / 304.8);
            var midPt = (min + max) / 2;

            // Направление разреза — по X
            var direction = XYZ.BasisX;
            var up = XYZ.BasisZ;
            var viewDir = direction.CrossProduct(up);

            var transform = Transform.Identity;
            transform.Origin = midPt;
            transform.BasisX = direction;
            transform.BasisY = up;
            transform.BasisZ = viewDir;

            var sectionBox = new BoundingBoxXYZ
            {
                Transform = transform,
                Min = new XYZ(-(max.X - min.X) / 2, -(max.Z - min.Z) / 2, 0),
                Max = new XYZ((max.X - min.X) / 2, (max.Z - min.Z) / 2, (max.Y - min.Y) / 2)
            };

            ViewSection view;
            using (var trans = new Transaction(doc, "STB2026: создание разреза"))
            {
                trans.Start();
                view = ViewSection.CreateSection(doc, vft.Id, sectionBox);
                if (!string.IsNullOrWhiteSpace(name)) view.Name = name;
                trans.Commit();
            }

            return new { action = "create_section", view_id = view.Id.IntegerValue, name = view.Name };
        }

        // ═══ create_3d ═══
        private static object Create3D(Document doc, JObject data)
        {
            string name = data.Value<string>("name") ?? "";
            bool isDefault = data.Value<bool?>("is_default") ?? true;

            var vft = new FilteredElementCollector(doc)
                .OfClass(typeof(ViewFamilyType))
                .Cast<ViewFamilyType>()
                .FirstOrDefault(v => v.ViewFamily == ViewFamily.ThreeDimensional);

            if (vft == null) return new { error = "Не найден тип для 3D вида" };

            View3D view;
            using (var trans = new Transaction(doc, "STB2026: 3D вид"))
            {
                trans.Start();
                view = isDefault
                    ? View3D.CreateIsometric(doc, vft.Id)
                    : View3D.CreatePerspective(doc, vft.Id);
                if (!string.IsNullOrWhiteSpace(name)) view.Name = name;
                trans.Commit();
            }

            return new { action = "create_3d", view_id = view.Id.IntegerValue, name = view.Name };
        }

        // ═══ duplicate_view ═══
        private static object DuplicateView(Document doc, JObject data)
        {
            int viewId = data.Value<int?>("view_id") ?? -1;
            string mode = data.Value<string>("mode") ?? "WithDetailing";
            string newName = data.Value<string>("new_name") ?? "";

            if (viewId <= 0) return new { error = "Укажите view_id" };
            var view = doc.GetElement(new ElementId(viewId)) as View;
            if (view == null) return new { error = $"Вид id={viewId} не найден" };

            if (!Enum.TryParse<ViewDuplicateOption>(mode, true, out var dupOption))
                dupOption = ViewDuplicateOption.WithDetailing;

            ElementId newId;
            using (var trans = new Transaction(doc, "STB2026: дублирование вида"))
            {
                trans.Start();
                newId = view.Duplicate(dupOption);
                if (!string.IsNullOrWhiteSpace(newName))
                {
                    var newView = doc.GetElement(newId) as View;
                    if (newView != null) newView.Name = newName;
                }
                trans.Commit();
            }

            var result = doc.GetElement(newId) as View;
            return new { action = "duplicate_view", new_view_id = newId.IntegerValue, name = result?.Name ?? "" };
        }

        // ═══ delete_view ═══
        private static object DeleteView(Document doc, JObject data)
        {
            int viewId = data.Value<int?>("view_id") ?? -1;
            if (viewId <= 0) return new { error = "Укажите view_id" };

            using (var trans = new Transaction(doc, "STB2026: удаление вида"))
            {
                trans.Start();
                doc.Delete(new ElementId(viewId));
                trans.Commit();
            }

            return new { action = "delete_view", deleted_id = viewId };
        }

        // ═══ rename_view ═══
        private static object RenameView(Document doc, JObject data)
        {
            int viewId = data.Value<int?>("view_id") ?? -1;
            string newName = data.Value<string>("new_name") ?? "";

            if (viewId <= 0) return new { error = "Укажите view_id" };
            if (string.IsNullOrWhiteSpace(newName)) return new { error = "Укажите new_name" };

            var view = doc.GetElement(new ElementId(viewId)) as View;
            if (view == null) return new { error = $"Вид id={viewId} не найден" };

            string oldName = view.Name;
            using (var trans = new Transaction(doc, "STB2026: переименование вида"))
            {
                trans.Start();
                view.Name = newName;
                trans.Commit();
            }

            return new { action = "rename_view", view_id = viewId, old_name = oldName, new_name = newName };
        }

        // ═══ set_active ═══
        private static object SetActiveView(UIDocument uidoc, JObject data)
        {
            int viewId = data.Value<int?>("view_id") ?? -1;
            if (viewId <= 0) return new { error = "Укажите view_id" };

            var view = uidoc.Document.GetElement(new ElementId(viewId)) as View;
            if (view == null) return new { error = $"Вид id={viewId} не найден" };

            uidoc.ActiveView = view;
            return new { action = "set_active", view_id = viewId, name = view.Name };
        }

        // ═══ set_scale ═══
        private static object SetScale(Document doc, JObject data)
        {
            int viewId = data.Value<int?>("view_id") ?? -1;
            int scale = data.Value<int?>("scale") ?? 100;
            if (viewId <= 0) return new { error = "Укажите view_id" };

            var view = doc.GetElement(new ElementId(viewId)) as View;
            if (view == null) return new { error = $"Вид id={viewId} не найден" };

            using (var trans = new Transaction(doc, "STB2026: масштаб вида"))
            {
                trans.Start();
                view.Scale = scale;
                trans.Commit();
            }

            return new { action = "set_scale", view_id = viewId, scale };
        }

        // ═══ set_template ═══
        private static object SetTemplate(Document doc, JObject data)
        {
            int viewId = data.Value<int?>("view_id") ?? -1;
            string templateName = data.Value<string>("template_name") ?? "";
            int templateId = data.Value<int?>("template_id") ?? -1;

            if (viewId <= 0) return new { error = "Укажите view_id" };
            var view = doc.GetElement(new ElementId(viewId)) as View;
            if (view == null) return new { error = $"Вид id={viewId} не найден" };

            View template = null;
            if (templateId > 0)
                template = doc.GetElement(new ElementId(templateId)) as View;
            else if (!string.IsNullOrWhiteSpace(templateName))
                template = new FilteredElementCollector(doc)
                    .OfClass(typeof(View))
                    .Cast<View>()
                    .FirstOrDefault(v => v.IsTemplate && v.Name.Equals(templateName, StringComparison.OrdinalIgnoreCase));

            if (template == null) return new { error = "Шаблон не найден" };

            using (var trans = new Transaction(doc, "STB2026: шаблон вида"))
            {
                trans.Start();
                view.ViewTemplateId = template.Id;
                trans.Commit();
            }

            return new { action = "set_template", view_id = viewId, template_name = template.Name };
        }

        // ═══ remove_template ═══
        private static object RemoveTemplate(Document doc, JObject data)
        {
            int viewId = data.Value<int?>("view_id") ?? -1;
            if (viewId <= 0) return new { error = "Укажите view_id" };

            var view = doc.GetElement(new ElementId(viewId)) as View;
            if (view == null) return new { error = $"Вид id={viewId} не найден" };

            using (var trans = new Transaction(doc, "STB2026: снять шаблон"))
            {
                trans.Start();
                view.ViewTemplateId = ElementId.InvalidElementId;
                trans.Commit();
            }

            return new { action = "remove_template", view_id = viewId };
        }

        // ═══ list_templates ═══
        private static object ListTemplates(Document doc, JObject data)
        {
            string typeFilter = data.Value<string>("type") ?? "";

            var templates = new FilteredElementCollector(doc)
                .OfClass(typeof(View))
                .Cast<View>()
                .Where(v => v.IsTemplate);

            if (!string.IsNullOrWhiteSpace(typeFilter))
                templates = templates.Where(v => v.ViewType.ToString().Equals(typeFilter, StringComparison.OrdinalIgnoreCase));

            var result = templates.Select(v => new
            {
                id = v.Id.IntegerValue,
                name = v.Name,
                type = v.ViewType.ToString()
            }).OrderBy(v => v.name).ToList();

            return new { action = "list_templates", count = result.Count, templates = result };
        }

        // ═══ set_crop_box ═══
        private static object SetCropBox(Document doc, JObject data)
        {
            int viewId = data.Value<int?>("view_id") ?? -1;
            if (viewId <= 0) return new { error = "Укажите view_id" };

            var view = doc.GetElement(new ElementId(viewId)) as View;
            if (view == null) return new { error = $"Вид id={viewId} не найден" };

            double minX = data.Value<double?>("min_x_mm") ?? 0;
            double minY = data.Value<double?>("min_y_mm") ?? 0;
            double minZ = data.Value<double?>("min_z_mm") ?? 0;
            double maxX = data.Value<double?>("max_x_mm") ?? 0;
            double maxY = data.Value<double?>("max_y_mm") ?? 0;
            double maxZ = data.Value<double?>("max_z_mm") ?? 0;

            using (var trans = new Transaction(doc, "STB2026: crop box"))
            {
                trans.Start();
                view.CropBoxActive = true;
                var bb = new BoundingBoxXYZ
                {
                    Min = new XYZ(minX / 304.8, minY / 304.8, minZ / 304.8),
                    Max = new XYZ(maxX / 304.8, maxY / 304.8, maxZ / 304.8)
                };
                view.CropBox = bb;
                trans.Commit();
            }

            return new { action = "set_crop_box", view_id = viewId };
        }

        // ═══ toggle_crop ═══
        private static object ToggleCrop(Document doc, JObject data)
        {
            int viewId = data.Value<int?>("view_id") ?? -1;
            bool enable = data.Value<bool?>("enable") ?? true;

            if (viewId <= 0) return new { error = "Укажите view_id" };
            var view = doc.GetElement(new ElementId(viewId)) as View;
            if (view == null) return new { error = $"Вид id={viewId} не найден" };

            using (var trans = new Transaction(doc, "STB2026: toggle crop"))
            {
                trans.Start();
                view.CropBoxActive = enable;
                trans.Commit();
            }

            return new { action = "toggle_crop", view_id = viewId, crop_active = enable };
        }

        // ═══ set_detail_level ═══
        private static object SetDetailLevel(Document doc, JObject data)
        {
            int viewId = data.Value<int?>("view_id") ?? -1;
            string level = data.Value<string>("level") ?? "medium";

            if (viewId <= 0) return new { error = "Укажите view_id" };
            var view = doc.GetElement(new ElementId(viewId)) as View;
            if (view == null) return new { error = $"Вид id={viewId} не найден" };

            var dl = level.ToLowerInvariant() switch
            {
                "coarse" => ViewDetailLevel.Coarse,
                "fine" => ViewDetailLevel.Fine,
                _ => ViewDetailLevel.Medium
            };

            using (var trans = new Transaction(doc, "STB2026: detail level"))
            {
                trans.Start();
                view.DetailLevel = dl;
                trans.Commit();
            }

            return new { action = "set_detail_level", view_id = viewId, level = dl.ToString() };
        }

        // ═══ set_display_style ═══
        private static object SetDisplayStyle(Document doc, JObject data)
        {
            int viewId = data.Value<int?>("view_id") ?? -1;
            string style = data.Value<string>("style") ?? "hidden";

            if (viewId <= 0) return new { error = "Укажите view_id" };
            var view = doc.GetElement(new ElementId(viewId)) as View;
            if (view == null) return new { error = $"Вид id={viewId} не найден" };

            var ds = style.ToLowerInvariant() switch
            {
                "wireframe" => DisplayStyle.Wireframe,
                "shaded" => DisplayStyle.Shading,
                "shadedwithedges" or "shaded_edges" => DisplayStyle.ShadingWithEdges,
                "realistic" => DisplayStyle.Realistic,
                _ => DisplayStyle.HLR // Hidden Line
            };

            using (var trans = new Transaction(doc, "STB2026: display style"))
            {
                trans.Start();
                view.DisplayStyle = ds;
                trans.Commit();
            }

            return new { action = "set_display_style", view_id = viewId, style = ds.ToString() };
        }

        // ═══ get_view_filters ═══
        private static object GetViewFilters(Document doc, JObject data)
        {
            int viewId = data.Value<int?>("view_id") ?? -1;
            if (viewId <= 0) return new { error = "Укажите view_id" };

            var view = doc.GetElement(new ElementId(viewId)) as View;
            if (view == null) return new { error = $"Вид id={viewId} не найден" };

            var filterIds = view.GetFilters();
            var filters = filterIds.Select(fid =>
            {
                var filter = doc.GetElement(fid);
                var isVisible = view.GetFilterVisibility(fid);
                return new
                {
                    filter_id = fid.IntegerValue,
                    name = filter?.Name ?? "",
                    is_visible = isVisible
                };
            }).ToList();

            return new { action = "get_view_filters", view_id = viewId, count = filters.Count, filters };
        }

        // ═══ add_view_filter ═══
        private static object AddViewFilter(Document doc, JObject data)
        {
            int viewId = data.Value<int?>("view_id") ?? -1;
            int filterId = data.Value<int?>("filter_id") ?? -1;
            string filterName = data.Value<string>("filter_name") ?? "";

            if (viewId <= 0) return new { error = "Укажите view_id" };
            var view = doc.GetElement(new ElementId(viewId)) as View;
            if (view == null) return new { error = $"Вид id={viewId} не найден" };

            ElementId fid;
            if (filterId > 0)
                fid = new ElementId(filterId);
            else if (!string.IsNullOrWhiteSpace(filterName))
            {
                var found = new FilteredElementCollector(doc)
                    .OfClass(typeof(ParameterFilterElement))
                    .FirstOrDefault(f => f.Name.Equals(filterName, StringComparison.OrdinalIgnoreCase));
                if (found == null) return new { error = $"Фильтр '{filterName}' не найден" };
                fid = found.Id;
            }
            else return new { error = "Укажите filter_id или filter_name" };

            using (var trans = new Transaction(doc, "STB2026: add view filter"))
            {
                trans.Start();
                view.AddFilter(fid);
                trans.Commit();
            }

            return new { action = "add_view_filter", view_id = viewId, filter_id = fid.IntegerValue };
        }

        // ═══ remove_view_filter ═══
        private static object RemoveViewFilter(Document doc, JObject data)
        {
            int viewId = data.Value<int?>("view_id") ?? -1;
            int filterId = data.Value<int?>("filter_id") ?? -1;

            if (viewId <= 0) return new { error = "Укажите view_id" };
            if (filterId <= 0) return new { error = "Укажите filter_id" };

            var view = doc.GetElement(new ElementId(viewId)) as View;
            if (view == null) return new { error = $"Вид id={viewId} не найден" };

            using (var trans = new Transaction(doc, "STB2026: remove view filter"))
            {
                trans.Start();
                view.RemoveFilter(new ElementId(filterId));
                trans.Commit();
            }

            return new { action = "remove_view_filter", view_id = viewId, filter_id = filterId };
        }

        // ═══ set_category_visibility ═══
        private static object SetCategoryVisibility(Document doc, JObject data)
        {
            int viewId = data.Value<int?>("view_id") ?? -1;
            string categoryName = data.Value<string>("category") ?? "";
            bool visible = data.Value<bool?>("visible") ?? true;

            if (viewId <= 0) return new { error = "Укажите view_id" };
            if (string.IsNullOrWhiteSpace(categoryName)) return new { error = "Укажите category" };

            var view = doc.GetElement(new ElementId(viewId)) as View;
            if (view == null) return new { error = $"Вид id={viewId} не найден" };

            // Ищем категорию
            Category cat = null;
            foreach (Category c in doc.Settings.Categories)
            {
                if (c.Name.Equals(categoryName, StringComparison.OrdinalIgnoreCase))
                {
                    cat = c;
                    break;
                }
            }
            if (cat == null) return new { error = $"Категория '{categoryName}' не найдена" };

            using (var trans = new Transaction(doc, "STB2026: category visibility"))
            {
                trans.Start();
                view.SetCategoryHidden(cat.Id, !visible);
                trans.Commit();
            }

            return new { action = "set_category_visibility", view_id = viewId, category = categoryName, visible };
        }

        // ═══ Helpers ═══
        private static bool IsOnSheet(Document doc, ElementId viewId)
        {
            return new FilteredElementCollector(doc)
                .OfClass(typeof(Viewport))
                .Cast<Viewport>()
                .Any(vp => vp.ViewId == viewId);
        }
    }
}
