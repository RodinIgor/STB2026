using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Newtonsoft.Json.Linq;

namespace STB2026.RevitBridge.Handlers
{
    /// <summary>
    /// manage_families â€” ÑƒÐ¿Ñ€Ð°Ð²Ð»ÐµÐ½Ð¸Ðµ ÑÐµÐ¼ÐµÐ¹ÑÑ‚Ð²Ð°Ð¼Ð¸ Ð¸ Ð¾Ð±Ñ‰Ð¸Ð¼Ð¸ Ð¿Ð°Ñ€Ð°Ð¼ÐµÑ‚Ñ€Ð°Ð¼Ð¸.
    /// 
    /// Ð”ÐµÐ¹ÑÑ‚Ð²Ð¸Ñ:
    ///   list_families        â€” ÑÐ¿Ð¸ÑÐ¾Ðº Ð·Ð°Ð³Ñ€ÑƒÐ¶ÐµÐ½Ð½Ñ‹Ñ… ÑÐµÐ¼ÐµÐ¹ÑÑ‚Ð² (Ð¿Ð¾ ÐºÐ°Ñ‚ÐµÐ³Ð¾Ñ€Ð¸Ð¸)
    ///   load_family           â€” Ð·Ð°Ð³Ñ€ÑƒÐ·Ð¸Ñ‚ÑŒ .rfa Ñ„Ð°Ð¹Ð» Ð² Ð¿Ñ€Ð¾ÐµÐºÑ‚
    ///   list_shared_params    â€” ÑÐ¿Ð¸ÑÐ¾Ðº Ð¾Ð±Ñ‰Ð¸Ñ… Ð¿Ð°Ñ€Ð°Ð¼ÐµÑ‚Ñ€Ð¾Ð² Ð¸Ð· Ð¤ÐžÐŸ (Ñ„Ð°Ð¹Ð»Ð° Ð¾Ð±Ñ‰Ð¸Ñ… Ð¿Ð°Ñ€Ð°Ð¼ÐµÑ‚Ñ€Ð¾Ð²)
    ///   add_shared_param      â€” Ð´Ð¾Ð±Ð°Ð²Ð¸Ñ‚ÑŒ Ð¾Ð±Ñ‰Ð¸Ð¹ Ð¿Ð°Ñ€Ð°Ð¼ÐµÑ‚Ñ€ Ðº ÐºÐ°Ñ‚ÐµÐ³Ð¾Ñ€Ð¸ÑÐ¼ Ð¿Ñ€Ð¾ÐµÐºÑ‚Ð°
    ///   remove_shared_param   â€” ÑƒÐ´Ð°Ð»Ð¸Ñ‚ÑŒ Ð¾Ð±Ñ‰Ð¸Ð¹ Ð¿Ð°Ñ€Ð°Ð¼ÐµÑ‚Ñ€ Ð¸Ð· Ð¿Ñ€Ð¸Ð²ÑÐ·Ð¾Ðº Ð¿Ñ€Ð¾ÐµÐºÑ‚Ð°
    ///   list_family_types     â€” Ñ‚Ð¸Ð¿Ð¾Ñ€Ð°Ð·Ð¼ÐµÑ€Ñ‹ ÑÐµÐ¼ÐµÐ¹ÑÑ‚Ð²Ð°
    ///   set_family_type_param â€” Ð¸Ð·Ð¼ÐµÐ½Ð¸Ñ‚ÑŒ Ð¿Ð°Ñ€Ð°Ð¼ÐµÑ‚Ñ€ Ñ‚Ð¸Ð¿Ð¾Ñ€Ð°Ð·Ð¼ÐµÑ€Ð°
    ///   duplicate_type        â€” Ð´ÑƒÐ±Ð»Ð¸Ñ€Ð¾Ð²Ð°Ñ‚ÑŒ Ñ‚Ð¸Ð¿Ð¾Ñ€Ð°Ð·Ð¼ÐµÑ€
    ///   edit_family           â€” Ð¾Ñ‚ÐºÑ€Ñ‹Ñ‚ÑŒ ÑÐµÐ¼ÐµÐ¹ÑÑ‚Ð²Ð¾, Ð¸Ð·Ð¼ÐµÐ½Ð¸Ñ‚ÑŒ/Ð´Ð¾Ð±Ð°Ð²Ð¸Ñ‚ÑŒ Ð¿Ð°Ñ€Ð°Ð¼ÐµÑ‚Ñ€Ñ‹ (Ð²ÐºÐ». Ð¾Ð±Ñ‰Ð¸Ðµ Ð¸Ð· Ð¤ÐžÐŸ), Ð·Ð°Ð³Ñ€ÑƒÐ·Ð¸Ñ‚ÑŒ Ð¾Ð±Ñ€Ð°Ñ‚Ð½Ð¾
    /// </summary>
    internal static class FamilyHandler
    {
        public static object Handle(UIApplication uiApp, Dictionary<string, object> p)
        {
            var uidoc = uiApp.ActiveUIDocument;
            if (uidoc == null)
                return new { error = "ÐÐµÑ‚ Ð¾Ñ‚ÐºÑ€Ñ‹Ñ‚Ð¾Ð³Ð¾ Ð´Ð¾ÐºÑƒÐ¼ÐµÐ½Ñ‚Ð°" };

            var doc = uidoc.Document;
            string action = p.TryGetValue("action", out var a) ? a?.ToString() ?? "" : "";
            string dataStr = p.TryGetValue("data", out var d) ? d?.ToString() ?? "{}" : "{}";

            JObject data;
            try { data = JObject.Parse(dataStr); }
            catch { return new { error = $"ÐÐµÐ²Ð°Ð»Ð¸Ð´Ð½Ñ‹Ð¹ JSON Ð² data: {dataStr}" }; }

            switch (action.ToLowerInvariant())
            {
                case "list_families": return ListFamilies(doc, data);
                case "load_family": return LoadFamily(doc, data);
                case "list_shared_params": return ListSharedParams(doc);
                case "add_shared_param": return AddSharedParam(doc, uiApp.Application, data);
                case "remove_shared_param": return RemoveSharedParam(doc, data);
                case "list_family_types": return ListFamilyTypes(doc, data);
                case "set_family_type_param": return SetFamilyTypeParam(doc, data);
                case "duplicate_type": return DuplicateType(doc, data);
                case "edit_family": return EditFamily(doc, uiApp, data);

                default:
                    return new
                    {
                        error = $"ÐÐµÐ¸Ð·Ð²ÐµÑÑ‚Ð½Ð¾Ðµ Ð´ÐµÐ¹ÑÑ‚Ð²Ð¸Ðµ: '{action}'",
                        available = new[]
                        {
                            "list_families", "load_family",
                            "list_shared_params", "add_shared_param", "remove_shared_param",
                            "list_family_types", "set_family_type_param", "duplicate_type",
                            "edit_family"
                        }
                    };
            }
        }

        // â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
        //  list_families â€” ÑÐ¿Ð¸ÑÐ¾Ðº Ð·Ð°Ð³Ñ€ÑƒÐ¶ÐµÐ½Ð½Ñ‹Ñ… ÑÐµÐ¼ÐµÐ¹ÑÑ‚Ð²
        // â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•

        /// <summary>
        /// data: { "category": "Ð’Ð¾Ð·Ð´ÑƒÑ…Ð¾Ð²Ð¾Ð´Ñ‹" } â€” Ñ„Ð¸Ð»ÑŒÑ‚Ñ€ Ð¿Ð¾ ÐºÐ°Ñ‚ÐµÐ³Ð¾Ñ€Ð¸Ð¸ (Ð½ÐµÐ¾Ð±ÑÐ·Ð°Ñ‚ÐµÐ»ÑŒÐ½Ð¾)
        /// </summary>
        private static object ListFamilies(Document doc, JObject data)
        {
            string catFilter = data.Value<string>("category") ?? "";

            var families = new FilteredElementCollector(doc)
                .OfClass(typeof(Family))
                .Cast<Family>()
                .Where(f => !f.IsInPlace)
                .ToList();

            if (!string.IsNullOrWhiteSpace(catFilter))
            {
                families = families
                    .Where(f => f.FamilyCategory != null &&
                           f.FamilyCategory.Name.IndexOf(catFilter, StringComparison.OrdinalIgnoreCase) >= 0)
                    .ToList();
            }

            var result = families.Select(f =>
            {
                var typeIds = f.GetFamilySymbolIds();
                return new
                {
                    id = f.Id.IntegerValue,
                    name = f.Name,
                    category = f.FamilyCategory?.Name ?? "â€”",
                    types_count = typeIds.Count,
                    type_names = typeIds
                        .Select(tid => doc.GetElement(tid)?.Name ?? "")
                        .Where(n => !string.IsNullOrEmpty(n))
                        .ToList(),
                    is_editable = f.IsEditable
                };
            }).OrderBy(f => f.category).ThenBy(f => f.name).ToList();

            return new
            {
                action = "list_families",
                category_filter = catFilter,
                total = result.Count,
                families = result
            };
        }

        // â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
        //  load_family â€” Ð·Ð°Ð³Ñ€ÑƒÐ·Ð¸Ñ‚ÑŒ .rfa Ð² Ð¿Ñ€Ð¾ÐµÐºÑ‚
        // â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•

        /// <summary>
        /// data: { "path": "C:\\Families\\MyFamily.rfa", "overwrite": true }
        /// </summary>
        private static object LoadFamily(Document doc, JObject data)
        {
            string path = data.Value<string>("path") ?? "";
            bool overwrite = data.Value<bool?>("overwrite") ?? false;

            if (string.IsNullOrWhiteSpace(path))
                return new { error = "Ð£ÐºÐ°Ð¶Ð¸Ñ‚Ðµ path â€” Ð¿ÑƒÑ‚ÑŒ Ðº .rfa Ñ„Ð°Ð¹Ð»Ñƒ" };

            if (!File.Exists(path))
                return new { error = $"Ð¤Ð°Ð¹Ð» Ð½Ðµ Ð½Ð°Ð¹Ð´ÐµÐ½: {path}" };

            using (var trans = new Transaction(doc, "STB2026: Load Family"))
            {
                trans.Start();

                bool loaded;
                Family family = null;

                if (overwrite)
                {
                    var options = new FamilyLoadOptions();
                    loaded = doc.LoadFamily(path, options, out family);
                }
                else
                {
                    loaded = doc.LoadFamily(path, out family);
                }

                if (!loaded && family == null)
                {
                    trans.RollBack();
                    return new { error = "ÐÐµ ÑƒÐ´Ð°Ð»Ð¾ÑÑŒ Ð·Ð°Ð³Ñ€ÑƒÐ·Ð¸Ñ‚ÑŒ ÑÐµÐ¼ÐµÐ¹ÑÑ‚Ð²Ð¾. Ð’Ð¾Ð·Ð¼Ð¾Ð¶Ð½Ð¾, ÑƒÐ¶Ðµ Ð·Ð°Ð³Ñ€ÑƒÐ¶ÐµÐ½Ð¾ (Ð¸ÑÐ¿Ð¾Ð»ÑŒÐ·ÑƒÐ¹Ñ‚Ðµ overwrite: true)." };
                }

                trans.Commit();

                return new
                {
                    action = "load_family",
                    loaded = true,
                    family_name = family?.Name ?? Path.GetFileNameWithoutExtension(path),
                    family_id = family?.Id.IntegerValue ?? -1,
                    category = family?.FamilyCategory?.Name ?? "â€”",
                    types_count = family?.GetFamilySymbolIds().Count ?? 0
                };
            }
        }

        // â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
        //  list_shared_params â€” Ñ‚ÐµÐºÑƒÑ‰Ð¸Ðµ Ð¿Ñ€Ð¸Ð²ÑÐ·ÐºÐ¸ Ð¾Ð±Ñ‰Ð¸Ñ… Ð¿Ð°Ñ€Ð°Ð¼ÐµÑ‚Ñ€Ð¾Ð²
        // â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•

        private static object ListSharedParams(Document doc)
        {
            var bindings = doc.ParameterBindings;
            var iterator = bindings.ForwardIterator();
            var result = new List<object>();

            while (iterator.MoveNext())
            {
                var def = iterator.Key;
                var binding = iterator.Current;

                string bindType = binding is InstanceBinding ? "Instance" : "Type";
                var categories = new List<string>();

                CategorySet catSet = null;
                if (binding is InstanceBinding ib) catSet = ib.Categories;
                else if (binding is TypeBinding tb) catSet = tb.Categories;

                if (catSet != null)
                {
                    foreach (Category c in catSet)
                    {
                        if (c != null) categories.Add(c.Name);
                    }
                }

                result.Add(new
                {
                    name = def.Name,
                    group = GetParamGroup(def),
                    bind_type = bindType,
                    categories = categories
                });
            }

            return new
            {
                action = "list_shared_params",
                total = result.Count,
                parameters = result.OrderBy(p => ((dynamic)p).name).ToList()
            };
        }

        // â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
        //  add_shared_param â€” Ð´Ð¾Ð±Ð°Ð²Ð¸Ñ‚ÑŒ Ð¾Ð±Ñ‰Ð¸Ð¹ Ð¿Ð°Ñ€Ð°Ð¼ÐµÑ‚Ñ€ Ðº ÐºÐ°Ñ‚ÐµÐ³Ð¾Ñ€Ð¸ÑÐ¼ Ð¿Ñ€Ð¾ÐµÐºÑ‚Ð°
        // â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•

        /// <summary>
        /// data: {
        ///   "param_name": "ADSK_Ð Ð°ÑÑ…Ð¾Ð´",
        ///   "group_name": "Ð“Ñ€ÑƒÐ¿Ð¿Ð° Ð² Ð¤ÐžÐŸ",      // Ð³Ñ€ÑƒÐ¿Ð¿Ð° Ð² Ñ„Ð°Ð¹Ð»Ðµ Ð¾Ð±Ñ‰Ð¸Ñ… Ð¿Ð°Ñ€Ð°Ð¼ÐµÑ‚Ñ€Ð¾Ð² (Ð¿Ð¾ ÑƒÐ¼Ð¾Ð»Ñ‡Ð°Ð½Ð¸ÑŽ "STB2026")
        ///   "categories": ["Ð’Ð¾Ð·Ð´ÑƒÑ…Ð¾Ð²Ð¾Ð´Ñ‹", "Ð’Ð¾Ð·Ð´ÑƒÑ…Ð¾Ñ€Ð°ÑÐ¿Ñ€ÐµÐ´ÐµÐ»Ð¸Ñ‚ÐµÐ»Ð¸"],
        ///   "is_instance": true,                // true=ÑÐºÐ·ÐµÐ¼Ð¿Ð»ÑÑ€, false=Ñ‚Ð¸Ð¿
        ///   "param_group": "ÐœÐµÑ…Ð°Ð½Ð¸Ð·Ð¼Ñ‹"          // Ð³Ñ€ÑƒÐ¿Ð¿Ð° Ð¾Ñ‚Ð¾Ð±Ñ€Ð°Ð¶ÐµÐ½Ð¸Ñ Ð² ÑÐ²Ð¾Ð¹ÑÑ‚Ð²Ð°Ñ…
        /// }
        /// Ð•ÑÐ»Ð¸ Ð¤ÐžÐŸ Ð½Ðµ Ð¿Ð¾Ð´ÐºÐ»ÑŽÑ‡Ñ‘Ð½, ÑÐ¾Ð·Ð´Ð°Ñ‘Ð¼ Ð²Ñ€ÐµÐ¼ÐµÐ½Ð½Ñ‹Ð¹.
        /// </summary>
        private static object AddSharedParam(Document doc,
            Autodesk.Revit.ApplicationServices.Application app, JObject data)
        {
            string paramName = data.Value<string>("param_name") ?? "";
            string groupName = data.Value<string>("group_name") ?? "STB2026";
            bool isInstance = data.Value<bool?>("is_instance") ?? true;
            string paramGroupName = data.Value<string>("param_group") ?? "";

            var categoriesArr = data["categories"] as JArray;
            if (string.IsNullOrWhiteSpace(paramName))
                return new { error = "Ð£ÐºÐ°Ð¶Ð¸Ñ‚Ðµ param_name" };
            if (categoriesArr == null || categoriesArr.Count == 0)
                return new { error = "Ð£ÐºÐ°Ð¶Ð¸Ñ‚Ðµ categories â€” Ð¼Ð°ÑÑÐ¸Ð² Ð¸Ð¼Ñ‘Ð½ ÐºÐ°Ñ‚ÐµÐ³Ð¾Ñ€Ð¸Ð¹" };

            // ÐŸÑ€Ð¾Ð²ÐµÑ€ÑÐµÐ¼ â€” Ð¼Ð¾Ð¶ÐµÑ‚ Ð¿Ð°Ñ€Ð°Ð¼ÐµÑ‚Ñ€ ÑƒÐ¶Ðµ Ð¿Ñ€Ð¸Ð²ÑÐ·Ð°Ð½
            var existingBinding = FindExistingBinding(doc, paramName);
            if (existingBinding != null)
            {
                return new
                {
                    error = $"ÐŸÐ°Ñ€Ð°Ð¼ÐµÑ‚Ñ€ '{paramName}' ÑƒÐ¶Ðµ Ð¿Ñ€Ð¸Ð²ÑÐ·Ð°Ð½ Ðº Ð¿Ñ€Ð¾ÐµÐºÑ‚Ñƒ",
                    existing_categories = GetBindingCategories(existingBinding)
                };
            }

            // Ð Ð°Ð±Ð¾Ñ‚Ð° Ñ Ð¤ÐžÐŸ
            string originalSpfPath = app.SharedParametersFilename;
            string tempSpfPath = null;
            bool usedTempSpf = false;

            try
            {
                DefinitionFile defFile = null;

                // ÐŸÑ€Ð¾Ð±ÑƒÐµÐ¼ Ð¸ÑÐ¿Ð¾Ð»ÑŒÐ·Ð¾Ð²Ð°Ñ‚ÑŒ Ñ‚ÐµÐºÑƒÑ‰Ð¸Ð¹ Ð¤ÐžÐŸ
                if (!string.IsNullOrWhiteSpace(originalSpfPath) && File.Exists(originalSpfPath))
                {
                    defFile = app.OpenSharedParameterFile();
                }

                // Ð•ÑÐ»Ð¸ Ð¤ÐžÐŸ Ð½ÐµÑ‚ â€” ÑÐ¾Ð·Ð´Ð°Ñ‘Ð¼ Ð²Ñ€ÐµÐ¼ÐµÐ½Ð½Ñ‹Ð¹
                if (defFile == null)
                {
                    tempSpfPath = Path.Combine(Path.GetTempPath(), "STB2026_SharedParams.txt");
                    if (!File.Exists(tempSpfPath))
                        File.WriteAllText(tempSpfPath, "# STB2026 Shared Parameters\r\n");

                    app.SharedParametersFilename = tempSpfPath;
                    defFile = app.OpenSharedParameterFile();
                    usedTempSpf = true;
                }

                if (defFile == null)
                    return new { error = "ÐÐµ ÑƒÐ´Ð°Ð»Ð¾ÑÑŒ Ð¾Ñ‚ÐºÑ€Ñ‹Ñ‚ÑŒ/ÑÐ¾Ð·Ð´Ð°Ñ‚ÑŒ Ñ„Ð°Ð¹Ð» Ð¾Ð±Ñ‰Ð¸Ñ… Ð¿Ð°Ñ€Ð°Ð¼ÐµÑ‚Ñ€Ð¾Ð² (Ð¤ÐžÐŸ)" };

                // ÐÐ°Ñ…Ð¾Ð´Ð¸Ð¼ Ð¸Ð»Ð¸ ÑÐ¾Ð·Ð´Ð°Ñ‘Ð¼ Ð³Ñ€ÑƒÐ¿Ð¿Ñƒ
                DefinitionGroup defGroup = defFile.Groups.get_Item(groupName);
                if (defGroup == null)
                    defGroup = defFile.Groups.Create(groupName);

                // ÐÐ°Ñ…Ð¾Ð´Ð¸Ð¼ Ð¸Ð»Ð¸ ÑÐ¾Ð·Ð´Ð°Ñ‘Ð¼ Ð¾Ð¿Ñ€ÐµÐ´ÐµÐ»ÐµÐ½Ð¸Ðµ Ð¿Ð°Ñ€Ð°Ð¼ÐµÑ‚Ñ€Ð°
                ExternalDefinition extDef = defGroup.Definitions.get_Item(paramName) as ExternalDefinition;
                if (extDef == null)
                {
                    var options = new ExternalDefinitionCreationOptions(paramName, SpecTypeId.String.Text);
                    options.UserModifiable = true;
                    options.Visible = true;

                    // Ð£Ð³Ð°Ð´Ñ‹Ð²Ð°ÐµÐ¼ Ñ‚Ð¸Ð¿ Ð¿Ð¾ Ð¸Ð¼ÐµÐ½Ð¸
                    if (paramName.IndexOf("Ð Ð°ÑÑ…Ð¾Ð´", StringComparison.OrdinalIgnoreCase) >= 0 ||
                        paramName.IndexOf("AIRFLOW", StringComparison.OrdinalIgnoreCase) >= 0)
                        options = new ExternalDefinitionCreationOptions(paramName, SpecTypeId.AirFlow);
                    else if (paramName.IndexOf("Ð¡ÐºÐ¾Ñ€Ð¾ÑÑ‚ÑŒ", StringComparison.OrdinalIgnoreCase) >= 0 ||
                             paramName.IndexOf("VELOCITY", StringComparison.OrdinalIgnoreCase) >= 0)
                        options = new ExternalDefinitionCreationOptions(paramName, SpecTypeId.HvacVelocity);
                    else if (paramName.IndexOf("Ð”Ð°Ð²Ð»ÐµÐ½", StringComparison.OrdinalIgnoreCase) >= 0 ||
                             paramName.IndexOf("PRESSURE", StringComparison.OrdinalIgnoreCase) >= 0)
                        options = new ExternalDefinitionCreationOptions(paramName, SpecTypeId.HvacPressure);
                    else if (paramName.IndexOf("ÐŸÐ»Ð¾Ñ‰Ð°Ð´", StringComparison.OrdinalIgnoreCase) >= 0)
                        options = new ExternalDefinitionCreationOptions(paramName, SpecTypeId.Area);
                    else if (paramName.IndexOf("Ð”Ð»Ð¸Ð½", StringComparison.OrdinalIgnoreCase) >= 0 ||
                             paramName.IndexOf("Ð’Ñ‹ÑÐ¾Ñ‚", StringComparison.OrdinalIgnoreCase) >= 0 ||
                             paramName.IndexOf("Ð¨Ð¸Ñ€Ð¸Ð½", StringComparison.OrdinalIgnoreCase) >= 0 ||
                             paramName.IndexOf("Ð”Ð¸Ð°Ð¼ÐµÑ‚Ñ€", StringComparison.OrdinalIgnoreCase) >= 0)
                        options = new ExternalDefinitionCreationOptions(paramName, SpecTypeId.Length);
                    else if (paramName.IndexOf("ÐœÐ°ÑÑ", StringComparison.OrdinalIgnoreCase) >= 0)
                        options = new ExternalDefinitionCreationOptions(paramName, SpecTypeId.Mass);

                    options.UserModifiable = true;
                    options.Visible = true;
                    extDef = defGroup.Definitions.Create(options) as ExternalDefinition;
                }

                if (extDef == null)
                    return new { error = $"ÐÐµ ÑƒÐ´Ð°Ð»Ð¾ÑÑŒ ÑÐ¾Ð·Ð´Ð°Ñ‚ÑŒ Ð¾Ð¿Ñ€ÐµÐ´ÐµÐ»ÐµÐ½Ð¸Ðµ Ð¿Ð°Ñ€Ð°Ð¼ÐµÑ‚Ñ€Ð° '{paramName}'" };

                // Ð¡Ð¾Ð±Ð¸Ñ€Ð°ÐµÐ¼ CategorySet
                var catSet = new CategorySet();
                var addedCategories = new List<string>();

                foreach (var catToken in categoriesArr)
                {
                    string catName = catToken.ToString();
                    Category cat = FindCategory(doc, catName);
                    if (cat != null && cat.AllowsBoundParameters)
                    {
                        catSet.Insert(cat);
                        addedCategories.Add(cat.Name);
                    }
                }

                if (catSet.Size == 0)
                    return new { error = "ÐÐ¸ Ð¾Ð´Ð½Ð° Ð¸Ð· ÑƒÐºÐ°Ð·Ð°Ð½Ð½Ñ‹Ñ… ÐºÐ°Ñ‚ÐµÐ³Ð¾Ñ€Ð¸Ð¹ Ð½Ðµ Ð½Ð°Ð¹Ð´ÐµÐ½Ð° Ð¸Ð»Ð¸ Ð½Ðµ Ð¿Ð¾Ð´Ð´ÐµÑ€Ð¶Ð¸Ð²Ð°ÐµÑ‚ Ð¿Ð°Ñ€Ð°Ð¼ÐµÑ‚Ñ€Ñ‹" };

                // Ð¡Ð¾Ð·Ð´Ð°Ñ‘Ð¼ Ð¿Ñ€Ð¸Ð²ÑÐ·ÐºÑƒ
                using (var trans = new Transaction(doc, $"STB2026: Add shared param '{paramName}'"))
                {
                    trans.Start();

                    ElementBinding binding = isInstance
                        ? (ElementBinding)doc.Application.Create.NewInstanceBinding(catSet)
                        : (ElementBinding)doc.Application.Create.NewTypeBinding(catSet);

                    ForgeTypeId groupTypeId = GroupTypeId.General;
                    if (!string.IsNullOrWhiteSpace(paramGroupName))
                        groupTypeId = GuessGroupTypeId(paramGroupName);

                    bool bound = doc.ParameterBindings.Insert(extDef, binding, groupTypeId);

                    if (!bound)
                    {
                        trans.RollBack();
                        return new { error = $"ÐÐµ ÑƒÐ´Ð°Ð»Ð¾ÑÑŒ Ð¿Ñ€Ð¸Ð²ÑÐ·Ð°Ñ‚ÑŒ Ð¿Ð°Ñ€Ð°Ð¼ÐµÑ‚Ñ€ '{paramName}' Ðº ÐºÐ°Ñ‚ÐµÐ³Ð¾Ñ€Ð¸ÑÐ¼" };
                    }

                    trans.Commit();
                }

                return new
                {
                    action = "add_shared_param",
                    param_name = paramName,
                    bind_type = isInstance ? "Instance" : "Type",
                    categories = addedCategories,
                    spf_group = groupName,
                    used_temp_spf = usedTempSpf
                };
            }
            finally
            {
                // Ð’Ð¾Ð·Ð²Ñ€Ð°Ñ‰Ð°ÐµÐ¼ Ð¾Ñ€Ð¸Ð³Ð¸Ð½Ð°Ð»ÑŒÐ½Ñ‹Ð¹ Ð¤ÐžÐŸ
                if (usedTempSpf && !string.IsNullOrWhiteSpace(originalSpfPath))
                    app.SharedParametersFilename = originalSpfPath;
            }
        }

        // â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
        //  remove_shared_param â€” ÑƒÐ´Ð°Ð»Ð¸Ñ‚ÑŒ Ð¿Ñ€Ð¸Ð²ÑÐ·ÐºÑƒ Ð¾Ð±Ñ‰ÐµÐ³Ð¾ Ð¿Ð°Ñ€Ð°Ð¼ÐµÑ‚Ñ€Ð°
        // â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•

        /// <summary>
        /// data: { "param_name": "ADSK_Ð Ð°ÑÑ…Ð¾Ð´" }
        /// </summary>
        private static object RemoveSharedParam(Document doc, JObject data)
        {
            string paramName = data.Value<string>("param_name") ?? "";
            if (string.IsNullOrWhiteSpace(paramName))
                return new { error = "Ð£ÐºÐ°Ð¶Ð¸Ñ‚Ðµ param_name" };

            var iterator = doc.ParameterBindings.ForwardIterator();
            Definition targetDef = null;

            while (iterator.MoveNext())
            {
                if (iterator.Key.Name == paramName)
                {
                    targetDef = iterator.Key;
                    break;
                }
            }

            if (targetDef == null)
                return new { error = $"ÐŸÐ°Ñ€Ð°Ð¼ÐµÑ‚Ñ€ '{paramName}' Ð½Ðµ Ð½Ð°Ð¹Ð´ÐµÐ½ Ð² Ð¿Ñ€Ð¸Ð²ÑÐ·ÐºÐ°Ñ…" };

            using (var trans = new Transaction(doc, $"STB2026: Remove param '{paramName}'"))
            {
                trans.Start();
                bool removed = doc.ParameterBindings.Remove(targetDef);
                if (removed)
                {
                    trans.Commit();
                    return new { action = "remove_shared_param", param_name = paramName, removed = true };
                }
                else
                {
                    trans.RollBack();
                    return new { error = $"ÐÐµ ÑƒÐ´Ð°Ð»Ð¾ÑÑŒ ÑƒÐ´Ð°Ð»Ð¸Ñ‚ÑŒ Ð¿Ñ€Ð¸Ð²ÑÐ·ÐºÑƒ '{paramName}'" };
                }
            }
        }

        // â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
        //  list_family_types â€” Ñ‚Ð¸Ð¿Ð¾Ñ€Ð°Ð·Ð¼ÐµÑ€Ñ‹ ÑÐµÐ¼ÐµÐ¹ÑÑ‚Ð²Ð°
        // â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•

        /// <summary>
        /// data: { "family_name": "ADSK_Ð”Ð¸Ñ„Ñ„ÑƒÐ·Ð¾Ñ€_ÐšÑ€ÑƒÐ³Ð»Ñ‹Ð¹_ÐŸÑ€Ð¸Ñ‚Ð¾Ñ‡Ð½Ñ‹Ð¹" }
        /// Ð¸Ð»Ð¸: { "family_id": 12345 }
        /// </summary>
        private static object ListFamilyTypes(Document doc, JObject data)
        {
            Family family = FindFamily(doc, data);
            if (family == null)
                return new { error = "Ð¡ÐµÐ¼ÐµÐ¹ÑÑ‚Ð²Ð¾ Ð½Ðµ Ð½Ð°Ð¹Ð´ÐµÐ½Ð¾. Ð£ÐºÐ°Ð¶Ð¸Ñ‚Ðµ family_name Ð¸Ð»Ð¸ family_id." };

            var types = family.GetFamilySymbolIds()
                .Select(id => doc.GetElement(id) as FamilySymbol)
                .Where(s => s != null)
                .Select(s =>
                {
                    var typeParams = new List<object>();
                    foreach (Parameter p in s.Parameters)
                    {
                        if (p?.Definition == null || !p.HasValue) continue;
                        if (p.IsReadOnly) continue;

                        typeParams.Add(new
                        {
                            name = p.Definition.Name,
                            value = ParamHelper.GetParamDisplayValue(p),
                            storage = p.StorageType.ToString()
                        });
                    }

                    return new
                    {
                        id = s.Id.IntegerValue,
                        name = s.Name,
                        is_active = s.IsActive,
                        editable_params = typeParams
                    };
                })
                .OrderBy(t => t.name)
                .ToList();

            return new
            {
                action = "list_family_types",
                family_name = family.Name,
                family_id = family.Id.IntegerValue,
                types_count = types.Count,
                types
            };
        }

        // â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
        //  set_family_type_param â€” Ð¸Ð·Ð¼ÐµÐ½Ð¸Ñ‚ÑŒ Ð¿Ð°Ñ€Ð°Ð¼ÐµÑ‚Ñ€ Ñ‚Ð¸Ð¿Ð¾Ñ€Ð°Ð·Ð¼ÐµÑ€Ð°
        // â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•

        /// <summary>
        /// data: { "type_id": 12345, "param_name": "Ð¨Ð¸Ñ€Ð¸Ð½Ð°", "param_value": "600" }
        /// Ð¸Ð»Ð¸: { "family_name": "...", "type_name": "Ð¤100", "param_name": "...", "param_value": "..." }
        /// </summary>
        private static object SetFamilyTypeParam(Document doc, JObject data)
        {
            FamilySymbol symbol = FindFamilySymbol(doc, data);
            if (symbol == null)
                return new { error = "Ð¢Ð¸Ð¿Ð¾Ñ€Ð°Ð·Ð¼ÐµÑ€ Ð½Ðµ Ð½Ð°Ð¹Ð´ÐµÐ½. Ð£ÐºÐ°Ð¶Ð¸Ñ‚Ðµ type_id Ð¸Ð»Ð¸ family_name+type_name." };

            string paramName = data.Value<string>("param_name") ?? "";
            string paramValue = data.Value<string>("param_value") ?? "";

            if (string.IsNullOrWhiteSpace(paramName))
                return new { error = "Ð£ÐºÐ°Ð¶Ð¸Ñ‚Ðµ param_name" };

            Parameter param = symbol.LookupParameter(paramName);
            if (param == null)
                return new { error = $"ÐŸÐ°Ñ€Ð°Ð¼ÐµÑ‚Ñ€ '{paramName}' Ð½Ðµ Ð½Ð°Ð¹Ð´ÐµÐ½ Ñƒ Ñ‚Ð¸Ð¿Ð¾Ñ€Ð°Ð·Ð¼ÐµÑ€Ð° '{symbol.Name}'" };

            if (param.IsReadOnly)
                return new { error = $"ÐŸÐ°Ñ€Ð°Ð¼ÐµÑ‚Ñ€ '{paramName}' Ñ‚Ð¾Ð»ÑŒÐºÐ¾ Ð´Ð»Ñ Ñ‡Ñ‚ÐµÐ½Ð¸Ñ" };

            using (var trans = new Transaction(doc, $"STB2026: Set type param '{paramName}'"))
            {
                trans.Start();
                try
                {
                    SetParamValue(param, paramValue);
                    trans.Commit();

                    return new
                    {
                        action = "set_family_type_param",
                        family = symbol.Family.Name,
                        type_name = symbol.Name,
                        param_name = paramName,
                        new_value = ParamHelper.GetParamDisplayValue(param)
                    };
                }
                catch (Exception ex)
                {
                    trans.RollBack();
                    return new { error = $"ÐžÑˆÐ¸Ð±ÐºÐ°: {ex.Message}" };
                }
            }
        }

        // â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
        //  duplicate_type â€” Ð´ÑƒÐ±Ð»Ð¸Ñ€Ð¾Ð²Ð°Ñ‚ÑŒ Ñ‚Ð¸Ð¿Ð¾Ñ€Ð°Ð·Ð¼ÐµÑ€
        // â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•

        /// <summary>
        /// data: { "family_name": "...", "type_name": "Ð¤100", "new_name": "Ð¤200" }
        /// </summary>
        private static object DuplicateType(Document doc, JObject data)
        {
            FamilySymbol symbol = FindFamilySymbol(doc, data);
            if (symbol == null)
                return new { error = "Ð¢Ð¸Ð¿Ð¾Ñ€Ð°Ð·Ð¼ÐµÑ€ Ð½Ðµ Ð½Ð°Ð¹Ð´ÐµÐ½" };

            string newName = data.Value<string>("new_name") ?? "";
            if (string.IsNullOrWhiteSpace(newName))
                return new { error = "Ð£ÐºÐ°Ð¶Ð¸Ñ‚Ðµ new_name â€” Ð¸Ð¼Ñ Ð½Ð¾Ð²Ð¾Ð³Ð¾ Ñ‚Ð¸Ð¿Ð¾Ñ€Ð°Ð·Ð¼ÐµÑ€Ð°" };

            using (var trans = new Transaction(doc, $"STB2026: Duplicate type '{newName}'"))
            {
                trans.Start();

                var newType = symbol.Duplicate(newName) as FamilySymbol;
                if (newType == null)
                {
                    trans.RollBack();
                    return new { error = $"ÐÐµ ÑƒÐ´Ð°Ð»Ð¾ÑÑŒ Ð´ÑƒÐ±Ð»Ð¸Ñ€Ð¾Ð²Ð°Ñ‚ÑŒ Ñ‚Ð¸Ð¿Ð¾Ñ€Ð°Ð·Ð¼ÐµÑ€" };
                }

                trans.Commit();

                return new
                {
                    action = "duplicate_type",
                    source = symbol.Name,
                    new_type_name = newType.Name,
                    new_type_id = newType.Id.IntegerValue,
                    family = symbol.Family.Name
                };
            }
        }

        // â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
        //  edit_family â€” Ñ€ÐµÐ´Ð°ÐºÑ‚Ð¸Ñ€Ð¾Ð²Ð°Ð½Ð¸Ðµ ÑÐµÐ¼ÐµÐ¹ÑÑ‚Ð²Ð°
        //  ÐŸÐ¾Ð´Ð´ÐµÑ€Ð¶Ð¸Ð²Ð°ÐµÑ‚ Ð´Ð¾Ð±Ð°Ð²Ð»ÐµÐ½Ð¸Ðµ Ð¾Ð±Ñ‰Ð¸Ñ… Ð¿Ð°Ñ€Ð°Ð¼ÐµÑ‚Ñ€Ð¾Ð² Ð¸Ð· Ð¤ÐžÐŸ (is_shared)
        // â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•

        /// <summary>
        /// ÐžÑ‚ÐºÑ€Ñ‹Ð²Ð°ÐµÑ‚ ÑÐµÐ¼ÐµÐ¹ÑÑ‚Ð²Ð¾ Ð´Ð»Ñ Ñ€ÐµÐ´Ð°ÐºÑ‚Ð¸Ñ€Ð¾Ð²Ð°Ð½Ð¸Ñ, Ð´Ð¾Ð±Ð°Ð²Ð»ÑÐµÑ‚/Ð¸Ð·Ð¼ÐµÐ½ÑÐµÑ‚ Ð¿Ð°Ñ€Ð°Ð¼ÐµÑ‚Ñ€Ñ‹, Ð·Ð°Ð³Ñ€ÑƒÐ¶Ð°ÐµÑ‚ Ð¾Ð±Ñ€Ð°Ñ‚Ð½Ð¾.
        /// data: {
        ///   "family_name": "ADSK_Ð”Ð¸Ñ„Ñ„ÑƒÐ·Ð¾Ñ€_ÐšÑ€ÑƒÐ³Ð»Ñ‹Ð¹_ÐŸÑ€Ð¸Ñ‚Ð¾Ñ‡Ð½Ñ‹Ð¹",
        ///   "params": [
        ///     { "name": "LIN_VE_PRESSURE_LOSS", "is_instance": true, "is_shared": true, "group": "ÐœÐµÑ…Ð°Ð½Ð¸Ð·Ð¼Ñ‹\u00A0â€” Ñ€Ð°ÑÑ…Ð¾Ð´" },
        ///     { "name": "CustomParam", "value": "500", "is_instance": true },
        ///     { "name": "ÐžÐ¿Ð¸ÑÐ°Ð½Ð¸Ðµ", "value": "ÐŸÑ€Ð¸Ñ‚Ð¾Ñ‡Ð½Ñ‹Ð¹ Ð´Ð¸Ñ„Ñ„ÑƒÐ·Ð¾Ñ€" }
        ///   ]
        /// }
        /// 
        /// ÐŸÐ°Ñ€Ð°Ð¼ÐµÑ‚Ñ€Ñ‹ ÐºÐ°Ð¶Ð´Ð¾Ð³Ð¾ ÑÐ»ÐµÐ¼ÐµÐ½Ñ‚Ð° Ð¼Ð°ÑÑÐ¸Ð²Ð° params:
        ///   name        â€” Ð¸Ð¼Ñ Ð¿Ð°Ñ€Ð°Ð¼ÐµÑ‚Ñ€Ð° (Ð¾Ð±ÑÐ·Ð°Ñ‚ÐµÐ»ÑŒÐ½Ð¾)
        ///   value       â€” Ð·Ð½Ð°Ñ‡ÐµÐ½Ð¸Ðµ (Ð½ÐµÐ¾Ð±ÑÐ·Ð°Ñ‚ÐµÐ»ÑŒÐ½Ð¾)
        ///   is_instance â€” true=ÑÐºÐ·ÐµÐ¼Ð¿Ð»ÑÑ€, false=Ñ‚Ð¸Ð¿ (Ð¿Ð¾ ÑƒÐ¼Ð¾Ð»Ñ‡Ð°Ð½Ð¸ÑŽ true)
        ///   is_shared   â€” true=Ð¸ÑÐºÐ°Ñ‚ÑŒ Ð² Ð¤ÐžÐŸ ÐºÐ°Ðº Ð¾Ð±Ñ‰Ð¸Ð¹ Ð¿Ð°Ñ€Ð°Ð¼ÐµÑ‚Ñ€ (Ð¿Ð¾ ÑƒÐ¼Ð¾Ð»Ñ‡Ð°Ð½Ð¸ÑŽ false)
        ///   group       â€” Ð³Ñ€ÑƒÐ¿Ð¿Ð° Ð¾Ñ‚Ð¾Ð±Ñ€Ð°Ð¶ÐµÐ½Ð¸Ñ: "ÐœÐµÑ…Ð°Ð½Ð¸Ð·Ð¼Ñ‹", "ÐœÐµÑ…Ð°Ð½Ð¸Ð·Ð¼Ñ‹â€” Ñ€Ð°ÑÑ…Ð¾Ð´", "Ð”Ð°Ð½Ð½Ñ‹Ðµ" Ð¸ Ñ‚.Ð´.
        ///   spec_type   â€” Ñ‚Ð¸Ð¿ Ð´Ð°Ð½Ð½Ñ‹Ñ… Ð´Ð»Ñ ÐÐ•Ð¾Ð±Ñ‰Ð¸Ñ… Ð¿Ð°Ñ€Ð°Ð¼ÐµÑ‚Ñ€Ð¾Ð²: "text","length","airflow","pressure" Ð¸ Ñ‚.Ð´.
        /// </summary>
        private static object EditFamily(Document doc, UIApplication uiApp, JObject data)
        {
            Family family = FindFamily(doc, data);
            if (family == null)
                return new { error = "Ð¡ÐµÐ¼ÐµÐ¹ÑÑ‚Ð²Ð¾ Ð½Ðµ Ð½Ð°Ð¹Ð´ÐµÐ½Ð¾" };

            if (!family.IsEditable)
                return new { error = $"Ð¡ÐµÐ¼ÐµÐ¹ÑÑ‚Ð²Ð¾ '{family.Name}' Ð½ÐµÐ´Ð¾ÑÑ‚ÑƒÐ¿Ð½Ð¾ Ð´Ð»Ñ Ñ€ÐµÐ´Ð°ÐºÑ‚Ð¸Ñ€Ð¾Ð²Ð°Ð½Ð¸Ñ" };

            var paramsArr = data["params"] as JArray;
            if (paramsArr == null || paramsArr.Count == 0)
                return new { error = "Ð£ÐºÐ°Ð¶Ð¸Ñ‚Ðµ params â€” Ð¼Ð°ÑÑÐ¸Ð² Ð¿Ð°Ñ€Ð°Ð¼ÐµÑ‚Ñ€Ð¾Ð² Ð´Ð»Ñ Ð¸Ð·Ð¼ÐµÐ½ÐµÐ½Ð¸Ñ/Ð´Ð¾Ð±Ð°Ð²Ð»ÐµÐ½Ð¸Ñ" };

            // â•â•â• ÐŸÐ¾Ð´Ñ…Ð¾Ð´ Ñ‡ÐµÑ€ÐµÐ· Ð²Ñ€ÐµÐ¼ÐµÐ½Ð½Ñ‹Ð¹ .rfa Ñ„Ð°Ð¹Ð» â•â•â•
            // doc.EditFamily() ÑÐ¾Ð·Ð´Ð°Ñ‘Ñ‚ in-memory Ð´Ð¾ÐºÑƒÐ¼ÐµÐ½Ñ‚, Ð² ÐºÐ¾Ñ‚Ð¾Ñ€Ð¾Ð¼ ExternalDefinition
            // Ð¸Ð· Ð¤ÐžÐŸ ÑÑ‚Ð°Ð½Ð¾Ð²Ð¸Ñ‚ÑÑ Ð½ÐµÐ²Ð°Ð»Ð¸Ð´Ð½Ñ‹Ð¼. ÐŸÐ¾ÑÑ‚Ð¾Ð¼Ñƒ:
            // 1. EditFamily â†’ Ð¿Ð¾Ð»ÑƒÑ‡Ð°ÐµÐ¼ famDoc
            // 2. Ð¡Ð¾Ñ…Ñ€Ð°Ð½ÑÐµÐ¼ famDoc Ð²Ð¾ Ð²Ñ€ÐµÐ¼ÐµÐ½Ð½Ñ‹Ð¹ .rfa
            // 3. Ð—Ð°ÐºÑ€Ñ‹Ð²Ð°ÐµÐ¼ famDoc
            // 4. ÐžÑ‚ÐºÑ€Ñ‹Ð²Ð°ÐµÐ¼ .rfa Ñ‡ÐµÑ€ÐµÐ· Application.OpenDocumentFile()
            // 5. Ð”Ð¾Ð±Ð°Ð²Ð»ÑÐµÐ¼ Ð¿Ð°Ñ€Ð°Ð¼ÐµÑ‚Ñ€Ñ‹ (ExternalDefinition Ð²Ð°Ð»Ð¸Ð´ÐµÐ½ Ð² ÑÑ‚Ð¾Ð¼ ÐºÐ¾Ð½Ñ‚ÐµÐºÑÑ‚Ðµ)
            // 6. Ð¡Ð¾Ñ…Ñ€Ð°Ð½ÑÐµÐ¼ .rfa
            // 7. Ð—Ð°ÐºÑ€Ñ‹Ð²Ð°ÐµÐ¼ .rfa
            // 8. Ð—Ð°Ð³Ñ€ÑƒÐ¶Ð°ÐµÐ¼ Ð² Ð¿Ñ€Ð¾ÐµÐºÑ‚ Ñ‡ÐµÑ€ÐµÐ· doc.LoadFamily(path)

            // Ð’Ñ€ÐµÐ¼ÐµÐ½Ð½Ñ‹Ð¹ .rfa â€” Ð’ÐÐ–ÐÐž: Ð¸Ð¼Ñ Ñ„Ð°Ð¹Ð»Ð° Ð”ÐžÐ›Ð–ÐÐž ÑÐ¾Ð²Ð¿Ð°Ð´Ð°Ñ‚ÑŒ Ñ Ð¸Ð¼ÐµÐ½ÐµÐ¼ ÑÐµÐ¼ÐµÐ¹ÑÑ‚Ð²Ð°,
            // Ð¸Ð½Ð°Ñ‡Ðµ doc.LoadFamily() ÑÐ¾Ð·Ð´Ð°ÑÑ‚ ÐÐžÐ’ÐžÐ• ÑÐµÐ¼ÐµÐ¹ÑÑ‚Ð²Ð¾ Ð²Ð¼ÐµÑÑ‚Ð¾ Ð¿ÐµÑ€ÐµÐ·Ð°Ð¿Ð¸ÑÐ¸
            string tempDir = Path.Combine(Path.GetTempPath(), "STB2026_Families");
            if (!Directory.Exists(tempDir)) Directory.CreateDirectory(tempDir);
            string tempRfaPath = Path.Combine(tempDir, $"{family.Name}.rfa");

            Document famDoc = null;
            Document openedFamDoc = null;

            try
            {
                // â•â•â• Ð¨Ð°Ð³ 1-3: ÐŸÐ¾Ð»ÑƒÑ‡Ð°ÐµÐ¼ .rfa Ð½Ð° Ð´Ð¸ÑÐº â•â•â•
                famDoc = doc.EditFamily(family);
                if (famDoc == null)
                    return new { error = "ÐÐµ ÑƒÐ´Ð°Ð»Ð¾ÑÑŒ Ð¾Ñ‚ÐºÑ€Ñ‹Ñ‚ÑŒ Ð´Ð¾ÐºÑƒÐ¼ÐµÐ½Ñ‚ ÑÐµÐ¼ÐµÐ¹ÑÑ‚Ð²Ð°" };

                var saveOpts = new SaveAsOptions { OverwriteExistingFile = true };
                famDoc.SaveAs(tempRfaPath, saveOpts);
                famDoc.Close(false);
                famDoc = null;

                // â•â•â• Ð¨Ð°Ð³ 4: ÐžÑ‚ÐºÑ€Ñ‹Ð²Ð°ÐµÐ¼ .rfa ÐºÐ°Ðº Ð¾Ð±Ñ‹Ñ‡Ð½Ñ‹Ð¹ Ð´Ð¾ÐºÑƒÐ¼ÐµÐ½Ñ‚ â•â•â•
                openedFamDoc = uiApp.Application.OpenDocumentFile(tempRfaPath);
                if (openedFamDoc == null)
                    return new { error = "ÐÐµ ÑƒÐ´Ð°Ð»Ð¾ÑÑŒ Ð¾Ñ‚ÐºÑ€Ñ‹Ñ‚ÑŒ Ð²Ñ€ÐµÐ¼ÐµÐ½Ð½Ñ‹Ð¹ .rfa Ñ„Ð°Ð¹Ð»" };

                var famMgr = openedFamDoc.FamilyManager;
                int changed = 0;
                int added = 0;
                var errors = new List<string>();
                var addedParams = new List<object>();

                // â•â•â• Ð¨Ð°Ð³ 5: Ð”Ð¾Ð±Ð°Ð²Ð»ÑÐµÐ¼/Ð¸Ð·Ð¼ÐµÐ½ÑÐµÐ¼ Ð¿Ð°Ñ€Ð°Ð¼ÐµÑ‚Ñ€Ñ‹ â•â•â•
                using (var trans = new Transaction(openedFamDoc, "STB2026: Edit Family"))
                {
                    trans.Start();

                    // Ð£Ð±ÐµÐ´Ð¸Ð¼ÑÑ Ñ‡Ñ‚Ð¾ ÐµÑÑ‚ÑŒ Ñ…Ð¾Ñ‚Ñ Ð±Ñ‹ Ð¾Ð´Ð¸Ð½ Ñ‚Ð¸Ð¿
                    if (famMgr.CurrentType == null)
                    {
                        if (famMgr.Types.Size == 0)
                        {
                            famMgr.NewType("Default");
                        }
                        else
                        {
                            var enumerator = famMgr.Types.ForwardIterator();
                            enumerator.MoveNext();
                            famMgr.CurrentType = enumerator.Current as FamilyType;
                        }
                    }

                    foreach (var paramToken in paramsArr)
                    {
                        string pName = paramToken.Value<string>("name") ?? "";
                        string pValue = paramToken.Value<string>("value") ?? "";
                        bool pInstance = paramToken.Value<bool?>("is_instance") ?? true;
                        bool pShared = paramToken.Value<bool?>("is_shared") ?? false;
                        string pGroup = paramToken.Value<string>("group") ?? "";
                        string pSpecType = paramToken.Value<string>("spec_type") ?? "";

                        if (string.IsNullOrWhiteSpace(pName)) continue;

                        // Ð˜Ñ‰ÐµÐ¼ ÑÑƒÑ‰ÐµÑÑ‚Ð²ÑƒÑŽÑ‰Ð¸Ð¹ Ð¿Ð°Ñ€Ð°Ð¼ÐµÑ‚Ñ€ Ð² ÑÐµÐ¼ÐµÐ¹ÑÑ‚Ð²Ðµ
                        FamilyParameter famParam = null;
                        foreach (FamilyParameter fp in famMgr.Parameters)
                        {
                            if (fp.Definition.Name == pName)
                            {
                                famParam = fp;
                                break;
                            }
                        }

                        // Ð•ÑÐ»Ð¸ Ð¿Ð°Ñ€Ð°Ð¼ÐµÑ‚Ñ€ ÑƒÐ¶Ðµ ÑÑƒÑ‰ÐµÑÑ‚Ð²ÑƒÐµÑ‚ â€” Ñ‚Ð¾Ð»ÑŒÐºÐ¾ Ð¼ÐµÐ½ÑÐµÐ¼ Ð·Ð½Ð°Ñ‡ÐµÐ½Ð¸Ðµ
                        if (famParam != null)
                        {
                            if (!string.IsNullOrWhiteSpace(pValue))
                            {
                                try
                                {
                                    SetFamilyParamValue(famMgr, famParam, pValue);
                                    changed++;
                                }
                                catch (Exception ex)
                                {
                                    errors.Add($"'{pName}': Ð¾ÑˆÐ¸Ð±ÐºÐ° ÑƒÑÑ‚Ð°Ð½Ð¾Ð²ÐºÐ¸ Ð·Ð½Ð°Ñ‡ÐµÐ½Ð¸Ñ: {ex.Message}");
                                }
                            }
                            else
                            {
                                errors.Add($"'{pName}': Ð¿Ð°Ñ€Ð°Ð¼ÐµÑ‚Ñ€ ÑƒÐ¶Ðµ ÑÑƒÑ‰ÐµÑÑ‚Ð²ÑƒÐµÑ‚ Ð² ÑÐµÐ¼ÐµÐ¹ÑÑ‚Ð²Ðµ");
                            }
                            continue;
                        }

                        // === ÐŸÐ°Ñ€Ð°Ð¼ÐµÑ‚Ñ€Ð° Ð½ÐµÑ‚ â€” ÑÐ¾Ð·Ð´Ð°Ñ‘Ð¼ ===

                        ForgeTypeId groupId = string.IsNullOrWhiteSpace(pGroup)
                            ? GroupTypeId.General
                            : GuessGroupTypeId(pGroup);

                        if (pShared)
                        {
                            // â”€â”€â”€ Ð”Ð¾Ð±Ð°Ð²Ð»ÐµÐ½Ð¸Ðµ ÐžÐ‘Ð©Ð•Ð“Ðž Ð¿Ð°Ñ€Ð°Ð¼ÐµÑ‚Ñ€Ð° Ð¸Ð· Ð¤ÐžÐŸ â”€â”€â”€
                            try
                            {
                                // Ð’ ÐºÐ¾Ð½Ñ‚ÐµÐºÑÑ‚Ðµ Ð¾Ñ‚Ð´ÐµÐ»ÑŒÐ½Ð¾ Ð¾Ñ‚ÐºÑ€Ñ‹Ñ‚Ð¾Ð³Ð¾ .rfa ExternalDefinition Ð²Ð°Ð»Ð¸Ð´ÐµÐ½
                                ExternalDefinition extDef = FindSharedParamDefinition(
                                    uiApp.Application, pName);

                                if (extDef == null)
                                {
                                    errors.Add($"'{pName}': Ð½Ðµ Ð½Ð°Ð¹Ð´ÐµÐ½ Ð² Ñ„Ð°Ð¹Ð»Ðµ Ð¾Ð±Ñ‰Ð¸Ñ… Ð¿Ð°Ñ€Ð°Ð¼ÐµÑ‚Ñ€Ð¾Ð² (Ð¤ÐžÐŸ). " +
                                               "ÐŸÑ€Ð¾Ð²ÐµÑ€ÑŒÑ‚Ðµ, Ñ‡Ñ‚Ð¾ Ð¤ÐžÐŸ Ð¿Ð¾Ð´ÐºÐ»ÑŽÑ‡Ñ‘Ð½ Ð¸ Ð¿Ð°Ñ€Ð°Ð¼ÐµÑ‚Ñ€ ÑÑƒÑ‰ÐµÑÑ‚Ð²ÑƒÐµÑ‚.");
                                    continue;
                                }

                                string dataType = "unknown";
                                try { dataType = extDef.GetDataType()?.TypeId ?? "unknown"; } catch { }

                                famParam = famMgr.AddParameter(extDef, groupId, pInstance);

                                if (famParam != null)
                                {
                                    added++;
                                    addedParams.Add(new
                                    {
                                        name = pName,
                                        type = "shared",
                                        is_instance = pInstance,
                                        group = GetGroupLabel(groupId),
                                        data_type = dataType
                                    });

                                    if (!string.IsNullOrWhiteSpace(pValue))
                                    {
                                        try
                                        {
                                            SetFamilyParamValue(famMgr, famParam, pValue);
                                            changed++;
                                        }
                                        catch (Exception ex)
                                        {
                                            errors.Add($"'{pName}': Ð´Ð¾Ð±Ð°Ð²Ð»ÐµÐ½, Ð½Ð¾ Ð¾ÑˆÐ¸Ð±ÐºÐ° Ð·Ð½Ð°Ñ‡ÐµÐ½Ð¸Ñ: {ex.Message}");
                                        }
                                    }
                                }
                                else
                                {
                                    errors.Add($"'{pName}': FamilyManager.AddParameter Ð²ÐµÑ€Ð½ÑƒÐ» null");
                                }
                            }
                            catch (Exception ex)
                            {
                                errors.Add($"'{pName}': Ð¾ÑˆÐ¸Ð±ÐºÐ° Ð´Ð¾Ð±Ð°Ð²Ð»ÐµÐ½Ð¸Ñ Ð¾Ð±Ñ‰ÐµÐ³Ð¾ Ð¿Ð°Ñ€Ð°Ð¼ÐµÑ‚Ñ€Ð°: {ex.Message}");
                            }
                        }
                        else
                        {
                            // â”€â”€â”€ Ð”Ð¾Ð±Ð°Ð²Ð»ÐµÐ½Ð¸Ðµ Ð¾Ð±Ñ‹Ñ‡Ð½Ð¾Ð³Ð¾ Ð¿Ð°Ñ€Ð°Ð¼ÐµÑ‚Ñ€Ð° ÑÐµÐ¼ÐµÐ¹ÑÑ‚Ð²Ð° â”€â”€â”€
                            try
                            {
                                ForgeTypeId specTypeId = GuessSpecTypeId(pSpecType, pName);

                                famParam = famMgr.AddParameter(pName, groupId, specTypeId, pInstance);

                                if (famParam != null)
                                {
                                    added++;
                                    addedParams.Add(new
                                    {
                                        name = pName,
                                        type = "family",
                                        is_instance = pInstance,
                                        group = GetGroupLabel(groupId),
                                        data_type = specTypeId?.TypeId ?? "text"
                                    });

                                    if (!string.IsNullOrWhiteSpace(pValue))
                                    {
                                        try
                                        {
                                            SetFamilyParamValue(famMgr, famParam, pValue);
                                            changed++;
                                        }
                                        catch (Exception ex)
                                        {
                                            errors.Add($"'{pName}': Ð´Ð¾Ð±Ð°Ð²Ð»ÐµÐ½, Ð½Ð¾ Ð¾ÑˆÐ¸Ð±ÐºÐ° Ð·Ð½Ð°Ñ‡ÐµÐ½Ð¸Ñ: {ex.Message}");
                                        }
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                errors.Add($"'{pName}': Ð½Ðµ ÑƒÐ´Ð°Ð»Ð¾ÑÑŒ ÑÐ¾Ð·Ð´Ð°Ñ‚ÑŒ: {ex.Message}");
                            }
                        }
                    }

                    trans.Commit();
                }

                // â•â•â• Ð¨Ð°Ð³ 6-7: Ð¡Ð¾Ñ…Ñ€Ð°Ð½ÑÐµÐ¼ Ð¸ Ð·Ð°ÐºÑ€Ñ‹Ð²Ð°ÐµÐ¼ .rfa â•â•â•
                openedFamDoc.Save();
                openedFamDoc.Close(false);
                openedFamDoc = null;

                // â•â•â• Ð¨Ð°Ð³ 8: Ð—Ð°Ð³Ñ€ÑƒÐ¶Ð°ÐµÐ¼ Ð¾Ð±Ð½Ð¾Ð²Ð»Ñ‘Ð½Ð½Ð¾Ðµ ÑÐµÐ¼ÐµÐ¹ÑÑ‚Ð²Ð¾ Ð² Ð¿Ñ€Ð¾ÐµÐºÑ‚ â•â•â•
                Family reloadedFamily = null;
                using (var trans = new Transaction(doc, "STB2026: Reload Family"))
                {
                    trans.Start();
                    var loadOpts = new FamilyLoadOptions();
                    doc.LoadFamily(tempRfaPath, loadOpts, out reloadedFamily);
                    trans.Commit();
                }

                return new
                {
                    action = "edit_family",
                    family_name = family.Name,
                    params_changed = changed,
                    params_added = added,
                    added_params = addedParams,
                    reloaded = reloadedFamily != null,
                    errors
                };
            }
            catch (Exception ex)
            {
                try { if (famDoc != null) famDoc.Close(false); } catch { }
                try { if (openedFamDoc != null) openedFamDoc.Close(false); } catch { }
                return new { error = $"ÐžÑˆÐ¸Ð±ÐºÐ° Ñ€ÐµÐ´Ð°ÐºÑ‚Ð¸Ñ€Ð¾Ð²Ð°Ð½Ð¸Ñ ÑÐµÐ¼ÐµÐ¹ÑÑ‚Ð²Ð°: {ex.Message}" };
            }
            finally
            {
                // Ð£Ð´Ð°Ð»ÑÐµÐ¼ Ð²Ñ€ÐµÐ¼ÐµÐ½Ð½Ñ‹Ð¹ Ñ„Ð°Ð¹Ð»
                try { if (File.Exists(tempRfaPath)) File.Delete(tempRfaPath); } catch { }
            }
        }

        // â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
        //  Ð’ÑÐ¿Ð¾Ð¼Ð¾Ð³Ð°Ñ‚ÐµÐ»ÑŒÐ½Ñ‹Ðµ Ð¼ÐµÑ‚Ð¾Ð´Ñ‹
        // â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•

        /// <summary>Ð˜Ñ‰ÐµÑ‚ ÑÐµÐ¼ÐµÐ¹ÑÑ‚Ð²Ð¾ Ð¿Ð¾ ID Ð¸Ð»Ð¸ Ð¸Ð¼ÐµÐ½Ð¸.</summary>
        private static Family FindFamily(Document doc, JObject data)
        {
            int? familyId = data.Value<int?>("family_id");
            string familyName = data.Value<string>("family_name") ?? "";

            if (familyId.HasValue && familyId > 0)
            {
                return doc.GetElement(new ElementId(familyId.Value)) as Family;
            }

            if (!string.IsNullOrWhiteSpace(familyName))
            {
                return new FilteredElementCollector(doc)
                    .OfClass(typeof(Family))
                    .Cast<Family>()
                    .FirstOrDefault(f => f.Name.IndexOf(familyName, StringComparison.OrdinalIgnoreCase) >= 0);
            }

            return null;
        }

        /// <summary>Ð˜Ñ‰ÐµÑ‚ FamilySymbol Ð¿Ð¾ type_id Ð¸Ð»Ð¸ family_name+type_name.</summary>
        private static FamilySymbol FindFamilySymbol(Document doc, JObject data)
        {
            int? typeId = data.Value<int?>("type_id");
            if (typeId.HasValue && typeId > 0)
                return doc.GetElement(new ElementId(typeId.Value)) as FamilySymbol;

            string familyName = data.Value<string>("family_name") ?? "";
            string typeName = data.Value<string>("type_name") ?? "";

            if (string.IsNullOrWhiteSpace(familyName)) return null;

            var family = FindFamily(doc, data);
            if (family == null) return null;

            return family.GetFamilySymbolIds()
                .Select(id => doc.GetElement(id) as FamilySymbol)
                .FirstOrDefault(s => s != null &&
                    (string.IsNullOrWhiteSpace(typeName) ||
                     s.Name.IndexOf(typeName, StringComparison.OrdinalIgnoreCase) >= 0));
        }

        /// <summary>Ð˜Ñ‰ÐµÑ‚ ÐºÐ°Ñ‚ÐµÐ³Ð¾Ñ€Ð¸ÑŽ Ð¿Ð¾ Ð¸Ð¼ÐµÐ½Ð¸ (Ñ‡Ð°ÑÑ‚Ð¸Ñ‡Ð½Ð¾Ðµ ÑÐ¾Ð²Ð¿Ð°Ð´ÐµÐ½Ð¸Ðµ).</summary>
        private static Category FindCategory(Document doc, string name)
        {
            foreach (Category c in doc.Settings.Categories)
            {
                if (c != null && c.Name.IndexOf(name, StringComparison.OrdinalIgnoreCase) >= 0)
                    return c;
            }
            return null;
        }

        /// <summary>Ð˜Ñ‰ÐµÑ‚ ÑÑƒÑ‰ÐµÑÑ‚Ð²ÑƒÑŽÑ‰ÑƒÑŽ Ð¿Ñ€Ð¸Ð²ÑÐ·ÐºÑƒ Ð¿Ð°Ñ€Ð°Ð¼ÐµÑ‚Ñ€Ð° Ð¿Ð¾ Ð¸Ð¼ÐµÐ½Ð¸.</summary>
        private static ElementBinding FindExistingBinding(Document doc, string paramName)
        {
            var iterator = doc.ParameterBindings.ForwardIterator();
            while (iterator.MoveNext())
            {
                if (iterator.Key.Name == paramName)
                    return iterator.Current as ElementBinding;
            }
            return null;
        }

        /// <summary>Ð’Ð¾Ð·Ð²Ñ€Ð°Ñ‰Ð°ÐµÑ‚ ÑÐ¿Ð¸ÑÐ¾Ðº ÐºÐ°Ñ‚ÐµÐ³Ð¾Ñ€Ð¸Ð¹ Ð¸Ð· Ð¿Ñ€Ð¸Ð²ÑÐ·ÐºÐ¸.</summary>
        private static List<string> GetBindingCategories(ElementBinding binding)
        {
            CategorySet cats = null;
            if (binding is InstanceBinding ib) cats = ib.Categories;
            else if (binding is TypeBinding tb) cats = tb.Categories;

            var result = new List<string>();
            if (cats != null)
                foreach (Category c in cats)
                    if (c != null) result.Add(c.Name);
            return result;
        }

        /// <summary>Ð’Ð¾Ð·Ð²Ñ€Ð°Ñ‰Ð°ÐµÑ‚ Ð¸Ð¼Ñ Ð³Ñ€ÑƒÐ¿Ð¿Ñ‹ Ð¿Ð°Ñ€Ð°Ð¼ÐµÑ‚Ñ€Ð°.</summary>
        private static string GetParamGroup(Definition def)
        {
            try
            {
                var groupId = def.GetGroupTypeId();
                return groupId != null ? LabelUtils.GetLabelForGroup(groupId) : "ÐŸÑ€Ð¾Ñ‡ÐµÐµ";
            }
            catch { return "ÐŸÑ€Ð¾Ñ‡ÐµÐµ"; }
        }

        /// <summary>
        /// Ð˜Ñ‰ÐµÑ‚ ExternalDefinition Ð¿Ð¾ Ð¸Ð¼ÐµÐ½Ð¸ Ð²Ð¾ Ð²ÑÐµÑ… Ð³Ñ€ÑƒÐ¿Ð¿Ð°Ñ… Ñ‚ÐµÐºÑƒÑ‰ÐµÐ³Ð¾ Ñ„Ð°Ð¹Ð»Ð° Ð¾Ð±Ñ‰Ð¸Ñ… Ð¿Ð°Ñ€Ð°Ð¼ÐµÑ‚Ñ€Ð¾Ð².
        /// Ð•ÑÐ»Ð¸ Ð¤ÐžÐŸ Ð½Ðµ Ð¿Ð¾Ð´ÐºÐ»ÑŽÑ‡Ñ‘Ð½ â€” Ð¿Ñ‹Ñ‚Ð°ÐµÑ‚ÑÑ Ð½Ð°Ð¹Ñ‚Ð¸ ÐµÐ³Ð¾ Ð¿Ð¾ ÑÑ‚Ð°Ð½Ð´Ð°Ñ€Ñ‚Ð½Ñ‹Ð¼ Ð¿ÑƒÑ‚ÑÐ¼.
        /// </summary>
        private static ExternalDefinition FindSharedParamDefinition(
            Autodesk.Revit.ApplicationServices.Application app, string paramName)
        {
            DefinitionFile defFile = app.OpenSharedParameterFile();

            if (defFile == null)
            {
                // ÐŸÑ€Ð¾Ð±ÑƒÐµÐ¼ Ð½Ð°Ð¹Ñ‚Ð¸ Ð¤ÐžÐŸ Ð¿Ð¾ ÑÑ‚Ð°Ð½Ð´Ð°Ñ€Ñ‚Ð½Ñ‹Ð¼ Ð¿ÑƒÑ‚ÑÐ¼
                string[] defaultPaths = new[]
                {
                    Path.Combine(
                        Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                        "Autodesk", "Revit", "SharedParams.txt"),
                    @"C:\ProgramData\Autodesk\Revit\SharedParams.txt",
                    Path.Combine(
                        Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
                        "ADSK_Shared_Parameters.txt"),
                    Path.Combine(
                        Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
                        "SharedParams.txt")
                };

                foreach (var path in defaultPaths)
                {
                    if (File.Exists(path))
                    {
                        try
                        {
                            app.SharedParametersFilename = path;
                            defFile = app.OpenSharedParameterFile();
                            if (defFile != null) break;
                        }
                        catch { /* Ð¿Ñ€Ð¾Ð±ÑƒÐµÐ¼ ÑÐ»ÐµÐ´ÑƒÑŽÑ‰Ð¸Ð¹ */ }
                    }
                }
            }

            if (defFile == null)
                return null;

            // Ð˜Ñ‰ÐµÐ¼ Ð¿Ð°Ñ€Ð°Ð¼ÐµÑ‚Ñ€ Ð²Ð¾ Ð²ÑÐµÑ… Ð³Ñ€ÑƒÐ¿Ð¿Ð°Ñ…
            foreach (DefinitionGroup group in defFile.Groups)
            {
                foreach (Definition def in group.Definitions)
                {
                    if (def.Name == paramName && def is ExternalDefinition extDef)
                        return extDef;
                }
            }

            return null;
        }

        /// <summary>ÐžÐ¿Ñ€ÐµÐ´ÐµÐ»ÑÐµÑ‚ Ð³Ñ€ÑƒÐ¿Ð¿Ñƒ Ð¾Ñ‚Ð¾Ð±Ñ€Ð°Ð¶ÐµÐ½Ð¸Ñ Ð¿Ð°Ñ€Ð°Ð¼ÐµÑ‚Ñ€Ð° Ð¿Ð¾ Ñ‚ÐµÐºÑÑ‚Ð¾Ð²Ð¾Ð¼Ñƒ Ð¸Ð¼ÐµÐ½Ð¸.</summary>
        private static ForgeTypeId GuessGroupTypeId(string groupName)
        {
            if (string.IsNullOrWhiteSpace(groupName))
                return GroupTypeId.General;

            string lower = groupName.ToLowerInvariant();

            // "ÐœÐµÑ…Ð°Ð½Ð¸Ð·Ð¼Ñ‹ â€” Ñ€Ð°ÑÑ…Ð¾Ð´" / "ÐœÐµÑ…Ð°Ð½Ð¸Ð·Ð¼Ñ‹\u00A0â€” Ñ€Ð°ÑÑ…Ð¾Ð´" (Ñ Ð½ÐµÑ€Ð°Ð·Ñ€Ñ‹Ð²Ð½Ñ‹Ð¼ Ð¿Ñ€Ð¾Ð±ÐµÐ»Ð¾Ð¼)
            if (lower.Contains("Ð¼ÐµÑ…Ð°Ð½") && (lower.Contains("Ñ€Ð°ÑÑ…Ð¾Ð´") || lower.Contains("flow")))
                return GroupTypeId.MechanicalAirflow;

            // "ÐœÐµÑ…Ð°Ð½Ð¸Ð·Ð¼Ñ‹ â€” Ð½Ð°Ð³Ñ€ÑƒÐ·ÐºÐ¸"
            if (lower.Contains("Ð¼ÐµÑ…Ð°Ð½") && (lower.Contains("Ð½Ð°Ð³Ñ€ÑƒÐ·Ðº") || lower.Contains("load")))
                return GroupTypeId.MechanicalLoads;

            // "ÐœÐµÑ…Ð°Ð½Ð¸Ð·Ð¼Ñ‹" (Ð¾Ð±Ñ‰ÐµÐµ)
            if (lower.Contains("Ð¼ÐµÑ…Ð°Ð½") || lower.Contains("hvac") || lower.Contains("mech"))
                return GroupTypeId.Mechanical;

            if (lower.Contains("Ñ€Ð°Ð·Ð¼") || lower.Contains("dimen") || lower.Contains("geom"))
                return GroupTypeId.Geometry;

            if (lower.Contains("Ð¸Ð´ÐµÐ½Ñ‚Ð¸Ñ„") || lower.Contains("ident"))
                return GroupTypeId.IdentityData;

            if (lower.Contains("Ð´Ð°Ð½Ð½") || lower.Contains("data"))
                return GroupTypeId.Data;

            if (lower.Contains("Ð·Ð°Ð²Ð¸Ñ") || lower.Contains("constr"))
                return GroupTypeId.Constraints;

            if (lower.Contains("ÑÑ‚Ñ€Ð¾Ð¸Ñ‚") || lower.Contains("construct"))
                return GroupTypeId.Construction;

            if (lower.Contains("Ñ‚ÐµÐºÑÑ‚") || lower.Contains("text"))
                return GroupTypeId.Text;

            if (lower.Contains("Ð¾Ð±Ñ‰") || lower.Contains("general"))
                return GroupTypeId.General;

            if (lower.Contains("ÑÐ»ÐµÐºÑ‚Ñ€") || lower.Contains("electr"))
                return GroupTypeId.Electrical;

            if (lower.Contains("ÑÐ°Ð½Ñ‚ÐµÑ…") || lower.Contains("plumb"))
                return GroupTypeId.Plumbing;

            if (lower.Contains("ifc"))
                return GroupTypeId.Ifc;

            return GroupTypeId.General;
        }

        /// <summary>ÐžÐ¿Ñ€ÐµÐ´ÐµÐ»ÑÐµÑ‚ Ñ‚Ð¸Ð¿ Ð´Ð°Ð½Ð½Ñ‹Ñ… Ð¿Ð°Ñ€Ð°Ð¼ÐµÑ‚Ñ€Ð° Ð¿Ð¾ ÑÐ²Ð½Ð¾Ð¼Ñƒ ÑƒÐºÐ°Ð·Ð°Ð½Ð¸ÑŽ Ð¸Ð»Ð¸ Ð¿Ð¾ Ð¸Ð¼ÐµÐ½Ð¸.</summary>
        private static ForgeTypeId GuessSpecTypeId(string specType, string paramName)
        {
            // Ð•ÑÐ»Ð¸ ÑÐ²Ð½Ð¾ ÑƒÐºÐ°Ð·Ð°Ð½ Ñ‚Ð¸Ð¿
            if (!string.IsNullOrWhiteSpace(specType))
            {
                switch (specType.ToLowerInvariant())
                {
                    case "text":
                    case "string":
                        return SpecTypeId.String.Text;
                    case "length":
                        return SpecTypeId.Length;
                    case "area":
                        return SpecTypeId.Area;
                    case "volume":
                        return SpecTypeId.Volume;
                    case "airflow":
                    case "air_flow":
                        return SpecTypeId.AirFlow;
                    case "pressure":
                    case "hvac_pressure":
                        return SpecTypeId.HvacPressure;
                    case "velocity":
                    case "hvac_velocity":
                        return SpecTypeId.HvacVelocity;
                    case "number":
                    case "real":
                        return SpecTypeId.Number;
                    case "integer":
                    case "int":
                        return SpecTypeId.Int.Integer;
                    case "yesno":
                    case "bool":
                        return SpecTypeId.Boolean.YesNo;
                    case "angle":
                        return SpecTypeId.Angle;
                    case "mass":
                        return SpecTypeId.Mass;
                    case "temperature":
                        return SpecTypeId.HvacTemperature;
                    case "power":
                        return SpecTypeId.HvacPower;
                }
            }

            // Ð£Ð³Ð°Ð´Ñ‹Ð²Ð°ÐµÐ¼ Ð¿Ð¾ Ð¸Ð¼ÐµÐ½Ð¸ Ð¿Ð°Ñ€Ð°Ð¼ÐµÑ‚Ñ€Ð°
            string lower = paramName.ToLowerInvariant();

            if (lower.Contains("Ñ€Ð°ÑÑ…Ð¾Ð´") || lower.Contains("airflow") || lower.Contains("air_flow"))
                return SpecTypeId.AirFlow;
            if (lower.Contains("Ð´Ð°Ð²Ð»ÐµÐ½") || lower.Contains("pressure"))
                return SpecTypeId.HvacPressure;
            if (lower.Contains("ÑÐºÐ¾Ñ€Ð¾ÑÑ‚") || lower.Contains("velocity"))
                return SpecTypeId.HvacVelocity;
            if (lower.Contains("Ð¿Ð»Ð¾Ñ‰Ð°Ð´") || lower.Contains("area"))
                return SpecTypeId.Area;
            if (lower.Contains("Ð´Ð»Ð¸Ð½") || lower.Contains("Ð²Ñ‹ÑÐ¾Ñ‚") || lower.Contains("ÑˆÐ¸Ñ€Ð¸Ð½") ||
                lower.Contains("Ð´Ð¸Ð°Ð¼ÐµÑ‚Ñ€") || lower.Contains("length") || lower.Contains("height") ||
                lower.Contains("width") || lower.Contains("diameter"))
                return SpecTypeId.Length;
            if (lower.Contains("Ð¼Ð°ÑÑ") || lower.Contains("mass") || lower.Contains("weight"))
                return SpecTypeId.Mass;
            if (lower.Contains("Ñ‚ÐµÐ¼Ð¿ÐµÑ€Ð°Ñ‚ÑƒÑ€") || lower.Contains("temperature"))
                return SpecTypeId.HvacTemperature;
            if (lower.Contains("Ð¼Ð¾Ñ‰Ð½Ð¾ÑÑ‚") || lower.Contains("power"))
                return SpecTypeId.HvacPower;
            if (lower.Contains("Ð¾Ð±ÑŠÑ‘Ð¼") || lower.Contains("Ð¾Ð±ÑŠÐµÐ¼") || lower.Contains("volume"))
                return SpecTypeId.Volume;

            return SpecTypeId.String.Text;
        }

        /// <summary>ÐŸÐ¾Ð»ÑƒÑ‡Ð°ÐµÑ‚ Ñ‡Ð¸Ñ‚Ð°ÐµÐ¼Ð¾Ðµ Ð¸Ð¼Ñ Ð³Ñ€ÑƒÐ¿Ð¿Ñ‹ Ð¿Ð°Ñ€Ð°Ð¼ÐµÑ‚Ñ€Ð¾Ð².</summary>
        private static string GetGroupLabel(ForgeTypeId groupId)
        {
            try
            {
                return LabelUtils.GetLabelForGroup(groupId);
            }
            catch
            {
                return groupId?.TypeId ?? "ÐŸÑ€Ð¾Ñ‡ÐµÐµ";
            }
        }

        /// <summary>Ð£ÑÑ‚Ð°Ð½Ð°Ð²Ð»Ð¸Ð²Ð°ÐµÑ‚ Ð·Ð½Ð°Ñ‡ÐµÐ½Ð¸Ðµ Ð¿Ð°Ñ€Ð°Ð¼ÐµÑ‚Ñ€Ð° ÑÐ»ÐµÐ¼ÐµÐ½Ñ‚Ð°.</summary>
        private static void SetParamValue(Parameter param, string value)
        {
            switch (param.StorageType)
            {
                case StorageType.String:
                    param.Set(value);
                    break;
                case StorageType.Integer:
                    if (int.TryParse(value, out int i)) param.Set(i);
                    else throw new ArgumentException($"ÐžÐ¶Ð¸Ð´Ð°Ð»Ð¾ÑÑŒ Ñ†ÐµÐ»Ð¾Ðµ Ñ‡Ð¸ÑÐ»Ð¾: {value}");
                    break;
                case StorageType.Double:
                    param.SetValueString(value);
                    break;
                case StorageType.ElementId:
                    if (int.TryParse(value, out int eid)) param.Set(new ElementId(eid));
                    else throw new ArgumentException($"ÐžÐ¶Ð¸Ð´Ð°Ð»ÑÑ ID: {value}");
                    break;
            }
        }

        /// <summary>Ð£ÑÑ‚Ð°Ð½Ð°Ð²Ð»Ð¸Ð²Ð°ÐµÑ‚ Ð·Ð½Ð°Ñ‡ÐµÐ½Ð¸Ðµ Ð¿Ð°Ñ€Ð°Ð¼ÐµÑ‚Ñ€Ð° ÑÐµÐ¼ÐµÐ¹ÑÑ‚Ð²Ð° Ñ‡ÐµÑ€ÐµÐ· FamilyManager.</summary>
        private static void SetFamilyParamValue(FamilyManager mgr, FamilyParameter param, string value)
        {
            switch (param.StorageType)
            {
                case StorageType.String:
                    mgr.Set(param, value);
                    break;
                case StorageType.Integer:
                    if (int.TryParse(value, out int i)) mgr.Set(param, i);
                    break;
                case StorageType.Double:
                    if (double.TryParse(value.Replace(",", "."),
                        System.Globalization.NumberStyles.Float,
                        System.Globalization.CultureInfo.InvariantCulture, out double d))
                        mgr.Set(param, d);
                    break;
                case StorageType.ElementId:
                    if (int.TryParse(value, out int eid)) mgr.Set(param, new ElementId(eid));
                    break;
            }
        }

        /// <summary>ÐžÐ¿Ñ†Ð¸Ð¸ Ð·Ð°Ð³Ñ€ÑƒÐ·ÐºÐ¸ ÑÐµÐ¼ÐµÐ¹ÑÑ‚Ð²Ð° Ñ Ð¿ÐµÑ€ÐµÐ·Ð°Ð¿Ð¸ÑÑŒÑŽ.</summary>
        private class FamilyLoadOptions : IFamilyLoadOptions
        {
            public bool OnFamilyFound(bool familyInUse, out bool overwriteParameterValues)
            {
                overwriteParameterValues = true;
                return true;
            }

            public bool OnSharedFamilyFound(Family sharedFamily, bool familyInUse,
                out FamilySource source, out bool overwriteParameterValues)
            {
                source = FamilySource.Project;
                overwriteParameterValues = true;
                return true;
            }
        }
    }
}
