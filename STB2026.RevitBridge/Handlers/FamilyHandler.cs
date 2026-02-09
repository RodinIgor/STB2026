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
    /// manage_families — управление семействами и общими параметрами.
    /// 
    /// Действия:
    ///   list_families        — список загруженных семейств (по категории)
    ///   load_family           — загрузить .rfa файл в проект
    ///   list_shared_params    — список общих параметров из ФОП (файла общих параметров)
    ///   add_shared_param      — добавить общий параметр к категориям
    ///   remove_shared_param   — удалить общий параметр из привязок
    ///   list_family_types     — типоразмеры семейства
    ///   set_family_type_param — изменить параметр типоразмера
    ///   duplicate_type        — дублировать типоразмер
    ///   edit_family           — открыть семейство, изменить параметры, загрузить обратно
    /// </summary>
    internal static class FamilyHandler
    {
        public static object Handle(UIApplication uiApp, Dictionary<string, object> p)
        {
            var uidoc = uiApp.ActiveUIDocument;
            if (uidoc == null)
                return new { error = "Нет открытого документа" };

            var doc = uidoc.Document;
            string action = p.TryGetValue("action", out var a) ? a?.ToString() ?? "" : "";
            string dataStr = p.TryGetValue("data", out var d) ? d?.ToString() ?? "{}" : "{}";

            JObject data;
            try { data = JObject.Parse(dataStr); }
            catch { return new { error = $"Невалидный JSON в data: {dataStr}" }; }

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
                        error = $"Неизвестное действие: '{action}'",
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

        // ═══════════════════════════════════════════════════════════
        //  list_families — список загруженных семейств
        // ═══════════════════════════════════════════════════════════

        /// <summary>
        /// data: { "category": "Воздуховоды" } — фильтр по категории (необязательно)
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
                    category = f.FamilyCategory?.Name ?? "—",
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

        // ═══════════════════════════════════════════════════════════
        //  load_family — загрузить .rfa в проект
        // ═══════════════════════════════════════════════════════════

        /// <summary>
        /// data: { "path": "C:\\Families\\MyFamily.rfa", "overwrite": true }
        /// </summary>
        private static object LoadFamily(Document doc, JObject data)
        {
            string path = data.Value<string>("path") ?? "";
            bool overwrite = data.Value<bool?>("overwrite") ?? false;

            if (string.IsNullOrWhiteSpace(path))
                return new { error = "Укажите path — путь к .rfa файлу" };

            if (!File.Exists(path))
                return new { error = $"Файл не найден: {path}" };

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
                    return new { error = "Не удалось загрузить семейство. Возможно, уже загружено (используйте overwrite: true)." };
                }

                trans.Commit();

                return new
                {
                    action = "load_family",
                    loaded = true,
                    family_name = family?.Name ?? Path.GetFileNameWithoutExtension(path),
                    family_id = family?.Id.IntegerValue ?? -1,
                    category = family?.FamilyCategory?.Name ?? "—",
                    types_count = family?.GetFamilySymbolIds().Count ?? 0
                };
            }
        }

        // ═══════════════════════════════════════════════════════════
        //  list_shared_params — текущие привязки общих параметров
        // ═══════════════════════════════════════════════════════════

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

        // ═══════════════════════════════════════════════════════════
        //  add_shared_param — добавить общий параметр
        // ═══════════════════════════════════════════════════════════

        /// <summary>
        /// data: {
        ///   "param_name": "ADSK_Расход",
        ///   "group_name": "Группа в ФОП",      // группа в файле общих параметров
        ///   "categories": ["Воздуховоды", "Воздухораспределители"],
        ///   "is_instance": true,                // true=экземпляр, false=тип
        ///   "param_group": "Механизмы"          // группа отображения в свойствах (необязательно)
        /// }
        /// 
        /// Если ФОП не подключён, подключаем временный.
        /// </summary>
        private static object AddSharedParam(Document doc, Autodesk.Revit.ApplicationServices.Application app, JObject data)
        {
            string paramName = data.Value<string>("param_name") ?? "";
            string groupName = data.Value<string>("group_name") ?? "STB2026";
            bool isInstance = data.Value<bool?>("is_instance") ?? true;
            string paramGroupName = data.Value<string>("param_group") ?? "";

            var categoriesArr = data["categories"] as JArray;
            if (string.IsNullOrWhiteSpace(paramName))
                return new { error = "Укажите param_name" };
            if (categoriesArr == null || categoriesArr.Count == 0)
                return new { error = "Укажите categories — массив имён категорий" };

            // Проверяем — может параметр уже привязан
            var existingBinding = FindExistingBinding(doc, paramName);
            if (existingBinding != null)
            {
                return new
                {
                    error = $"Параметр '{paramName}' уже привязан к проекту",
                    existing_categories = GetBindingCategories(existingBinding)
                };
            }

            // Работа с ФОП
            string originalSpfPath = app.SharedParametersFilename;
            string tempSpfPath = null;
            bool usedTempSpf = false;

            try
            {
                DefinitionFile defFile = null;

                // Пробуем использовать текущий ФОП
                if (!string.IsNullOrWhiteSpace(originalSpfPath) && File.Exists(originalSpfPath))
                {
                    defFile = app.OpenSharedParameterFile();
                }

                // Если ФОП нет — создаём временный
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
                    return new { error = "Не удалось открыть/создать файл общих параметров (ФОП)" };

                // Находим или создаём группу
                DefinitionGroup defGroup = defFile.Groups.get_Item(groupName);
                if (defGroup == null)
                    defGroup = defFile.Groups.Create(groupName);

                // Находим или создаём определение параметра
                ExternalDefinition extDef = defGroup.Definitions.get_Item(paramName) as ExternalDefinition;
                if (extDef == null)
                {
                    var options = new ExternalDefinitionCreationOptions(paramName, SpecTypeId.String.Text);
                    options.UserModifiable = true;
                    options.Visible = true;

                    // Угадываем тип по имени
                    if (paramName.Contains("Расход", StringComparison.OrdinalIgnoreCase))
                        options = new ExternalDefinitionCreationOptions(paramName, SpecTypeId.AirFlow);
                    else if (paramName.Contains("Скорость", StringComparison.OrdinalIgnoreCase))
                        options = new ExternalDefinitionCreationOptions(paramName, SpecTypeId.HvacVelocity);
                    else if (paramName.Contains("Давлен", StringComparison.OrdinalIgnoreCase))
                        options = new ExternalDefinitionCreationOptions(paramName, SpecTypeId.HvacPressure);
                    else if (paramName.Contains("Площад", StringComparison.OrdinalIgnoreCase))
                        options = new ExternalDefinitionCreationOptions(paramName, SpecTypeId.Area);
                    else if (paramName.Contains("Длин", StringComparison.OrdinalIgnoreCase) ||
                             paramName.Contains("Высот", StringComparison.OrdinalIgnoreCase) ||
                             paramName.Contains("Ширин", StringComparison.OrdinalIgnoreCase) ||
                             paramName.Contains("Диаметр", StringComparison.OrdinalIgnoreCase))
                        options = new ExternalDefinitionCreationOptions(paramName, SpecTypeId.Length);
                    else if (paramName.Contains("Масс", StringComparison.OrdinalIgnoreCase))
                        options = new ExternalDefinitionCreationOptions(paramName, SpecTypeId.Mass);

                    extDef = defGroup.Definitions.Create(options) as ExternalDefinition;
                }

                if (extDef == null)
                    return new { error = $"Не удалось создать определение параметра '{paramName}'" };

                // Собираем CategorySet
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
                    return new { error = "Ни одна из указанных категорий не найдена или не поддерживает параметры" };

                // Создаём привязку
                using (var trans = new Transaction(doc, $"STB2026: Add shared param '{paramName}'"))
                {
                    trans.Start();

                    ElementBinding binding = isInstance
                        ? (ElementBinding)doc.Application.Create.NewInstanceBinding(catSet)
                        : (ElementBinding)doc.Application.Create.NewTypeBinding(catSet);

                    // Группа отображения
                    ForgeTypeId groupTypeId = GroupTypeId.General;
                    if (!string.IsNullOrWhiteSpace(paramGroupName))
                        groupTypeId = GuessGroupTypeId(paramGroupName);

                    bool bound = doc.ParameterBindings.Insert(extDef, binding, groupTypeId);

                    if (!bound)
                    {
                        trans.RollBack();
                        return new { error = $"Не удалось привязать параметр '{paramName}' к категориям" };
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
                // Возвращаем оригинальный ФОП
                if (usedTempSpf && !string.IsNullOrWhiteSpace(originalSpfPath))
                    app.SharedParametersFilename = originalSpfPath;
            }
        }

        // ═══════════════════════════════════════════════════════════
        //  remove_shared_param — удалить привязку общего параметра
        // ═══════════════════════════════════════════════════════════

        /// <summary>
        /// data: { "param_name": "ADSK_Расход" }
        /// </summary>
        private static object RemoveSharedParam(Document doc, JObject data)
        {
            string paramName = data.Value<string>("param_name") ?? "";
            if (string.IsNullOrWhiteSpace(paramName))
                return new { error = "Укажите param_name" };

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
                return new { error = $"Параметр '{paramName}' не найден в привязках" };

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
                    return new { error = $"Не удалось удалить привязку '{paramName}'" };
                }
            }
        }

        // ═══════════════════════════════════════════════════════════
        //  list_family_types — типоразмеры семейства
        // ═══════════════════════════════════════════════════════════

        /// <summary>
        /// data: { "family_name": "ADSK_Диффузор_Круглый_Приточный" }
        /// или: { "family_id": 12345 }
        /// </summary>
        private static object ListFamilyTypes(Document doc, JObject data)
        {
            Family family = FindFamily(doc, data);
            if (family == null)
                return new { error = "Семейство не найдено. Укажите family_name или family_id." };

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

        // ═══════════════════════════════════════════════════════════
        //  set_family_type_param — изменить параметр типоразмера
        // ═══════════════════════════════════════════════════════════

        /// <summary>
        /// data: { "type_id": 12345, "param_name": "Ширина", "param_value": "600" }
        /// или: { "family_name": "...", "type_name": "Ф100", "param_name": "...", "param_value": "..." }
        /// </summary>
        private static object SetFamilyTypeParam(Document doc, JObject data)
        {
            FamilySymbol symbol = FindFamilySymbol(doc, data);
            if (symbol == null)
                return new { error = "Типоразмер не найден. Укажите type_id или family_name+type_name." };

            string paramName = data.Value<string>("param_name") ?? "";
            string paramValue = data.Value<string>("param_value") ?? "";

            if (string.IsNullOrWhiteSpace(paramName))
                return new { error = "Укажите param_name" };

            var param = symbol.LookupParameter(paramName);
            if (param == null)
                return new { error = $"Параметр '{paramName}' не найден в типоразмере '{symbol.Name}'" };

            if (param.IsReadOnly)
                return new { error = $"Параметр '{paramName}' доступен только для чтения" };

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
                    return new { error = $"Ошибка установки параметра: {ex.Message}" };
                }
            }
        }

        // ═══════════════════════════════════════════════════════════
        //  duplicate_type — дублировать типоразмер
        // ═══════════════════════════════════════════════════════════

        /// <summary>
        /// data: { "type_id": 12345, "new_name": "Ф150" }
        /// или: { "family_name": "...", "type_name": "Ф100", "new_name": "Ф150" }
        /// </summary>
        private static object DuplicateType(Document doc, JObject data)
        {
            FamilySymbol symbol = FindFamilySymbol(doc, data);
            if (symbol == null)
                return new { error = "Типоразмер не найден" };

            string newName = data.Value<string>("new_name") ?? "";
            if (string.IsNullOrWhiteSpace(newName))
                return new { error = "Укажите new_name — имя нового типоразмера" };

            using (var trans = new Transaction(doc, $"STB2026: Duplicate type '{newName}'"))
            {
                trans.Start();

                var newType = symbol.Duplicate(newName) as FamilySymbol;
                if (newType == null)
                {
                    trans.RollBack();
                    return new { error = $"Не удалось дублировать типоразмер" };
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

        // ═══════════════════════════════════════════════════════════
        //  edit_family — редактирование семейства
        // ═══════════════════════════════════════════════════════════

        /// <summary>
        /// Открывает семейство для редактирования, меняет параметры, загружает обратно.
        /// data: {
        ///   "family_name": "ADSK_Диффузор_Круглый_Приточный",
        ///   "params": [
        ///     { "name": "ОВ_Расход_Номинальный", "value": "500", "is_instance": true },
        ///     { "name": "Описание", "value": "Приточный диффузор" }
        ///   ]
        /// }
        /// </summary>
        private static object EditFamily(Document doc, UIApplication uiApp, JObject data)
        {
            Family family = FindFamily(doc, data);
            if (family == null)
                return new { error = "Семейство не найдено" };

            if (!family.IsEditable)
                return new { error = $"Семейство '{family.Name}' недоступно для редактирования" };

            var paramsArr = data["params"] as JArray;
            if (paramsArr == null || paramsArr.Count == 0)
                return new { error = "Укажите params — массив параметров для изменения" };

            // Открываем документ семейства
            Document famDoc = doc.EditFamily(family);
            if (famDoc == null)
                return new { error = "Не удалось открыть документ семейства" };

            try
            {
                var famMgr = famDoc.FamilyManager;
                int changed = 0;
                int added = 0;
                var errors = new List<string>();

                using (var trans = new Transaction(famDoc, "STB2026: Edit Family"))
                {
                    trans.Start();

                    foreach (var paramToken in paramsArr)
                    {
                        string pName = paramToken.Value<string>("name") ?? "";
                        string pValue = paramToken.Value<string>("value") ?? "";
                        bool pInstance = paramToken.Value<bool?>("is_instance") ?? true;

                        if (string.IsNullOrWhiteSpace(pName)) continue;

                        // Ищем параметр
                        FamilyParameter famParam = null;
                        foreach (FamilyParameter fp in famMgr.Parameters)
                        {
                            if (fp.Definition.Name == pName)
                            {
                                famParam = fp;
                                break;
                            }
                        }

                        // Если параметра нет — создаём
                        if (famParam == null)
                        {
                            try
                            {
                                famParam = famMgr.AddParameter(
                                    pName,
                                    GroupTypeId.General,
                                    SpecTypeId.String.Text,
                                    pInstance);
                                added++;
                            }
                            catch (Exception ex)
                            {
                                errors.Add($"'{pName}': не удалось создать: {ex.Message}");
                                continue;
                            }
                        }

                        // Устанавливаем значение
                        if (!string.IsNullOrWhiteSpace(pValue))
                        {
                            try
                            {
                                SetFamilyParamValue(famMgr, famParam, pValue);
                                changed++;
                            }
                            catch (Exception ex)
                            {
                                errors.Add($"'{pName}': {ex.Message}");
                            }
                        }
                    }

                    trans.Commit();
                }

                // Загружаем обратно в проект
                var loadOptions = new FamilyLoadOptions();
                Family reloaded = famDoc.LoadFamily(doc, loadOptions);

                famDoc.Close(false);

                return new
                {
                    action = "edit_family",
                    family_name = family.Name,
                    params_changed = changed,
                    params_added = added,
                    reloaded = reloaded != null,
                    errors
                };
            }
            catch (Exception ex)
            {
                try { famDoc.Close(false); } catch { }
                return new { error = $"Ошибка редактирования семейства: {ex.Message}" };
            }
        }

        // ═══════════════════════════════════════════════════════════
        //  Вспомогательные методы
        // ═══════════════════════════════════════════════════════════

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

        private static Category FindCategory(Document doc, string name)
        {
            foreach (Category c in doc.Settings.Categories)
            {
                if (c != null && c.Name.IndexOf(name, StringComparison.OrdinalIgnoreCase) >= 0)
                    return c;
            }
            return null;
        }

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

        private static string GetParamGroup(Definition def)
        {
            try
            {
                var groupId = def.GetGroupTypeId();
                return groupId != null ? LabelUtils.GetLabelForGroup(groupId) : "Прочее";
            }
            catch { return "Прочее"; }
        }

        private static ForgeTypeId GuessGroupTypeId(string groupName)
        {
            string lower = groupName.ToLowerInvariant();

            if (lower.Contains("механ") || lower.Contains("hvac") || lower.Contains("mech"))
                return GroupTypeId.Mechanical;
            if (lower.Contains("разм") || lower.Contains("dimen"))
                return GroupTypeId.Geometry;
            if (lower.Contains("идентиф") || lower.Contains("ident"))
                return GroupTypeId.IdentityData;
            if (lower.Contains("данн") || lower.Contains("data"))
                return GroupTypeId.Data;
            if (lower.Contains("завис") || lower.Contains("constr"))
                return GroupTypeId.Constraints;

            return GroupTypeId.General;
        }

        private static void SetParamValue(Parameter param, string value)
        {
            switch (param.StorageType)
            {
                case StorageType.String:
                    param.Set(value);
                    break;
                case StorageType.Integer:
                    if (int.TryParse(value, out int i)) param.Set(i);
                    else throw new ArgumentException($"Ожидалось целое число: {value}");
                    break;
                case StorageType.Double:
                    param.SetValueString(value);
                    break;
                case StorageType.ElementId:
                    if (int.TryParse(value, out int eid)) param.Set(new ElementId(eid));
                    else throw new ArgumentException($"Ожидался ID: {value}");
                    break;
            }
        }

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

        /// <summary>Опции загрузки семейства с перезаписью.</summary>
        private class FamilyLoadOptions : IFamilyLoadOptions
        {
            public bool OnFamilyFound(bool familyInUse, out bool overwriteParameterValues)
            {
                overwriteParameterValues = true;
                return true; // Перезаписать
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
