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
    ///   add_shared_param      — добавить общий параметр к категориям проекта
    ///   remove_shared_param   — удалить общий параметр из привязок проекта
    ///   list_family_types     — типоразмеры семейства
    ///   set_family_type_param — изменить параметр типоразмера
    ///   duplicate_type        — дублировать типоразмер
    ///   edit_family           — открыть семейство, изменить/добавить параметры (вкл. общие из ФОП), загрузить обратно
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
        //  add_shared_param — добавить общий параметр к категориям проекта
        // ═══════════════════════════════════════════════════════════

        /// <summary>
        /// data: {
        ///   "param_name": "ADSK_Расход",
        ///   "group_name": "Группа в ФОП",      // группа в файле общих параметров (по умолчанию "STB2026")
        ///   "categories": ["Воздуховоды", "Воздухораспределители"],
        ///   "is_instance": true,                // true=экземпляр, false=тип
        ///   "param_group": "Механизмы"          // группа отображения в свойствах
        /// }
        /// Если ФОП не подключён, создаём временный.
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
                    if (paramName.IndexOf("Расход", StringComparison.OrdinalIgnoreCase) >= 0 ||
                        paramName.IndexOf("AIRFLOW", StringComparison.OrdinalIgnoreCase) >= 0)
                        options = new ExternalDefinitionCreationOptions(paramName, SpecTypeId.AirFlow);
                    else if (paramName.IndexOf("Скорость", StringComparison.OrdinalIgnoreCase) >= 0 ||
                             paramName.IndexOf("VELOCITY", StringComparison.OrdinalIgnoreCase) >= 0)
                        options = new ExternalDefinitionCreationOptions(paramName, SpecTypeId.HvacVelocity);
                    else if (paramName.IndexOf("Давлен", StringComparison.OrdinalIgnoreCase) >= 0 ||
                             paramName.IndexOf("PRESSURE", StringComparison.OrdinalIgnoreCase) >= 0)
                        options = new ExternalDefinitionCreationOptions(paramName, SpecTypeId.HvacPressure);
                    else if (paramName.IndexOf("Площад", StringComparison.OrdinalIgnoreCase) >= 0)
                        options = new ExternalDefinitionCreationOptions(paramName, SpecTypeId.Area);
                    else if (paramName.IndexOf("Длин", StringComparison.OrdinalIgnoreCase) >= 0 ||
                             paramName.IndexOf("Высот", StringComparison.OrdinalIgnoreCase) >= 0 ||
                             paramName.IndexOf("Ширин", StringComparison.OrdinalIgnoreCase) >= 0 ||
                             paramName.IndexOf("Диаметр", StringComparison.OrdinalIgnoreCase) >= 0)
                        options = new ExternalDefinitionCreationOptions(paramName, SpecTypeId.Length);
                    else if (paramName.IndexOf("Масс", StringComparison.OrdinalIgnoreCase) >= 0)
                        options = new ExternalDefinitionCreationOptions(paramName, SpecTypeId.Mass);

                    options.UserModifiable = true;
                    options.Visible = true;
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

            Parameter param = symbol.LookupParameter(paramName);
            if (param == null)
                return new { error = $"Параметр '{paramName}' не найден у типоразмера '{symbol.Name}'" };

            if (param.IsReadOnly)
                return new { error = $"Параметр '{paramName}' только для чтения" };

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
                    return new { error = $"Ошибка: {ex.Message}" };
                }
            }
        }

        // ═══════════════════════════════════════════════════════════
        //  duplicate_type — дублировать типоразмер
        // ═══════════════════════════════════════════════════════════

        /// <summary>
        /// data: { "family_name": "...", "type_name": "Ф100", "new_name": "Ф200" }
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
        //  Поддерживает добавление общих параметров из ФОП (is_shared)
        // ═══════════════════════════════════════════════════════════

        /// <summary>
        /// Открывает семейство для редактирования, добавляет/изменяет параметры, загружает обратно.
        /// data: {
        ///   "family_name": "ADSK_Диффузор_Круглый_Приточный",
        ///   "params": [
        ///     { "name": "LIN_VE_PRESSURE_LOSS", "is_instance": true, "is_shared": true, "group": "Механизмы\u00A0— расход" },
        ///     { "name": "CustomParam", "value": "500", "is_instance": true },
        ///     { "name": "Описание", "value": "Приточный диффузор" }
        ///   ]
        /// }
        /// 
        /// Параметры каждого элемента массива params:
        ///   name        — имя параметра (обязательно)
        ///   value       — значение (необязательно)
        ///   is_instance — true=экземпляр, false=тип (по умолчанию true)
        ///   is_shared   — true=искать в ФОП как общий параметр (по умолчанию false)
        ///   group       — группа отображения: "Механизмы", "Механизмы— расход", "Данные" и т.д.
        ///   spec_type   — тип данных для НЕобщих параметров: "text","length","airflow","pressure" и т.д.
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
                return new { error = "Укажите params — массив параметров для изменения/добавления" };

            // ═══ Подход через временный .rfa файл ═══
            // doc.EditFamily() создаёт in-memory документ, в котором ExternalDefinition
            // из ФОП становится невалидным. Поэтому:
            // 1. EditFamily → получаем famDoc
            // 2. Сохраняем famDoc во временный .rfa
            // 3. Закрываем famDoc
            // 4. Открываем .rfa через Application.OpenDocumentFile()
            // 5. Добавляем параметры (ExternalDefinition валиден в этом контексте)
            // 6. Сохраняем .rfa
            // 7. Закрываем .rfa
            // 8. Загружаем в проект через doc.LoadFamily(path)

            // Временный .rfa — ВАЖНО: имя файла ДОЛЖНО совпадать с именем семейства,
            // иначе doc.LoadFamily() создаст НОВОЕ семейство вместо перезаписи
            string tempDir = Path.Combine(Path.GetTempPath(), "STB2026_Families");
            if (!Directory.Exists(tempDir)) Directory.CreateDirectory(tempDir);
            string tempRfaPath = Path.Combine(tempDir, $"{family.Name}.rfa");

            Document famDoc = null;
            Document openedFamDoc = null;

            try
            {
                // ═══ Шаг 1-3: Получаем .rfa на диск ═══
                famDoc = doc.EditFamily(family);
                if (famDoc == null)
                    return new { error = "Не удалось открыть документ семейства" };

                var saveOpts = new SaveAsOptions { OverwriteExistingFile = true };
                famDoc.SaveAs(tempRfaPath, saveOpts);
                famDoc.Close(false);
                famDoc = null;

                // ═══ Шаг 4: Открываем .rfa как обычный документ ═══
                openedFamDoc = uiApp.Application.OpenDocumentFile(tempRfaPath);
                if (openedFamDoc == null)
                    return new { error = "Не удалось открыть временный .rfa файл" };

                var famMgr = openedFamDoc.FamilyManager;
                int changed = 0;
                int added = 0;
                var errors = new List<string>();
                var addedParams = new List<object>();

                // ═══ Шаг 5: Добавляем/изменяем параметры ═══
                using (var trans = new Transaction(openedFamDoc, "STB2026: Edit Family"))
                {
                    trans.Start();

                    // Убедимся что есть хотя бы один тип
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

                        // Ищем существующий параметр в семействе
                        FamilyParameter famParam = null;
                        foreach (FamilyParameter fp in famMgr.Parameters)
                        {
                            if (fp.Definition.Name == pName)
                            {
                                famParam = fp;
                                break;
                            }
                        }

                        // Если параметр уже существует — только меняем значение
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
                                    errors.Add($"'{pName}': ошибка установки значения: {ex.Message}");
                                }
                            }
                            else
                            {
                                errors.Add($"'{pName}': параметр уже существует в семействе");
                            }
                            continue;
                        }

                        // === Параметра нет — создаём ===

                        ForgeTypeId groupId = string.IsNullOrWhiteSpace(pGroup)
                            ? GroupTypeId.General
                            : GuessGroupTypeId(pGroup);

                        if (pShared)
                        {
                            // ─── Добавление ОБЩЕГО параметра из ФОП ───
                            try
                            {
                                // В контексте отдельно открытого .rfa ExternalDefinition валиден
                                ExternalDefinition extDef = FindSharedParamDefinition(
                                    uiApp.Application, pName);

                                if (extDef == null)
                                {
                                    errors.Add($"'{pName}': не найден в файле общих параметров (ФОП). " +
                                               "Проверьте, что ФОП подключён и параметр существует.");
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
                                            errors.Add($"'{pName}': добавлен, но ошибка значения: {ex.Message}");
                                        }
                                    }
                                }
                                else
                                {
                                    errors.Add($"'{pName}': FamilyManager.AddParameter вернул null");
                                }
                            }
                            catch (Exception ex)
                            {
                                errors.Add($"'{pName}': ошибка добавления общего параметра: {ex.Message}");
                            }
                        }
                        else
                        {
                            // ─── Добавление обычного параметра семейства ───
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
                                            errors.Add($"'{pName}': добавлен, но ошибка значения: {ex.Message}");
                                        }
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                errors.Add($"'{pName}': не удалось создать: {ex.Message}");
                            }
                        }
                    }

                    trans.Commit();
                }

                // ═══ Шаг 6-7: Сохраняем и закрываем .rfa ═══
                openedFamDoc.Save();
                openedFamDoc.Close(false);
                openedFamDoc = null;

                // ═══ Шаг 8: Загружаем обновлённое семейство в проект ═══
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
                return new { error = $"Ошибка редактирования семейства: {ex.Message}" };
            }
            finally
            {
                // Удаляем временный файл
                try { if (File.Exists(tempRfaPath)) File.Delete(tempRfaPath); } catch { }
            }
        }

        // ═══════════════════════════════════════════════════════════
        //  Вспомогательные методы
        // ═══════════════════════════════════════════════════════════

        /// <summary>Ищет семейство по ID или имени.</summary>
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

        /// <summary>Ищет FamilySymbol по type_id или family_name+type_name.</summary>
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

        /// <summary>Ищет категорию по имени (частичное совпадение).</summary>
        private static Category FindCategory(Document doc, string name)
        {
            foreach (Category c in doc.Settings.Categories)
            {
                if (c != null && c.Name.IndexOf(name, StringComparison.OrdinalIgnoreCase) >= 0)
                    return c;
            }
            return null;
        }

        /// <summary>Ищет существующую привязку параметра по имени.</summary>
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

        /// <summary>Возвращает список категорий из привязки.</summary>
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

        /// <summary>Возвращает имя группы параметра.</summary>
        private static string GetParamGroup(Definition def)
        {
            try
            {
                var groupId = def.GetGroupTypeId();
                return groupId != null ? LabelUtils.GetLabelForGroup(groupId) : "Прочее";
            }
            catch { return "Прочее"; }
        }

        /// <summary>
        /// Ищет ExternalDefinition по имени во всех группах текущего файла общих параметров.
        /// Если ФОП не подключён — пытается найти его по стандартным путям.
        /// </summary>
        private static ExternalDefinition FindSharedParamDefinition(
            Autodesk.Revit.ApplicationServices.Application app, string paramName)
        {
            DefinitionFile defFile = app.OpenSharedParameterFile();

            if (defFile == null)
            {
                // Пробуем найти ФОП по стандартным путям
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
                        catch { /* пробуем следующий */ }
                    }
                }
            }

            if (defFile == null)
                return null;

            // Ищем параметр во всех группах
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

        /// <summary>Определяет группу отображения параметра по текстовому имени.</summary>
        private static ForgeTypeId GuessGroupTypeId(string groupName)
        {
            if (string.IsNullOrWhiteSpace(groupName))
                return GroupTypeId.General;

            string lower = groupName.ToLowerInvariant();

            // "Механизмы — расход" / "Механизмы\u00A0— расход" (с неразрывным пробелом)
            if (lower.Contains("механ") && (lower.Contains("расход") || lower.Contains("flow")))
                return GroupTypeId.MechanicalAirflow;

            // "Механизмы — нагрузки"
            if (lower.Contains("механ") && (lower.Contains("нагрузк") || lower.Contains("load")))
                return GroupTypeId.MechanicalLoads;

            // "Механизмы" (общее)
            if (lower.Contains("механ") || lower.Contains("hvac") || lower.Contains("mech"))
                return GroupTypeId.Mechanical;

            if (lower.Contains("разм") || lower.Contains("dimen") || lower.Contains("geom"))
                return GroupTypeId.Geometry;

            if (lower.Contains("идентиф") || lower.Contains("ident"))
                return GroupTypeId.IdentityData;

            if (lower.Contains("данн") || lower.Contains("data"))
                return GroupTypeId.Data;

            if (lower.Contains("завис") || lower.Contains("constr"))
                return GroupTypeId.Constraints;

            if (lower.Contains("строит") || lower.Contains("construct"))
                return GroupTypeId.Construction;

            if (lower.Contains("текст") || lower.Contains("text"))
                return GroupTypeId.Text;

            if (lower.Contains("общ") || lower.Contains("general"))
                return GroupTypeId.General;

            if (lower.Contains("электр") || lower.Contains("electr"))
                return GroupTypeId.Electrical;

            if (lower.Contains("сантех") || lower.Contains("plumb"))
                return GroupTypeId.Plumbing;

            if (lower.Contains("ifc"))
                return GroupTypeId.Ifc;

            return GroupTypeId.General;
        }

        /// <summary>Определяет тип данных параметра по явному указанию или по имени.</summary>
        private static ForgeTypeId GuessSpecTypeId(string specType, string paramName)
        {
            // Если явно указан тип
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

            // Угадываем по имени параметра
            string lower = paramName.ToLowerInvariant();

            if (lower.Contains("расход") || lower.Contains("airflow") || lower.Contains("air_flow"))
                return SpecTypeId.AirFlow;
            if (lower.Contains("давлен") || lower.Contains("pressure"))
                return SpecTypeId.HvacPressure;
            if (lower.Contains("скорост") || lower.Contains("velocity"))
                return SpecTypeId.HvacVelocity;
            if (lower.Contains("площад") || lower.Contains("area"))
                return SpecTypeId.Area;
            if (lower.Contains("длин") || lower.Contains("высот") || lower.Contains("ширин") ||
                lower.Contains("диаметр") || lower.Contains("length") || lower.Contains("height") ||
                lower.Contains("width") || lower.Contains("diameter"))
                return SpecTypeId.Length;
            if (lower.Contains("масс") || lower.Contains("mass") || lower.Contains("weight"))
                return SpecTypeId.Mass;
            if (lower.Contains("температур") || lower.Contains("temperature"))
                return SpecTypeId.HvacTemperature;
            if (lower.Contains("мощност") || lower.Contains("power"))
                return SpecTypeId.HvacPower;
            if (lower.Contains("объём") || lower.Contains("объем") || lower.Contains("volume"))
                return SpecTypeId.Volume;

            return SpecTypeId.String.Text;
        }

        /// <summary>Получает читаемое имя группы параметров.</summary>
        private static string GetGroupLabel(ForgeTypeId groupId)
        {
            try
            {
                return LabelUtils.GetLabelForGroup(groupId);
            }
            catch
            {
                return groupId?.TypeId ?? "Прочее";
            }
        }

        /// <summary>Устанавливает значение параметра элемента.</summary>
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

        /// <summary>Устанавливает значение параметра семейства через FamilyManager.</summary>
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
