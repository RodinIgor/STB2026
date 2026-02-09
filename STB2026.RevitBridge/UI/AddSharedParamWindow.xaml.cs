using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace STB2026.RevitBridge.UI
{
    public partial class AddSharedParamWindow : Window
    {
        private readonly UIApplication _uiApp;
        private readonly Document _doc;

        private List<SpfGroupInfo> _spfGroups = new List<SpfGroupInfo>();
        private List<SpfParamInfo> _spfParams = new List<SpfParamInfo>();
        private List<CategoryInfo> _categories = new List<CategoryInfo>();
        private List<FamilyCheckItem> _familyCheckItems = new List<FamilyCheckItem>();
        private List<string> _familyParams = new List<string>();

        public AddSharedParamWindow(UIApplication uiApp)
        {
            InitializeComponent();
            _uiApp = uiApp;
            _doc = uiApp.ActiveUIDocument.Document;
            LoadSpfGroups();
            LoadCategories();
            LoadDisplayGroups();
        }

        // ═══════════════════════════════════════════════════════════
        //  Загрузка данных
        // ═══════════════════════════════════════════════════════════

        private void LoadSpfGroups()
        {
            _spfGroups.Clear();
            DefinitionFile defFile = null;
            try { defFile = _uiApp.Application.OpenSharedParameterFile(); } catch { }

            if (defFile == null)
            {
                cmbSpfGroup.ItemsSource = null;
                System.Windows.MessageBox.Show(
                    "Файл общих параметров (ФОП) не подключён.\n\n" +
                    "Подключите ФОП: Управление → Файл общих параметров",
                    "ФОП не найден", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            foreach (DefinitionGroup group in defFile.Groups)
            {
                _spfGroups.Add(new SpfGroupInfo
                {
                    Name = group.Name,
                    Group = group,
                    Count = group.Definitions.Size
                });
            }
            _spfGroups = _spfGroups.OrderBy(g => g.Name).ToList();
            cmbSpfGroup.ItemsSource = _spfGroups;
            cmbSpfGroup.DisplayMemberPath = "DisplayText";
        }

        private void LoadCategories()
        {
            _categories.Clear();

            // Собираем все категории, в которых есть загруженные семейства
            var familyCategories = new FilteredElementCollector(_doc)
                .OfClass(typeof(Family))
                .Cast<Family>()
                .Where(f => f.FamilyCategory != null && !f.IsInPlace && f.IsEditable)
                .Select(f => f.FamilyCategory)
                .GroupBy(c => c.Id.Value)
                .Select(g => g.First())
                .OrderBy(c => c.Name);

            foreach (var cat in familyCategories)
            {
                _categories.Add(new CategoryInfo
                {
                    BuiltIn = (BuiltInCategory)cat.Id.Value,
                    Category = cat,
                    DisplayName = cat.Name
                });
            }

            cmbCategory.ItemsSource = _categories;
            cmbCategory.DisplayMemberPath = "DisplayName";
        }

        private void LoadDisplayGroups()
        {
            var groups = new List<string>
            {
                "Общие", "Данные", "Идентификация", "Размеры",
                "Зависимости", "Строительство", "Механизмы",
                "Механизмы\u00A0— расход", "Механизмы\u00A0— нагрузки",
                "Электросети", "Сантехника", "Текст", "IFC-параметры"
            };
            cmbDisplayGroup.ItemsSource = groups;
            cmbDisplayGroup.SelectedIndex = 0;
        }

        private void LoadFamiliesByCategory(BuiltInCategory bic)
        {
            _familyCheckItems.Clear();
            pnlFamilies.Children.Clear();

            var families = new FilteredElementCollector(_doc)
                .OfClass(typeof(Family))
                .Cast<Family>()
                .Where(f => f.FamilyCategory != null
                    && (BuiltInCategory)f.FamilyCategory.Id.Value == bic
                    && f.IsEditable && !f.IsInPlace)
                .OrderBy(f => f.Name)
                .ToList();

            foreach (var fam in families)
            {
                var item = new FamilyCheckItem
                {
                    FamilyName = fam.Name,
                    FamilyId = fam.Id.Value,
                    IsChecked = false
                };

                var cb = new CheckBox
                {
                    Content = fam.Name,
                    Tag = item,
                    FontSize = 12,
                    Margin = new Thickness(4, 2, 4, 2),
                    IsChecked = chkAllFamilies.IsChecked == true
                };
                cb.Checked += FamilyCb_Changed;
                cb.Unchecked += FamilyCb_Changed;

                item.IsChecked = chkAllFamilies.IsChecked == true;
                item.CheckBox = cb;
                _familyCheckItems.Add(item);
                pnlFamilies.Children.Add(cb);
            }

            // Загружаем параметры первого семейства для связки
            if (families.Count > 0)
                LoadFamilyParamsForLink(families[0]);
            else
            {
                _familyParams.Clear();
                _familyParams.Add("(не связывать)");
                cmbLinkParam.ItemsSource = _familyParams;
                cmbLinkParam.SelectedIndex = 0;
            }

            UpdateAddButton();
        }

        private void LoadFamilyParamsForLink(Family family)
        {
            _familyParams.Clear();
            _familyParams.Add("(не связывать)");
            try
            {
                Document famDoc = _doc.EditFamily(family);
                if (famDoc != null)
                {
                    var famMgr = famDoc.FamilyManager;
                    var sorted = new List<string>();
                    foreach (FamilyParameter fp in famMgr.Parameters)
                    {
                        if (fp?.Definition?.Name != null)
                        {
                            string suffix = fp.IsInstance ? " [экз]" : " [тип]";
                            string storage = fp.StorageType.ToString();
                            sorted.Add($"{fp.Definition.Name}{suffix}  ({storage})");
                        }
                    }
                    sorted.Sort();
                    _familyParams.AddRange(sorted);
                    famDoc.Close(false);
                }
            }
            catch { }

            cmbLinkParam.ItemsSource = _familyParams;
            cmbLinkParam.SelectedIndex = 0;
        }

        // ═══════════════════════════════════════════════════════════
        //  Обработчики UI
        // ═══════════════════════════════════════════════════════════

        private void CmbSpfGroup_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (cmbSpfGroup.SelectedItem is SpfGroupInfo gi)
            {
                _spfParams.Clear();
                foreach (Definition def in gi.Group.Definitions)
                {
                    if (def is ExternalDefinition extDef)
                    {
                        string dataType = "text";
                        try { dataType = extDef.GetDataType()?.TypeId ?? "text"; } catch { }

                        string shortType = dataType;
                        int lastColon = dataType.LastIndexOf(':');
                        if (lastColon >= 0 && lastColon < dataType.Length - 1)
                            shortType = dataType.Substring(lastColon + 1).Replace("-2.0.0", "");

                        string guidStr = "";
                        try { guidStr = extDef.GUID.ToString(); } catch { }

                        _spfParams.Add(new SpfParamInfo
                        {
                            Name = def.Name,
                            GuidString = guidStr,
                            DataType = dataType,
                            ShortType = shortType,
                            GroupName = gi.Name,
                            DisplayText = $"{def.Name}  ({shortType})"
                        });
                    }
                }
                // Сортировка по алфавиту
                _spfParams = _spfParams.OrderBy(p => p.Name).ToList();
                cmbParam.ItemsSource = _spfParams;
                cmbParam.DisplayMemberPath = "DisplayText";
            }
            UpdateAddButton();
        }

        private void CmbParam_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (cmbParam.SelectedItem is SpfParamInfo pi)
            {
                pnlParamInfo.Visibility = System.Windows.Visibility.Visible;
                txtParamInfo.Text = $"GUID: {pi.GuidString}\nТип данных: {pi.DataType}";

                string lower = pi.Name.ToLowerInvariant();
                if (lower.Contains("расход") || lower.Contains("airflow") || lower.Contains("flow"))
                    cmbDisplayGroup.SelectedItem = "Механизмы\u00A0— расход";
                else if (lower.Contains("давлен") || lower.Contains("pressure"))
                    cmbDisplayGroup.SelectedItem = "Механизмы\u00A0— расход";
                else if (lower.Contains("механ") || lower.Contains("hvac"))
                    cmbDisplayGroup.SelectedItem = "Механизмы";
                else if (lower.Contains("идентиф") || lower.Contains("марк"))
                    cmbDisplayGroup.SelectedItem = "Идентификация";
            }
            else
            {
                pnlParamInfo.Visibility = System.Windows.Visibility.Collapsed;
            }
            UpdateAddButton();
        }

        private void CmbCategory_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (cmbCategory.SelectedItem is CategoryInfo ci)
                LoadFamiliesByCategory(ci.BuiltIn);
            UpdateAddButton();
        }

        private void ChkAllFamilies_Changed(object sender, RoutedEventArgs e)
        {
            bool isAll = chkAllFamilies.IsChecked == true;
            foreach (var item in _familyCheckItems)
            {
                item.IsChecked = isAll;
                if (item.CheckBox != null)
                    item.CheckBox.IsChecked = isAll;
            }
            UpdateAddButton();
        }

        private void FamilyCb_Changed(object sender, RoutedEventArgs e)
        {
            if (sender is CheckBox cb && cb.Tag is FamilyCheckItem item)
                item.IsChecked = cb.IsChecked == true;
            UpdateAddButton();
        }

        private void UpdateAddButton()
        {
            bool hasParam = cmbParam.SelectedItem is SpfParamInfo;
            bool hasFamilies = _familyCheckItems.Any(f => f.IsChecked);
            btnAdd.IsEnabled = hasParam && hasFamilies;

            int count = _familyCheckItems.Count(f => f.IsChecked);
            txtStatus.Text = count > 0 ? $"Выбрано семейств: {count}" : "";
        }

        // ═══════════════════════════════════════════════════════════
        //  Кнопка "Добавить"
        // ═══════════════════════════════════════════════════════════

        private void BtnAdd_Click(object sender, RoutedEventArgs e)
        {
            if (!(cmbParam.SelectedItem is SpfParamInfo pi)) return;

            bool isInstance = rbInstance.IsChecked == true;
            string displayGroup = cmbDisplayGroup.SelectedItem?.ToString() ?? "Общие";

            // Собираем данные в простые строки ДО любых операций с Revit
            string paramName = pi.Name;
            string paramGroupName = pi.GroupName;
            string paramGuid = pi.GuidString;

            var selectedFamilies = _familyCheckItems
                .Where(f => f.IsChecked)
                .Select(f => new { f.FamilyName, f.FamilyId })
                .ToList();

            if (selectedFamilies.Count == 0) return;

            string linkedParamName = null;
            if (cmbLinkParam.SelectedIndex > 0 && cmbLinkParam.SelectedItem is string linkStr)
            {
                int bracketIdx = linkStr.IndexOf(" [");
                linkedParamName = bracketIdx > 0 ? linkStr.Substring(0, bracketIdx) : linkStr;
            }

            // ═══════════════════════════════════════════════════════
            // FIX: Закрываем окно ПЕРЕД операцией.
            // После EditFamily/OpenDocumentFile/LoadFamily все объекты
            // WPF-привязок (Family, DefinitionGroup) становятся невалидными.
            // Если окно открыто — WPF пытается обновить биндинги и падает.
            // ═══════════════════════════════════════════════════════
            this.Hide();

            var results = new List<string>();
            int successCount = 0;
            int failCount = 0;

            foreach (var sel in selectedFamilies)
            {
                try
                {
                    // Ищем семейство по ID заново — после предыдущих LoadFamily
                    Family family = _doc.GetElement(new ElementId(sel.FamilyId)) as Family;
                    if (family == null)
                    {
                        results.Add($"✗ {sel.FamilyName}: семейство не найдено");
                        failCount++;
                        continue;
                    }

                    string result = AddSharedParamToFamily(
                        family, paramName, paramGroupName,
                        isInstance, displayGroup, linkedParamName);

                    results.Add(result);
                    if (result.StartsWith("✓")) successCount++;
                    else failCount++;
                }
                catch (Exception ex)
                {
                    results.Add($"⚠ {sel.FamilyName}: {ex.Message}");
                    failCount++;
                }
            }

            // Показываем результат обычным TaskDialog (после закрытия окна)
            string summary = $"Готово: {successCount} успешно, {failCount} ошибок\n" +
                             $"Параметр: {paramName}\n" +
                             $"GUID: {paramGuid}\n\n" +
                             string.Join("\n", results);

            TaskDialog.Show("STB2026 — Результат", summary);
            this.Close();
        }

        private void BtnClose_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        // ═══════════════════════════════════════════════════════════
        //  Основной метод — через временный .rfa файл
        // ═══════════════════════════════════════════════════════════

        private string AddSharedParamToFamily(
            Family family, string sharedParamName, string spfGroupName,
            bool isInstance, string displayGroup, string linkedParamName)
        {
            string familyName = family.Name;
            string tempDir = Path.Combine(Path.GetTempPath(), "STB2026_Families");
            Directory.CreateDirectory(tempDir);
            string tempRfaPath = Path.Combine(tempDir, $"{familyName}.rfa");

            Document famDoc = null;
            Document openedFamDoc = null;

            try
            {
                // 1. Экспортируем во временный .rfa
                famDoc = _doc.EditFamily(family);
                if (famDoc == null)
                    return $"✗ {familyName}: не удалось открыть";

                famDoc.SaveAs(tempRfaPath, new SaveAsOptions { OverwriteExistingFile = true });
                famDoc.Close(false);
                famDoc = null;

                // 2. Открываем .rfa
                openedFamDoc = _uiApp.Application.OpenDocumentFile(tempRfaPath);
                if (openedFamDoc == null)
                    return $"✗ {familyName}: не удалось открыть .rfa";

                var famMgr = openedFamDoc.FamilyManager;

                // 3. Проверяем дубликат
                foreach (FamilyParameter fp in famMgr.Parameters)
                {
                    if (fp.Definition.Name == sharedParamName)
                    {
                        openedFamDoc.Close(false);
                        return $"— {familyName}: параметр уже существует";
                    }
                }

                // 4. Получаем ExternalDefinition в контексте famDoc
                ExternalDefinition freshExtDef = null;
                DefinitionFile defFile = null;
                try { defFile = openedFamDoc.Application.OpenSharedParameterFile(); } catch { }
                if (defFile == null)
                    try { defFile = _uiApp.Application.OpenSharedParameterFile(); } catch { }

                if (defFile != null)
                {
                    DefinitionGroup targetGroup = defFile.Groups.get_Item(spfGroupName);
                    if (targetGroup != null)
                    {
                        Definition def = targetGroup.Definitions.get_Item(sharedParamName);
                        if (def is ExternalDefinition ed)
                            freshExtDef = ed;
                    }
                    if (freshExtDef == null)
                    {
                        foreach (DefinitionGroup group in defFile.Groups)
                        {
                            foreach (Definition def in group.Definitions)
                            {
                                if (def.Name == sharedParamName && def is ExternalDefinition ed)
                                { freshExtDef = ed; break; }
                            }
                            if (freshExtDef != null) break;
                        }
                    }
                }

                if (freshExtDef == null)
                {
                    openedFamDoc.Close(false);
                    return $"✗ {familyName}: параметр не найден в ФОП";
                }

                // 5. Группа отображения
                ForgeTypeId groupId = ResolveGroupTypeId(displayGroup);

                // 6. Добавляем
                using (var trans = new Transaction(openedFamDoc, "STB2026: Add Shared Param"))
                {
                    trans.Start();

                    if (famMgr.CurrentType == null)
                    {
                        if (famMgr.Types.Size == 0)
                            famMgr.NewType("Default");
                        else
                        {
                            var en = famMgr.Types.ForwardIterator();
                            en.MoveNext();
                            famMgr.CurrentType = en.Current as FamilyType;
                        }
                    }

                    FamilyParameter newParam = famMgr.AddParameter(freshExtDef, groupId, isInstance);
                    if (newParam == null)
                        throw new InvalidOperationException("AddParameter вернул null");

                    if (!string.IsNullOrWhiteSpace(linkedParamName))
                    {
                        try { famMgr.SetFormula(newParam, linkedParamName); }
                        catch { }
                    }

                    trans.Commit();
                }

                // 7. Сохраняем и закрываем
                openedFamDoc.Save();
                openedFamDoc.Close(false);
                openedFamDoc = null;

                // 8. Загружаем обратно
                Family reloadedFamily = null;
                using (var trans = new Transaction(_doc, $"STB2026: Reload {familyName}"))
                {
                    trans.Start();
                    _doc.LoadFamily(tempRfaPath, new OverwriteFamilyLoadOptions(), out reloadedFamily);
                    trans.Commit();
                }

                string linkedInfo = string.IsNullOrWhiteSpace(linkedParamName)
                    ? "" : $" (формула: {linkedParamName})";

                return $"✓ {familyName}: добавлен{linkedInfo}";
            }
            catch (Exception ex)
            {
                try { famDoc?.Close(false); } catch { }
                try { openedFamDoc?.Close(false); } catch { }
                return $"⚠ {familyName}: {ex.Message}";
            }
            finally
            {
                try { if (File.Exists(tempRfaPath)) File.Delete(tempRfaPath); } catch { }
            }
        }

        // ═══════════════════════════════════════════════════════════
        //  Вспомогательные методы
        // ═══════════════════════════════════════════════════════════

        private static ForgeTypeId ResolveGroupTypeId(string groupName)
        {
            if (string.IsNullOrWhiteSpace(groupName))
                return GroupTypeId.General;

            string lower = groupName.ToLowerInvariant();

            if (lower.Contains("механ") && (lower.Contains("расход") || lower.Contains("flow")))
                return GroupTypeId.MechanicalAirflow;
            if (lower.Contains("механ") && (lower.Contains("нагрузк") || lower.Contains("load")))
                return GroupTypeId.MechanicalLoads;
            if (lower.Contains("механ") || lower.Contains("hvac"))
                return GroupTypeId.Mechanical;
            if (lower.Contains("разм") || lower.Contains("dimen"))
                return GroupTypeId.Geometry;
            if (lower.Contains("идентиф") || lower.Contains("ident"))
                return GroupTypeId.IdentityData;
            if (lower.Contains("данн") || lower.Contains("data"))
                return GroupTypeId.Data;
            if (lower.Contains("завис") || lower.Contains("constr"))
                return GroupTypeId.Constraints;
            if (lower.Contains("строит"))
                return GroupTypeId.Construction;
            if (lower.Contains("текст") || lower.Contains("text"))
                return GroupTypeId.Text;
            if (lower.Contains("электр"))
                return GroupTypeId.Electrical;
            if (lower.Contains("сантех") || lower.Contains("plumb"))
                return GroupTypeId.Plumbing;
            if (lower.Contains("ifc"))
                return GroupTypeId.Ifc;

            return GroupTypeId.General;
        }

        // ═══════════════════════════════════════════════════════════
        //  Внутренние классы
        // ═══════════════════════════════════════════════════════════

        internal class SpfGroupInfo
        {
            public string Name { get; set; }
            public DefinitionGroup Group { get; set; }
            public int Count { get; set; }
            public string DisplayText => $"{Name}  ({Count} пар.)";
        }

        internal class SpfParamInfo
        {
            public string Name { get; set; }
            public string GuidString { get; set; }
            public string DataType { get; set; }
            public string ShortType { get; set; }
            public string GroupName { get; set; }
            public string DisplayText { get; set; }
        }

        internal class CategoryInfo
        {
            public BuiltInCategory BuiltIn { get; set; }
            public Category Category { get; set; }
            public string DisplayName { get; set; }
        }

        internal class FamilyCheckItem
        {
            public string FamilyName { get; set; }
            public long FamilyId { get; set; }
            public bool IsChecked { get; set; }
            public CheckBox CheckBox { get; set; }
        }

        internal class OverwriteFamilyLoadOptions : IFamilyLoadOptions
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
