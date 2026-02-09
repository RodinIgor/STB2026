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

        private List<FamilyInfo> _families = new List<FamilyInfo>();
        private List<SpfGroupInfo> _spfGroups = new List<SpfGroupInfo>();
        private List<SpfParamInfo> _spfParams = new List<SpfParamInfo>();
        private List<string> _familyParams = new List<string>();

        public AddSharedParamWindow(UIApplication uiApp)
        {
            InitializeComponent();
            _uiApp = uiApp;
            _doc = uiApp.ActiveUIDocument.Document;
            LoadFamilies();
            LoadSpfGroups();
            LoadDisplayGroups();
        }

        // ═══════════════════════════════════════════════════════════
        //  Загрузка данных
        // ═══════════════════════════════════════════════════════════

        private void LoadFamilies()
        {
            _families = new FilteredElementCollector(_doc)
                .OfClass(typeof(Family))
                .Cast<Family>()
                .Where(f => !f.IsInPlace && f.IsEditable)
                .Select(f => new FamilyInfo
                {
                    Family = f,
                    DisplayName = $"{f.Name}  [{f.FamilyCategory?.Name ?? "—"}]"
                })
                .OrderBy(f => f.Family.FamilyCategory?.Name ?? "")
                .ThenBy(f => f.Family.Name)
                .ToList();

            cmbFamily.ItemsSource = _families;
            cmbFamily.DisplayMemberPath = "DisplayName";
        }

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

        private void LoadDisplayGroups()
        {
            var groups = new List<string>
            {
                "Общие",
                "Данные",
                "Идентификация",
                "Размеры",
                "Зависимости",
                "Строительство",
                "Механизмы",
                "Механизмы\u00A0— расход",
                "Механизмы\u00A0— нагрузки",
                "Электросети",
                "Сантехника",
                "Текст",
                "IFC-параметры"
            };
            cmbDisplayGroup.ItemsSource = groups;
            cmbDisplayGroup.SelectedIndex = 0;
        }

        private void LoadFamilyParameters(Family family)
        {
            _familyParams.Clear();
            _familyParams.Add("(не связывать)");
            try
            {
                Document famDoc = _doc.EditFamily(family);
                if (famDoc != null)
                {
                    var famMgr = famDoc.FamilyManager;
                    foreach (FamilyParameter fp in famMgr.Parameters)
                    {
                        if (fp?.Definition?.Name != null)
                        {
                            string suffix = fp.IsInstance ? " [экз]" : " [тип]";
                            string storage = fp.StorageType.ToString();
                            _familyParams.Add($"{fp.Definition.Name}{suffix}  ({storage})");
                        }
                    }
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

        private void CmbFamily_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (cmbFamily.SelectedItem is FamilyInfo fi)
                LoadFamilyParameters(fi.Family);
            UpdateAddButton();
        }

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

                        // Сохраняем GUID как строку сразу — объект может стать невалидным
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

        private void UpdateAddButton()
        {
            btnAdd.IsEnabled = cmbFamily.SelectedItem is FamilyInfo
                            && cmbParam.SelectedItem is SpfParamInfo;
        }

        // ═══════════════════════════════════════════════════════════
        //  Кнопка "Добавить"
        // ═══════════════════════════════════════════════════════════

        private void BtnAdd_Click(object sender, RoutedEventArgs e)
        {
            if (!(cmbFamily.SelectedItem is FamilyInfo fi)) return;
            if (!(cmbParam.SelectedItem is SpfParamInfo pi)) return;

            bool isInstance = rbInstance.IsChecked == true;
            string displayGroup = cmbDisplayGroup.SelectedItem?.ToString() ?? "Общие";

            // ══════════════════════════════════════════════════════
            // FIX: Сохраняем ВСЕ данные в локальные строки ДО вызова,
            // т.к. после операций с семейством объекты SpfParamInfo
            // и ExternalDefinition могут стать невалидными
            // ══════════════════════════════════════════════════════
            string familyName = fi.Family.Name;
            string paramName = pi.Name;
            string paramGroupName = pi.GroupName;
            string paramGuid = pi.GuidString;
            string paramDataType = pi.DataType;

            string linkedParamName = null;
            if (cmbLinkParam.SelectedIndex > 0 && cmbLinkParam.SelectedItem is string linkStr)
            {
                int bracketIdx = linkStr.IndexOf(" [");
                linkedParamName = bracketIdx > 0 ? linkStr.Substring(0, bracketIdx) : linkStr;
            }

            btnAdd.IsEnabled = false;
            btnAdd.Content = "Выполняю...";

            try
            {
                string result = AddSharedParamToFamily(
                    fi.Family, paramName, paramGroupName,
                    isInstance, displayGroup, linkedParamName);

                System.Windows.MessageBox.Show(result, "Результат",
                    MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                // FIX: Используем только локальные строки, не обращаемся к pi/fi
                System.Windows.MessageBox.Show(
                    $"Параметр '{paramName}' мог быть добавлен в семейство '{familyName}',\n" +
                    $"но возникла ошибка при завершении операции:\n\n{ex.Message}\n\n" +
                    $"Проверьте свойства элемента — параметр может уже работать.",
                    "Внимание", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
            finally
            {
                btnAdd.IsEnabled = true;
                btnAdd.Content = "Добавить";
            }
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
            string tempDir = Path.Combine(Path.GetTempPath(), "STB2026_Families");
            Directory.CreateDirectory(tempDir);
            string tempRfaPath = Path.Combine(tempDir, $"{family.Name}.rfa");

            // Сохраняем имя до любых операций
            string familyName = family.Name;

            Document famDoc = null;
            Document openedFamDoc = null;

            try
            {
                // 1. Экспортируем семейство во временный .rfa
                famDoc = _doc.EditFamily(family);
                if (famDoc == null)
                    return "Не удалось открыть документ семейства";

                famDoc.SaveAs(tempRfaPath, new SaveAsOptions { OverwriteExistingFile = true });
                famDoc.Close(false);
                famDoc = null;

                // 2. Открываем .rfa как отдельный документ
                openedFamDoc = _uiApp.Application.OpenDocumentFile(tempRfaPath);
                if (openedFamDoc == null)
                    return "Не удалось открыть временный .rfa файл";

                var famMgr = openedFamDoc.FamilyManager;

                // 3. Проверяем что параметр не существует
                foreach (FamilyParameter fp in famMgr.Parameters)
                {
                    if (fp.Definition.Name == sharedParamName)
                    {
                        openedFamDoc.Close(false);
                        return $"Параметр '{sharedParamName}' уже существует в '{familyName}'";
                    }
                }

                // 4. Получаем СВЕЖИЙ ExternalDefinition в контексте famDoc
                ExternalDefinition freshExtDef = null;
                DefinitionFile defFile = null;
                try { defFile = openedFamDoc.Application.OpenSharedParameterFile(); } catch { }
                if (defFile == null)
                    try { defFile = _uiApp.Application.OpenSharedParameterFile(); } catch { }

                if (defFile != null)
                {
                    // Сначала ищем в конкретной группе
                    DefinitionGroup targetGroup = defFile.Groups.get_Item(spfGroupName);
                    if (targetGroup != null)
                    {
                        Definition def = targetGroup.Definitions.get_Item(sharedParamName);
                        if (def is ExternalDefinition ed)
                            freshExtDef = ed;
                    }

                    // Если не нашли — поиск по всем группам
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
                    return $"Параметр '{sharedParamName}' не найден в ФОП";
                }

                // 5. Группа отображения
                ForgeTypeId groupId = ResolveGroupTypeId(displayGroup);

                // 6. Добавляем параметр
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
                        throw new InvalidOperationException("FamilyManager.AddParameter вернул null");

                    // 7. Связать формулой
                    if (!string.IsNullOrWhiteSpace(linkedParamName))
                    {
                        try { famMgr.SetFormula(newParam, linkedParamName); }
                        catch { }
                    }

                    trans.Commit();
                }

                // 8. Сохраняем и закрываем
                openedFamDoc.Save();
                openedFamDoc.Close(false);
                openedFamDoc = null;

                // 9. Загружаем обратно в проект
                Family reloadedFamily = null;
                using (var trans = new Transaction(_doc, $"STB2026: Reload {familyName}"))
                {
                    trans.Start();
                    _doc.LoadFamily(tempRfaPath, new OverwriteFamilyLoadOptions(), out reloadedFamily);
                    trans.Commit();
                }

                // FIX: Используем только локальные строки — никаких обращений к объектам Revit
                string linkedInfo = string.IsNullOrWhiteSpace(linkedParamName)
                    ? "" : $"\nСвязан формулой с: {linkedParamName}";

                return $"✓ Параметр '{sharedParamName}' успешно добавлен\n" +
                       $"  в семейство '{familyName}'\n" +
                       $"  Тип: {(isInstance ? "Экземпляр" : "Тип")}\n" +
                       $"  Группа: {displayGroup}" + linkedInfo;
            }
            catch (Exception ex)
            {
                try { famDoc?.Close(false); } catch { }
                try { openedFamDoc?.Close(false); } catch { }
                throw new InvalidOperationException(
                    $"Ошибка при добавлении '{sharedParamName}' в '{familyName}': {ex.Message}", ex);
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

        internal class FamilyInfo
        {
            public Family Family { get; set; }
            public string DisplayName { get; set; }
        }

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
