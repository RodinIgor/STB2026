using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
using System;
using System.Collections.Generic;
using System.Linq;

namespace STB2026.Services
{
    /// <summary>
    /// Результат валидации системы вентиляции
    /// </summary>
    public class ValidationResult
    {
        public int TotalElements { get; set; }
        public int ZeroFlowCount => ZeroFlowIds.Count;
        public int DisconnectedCount => DisconnectedIds.Count;
        public int NoSystemCount => NoSystemIds.Count;

        public List<int> ZeroFlowIds { get; set; } = new List<int>();
        public List<int> DisconnectedIds { get; set; } = new List<int>();
        public List<int> NoSystemIds { get; set; } = new List<int>();

        public bool HasErrors => ZeroFlowCount > 0 || DisconnectedCount > 0 || NoSystemCount > 0;
    }

    /// <summary>
    /// Сервис валидации системы вентиляции.
    /// Проверяет элементы ТОЛЬКО на текущем виде.
    /// Проверяет: нулевые расходы, отключённые элементы, неназначенные системы.
    /// Автоматически подсвечивает проблемные элементы при проверке.
    /// </summary>
    public class SystemValidatorService
    {
        private readonly Document _doc;
        private readonly View _view;

        private static readonly Color ColorZeroFlow = new Color(255, 0, 0);       // красный
        private static readonly Color ColorDisconnected = new Color(255, 200, 0);  // жёлтый
        private static readonly Color ColorNoSystem = new Color(128, 128, 128);    // серый

        public SystemValidatorService(Document doc, View view)
        {
            _doc = doc;
            _view = view;
        }

        /// <summary>
        /// Конструктор для обратной совместимости (без вида — не подсвечивает).
        /// </summary>
        public SystemValidatorService(Document doc)
        {
            _doc = doc;
            _view = doc.ActiveView;
        }

        /// <summary>
        /// Валидация + автоматическая подсветка проблемных элементов.
        /// </summary>
        public ValidationResult ValidateAndColorize()
        {
            var result = Validate();

            if (result.HasErrors && _view != null)
            {
                ColorizeProblems(result);
            }

            return result;
        }

        public ValidationResult Validate()
        {
            var result = new ValidationResult();

            // Собираем элементы ТОЛЬКО на текущем виде
            var ducts = new FilteredElementCollector(_doc, _view.Id)
                .OfClass(typeof(Duct))
                .WhereElementIsNotElementType()
                .ToList();

            var fittings = new FilteredElementCollector(_doc, _view.Id)
                .OfCategory(BuiltInCategory.OST_DuctFitting)
                .WhereElementIsNotElementType()
                .ToList();

            var terminals = new FilteredElementCollector(_doc, _view.Id)
                .OfCategory(BuiltInCategory.OST_DuctTerminal)
                .WhereElementIsNotElementType()
                .ToList();

            var equipment = new FilteredElementCollector(_doc, _view.Id)
                .OfCategory(BuiltInCategory.OST_MechanicalEquipment)
                .WhereElementIsNotElementType()
                .ToList();

            var allElements = ducts
                .Concat(fittings)
                .Concat(terminals)
                .Concat(equipment)
                .ToList();

            result.TotalElements = allElements.Count;

            foreach (var element in allElements)
            {
                int id = element.Id.IntegerValue;

                // 1. Проверка нулевого расхода (только для воздуховодов)
                if (element is Duct)
                {
                    if (HasZeroFlow(element))
                    {
                        result.ZeroFlowIds.Add(id);
                    }
                }

                // 2. Проверка подключения к системе
                if (!HasSystem(element))
                {
                    result.NoSystemIds.Add(id);
                }

                // 3. Проверка физического подключения (коннекторы)
                if (IsDisconnected(element))
                {
                    result.DisconnectedIds.Add(id);
                }
            }

            return result;
        }

        /// <summary>
        /// Подсветка проблемных элементов на текущем виде.
        /// Вызывается автоматически из ValidateAndColorize()
        /// и может вызываться отдельно из Command.
        /// </summary>
        public void ColorizeProblems(ValidationResult result)
        {
            using (Transaction tx = new Transaction(_doc, "STB2026: Подсветка проблем"))
            {
                tx.Start();

                foreach (int id in result.ZeroFlowIds)
                {
                    try { ApplyColor(new ElementId(id), ColorZeroFlow); } catch { }
                }
                foreach (int id in result.DisconnectedIds)
                {
                    try { ApplyColor(new ElementId(id), ColorDisconnected); } catch { }
                }
                foreach (int id in result.NoSystemIds)
                {
                    try { ApplyColor(new ElementId(id), ColorNoSystem); } catch { }
                }

                tx.Commit();
            }
        }

        private void ApplyColor(ElementId elementId, Color color)
        {
            var ogs = new OverrideGraphicSettings();
            ogs.SetProjectionLineColor(color);

            var solidPattern = FillPatternElement.GetFillPatternElementByName(
                _doc, FillPatternTarget.Drafting, "<Solid fill>");
            if (solidPattern != null)
            {
                ogs.SetSurfaceForegroundPatternId(solidPattern.Id);
                ogs.SetSurfaceForegroundPatternColor(color);
            }

            _view.SetElementOverrides(elementId, ogs);
        }

        private bool HasZeroFlow(Element element)
        {
            Parameter flowParam = element.get_Parameter(
                BuiltInParameter.RBS_DUCT_FLOW_PARAM);

            if (flowParam == null || !flowParam.HasValue)
                return true;

            return Math.Abs(flowParam.AsDouble()) < 1e-6;
        }

        private bool HasSystem(Element element)
        {
            Parameter sysParam = element.get_Parameter(
                BuiltInParameter.RBS_SYSTEM_NAME_PARAM);

            if (sysParam == null)
                return true;

            string sysName = sysParam.AsString();
            return !string.IsNullOrWhiteSpace(sysName);
        }

        private bool IsDisconnected(Element element)
        {
            ConnectorSet connectors = null;

            try
            {
                if (element is MEPCurve mepCurve)
                {
                    var connMgr = mepCurve.ConnectorManager;
                    connectors = connMgr?.Connectors;
                }
                else if (element is FamilyInstance fi)
                {
                    var mepModel = fi.MEPModel;
                    connectors = mepModel?.ConnectorManager?.Connectors;
                }
            }
            catch
            {
                return false;
            }

            if (connectors == null || connectors.Size == 0)
                return false;

            foreach (Connector connector in connectors)
            {
                try
                {
                    if (connector.ConnectorType != ConnectorType.End &&
                        connector.ConnectorType != ConnectorType.Curve)
                        continue;

                    if (connector.Domain != Domain.DomainHvac)
                        continue;

                    if (!connector.IsConnected)
                        return true;
                }
                catch
                {
                    continue;
                }
            }

            return false;
        }
    }
}
