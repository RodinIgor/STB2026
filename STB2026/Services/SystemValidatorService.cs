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
    /// Проверяет: нулевые расходы, отключённые элементы, неназначенные системы.
    /// </summary>
    public class SystemValidatorService
    {
        private readonly Document _doc;

        public SystemValidatorService(Document doc)
        {
            _doc = doc;
        }

        public ValidationResult Validate()
        {
            var result = new ValidationResult();

            // Собираем все элементы воздуховодов
            var ducts = new FilteredElementCollector(_doc)
                .OfClass(typeof(Duct))
                .WhereElementIsNotElementType()
                .ToList();

            // Собираем фитинги воздуховодов
            var fittings = new FilteredElementCollector(_doc)
                .OfCategory(BuiltInCategory.OST_DuctFitting)
                .WhereElementIsNotElementType()
                .ToList();

            // Собираем воздухораспределители (терминалы)
            var terminals = new FilteredElementCollector(_doc)
                .OfCategory(BuiltInCategory.OST_DuctTerminal)
                .WhereElementIsNotElementType()
                .ToList();

            // Собираем оборудование
            var equipment = new FilteredElementCollector(_doc)
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
        /// Проверяет, что расход воздуха = 0
        /// </summary>
        private bool HasZeroFlow(Element element)
        {
            Parameter flowParam = element.get_Parameter(
                BuiltInParameter.RBS_DUCT_FLOW_PARAM);

            if (flowParam == null || !flowParam.HasValue)
                return true; // Нет параметра — считаем проблемой

            return Math.Abs(flowParam.AsDouble()) < 1e-6;
        }

        /// <summary>
        /// Проверяет, назначена ли система
        /// </summary>
        private bool HasSystem(Element element)
        {
            Parameter sysParam = element.get_Parameter(
                BuiltInParameter.RBS_SYSTEM_NAME_PARAM);

            if (sysParam == null)
                return true; // Параметра нет — не проверяем (оборудование и т.д.)

            string sysName = sysParam.AsString();
            return !string.IsNullOrWhiteSpace(sysName);
        }

        /// <summary>
        /// Проверяет, есть ли неподключённые коннекторы.
        /// Воздуховод считается «отключённым», если хотя бы один
        /// из его коннекторов не подсоединён к другому элементу.
        /// </summary>
        private bool IsDisconnected(Element element)
        {
            ConnectorSet connectors = null;

            try
            {
                // Для воздуховодов
                if (element is MEPCurve mepCurve)
                {
                    var connMgr = mepCurve.ConnectorManager;
                    connectors = connMgr?.Connectors;
                }
                // Для FamilyInstance (фитинги, терминалы, оборудование)
                else if (element is FamilyInstance fi)
                {
                    var mepModel = fi.MEPModel;
                    connectors = mepModel?.ConnectorManager?.Connectors;
                }
            }
            catch
            {
                return false; // Не удалось получить коннекторы — пропускаем
            }

            if (connectors == null || connectors.Size == 0)
                return false;

            // Проверяем каждый коннектор
            foreach (Connector connector in connectors)
            {
                try
                {
                    // Проверяем только физические коннекторы (не логические)
                    if (connector.ConnectorType != ConnectorType.End &&
                        connector.ConnectorType != ConnectorType.Curve)
                        continue;

                    // Только для MEP (ductwork)
                    if (connector.Domain != Domain.DomainHvac)
                        continue;

                    // Проверяем подключение
                    if (!connector.IsConnected)
                        return true; // Найден неподключённый коннектор
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
