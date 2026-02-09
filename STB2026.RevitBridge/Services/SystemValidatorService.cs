using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
using System;
using System.Collections.Generic;
using System.Linq;

namespace STB2026.RevitBridge.Services
{
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

            var ducts = new FilteredElementCollector(_doc)
                .OfClass(typeof(Duct))
                .WhereElementIsNotElementType()
                .ToList();

            var fittings = new FilteredElementCollector(_doc)
                .OfCategory(BuiltInCategory.OST_DuctFitting)
                .WhereElementIsNotElementType()
                .ToList();

            var terminals = new FilteredElementCollector(_doc)
                .OfCategory(BuiltInCategory.OST_DuctTerminal)
                .WhereElementIsNotElementType()
                .ToList();

            var equipment = new FilteredElementCollector(_doc)
                .OfCategory(BuiltInCategory.OST_MechanicalEquipment)
                .WhereElementIsNotElementType()
                .ToList();

            var allElements = ducts.Concat(fittings).Concat(terminals).Concat(equipment).ToList();
            result.TotalElements = allElements.Count;

            foreach (var element in allElements)
            {
                int id = element.Id.IntegerValue;

                if (element is Duct)
                {
                    if (HasZeroFlow(element))
                        result.ZeroFlowIds.Add(id);
                }

                if (!HasSystem(element))
                    result.NoSystemIds.Add(id);

                if (IsDisconnected(element))
                    result.DisconnectedIds.Add(id);
            }

            return result;
        }

        private bool HasZeroFlow(Element element)
        {
            Parameter flowParam = element.get_Parameter(BuiltInParameter.RBS_DUCT_FLOW_PARAM);
            if (flowParam == null || !flowParam.HasValue) return true;
            return Math.Abs(flowParam.AsDouble()) < 1e-6;
        }

        private bool HasSystem(Element element)
        {
            Parameter sysParam = element.get_Parameter(BuiltInParameter.RBS_SYSTEM_NAME_PARAM);
            if (sysParam == null) return true;
            string sysName = sysParam.AsString();
            return !string.IsNullOrWhiteSpace(sysName);
        }

        private bool IsDisconnected(Element element)
        {
            ConnectorSet connectors = null;
            try
            {
                if (element is MEPCurve mepCurve)
                    connectors = mepCurve.ConnectorManager?.Connectors;
                else if (element is FamilyInstance fi)
                    connectors = fi.MEPModel?.ConnectorManager?.Connectors;
            }
            catch { return false; }

            if (connectors == null || connectors.Size == 0) return false;

            foreach (Connector connector in connectors)
            {
                try
                {
                    if (connector.ConnectorType != ConnectorType.End &&
                        connector.ConnectorType != ConnectorType.Curve)
                        continue;
                    if (connector.Domain != Domain.DomainHvac) continue;
                    if (!connector.IsConnected) return true;
                }
                catch { continue; }
            }
            return false;
        }
    }
}
