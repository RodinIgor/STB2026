using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
using STB2026.Models;
using System;
using System.Collections.Generic;
using System.Linq;

namespace STB2026.Services
{
    public class VelocityCheckResult
    {
        public int Total { get; set; }
        public int Normal { get; set; }
        public int Warning { get; set; }
        public int Exceeded { get; set; }
        public int NoData { get; set; }
        public Dictionary<string, int> RangeUsage { get; set; } = new Dictionary<string, int>();
    }

    public class VelocityCheckerService
    {
        private readonly Document _doc;
        private readonly View _view;

        private static readonly Color ColorNormal = new Color(0, 180, 0);
        private static readonly Color ColorWarning = new Color(255, 200, 0);
        private static readonly Color ColorExceeded = new Color(200, 80, 0); // тёмно-оранжевый (не красный — конфликт с приточкой)
        private static readonly Color ColorNoData = new Color(180, 180, 180);

        public VelocityCheckerService(Document doc, View view)
        {
            _doc = doc;
            _view = view;
        }

        public VelocityCheckResult CheckAndColorize()
        {
            var result = new VelocityCheckResult();

            var ducts = new FilteredElementCollector(_doc, _view.Id)
                .OfClass(typeof(Duct))
                .WhereElementIsNotElementType()
                .Cast<Duct>()
                .ToList();

            result.Total = ducts.Count;
            if (ducts.Count == 0) return result;

            var systemFlowCache = new Dictionary<ElementId, double>();

            using (Transaction tx = new Transaction(_doc, "STB2026: Проверка скоростей"))
            {
                tx.Start();

                foreach (var duct in ducts)
                {
                    try
                    {
                        double velocity = GetVelocity(duct);
                        bool isSupply = IsSupplySystem(duct);
                        double systemFlow = GetSystemFlow(duct, systemFlowCache);

                        VelocityNorms.VelocityStatus status;
                        if (velocity <= 0)
                        {
                            status = VelocityNorms.VelocityStatus.NoData;
                        }
                        else
                        {
                            status = VelocityNorms.CheckVelocitySimple(velocity, systemFlow, isSupply);
                        }

                        // Учитываем использованный диапазон
                        var range = VelocityNorms.GetRange(systemFlow, isSupply);
                        string rangeDesc = $"{range.Description}: {range.Min}–{range.Max} м/с";
                        if (result.RangeUsage.ContainsKey(rangeDesc))
                            result.RangeUsage[rangeDesc]++;
                        else
                            result.RangeUsage[rangeDesc] = 1;

                        Color color = GetStatusColor(status);
                        ApplyColor(duct.Id, color);

                        switch (status)
                        {
                            case VelocityNorms.VelocityStatus.Normal: result.Normal++; break;
                            case VelocityNorms.VelocityStatus.Warning: result.Warning++; break;
                            case VelocityNorms.VelocityStatus.Exceeded: result.Exceeded++; break;
                            case VelocityNorms.VelocityStatus.NoData: result.NoData++; break;
                        }
                    }
                    catch { result.NoData++; }
                }

                tx.Commit();
            }

            return result;
        }

        private double GetVelocity(Duct duct)
        {
            Parameter velParam = duct.get_Parameter(BuiltInParameter.RBS_VELOCITY);
            if (velParam != null && velParam.HasValue)
            {
                double velFtPerSec = velParam.AsDouble();
                return velFtPerSec * 0.3048;
            }

            double flow = GetFlow(duct);
            double area = GetCrossSectionArea(duct);
            if (flow > 0 && area > 0)
                return (flow / 3600.0) / area;

            return 0;
        }

        private double GetFlow(Duct duct)
        {
            Parameter flowParam = duct.get_Parameter(BuiltInParameter.RBS_DUCT_FLOW_PARAM);
            if (flowParam != null && flowParam.HasValue)
            {
                double cfm = flowParam.AsDouble();
                return cfm * 1.699;
            }
            return 0;
        }

        private double GetCrossSectionArea(Duct duct)
        {
            Parameter widthParam = duct.get_Parameter(BuiltInParameter.RBS_CURVE_WIDTH_PARAM);
            Parameter heightParam = duct.get_Parameter(BuiltInParameter.RBS_CURVE_HEIGHT_PARAM);
            Parameter diamParam = duct.get_Parameter(BuiltInParameter.RBS_CURVE_DIAMETER_PARAM);

            if (diamParam != null && diamParam.HasValue)
            {
                double d = diamParam.AsDouble() * 0.3048;
                return Math.PI * d * d / 4.0;
            }
            if (widthParam != null && widthParam.HasValue &&
                heightParam != null && heightParam.HasValue)
            {
                double w = widthParam.AsDouble() * 0.3048;
                double h = heightParam.AsDouble() * 0.3048;
                return w * h;
            }
            return 0;
        }

        private bool IsSupplySystem(Duct duct)
        {
            Parameter sysTypeParam = duct.get_Parameter(BuiltInParameter.RBS_DUCT_SYSTEM_TYPE_PARAM);
            if (sysTypeParam != null && sysTypeParam.HasValue)
            {
                string sysType = sysTypeParam.AsValueString() ?? "";
                return sysType.Contains("Приточ", StringComparison.OrdinalIgnoreCase) ||
                       sysType.Contains("Supply", StringComparison.OrdinalIgnoreCase) ||
                       sysType.Contains("П ", StringComparison.Ordinal);
            }
            return true;
        }

        private double GetSystemFlow(Duct duct, Dictionary<ElementId, double> cache)
        {
            double flow = GetFlow(duct);
            return flow > 0 ? flow : 5000;
        }

        private Color GetStatusColor(VelocityNorms.VelocityStatus status)
        {
            switch (status)
            {
                case VelocityNorms.VelocityStatus.Normal: return ColorNormal;
                case VelocityNorms.VelocityStatus.Warning: return ColorWarning;
                case VelocityNorms.VelocityStatus.Exceeded: return ColorExceeded;
                default: return ColorNoData;
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
    }
}