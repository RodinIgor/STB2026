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

        /// <summary>
        /// Использованные диапазоны: ключ — описание, значение — количество
        /// </summary>
        public Dictionary<string, int> RangeUsage { get; set; } = new Dictionary<string, int>();

        /// <summary>
        /// Детали по каждому воздуховоду для расширенного отчёта
        /// </summary>
        public List<DuctVelocityDetail> Details { get; set; } = new List<DuctVelocityDetail>();
    }

    public class DuctVelocityDetail
    {
        public int DuctId { get; set; }
        public string SystemName { get; set; } = "";
        public string Size { get; set; } = "";
        public double FlowM3h { get; set; }
        public double VelocityMs { get; set; }
        public string RangeDesc { get; set; } = "";
        public VelocityNorms.VelocityStatus Status { get; set; }
    }

    public class VelocityCheckerService
    {
        private readonly Document _doc;
        private readonly View _view;

        private static readonly Color ColorNormal = new Color(0, 180, 0);
        private static readonly Color ColorWarning = new Color(255, 200, 0);
        private static readonly Color ColorExceeded = new Color(200, 80, 0);
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

            using (Transaction tx = new Transaction(_doc, "STB2026: Проверка скоростей"))
            {
                tx.Start();

                foreach (var duct in ducts)
                {
                    try
                    {
                        double velocity = GetVelocity(duct);
                        bool isSupply = IsSupplySystem(duct);
                        double flow = GetFlow(duct);

                        VelocityNorms.VelocityStatus status;
                        if (velocity <= 0)
                        {
                            status = VelocityNorms.VelocityStatus.NoData;
                        }
                        else
                        {
                            status = VelocityNorms.CheckVelocitySimple(velocity, flow, isSupply);
                        }

                        // Учитываем использованный диапазон
                        var range = VelocityNorms.GetRange(flow, isSupply);
                        string rangeDesc = $"{range.Description}: {range.Min}–{range.Max} м/с";
                        if (result.RangeUsage.ContainsKey(rangeDesc))
                            result.RangeUsage[rangeDesc]++;
                        else
                            result.RangeUsage[rangeDesc] = 1;

                        // Детали
                        result.Details.Add(new DuctVelocityDetail
                        {
                            DuctId = duct.Id.IntegerValue,
                            SystemName = GetSystemName(duct),
                            Size = GetDuctSize(duct),
                            FlowM3h = flow,
                            VelocityMs = Math.Round(velocity, 2),
                            RangeDesc = rangeDesc,
                            Status = status
                        });

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

        /// <summary>
        /// Определяет тип системы по параметру «Классификация систем» (RBS_SYSTEM_CLASSIFICATION_PARAM).
        /// Возвращает: «Приточный воздух» / «Supply Air» → true, иначе false.
        /// Fallback: проверка имени системы и типа системы.
        /// </summary>
        private bool IsSupplySystem(Duct duct)
        {
            // 1. Приоритетный параметр — Классификация систем (-1140325)
            Parameter classParam = duct.get_Parameter(BuiltInParameter.RBS_SYSTEM_CLASSIFICATION_PARAM);
            if (classParam != null && classParam.HasValue)
            {
                string classification = classParam.AsString() ?? "";
                if (!string.IsNullOrEmpty(classification))
                {
                    return classification.IndexOf("Приточн", StringComparison.OrdinalIgnoreCase) >= 0 ||
                           classification.IndexOf("Supply", StringComparison.OrdinalIgnoreCase) >= 0;
                }
            }

            // 2. Fallback — Имя системы
            Parameter sysNameParam = duct.get_Parameter(BuiltInParameter.RBS_SYSTEM_NAME_PARAM);
            if (sysNameParam != null && sysNameParam.HasValue)
            {
                string sysName = sysNameParam.AsString() ?? "";
                if (sysName.StartsWith("П ", StringComparison.Ordinal) ||
                    sysName.StartsWith("П", StringComparison.Ordinal) && sysName.Length > 1 && char.IsDigit(sysName[1]))
                {
                    return true;
                }
            }

            // 3. Fallback — Тип системы (ElementId → AsValueString)
            Parameter sysTypeParam = duct.get_Parameter(BuiltInParameter.RBS_DUCT_SYSTEM_TYPE_PARAM);
            if (sysTypeParam != null && sysTypeParam.HasValue)
            {
                string sysType = sysTypeParam.AsValueString() ?? "";
                return sysType.IndexOf("Приток", StringComparison.OrdinalIgnoreCase) >= 0 ||
                       sysType.IndexOf("Supply", StringComparison.OrdinalIgnoreCase) >= 0;
            }

            return true; // По умолчанию — приточная
        }

        private string GetSystemName(Duct duct)
        {
            Parameter p = duct.get_Parameter(BuiltInParameter.RBS_SYSTEM_NAME_PARAM);
            return p?.AsString() ?? "N/A";
        }

        private string GetDuctSize(Duct duct)
        {
            Parameter p = duct.get_Parameter(BuiltInParameter.RBS_CALCULATED_SIZE);
            return p?.AsString() ?? "N/A";
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

        /// <summary>
        /// Полный справочник норм СП 60.13330.2020 для отображения в диалоге.
        /// </summary>
        public static string GetFullNormsReference()
        {
            return
                "═══ СП 60.13330.2020, Приложение Л ═══\n\n" +
                "ПРИТОЧНАЯ вентиляция (общественные здания):\n" +
                "  до 3 000 м³/ч:     3,5 – 4,5 м/с\n" +
                "  3 000–10 000 м³/ч:  4,5 – 5,5 м/с\n" +
                "  свыше 10 000 м³/ч:  5,0 – 6,0 м/с\n\n" +
                "ВЫТЯЖНАЯ вентиляция (общественные здания):\n" +
                "  до 500 м³/ч:        2,5 – 3,5 м/с\n" +
                "  500–2 000 м³/ч:     3,5 – 4,5 м/с\n" +
                "  2 000–5 000 м³/ч:   4,0 – 5,0 м/с\n" +
                "  свыше 5 000 м³/ч:   4,5 – 5,5 м/с\n\n" +
                "ЖИЛЫЕ здания (мех. побуждение):\n" +
                "  в помещении:        1,5 – 2,5 м/с\n" +
                "  вне помещений:      2,0 – 4,0 м/с\n\n" +
                "ПРОИЗВОДСТВЕННЫЕ здания:\n" +
                "  приточная:          4,0 – 7,0 м/с\n" +
                "  вытяжная:           4,0 – 8,0 м/с";
        }
    }
}
