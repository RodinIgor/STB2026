using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
using STB2026.Models;
using System;
using System.Collections.Generic;
using System.Linq;

namespace STB2026.Services
{
    /// <summary>
    /// Результат проверки скоростей
    /// </summary>
    public class VelocityCheckResult
    {
        public int Total { get; set; }
        public int Normal { get; set; }
        public int Warning { get; set; }
        public int Exceeded { get; set; }
        public int NoData { get; set; }
    }

    /// <summary>
    /// Сервис проверки скоростей воздуха по СП 60.13330.2020 (Прил. Л).
    /// Применяет цветовую карту к воздуховодам на текущем виде.
    /// </summary>
    public class VelocityCheckerService
    {
        private readonly Document _doc;
        private readonly View _view;

        // Цвета статусов
        private static readonly Color ColorNormal   = new Color(0, 180, 0);     // зелёный
        private static readonly Color ColorWarning  = new Color(255, 200, 0);   // жёлтый
        private static readonly Color ColorExceeded = new Color(255, 0, 0);     // красный
        private static readonly Color ColorNoData   = new Color(180, 180, 180); // серый

        public VelocityCheckerService(Document doc, View view)
        {
            _doc = doc;
            _view = view;
        }

        public VelocityCheckResult CheckAndColorize()
        {
            var result = new VelocityCheckResult();

            // Собираем воздуховоды на виде
            var ducts = new FilteredElementCollector(_doc, _view.Id)
                .OfClass(typeof(Duct))
                .WhereElementIsNotElementType()
                .Cast<Duct>()
                .ToList();

            result.Total = ducts.Count;

            if (ducts.Count == 0)
                return result;

            // Кэш: система → расход системы (для определения диапазона норм)
            var systemFlowCache = new Dictionary<ElementId, double>();

            using (Transaction tx = new Transaction(_doc, "STB2026: Проверка скоростей"))
            {
                tx.Start();

                foreach (var duct in ducts)
                {
                    try
                    {
                        // Получаем скорость
                        double velocity = GetVelocity(duct);

                        // Определяем тип системы и расход
                        bool isSupply = IsSupplySystem(duct);
                        double systemFlow = GetSystemFlow(duct, systemFlowCache);

                        // Проверяем по нормам СП 60
                        VelocityNorms.VelocityStatus status;

                        if (velocity <= 0)
                        {
                            status = VelocityNorms.VelocityStatus.NoData;
                        }
                        else
                        {
                            status = VelocityNorms.CheckVelocitySimple(
                                velocity, systemFlow, isSupply);
                        }

                        // Применяем цвет
                        Color color = GetStatusColor(status);
                        ApplyColor(duct.Id, color);

                        // Считаем статистику
                        switch (status)
                        {
                            case VelocityNorms.VelocityStatus.Normal: result.Normal++; break;
                            case VelocityNorms.VelocityStatus.Warning: result.Warning++; break;
                            case VelocityNorms.VelocityStatus.Exceeded: result.Exceeded++; break;
                            case VelocityNorms.VelocityStatus.NoData: result.NoData++; break;
                        }
                    }
                    catch
                    {
                        result.NoData++;
                    }
                }

                tx.Commit();
            }

            return result;
        }

        /// <summary>
        /// Получает скорость воздуха в воздуховоде, м/с
        /// </summary>
        private double GetVelocity(Duct duct)
        {
            // Пробуем встроенный параметр Velocity
            Parameter velParam = duct.get_Parameter(
                BuiltInParameter.RBS_VELOCITY);

            if (velParam != null && velParam.HasValue)
            {
                // Revit хранит в футах/с, конвертируем в м/с
                double velFtPerSec = velParam.AsDouble();
                return velFtPerSec * 0.3048;
            }

            // Вычисляем: V = Q / A
            double flow = GetFlow(duct); // м³/ч
            double area = GetCrossSectionArea(duct); // м²

            if (flow > 0 && area > 0)
            {
                return (flow / 3600.0) / area; // м/с
            }

            return 0;
        }

        /// <summary>
        /// Получает расход воздуха, м³/ч
        /// </summary>
        private double GetFlow(Duct duct)
        {
            Parameter flowParam = duct.get_Parameter(
                BuiltInParameter.RBS_DUCT_FLOW_PARAM);

            if (flowParam != null && flowParam.HasValue)
            {
                // Revit хранит в куб. футах/мин → конвертируем в м³/ч
                double cfm = flowParam.AsDouble();
                return cfm * 1.699; // 1 CFM = 1.699 м³/ч
            }
            return 0;
        }

        /// <summary>
        /// Получает площадь поперечного сечения, м²
        /// </summary>
        private double GetCrossSectionArea(Duct duct)
        {
            // Вычисляем площадь из размеров сечения
            Parameter widthParam = duct.get_Parameter(BuiltInParameter.RBS_CURVE_WIDTH_PARAM);
            Parameter heightParam = duct.get_Parameter(BuiltInParameter.RBS_CURVE_HEIGHT_PARAM);
            Parameter diamParam = duct.get_Parameter(BuiltInParameter.RBS_CURVE_DIAMETER_PARAM);

            if (diamParam != null && diamParam.HasValue)
            {
                double d = diamParam.AsDouble() * 0.3048; // футы → м
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
        /// Определяет, является ли система приточной
        /// </summary>
        private bool IsSupplySystem(Duct duct)
        {
            Parameter sysTypeParam = duct.get_Parameter(
                BuiltInParameter.RBS_DUCT_SYSTEM_TYPE_PARAM);

            if (sysTypeParam != null && sysTypeParam.HasValue)
            {
                string sysType = sysTypeParam.AsValueString() ?? "";
                // "Приточный воздух" / "Supply Air" → true
                // "Вытяжной воздух" / "Return Air" / "Exhaust Air" → false
                return sysType.Contains("Приточ", StringComparison.OrdinalIgnoreCase) ||
                       sysType.Contains("Supply", StringComparison.OrdinalIgnoreCase) ||
                       sysType.Contains("П ", StringComparison.Ordinal);
            }
            return true; // По умолчанию — приточная
        }

        /// <summary>
        /// Получает суммарный расход системы для определения нормы
        /// </summary>
        private double GetSystemFlow(Duct duct, Dictionary<ElementId, double> cache)
        {
            // Пробуем найти систему
            Parameter sysParam = duct.get_Parameter(
                BuiltInParameter.RBS_SYSTEM_NAME_PARAM);

            // Для упрощения используем расход самого воздуховода
            // (в реальности нужен максимальный расход в системе)
            double flow = GetFlow(duct);
            return flow > 0 ? flow : 5000; // fallback
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

            // Заливка проекции для наглядности
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
