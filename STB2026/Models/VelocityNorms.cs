using System.Collections.Generic;

namespace STB2026.Models
{
    /// <summary>
    /// Нормы скоростей воздуха в воздуховодах по СП 60.13330.2020, Приложение Л.
    /// Таблицы Л.1–Л.4: рекомендуемые средние скорости движения воздуха.
    /// </summary>
    public static class VelocityNorms
    {
        /// <summary>
        /// Результат проверки скорости
        /// </summary>
        public enum VelocityStatus
        {
            /// <summary>Скорость в норме (зелёный)</summary>
            Normal,
            /// <summary>Скорость на верхней границе нормы (жёлтый)</summary>
            Warning,
            /// <summary>Скорость превышает норму (красный)</summary>
            Exceeded,
            /// <summary>Нет данных для проверки (серый)</summary>
            NoData
        }

        /// <summary>
        /// Диапазон допустимых скоростей, м/с
        /// </summary>
        public class SpeedRange
        {
            public double Min { get; set; }
            public double Max { get; set; }
            public string Description { get; set; } = "";

            public SpeedRange(double min, double max, string description = "")
            {
                Min = min;
                Max = max;
                Description = description;
            }
        }

        // ═══════════════════════════════════════════════════
        // Таблица Л.2 — Приточная вентиляция, общественные здания
        // (основной сценарий для проекта)
        // ═══════════════════════════════════════════════════

        /// <summary>
        /// Приточные системы общественных зданий.
        /// Ключ — производительность системы (м³/ч), значение — диапазон скоростей.
        /// Для часов работы 2000–4000 ч/год (типичный случай).
        /// </summary>
        public static readonly List<(double MaxFlow, SpeedRange Range)> SupplyPublic = new()
        {
            (3000,   new SpeedRange(3.5, 4.5, "Приточная до 3000 м³/ч")),
            (10000,  new SpeedRange(4.5, 5.5, "Приточная 3000–10000 м³/ч")),
            (double.MaxValue, new SpeedRange(5.0, 6.0, "Приточная свыше 10000 м³/ч")),
        };

        // ═══════════════════════════════════════════════════
        // Таблица Л.1 — Вытяжная вентиляция, общественные здания
        // ═══════════════════════════════════════════════════

        public static readonly List<(double MaxFlow, SpeedRange Range)> ExhaustPublic = new()
        {
            (500,    new SpeedRange(2.5, 3.5, "Вытяжная до 500 м³/ч")),
            (2000,   new SpeedRange(3.5, 4.5, "Вытяжная 500–2000 м³/ч")),
            (5000,   new SpeedRange(4.0, 5.0, "Вытяжная 2000–5000 м³/ч")),
            (double.MaxValue, new SpeedRange(4.5, 5.5, "Вытяжная свыше 5000 м³/ч")),
        };

        // ═══════════════════════════════════════════════════
        // Таблица Л.4 — Производственные здания
        // ═══════════════════════════════════════════════════

        public static readonly SpeedRange SupplyIndustrial =
            new SpeedRange(4.0, 7.0, "Приточная мех. вентиляция, произв. здания");

        public static readonly SpeedRange ExhaustIndustrial =
            new SpeedRange(4.0, 8.0, "Вытяжная мех. вентиляция, произв. здания");

        // ═══════════════════════════════════════════════════
        // Таблица Л.3 — Жилые здания
        // ═══════════════════════════════════════════════════

        public static readonly SpeedRange ResidentialInRoom =
            new SpeedRange(1.5, 2.5, "Жилые, мех. побуждение, в помещении");

        public static readonly SpeedRange ResidentialOutOfRoom =
            new SpeedRange(2.0, 4.0, "Жилые, мех. побуждение, вне помещений");

        // ═══════════════════════════════════════════════════
        // Универсальные пороги (упрощённая проверка)
        // ═══════════════════════════════════════════════════

        /// <summary>Максимум для приточных систем общественных зданий</summary>
        public const double DefaultMaxSupply = 6.0;

        /// <summary>Максимум для вытяжных систем общественных зданий</summary>
        public const double DefaultMaxExhaust = 5.5;

        /// <summary>Жёлтая зона: скорость > Max * WarningThreshold</summary>
        public const double WarningThreshold = 0.85;

        /// <summary>
        /// Определяет диапазон допустимой скорости по расходу системы и типу.
        /// </summary>
        /// <param name="systemFlowM3h">Расход системы, м³/ч</param>
        /// <param name="isSupply">true — приточная, false — вытяжная</param>
        /// <returns>Рекомендуемый диапазон скоростей</returns>
        public static SpeedRange GetRange(double systemFlowM3h, bool isSupply)
        {
            var table = isSupply ? SupplyPublic : ExhaustPublic;

            foreach (var (maxFlow, range) in table)
            {
                if (systemFlowM3h <= maxFlow)
                    return range;
            }

            // Fallback — последний диапазон
            return table[table.Count - 1].Range;
        }

        /// <summary>
        /// Проверяет скорость и возвращает статус.
        /// </summary>
        /// <param name="velocityMs">Фактическая скорость, м/с</param>
        /// <param name="range">Допустимый диапазон</param>
        /// <returns>Статус проверки</returns>
        public static VelocityStatus CheckVelocity(double velocityMs, SpeedRange range)
        {
            if (range == null || velocityMs <= 0)
                return VelocityStatus.NoData;

            if (velocityMs <= range.Max * WarningThreshold)
                return VelocityStatus.Normal;

            if (velocityMs <= range.Max)
                return VelocityStatus.Warning;

            return VelocityStatus.Exceeded;
        }

        /// <summary>
        /// Упрощённая проверка по умолчанию (общественные здания, 2000–4000 ч/год).
        /// </summary>
        public static VelocityStatus CheckVelocitySimple(double velocityMs, double systemFlowM3h, bool isSupply)
        {
            var range = GetRange(systemFlowM3h, isSupply);
            return CheckVelocity(velocityMs, range);
        }
    }
}
