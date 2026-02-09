using System.Collections.Generic;

namespace STB2026.RevitBridge.Models
{
    public static class VelocityNorms
    {
        public enum VelocityStatus { Normal, Warning, Exceeded, NoData }

        public class SpeedRange
        {
            public double Min { get; set; }
            public double Max { get; set; }
            public string Description { get; set; } = "";
            public SpeedRange(double min, double max, string description = "")
            {
                Min = min; Max = max; Description = description;
            }
        }

        public static readonly List<(double MaxFlow, SpeedRange Range)> SupplyPublic = new()
        {
            (3000, new SpeedRange(3.5, 4.5, "Приточная до 3000 м³/ч")),
            (10000, new SpeedRange(4.5, 5.5, "Приточная 3000–10000 м³/ч")),
            (double.MaxValue, new SpeedRange(5.0, 6.0, "Приточная свыше 10000 м³/ч")),
        };

        public static readonly List<(double MaxFlow, SpeedRange Range)> ExhaustPublic = new()
        {
            (500, new SpeedRange(2.5, 3.5, "Вытяжная до 500 м³/ч")),
            (2000, new SpeedRange(3.5, 4.5, "Вытяжная 500–2000 м³/ч")),
            (5000, new SpeedRange(4.0, 5.0, "Вытяжная 2000–5000 м³/ч")),
            (double.MaxValue, new SpeedRange(4.5, 5.5, "Вытяжная свыше 5000 м³/ч")),
        };

        public static readonly SpeedRange SupplyIndustrial = new(4.0, 7.0, "Приточная мех., произв.");
        public static readonly SpeedRange ExhaustIndustrial = new(4.0, 8.0, "Вытяжная мех., произв.");
        public static readonly SpeedRange ResidentialInRoom = new(1.5, 2.5, "Жилые, в помещении");
        public static readonly SpeedRange ResidentialOutOfRoom = new(2.0, 4.0, "Жилые, вне помещений");

        public const double DefaultMaxSupply = 6.0;
        public const double DefaultMaxExhaust = 5.5;
        public const double WarningThreshold = 0.85;

        public static SpeedRange GetRange(double systemFlowM3h, bool isSupply)
        {
            var table = isSupply ? SupplyPublic : ExhaustPublic;
            foreach (var (maxFlow, range) in table)
            {
                if (systemFlowM3h <= maxFlow) return range;
            }
            return table[table.Count - 1].Range;
        }

        public static VelocityStatus CheckVelocity(double velocityMs, SpeedRange range)
        {
            if (range == null || velocityMs <= 0) return VelocityStatus.NoData;
            if (velocityMs <= range.Max * WarningThreshold) return VelocityStatus.Normal;
            if (velocityMs <= range.Max) return VelocityStatus.Warning;
            return VelocityStatus.Exceeded;
        }

        public static VelocityStatus CheckVelocitySimple(double velocityMs, double systemFlowM3h, bool isSupply)
        {
            var range = GetRange(systemFlowM3h, isSupply);
            return CheckVelocity(velocityMs, range);
        }
    }
}
