using System;
using System.Drawing;

namespace grbloxy
{
    internal static class HeatmapColorMapper
    {
        private static readonly (double Stop, Color Color)[] PaletteStops =
        {
            (0.00, Color.FromArgb(26, 49, 160)),
            (0.25, Color.FromArgb(40, 166, 230)),
            (0.50, Color.FromArgb(45, 184, 93)),
            (0.75, Color.FromArgb(242, 202, 68)),
            (1.00, Color.FromArgb(225, 63, 45))
        };

        public static Color GetHeatmapColor(double value, double minValue, double maxValue)
        {
            if (double.IsNaN(value) || double.IsInfinity(value))
            {
                return Color.FromArgb(220, 225, 232);
            }

            double range = maxValue - minValue;
            if (Math.Abs(range) < 1e-12)
            {
                return PaletteStops[PaletteStops.Length / 2].Color;
            }

            double normalized = (value - minValue) / range;
            return GetColorFromNormalized(normalized);
        }

        public static Color GetColorFromNormalized(double normalized)
        {
            normalized = Math.Max(0, Math.Min(1, normalized));

            for (int index = 0; index < PaletteStops.Length - 1; index++)
            {
                if (normalized <= PaletteStops[index + 1].Stop)
                {
                    double localT = (normalized - PaletteStops[index].Stop) /
                        Math.Max(1e-12, PaletteStops[index + 1].Stop - PaletteStops[index].Stop);
                    return InterpolateColor(PaletteStops[index].Color, PaletteStops[index + 1].Color, localT);
                }
            }

            return PaletteStops[PaletteStops.Length - 1].Color;
        }

        private static Color InterpolateColor(Color start, Color end, double t)
        {
            t = Math.Max(0, Math.Min(1, t));
            int r = (int)Math.Round(start.R + (end.R - start.R) * t);
            int g = (int)Math.Round(start.G + (end.G - start.G) * t);
            int b = (int)Math.Round(start.B + (end.B - start.B) * t);
            return Color.FromArgb(r, g, b);
        }
    }
}
