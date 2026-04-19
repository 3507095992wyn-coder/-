using System;
using System.Collections.Generic;
using System.Linq;

namespace grbloxy
{
    internal static class PlotDataBuilder
    {
        private const double AxisTolerance = 1e-6;

        public static SurfaceMesh BuildSurfaceMesh(IEnumerable<Control.PointData> points, string title, string subtitle)
        {
            List<Control.PointData> validPoints = (points ?? Enumerable.Empty<Control.PointData>())
                .Where(point => point.Order >= 0)
                .OrderBy(point => point.Y)
                .ThenBy(point => point.X)
                .ToList();

            var mesh = new SurfaceMesh
            {
                Title = title ?? string.Empty,
                Subtitle = subtitle ?? string.Empty,
                XLabel = "X (mm)",
                YLabel = "Y (mm)",
                ZLabel = "B"
            };

            if (validPoints.Count == 0)
            {
                return mesh;
            }

            Dictionary<string, SurfaceVertex> lookup = validPoints
                .GroupBy(point => BuildAxisKey(point.X, point.Y))
                .Select(group => new SurfaceVertex(group.First().X, group.First().Y, group.Average(item => item.B)))
                .ToDictionary(vertex => BuildAxisKey(vertex.X, vertex.Y));

            mesh.Vertices.AddRange(lookup.Values.OrderBy(v => v.Y).ThenBy(v => v.X));

            List<double> xValues = mesh.Vertices.Select(v => v.X).Distinct().OrderBy(v => v).ToList();
            List<double> yValues = mesh.Vertices.Select(v => v.Y).Distinct().OrderBy(v => v).ToList();

            for (int yIndex = 0; yIndex < yValues.Count - 1; yIndex++)
            {
                for (int xIndex = 0; xIndex < xValues.Count - 1; xIndex++)
                {
                    if (TryGetVertex(lookup, xValues[xIndex], yValues[yIndex], out SurfaceVertex v00) &&
                        TryGetVertex(lookup, xValues[xIndex + 1], yValues[yIndex], out SurfaceVertex v10) &&
                        TryGetVertex(lookup, xValues[xIndex + 1], yValues[yIndex + 1], out SurfaceVertex v11) &&
                        TryGetVertex(lookup, xValues[xIndex], yValues[yIndex + 1], out SurfaceVertex v01))
                    {
                        mesh.Quads.Add(new SurfaceQuad(v00, v10, v11, v01));
                    }
                }
            }

            return mesh;
        }

        public static SurfaceMesh BuildAverageAlongZSurface(IEnumerable<Control.PointData> points)
        {
            List<Control.PointData> averagedPoints = (points ?? Enumerable.Empty<Control.PointData>())
                .Where(point => point.Order >= 0)
                .GroupBy(point => BuildAxisKey(point.X, point.Y))
                .Select((group, index) =>
                {
                    Control.PointData seed = group.First();
                    return new Control.PointData(index, seed.X, seed.Y, 0d, Control.ConvertBToVoltage(group.Average(item => item.B)));
                })
                .OrderBy(point => point.Y)
                .ThenBy(point => point.X)
                .ToList();

            SurfaceMesh mesh = BuildSurfaceMesh(
                averagedPoints,
                "沿 Z 平均三维图",
                "对相同 (X, Y) 点沿 Z 方向取平均后的 B");
            mesh.ZLabel = "Avg B";
            return mesh;
        }

        public static PlotCurveData BuildAverageCurve(IEnumerable<Control.PointData> points, Func<Control.PointData, double> axisSelector, string axisName, string title)
        {
            List<Control.PointData> validPoints = (points ?? Enumerable.Empty<Control.PointData>())
                .Where(point => point.Order >= 0)
                .ToList();

            List<PlotCurvePoint> curvePoints = BuildCurvePoints(validPoints, axisSelector, ResolveAxisTolerance(validPoints.Select(axisSelector)));

            return new PlotCurveData
            {
                Title = title,
                XLabel = $"{axisName} (mm)",
                YLabel = "B",
                Points = curvePoints
            };
        }

        public static PlotCurveCollectionData BuildLayeredCurve(
            IEnumerable<PlotCurveLayerSource> layers,
            Func<Control.PointData, double> axisSelector,
            string axisName,
            string title)
        {
            List<PlotCurveLayerSource> validLayers = (layers ?? Enumerable.Empty<PlotCurveLayerSource>())
                .Where(layer => layer != null)
                .OrderBy(layer => layer.LayerIndex)
                .ToList();

            List<Control.PointData> allPoints = validLayers
                .SelectMany(layer => layer.Points ?? Array.Empty<Control.PointData>())
                .Where(point => point.Order >= 0)
                .ToList();

            var data = new PlotCurveCollectionData
            {
                Title = title ?? string.Empty,
                XLabel = $"{axisName} (mm)",
                YLabel = "B"
            };

            if (allPoints.Count == 0)
            {
                return data;
            }

            double axisTolerance = ResolveAxisTolerance(allPoints.Select(axisSelector));
            List<PlotCurveSeries> layerCurves = new List<PlotCurveSeries>();

            foreach (PlotCurveLayerSource layer in validLayers)
            {
                List<Control.PointData> layerPoints = (layer.Points ?? Array.Empty<Control.PointData>())
                    .Where(point => point.Order >= 0)
                    .ToList();

                List<PlotCurvePoint> curvePoints = BuildCurvePoints(layerPoints, axisSelector, axisTolerance);
                if (curvePoints.Count == 0)
                {
                    continue;
                }

                layerCurves.Add(new PlotCurveSeries
                {
                    LayerIndex = layer.LayerIndex,
                    ZValue = NormalizeAxis(layer.ZValue),
                    LegendText = $"z={FormatLegendValue(layer.ZValue)}",
                    Points = curvePoints
                });
            }

            List<PlotCurvePoint> averageCurve = BuildCurvePoints(allPoints, axisSelector, axisTolerance);
            data.LayerCurves = layerCurves;
            data.AverageCurve = averageCurve.Count == 0
                ? null
                : new PlotCurveSeries
                {
                    LegendText = "平均值",
                    Points = averageCurve
                };

            return data;
        }

        public static HeatmapGridData BuildHeatmapGrid(IEnumerable<Control.PointData> points, string title, string subtitle)
        {
            List<Control.PointData> validPoints = (points ?? Enumerable.Empty<Control.PointData>())
                .Where(point => point.Order >= 0)
                .OrderBy(point => point.Y)
                .ThenBy(point => point.X)
                .ToList();

            var grid = new HeatmapGridData
            {
                Title = title ?? string.Empty,
                Subtitle = subtitle ?? string.Empty,
                XLabel = "X (mm)",
                YLabel = "Y (mm)",
                ValueLabel = "B"
            };

            if (validPoints.Count == 0)
            {
                return grid;
            }

            Dictionary<string, double> valueLookup = validPoints
                .GroupBy(point => BuildAxisKey(point.X, point.Y))
                .ToDictionary(group => group.Key, group => group.Average(item => item.B));

            List<double> xValues = validPoints.Select(point => NormalizeAxis(point.X)).Distinct().OrderBy(value => value).ToList();
            List<double> yValues = validPoints.Select(point => NormalizeAxis(point.Y)).Distinct().OrderBy(value => value).ToList();
            double[,] values = new double[yValues.Count, xValues.Count];

            for (int yIndex = 0; yIndex < yValues.Count; yIndex++)
            {
                for (int xIndex = 0; xIndex < xValues.Count; xIndex++)
                {
                    values[yIndex, xIndex] = valueLookup.TryGetValue(BuildAxisKey(xValues[xIndex], yValues[yIndex]), out double value)
                        ? value
                        : double.NaN;
                }
            }

            grid.XValues = xValues.ToArray();
            grid.YValues = yValues.ToArray();
            grid.Values = values;
            return grid;
        }

        public static HeatmapGridData BuildAverageAlongZHeatmap(IEnumerable<Control.PointData> points)
        {
            List<Control.PointData> averagedPoints = (points ?? Enumerable.Empty<Control.PointData>())
                .Where(point => point.Order >= 0)
                .GroupBy(point => BuildAxisKey(point.X, point.Y))
                .Select((group, index) =>
                {
                    Control.PointData seed = group.First();
                    return new Control.PointData(index, seed.X, seed.Y, 0d, Control.ConvertBToVoltage(group.Average(item => item.B)));
                })
                .OrderBy(point => point.Y)
                .ThenBy(point => point.X)
                .ToList();

            return BuildHeatmapGrid(
                averagedPoints,
                "沿 Z 平均静态热力图",
                "对相同 (X, Y) 点沿 Z 方向取平均后的 B");
        }

        private static bool TryGetVertex(Dictionary<string, SurfaceVertex> lookup, double x, double y, out SurfaceVertex vertex)
        {
            return lookup.TryGetValue(BuildAxisKey(x, y), out vertex);
        }

        private static string BuildAxisKey(double x, double y)
        {
            return $"{NormalizeAxis(x):F6}_{NormalizeAxis(y):F6}";
        }

        private static List<PlotCurvePoint> BuildCurvePoints(
            IEnumerable<Control.PointData> points,
            Func<Control.PointData, double> axisSelector,
            double axisTolerance)
        {
            return (points ?? Enumerable.Empty<Control.PointData>())
                .Where(point => point.Order >= 0)
                .GroupBy(point => NormalizeAxis(axisSelector(point), axisTolerance))
                .Select(group => new PlotCurvePoint
                {
                    AxisValue = group.Key,
                    B = group.Average(item => item.B)
                })
                .OrderBy(point => point.AxisValue)
                .ToList();
        }

        private static double ResolveAxisTolerance(IEnumerable<double> values)
        {
            List<double> normalizedValues = (values ?? Enumerable.Empty<double>())
                .Select(value => NormalizeAxis(value))
                .Distinct()
                .OrderBy(value => value)
                .ToList();

            if (normalizedValues.Count < 2)
            {
                return AxisTolerance;
            }

            List<double> positiveDiffs = normalizedValues
                .Zip(normalizedValues.Skip(1), (previous, next) => next - previous)
                .Where(diff => diff > AxisTolerance)
                .ToList();

            if (positiveDiffs.Count == 0)
            {
                return AxisTolerance;
            }

            double step = positiveDiffs.Min();
            double derivedTolerance = step / 4d;
            return Math.Max(AxisTolerance, Math.Round(derivedTolerance, 6));
        }

        private static string FormatLegendValue(double value)
        {
            double normalized = NormalizeAxis(value);
            return normalized.ToString("0.###");
        }

        private static double NormalizeAxis(double value, double tolerance)
        {
            double safeTolerance = tolerance > AxisTolerance ? tolerance : AxisTolerance;
            double snapped = Math.Round(value / safeTolerance) * safeTolerance;
            return NormalizeAxis(snapped);
        }

        private static double NormalizeAxis(double value)
        {
            return Math.Abs(value) < AxisTolerance ? 0d : Math.Round(value, 6);
        }
    }
}
