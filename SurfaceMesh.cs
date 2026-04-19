using System;
using System.Collections.Generic;
using System.Linq;

namespace grbloxy
{
    internal sealed class SurfaceMesh
    {
        public string Title { get; set; } = string.Empty;
        public string Subtitle { get; set; } = string.Empty;
        public string XLabel { get; set; } = "X";
        public string YLabel { get; set; } = "Y";
        public string ZLabel { get; set; } = "B";
        public List<SurfaceVertex> Vertices { get; } = new List<SurfaceVertex>();
        public List<SurfaceQuad> Quads { get; } = new List<SurfaceQuad>();

        public bool HasData => Vertices.Count > 0;

        public double MinX => HasData ? Vertices.Min(v => v.X) : 0;
        public double MaxX => HasData ? Vertices.Max(v => v.X) : 1;
        public double MinY => HasData ? Vertices.Min(v => v.Y) : 0;
        public double MaxY => HasData ? Vertices.Max(v => v.Y) : 1;
        public double MinZ => HasData ? Vertices.Min(v => v.Z) : 0;
        public double MaxZ => HasData ? Vertices.Max(v => v.Z) : 1;

        public double XSpan => Math.Max(1e-9, MaxX - MinX);
        public double YSpan => Math.Max(1e-9, MaxY - MinY);
        public double ZSpan => Math.Max(1e-9, MaxZ - MinZ);
    }

    internal sealed class SurfaceVertex
    {
        public SurfaceVertex(double x, double y, double z)
        {
            X = x;
            Y = y;
            Z = z;
        }

        public double X { get; }
        public double Y { get; }
        public double Z { get; }
    }

    internal sealed class SurfaceQuad
    {
        public SurfaceQuad(SurfaceVertex v00, SurfaceVertex v10, SurfaceVertex v11, SurfaceVertex v01)
        {
            V00 = v00;
            V10 = v10;
            V11 = v11;
            V01 = v01;
        }

        public SurfaceVertex V00 { get; }
        public SurfaceVertex V10 { get; }
        public SurfaceVertex V11 { get; }
        public SurfaceVertex V01 { get; }
    }

    internal sealed class PlotCurvePoint
    {
        public double AxisValue { get; set; }
        public double B { get; set; }
    }

    internal sealed class PlotCurveData
    {
        public string Title { get; set; } = string.Empty;
        public string XLabel { get; set; } = string.Empty;
        public string YLabel { get; set; } = "B";
        public IReadOnlyList<PlotCurvePoint> Points { get; set; } = Array.Empty<PlotCurvePoint>();
    }

    internal sealed class PlotCurveLayerSource
    {
        public int LayerIndex { get; set; }
        public double ZValue { get; set; }
        public IReadOnlyList<Control.PointData> Points { get; set; } = Array.Empty<Control.PointData>();
    }

    internal sealed class PlotCurveSeries
    {
        public int? LayerIndex { get; set; }
        public double? ZValue { get; set; }
        public string LegendText { get; set; } = string.Empty;
        public IReadOnlyList<PlotCurvePoint> Points { get; set; } = Array.Empty<PlotCurvePoint>();

        public bool HasData => Points != null && Points.Count > 0;
    }

    internal sealed class PlotCurveCollectionData
    {
        public string Title { get; set; } = string.Empty;
        public string XLabel { get; set; } = string.Empty;
        public string YLabel { get; set; } = "B";
        public IReadOnlyList<PlotCurveSeries> LayerCurves { get; set; } = Array.Empty<PlotCurveSeries>();
        public PlotCurveSeries AverageCurve { get; set; }

        public bool HasData => (LayerCurves?.Any(series => series != null && series.HasData) ?? false) ||
            (AverageCurve?.HasData ?? false);
    }
}
