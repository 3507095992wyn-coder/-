using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Linq;
using System.Windows.Forms;

namespace grbloxy
{
    internal sealed class HeatmapViewControl : UserControl
    {
        private readonly Label labelTitle;
        private readonly Label labelHint;
        private readonly Panel canvasPanel;
        private readonly HeatmapLegendControl legendControl;
        private HeatmapGridData currentGrid;
        private Bitmap cachedBitmap;
        private Size cachedBitmapSize = Size.Empty;

        public HeatmapViewControl()
        {
            BackColor = Color.White;

            labelTitle = new Label
            {
                AutoSize = false,
                Dock = DockStyle.Top,
                Height = 28,
                TextAlign = ContentAlignment.MiddleLeft,
                Padding = new Padding(8, 0, 8, 0),
                Font = new Font("Microsoft YaHei UI", 10.5f, FontStyle.Bold),
                BackColor = Color.FromArgb(245, 247, 250),
                ForeColor = Color.FromArgb(45, 62, 80),
                Text = "暂无热力图数据"
            };

            labelHint = new Label
            {
                AutoSize = false,
                Dock = DockStyle.Bottom,
                Height = 26,
                TextAlign = ContentAlignment.MiddleLeft,
                Padding = new Padding(8, 0, 8, 0),
                Font = new Font("Microsoft YaHei UI", 8.5f, FontStyle.Regular),
                BackColor = Color.FromArgb(245, 247, 250),
                ForeColor = Color.FromArgb(80, 80, 80),
                Text = "静态热力图视图：颜色表示 B 值大小"
            };

            legendControl = new HeatmapLegendControl
            {
                Dock = DockStyle.Right,
                Width = 142,
                Visible = false
            };

            canvasPanel = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = Color.White
            };
            canvasPanel.Paint += CanvasPanel_Paint;
            canvasPanel.Resize += CanvasPanel_Resize;

            Controls.Add(canvasPanel);
            Controls.Add(legendControl);
            Controls.Add(labelHint);
            Controls.Add(labelTitle);
        }

        public void ShowHeatmap(HeatmapGridData grid)
        {
            currentGrid = grid;
            labelTitle.Text = grid != null && grid.HasData
                ? string.IsNullOrWhiteSpace(grid.Subtitle) ? grid.Title : $"{grid.Title} | {grid.Subtitle}"
                : "暂无热力图数据";

            if (grid != null && grid.HasData)
            {
                legendControl.SetScale(grid.ValueLabel, grid.MinValue, grid.MaxValue);
                legendControl.Visible = true;
            }
            else
            {
                legendControl.Visible = false;
            }

            InvalidateHeatmap();
            canvasPanel.Invalidate();
        }

        public void ClearView(string message)
        {
            currentGrid = null;
            labelTitle.Text = message;
            legendControl.Visible = false;
            InvalidateHeatmap();
            canvasPanel.Invalidate();
        }

        public Bitmap CaptureBitmap(int width, int height)
        {
            int renderWidth = Math.Max(1, width);
            int renderHeight = Math.Max(1, height);
            Bitmap bitmap = new Bitmap(renderWidth, renderHeight);
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.Clear(Color.White);
                RenderFullControl(g, new Rectangle(0, 0, renderWidth, renderHeight));
            }

            return bitmap;
        }

        private void CanvasPanel_Resize(object sender, EventArgs e)
        {
            InvalidateHeatmap();
        }

        private void CanvasPanel_Paint(object sender, PaintEventArgs e)
        {
            e.Graphics.Clear(Color.White);
            e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
            e.Graphics.InterpolationMode = InterpolationMode.HighQualityBilinear;
            e.Graphics.PixelOffsetMode = PixelOffsetMode.HighQuality;
            RenderCanvas(e.Graphics, canvasPanel.ClientRectangle);
        }

        private void RenderFullControl(Graphics g, Rectangle bounds)
        {
            int titleHeight = 28;
            int hintHeight = 26;
            int legendWidth = currentGrid != null && currentGrid.HasData ? 142 : 0;

            Rectangle titleRect = new Rectangle(bounds.Left, bounds.Top, bounds.Width, titleHeight);
            Rectangle hintRect = new Rectangle(bounds.Left, bounds.Bottom - hintHeight, bounds.Width, hintHeight);
            Rectangle legendRect = new Rectangle(bounds.Right - legendWidth, titleRect.Bottom, legendWidth, bounds.Height - titleHeight - hintHeight);
            Rectangle canvasRect = new Rectangle(bounds.Left, titleRect.Bottom, bounds.Width - legendWidth, bounds.Height - titleHeight - hintHeight);

            using (SolidBrush headerBrush = new SolidBrush(Color.FromArgb(245, 247, 250)))
            using (SolidBrush textBrush = new SolidBrush(Color.FromArgb(45, 62, 80)))
            using (SolidBrush hintBrush = new SolidBrush(Color.FromArgb(80, 80, 80)))
            using (Font titleFont = new Font("Microsoft YaHei UI", 10.5f, FontStyle.Bold))
            using (Font hintFont = new Font("Microsoft YaHei UI", 8.5f))
            {
                g.FillRectangle(headerBrush, titleRect);
                g.DrawString(labelTitle.Text, titleFont, textBrush, new RectangleF(titleRect.Left + 8, titleRect.Top + 4, titleRect.Width - 16, titleRect.Height - 8));
                g.FillRectangle(headerBrush, hintRect);
                g.DrawString(labelHint.Text, hintFont, hintBrush, new RectangleF(hintRect.Left + 8, hintRect.Top + 5, hintRect.Width - 16, hintRect.Height - 8));
            }

            RenderCanvas(g, canvasRect);
            if (legendWidth > 0)
            {
                legendControl.RenderToGraphics(g, legendRect);
            }
        }

        private void RenderCanvas(Graphics g, Rectangle bounds)
        {
            if (bounds.Width <= 0 || bounds.Height <= 0)
            {
                return;
            }

            using (Pen borderPen = new Pen(Color.FromArgb(210, 217, 224)))
            using (SolidBrush messageBrush = new SolidBrush(Color.FromArgb(80, 80, 80)))
            using (Font axisFont = new Font("Microsoft YaHei UI", 8.5f))
            using (Font labelFont = new Font("Microsoft YaHei UI", 9f, FontStyle.Bold))
            {
                g.FillRectangle(Brushes.White, bounds);

                if (currentGrid == null || !currentGrid.HasData)
                {
                    string message = "暂无热力图数据";
                    SizeF size = g.MeasureString(message, labelFont);
                    g.DrawString(message, labelFont, messageBrush,
                        bounds.Left + (bounds.Width - size.Width) / 2f,
                        bounds.Top + (bounds.Height - size.Height) / 2f);
                    return;
                }

                Rectangle plotArea = new Rectangle(
                    bounds.Left + 54,
                    bounds.Top + 20,
                    Math.Max(80, bounds.Width - 92),
                    Math.Max(80, bounds.Height - 70));

                using (Bitmap heatmapBitmap = GetOrCreateHeatmapBitmap(plotArea.Size))
                {
                    g.DrawImage(heatmapBitmap, plotArea);
                }

                g.DrawRectangle(borderPen, plotArea);
                DrawAxes(g, plotArea, axisFont, labelFont, borderPen);
            }
        }

        private void DrawAxes(Graphics g, Rectangle plotArea, Font axisFont, Font labelFont, Pen borderPen)
        {
            double[] xValues = currentGrid.XValues;
            double[] yValues = currentGrid.YValues;

            for (int tickIndex = 0; tickIndex < 5; tickIndex++)
            {
                float t = tickIndex / 4f;
                float x = plotArea.Left + plotArea.Width * t;
                float y = plotArea.Bottom - plotArea.Height * t;

                g.DrawLine(borderPen, x, plotArea.Bottom, x, plotArea.Bottom + 4);
                g.DrawLine(borderPen, plotArea.Left - 4, y, plotArea.Left, y);

                double xValue = InterpolateAxisValue(xValues, t);
                double yValue = InterpolateAxisValue(yValues, t);

                string xText = FormatAxisValue(xValue);
                string yText = FormatAxisValue(yValue);

                SizeF xSize = g.MeasureString(xText, axisFont);
                SizeF ySize = g.MeasureString(yText, axisFont);

                g.DrawString(xText, axisFont, Brushes.Black, x - xSize.Width / 2f, plotArea.Bottom + 6);
                g.DrawString(yText, axisFont, Brushes.Black, plotArea.Left - ySize.Width - 8, y - ySize.Height / 2f);
            }

            StringFormat centerFormat = new StringFormat
            {
                Alignment = StringAlignment.Center,
                LineAlignment = StringAlignment.Center
            };

            g.DrawString(currentGrid.XLabel, labelFont, Brushes.Black,
                new RectangleF(plotArea.Left, plotArea.Bottom + 24, plotArea.Width, 20), centerFormat);

            GraphicsState state = g.Save();
            g.TranslateTransform(plotArea.Left - 40, plotArea.Top + plotArea.Height / 2f);
            g.RotateTransform(-90f);
            g.DrawString(currentGrid.YLabel, labelFont, Brushes.Black,
                new RectangleF(-plotArea.Height / 2f, -10f, plotArea.Height, 20f), centerFormat);
            g.Restore(state);
        }

        private Bitmap GetOrCreateHeatmapBitmap(Size targetSize)
        {
            if (cachedBitmap != null && cachedBitmapSize == targetSize)
            {
                return (Bitmap)cachedBitmap.Clone();
            }

            InvalidateHeatmap();
            cachedBitmapSize = targetSize;
            cachedBitmap = BuildHeatmapBitmap(targetSize);
            return (Bitmap)cachedBitmap.Clone();
        }

        private Bitmap BuildHeatmapBitmap(Size targetSize)
        {
            int width = Math.Max(1, targetSize.Width);
            int height = Math.Max(1, targetSize.Height);
            Bitmap bitmap = new Bitmap(width, height, PixelFormat.Format24bppRgb);
            BitmapData data = bitmap.LockBits(new Rectangle(0, 0, width, height), ImageLockMode.WriteOnly, bitmap.PixelFormat);
            byte[] pixelBytes = new byte[Math.Abs(data.Stride) * height];

            try
            {
                for (int py = 0; py < height; py++)
                {
                    double normalizedY = height == 1 ? 0 : py / (double)(height - 1);
                    int rowStart = py * data.Stride;

                    for (int px = 0; px < width; px++)
                    {
                        double normalizedX = width == 1 ? 0 : px / (double)(width - 1);
                        double value = SampleValue(normalizedX, 1d - normalizedY);
                        Color color = HeatmapColorMapper.GetHeatmapColor(value, currentGrid.MinValue, currentGrid.MaxValue);

                        int offset = rowStart + px * 3;
                        pixelBytes[offset] = color.B;
                        pixelBytes[offset + 1] = color.G;
                        pixelBytes[offset + 2] = color.R;
                    }
                }

                Marshal.Copy(pixelBytes, 0, data.Scan0, pixelBytes.Length);
            }
            finally
            {
                bitmap.UnlockBits(data);
            }

            return bitmap;
        }

        private double SampleValue(double normalizedX, double normalizedY)
        {
            if (currentGrid.XValues.Length == 1 && currentGrid.YValues.Length == 1)
            {
                return currentGrid.Values[0, 0];
            }

            double xPosition = normalizedX * Math.Max(0, currentGrid.XValues.Length - 1);
            double yPosition = normalizedY * Math.Max(0, currentGrid.YValues.Length - 1);

            int x0 = (int)Math.Floor(xPosition);
            int y0 = (int)Math.Floor(yPosition);
            int x1 = Math.Min(currentGrid.XValues.Length - 1, x0 + 1);
            int y1 = Math.Min(currentGrid.YValues.Length - 1, y0 + 1);

            double tx = xPosition - x0;
            double ty = yPosition - y0;

            double v00 = currentGrid.Values[y0, x0];
            double v10 = currentGrid.Values[y0, x1];
            double v01 = currentGrid.Values[y1, x0];
            double v11 = currentGrid.Values[y1, x1];

            if (double.IsNaN(v00)) v00 = FindNearestValidValue(x0, y0);
            if (double.IsNaN(v10)) v10 = FindNearestValidValue(x1, y0);
            if (double.IsNaN(v01)) v01 = FindNearestValidValue(x0, y1);
            if (double.IsNaN(v11)) v11 = FindNearestValidValue(x1, y1);

            double top = v00 + (v10 - v00) * tx;
            double bottom = v01 + (v11 - v01) * tx;
            return top + (bottom - top) * ty;
        }

        private double FindNearestValidValue(int x, int y)
        {
            double fallback = currentGrid.MinValue;
            double bestDistance = double.MaxValue;

            for (int row = 0; row < currentGrid.YValues.Length; row++)
            {
                for (int column = 0; column < currentGrid.XValues.Length; column++)
                {
                    double candidate = currentGrid.Values[row, column];
                    if (double.IsNaN(candidate))
                    {
                        continue;
                    }

                    double distance = Math.Pow(column - x, 2) + Math.Pow(row - y, 2);
                    if (distance < bestDistance)
                    {
                        bestDistance = distance;
                        fallback = candidate;
                    }
                }
            }

            return fallback;
        }

        private static double InterpolateAxisValue(double[] axisValues, float normalized)
        {
            if (axisValues == null || axisValues.Length == 0)
            {
                return 0;
            }

            if (axisValues.Length == 1)
            {
                return axisValues[0];
            }

            double position = normalized * (axisValues.Length - 1);
            int left = (int)Math.Floor(position);
            int right = Math.Min(axisValues.Length - 1, left + 1);
            double t = position - left;
            return axisValues[left] + (axisValues[right] - axisValues[left]) * t;
        }

        private static string FormatAxisValue(double value)
        {
            if (Math.Abs(value) < 1e-9)
            {
                value = 0;
            }

            return value.ToString("0.###", CultureInfo.InvariantCulture);
        }

        private void InvalidateHeatmap()
        {
            cachedBitmap?.Dispose();
            cachedBitmap = null;
            cachedBitmapSize = Size.Empty;
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                cachedBitmap?.Dispose();
            }

            base.Dispose(disposing);
        }

        private sealed class HeatmapLegendControl : System.Windows.Forms.Control
        {
            private string valueLabel = "B";
            private double minValue;
            private double maxValue = 1;

            public HeatmapLegendControl()
            {
                BackColor = Color.FromArgb(246, 248, 250);
                DoubleBuffered = true;
            }

            public void SetScale(string label, double min, double max)
            {
                valueLabel = string.IsNullOrWhiteSpace(label) ? "B" : label;
                minValue = min;
                maxValue = max;
                Invalidate();
            }

            public void RenderToGraphics(Graphics g, Rectangle bounds)
            {
                using (SolidBrush backgroundBrush = new SolidBrush(BackColor))
                using (Pen borderPen = new Pen(Color.FromArgb(210, 217, 224)))
                using (Font titleFont = new Font("Microsoft YaHei UI", 9f, FontStyle.Bold))
                using (Font tickFont = new Font("Microsoft YaHei UI", 8.2f, FontStyle.Regular))
                {
                    g.FillRectangle(backgroundBrush, bounds);
                    g.DrawRectangle(borderPen, bounds);

                    Rectangle barRect = new Rectangle(bounds.Left + 36, bounds.Top + 38, 28, Math.Max(120, bounds.Height - 86));
                    for (int y = 0; y < barRect.Height; y++)
                    {
                        double normalized = 1d - y / (double)Math.Max(1, barRect.Height - 1);
                        using (Pen linePen = new Pen(HeatmapColorMapper.GetColorFromNormalized(normalized)))
                        {
                            g.DrawLine(linePen, barRect.Left, barRect.Top + y, barRect.Right, barRect.Top + y);
                        }
                    }

                    g.DrawRectangle(Pens.Gray, barRect);
                    g.DrawString("Color Bar", titleFont, Brushes.Black, bounds.Left + 18, bounds.Top + 10);
                    g.DrawString(valueLabel, titleFont, Brushes.Black, bounds.Left + 72, bounds.Top + 42);
                    g.DrawString(maxValue.ToString("0.###", CultureInfo.InvariantCulture), tickFont, Brushes.Black, bounds.Left + 72, barRect.Top - 4);
                    g.DrawString(((minValue + maxValue) / 2d).ToString("0.###", CultureInfo.InvariantCulture), tickFont, Brushes.Black, bounds.Left + 72, barRect.Top + barRect.Height / 2f - 8);
                    g.DrawString(minValue.ToString("0.###", CultureInfo.InvariantCulture), tickFont, Brushes.Black, bounds.Left + 72, barRect.Bottom - 18);
                }
            }

            protected override void OnPaint(PaintEventArgs e)
            {
                base.OnPaint(e);
                RenderToGraphics(e.Graphics, ClientRectangle);
            }
        }
    }

    internal sealed class HeatmapGridData
    {
        public string Title { get; set; } = string.Empty;
        public string Subtitle { get; set; } = string.Empty;
        public string XLabel { get; set; } = "X (mm)";
        public string YLabel { get; set; } = "Y (mm)";
        public string ValueLabel { get; set; } = "B";
        public double[] XValues { get; set; } = Array.Empty<double>();
        public double[] YValues { get; set; } = Array.Empty<double>();
        public double[,] Values { get; set; } = new double[0, 0];

        public bool HasData => XValues.Length > 0 && YValues.Length > 0 && Values.Length > 0;

        public double MinValue
        {
            get
            {
                return EnumerateValidValues().DefaultIfEmpty(0d).Min();
            }
        }

        public double MaxValue
        {
            get
            {
                return EnumerateValidValues().DefaultIfEmpty(1d).Max();
            }
        }

        private System.Collections.Generic.IEnumerable<double> EnumerateValidValues()
        {
            foreach (double value in Values)
            {
                if (!double.IsNaN(value) && !double.IsInfinity(value))
                {
                    yield return value;
                }
            }
        }
    }
}
