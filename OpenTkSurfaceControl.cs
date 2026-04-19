using OpenTK;
using OpenTK.Graphics;
using OpenTK.Graphics.OpenGL;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.Globalization;
using System.Linq;
using System.Windows.Forms;

namespace grbloxy
{
    internal sealed class OpenTkSurfaceControl : UserControl
    {
        private readonly Panel contentPanel;
        private readonly Panel viewportHost;
        private readonly GLControl glControl;
        private readonly AxisOverlayControl axisOverlay;
        private readonly Label labelTitle;
        private readonly Label labelHint;
        private readonly SurfaceColorBarControl colorBar;

        private SurfaceMesh currentMesh;
        private PreparedMesh preparedMesh;
        private Point lastMouseLocation;
        private bool isRotating;
        private bool isPanning;
        private float yawDegrees = -40f;
        private float pitchDegrees = 28f;
        private float zoomDistance = 5.2f;
        private float panX;
        private float panY;
        private bool glReady;
        private Matrix4 lastProjectionMatrix;
        private Matrix4 lastModelViewMatrix;
        private int lastViewportWidth;
        private int lastViewportHeight;
        private IReadOnlyList<AxisTick> xTicks = Array.Empty<AxisTick>();
        private IReadOnlyList<AxisTick> yTicks = Array.Empty<AxisTick>();
        private IReadOnlyList<AxisTick> zTicks = Array.Empty<AxisTick>();

        public OpenTkSurfaceControl()
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
                Text = "暂无三维数据"
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
                Text = "左键旋转 | 右键平移 | 滚轮缩放 | 双击复位视角"
            };

            contentPanel = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = Color.White
            };

            colorBar = new SurfaceColorBarControl
            {
                Dock = DockStyle.Right,
                Width = 142,
                Visible = false
            };

            viewportHost = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = Color.White
            };

            glControl = new GLControl(new GraphicsMode(32, 24, 0, 4))
            {
                Dock = DockStyle.Fill,
                BackColor = Color.White,
                VSync = true
            };

            axisOverlay = new AxisOverlayControl
            {
                Dock = DockStyle.Fill,
                BackColor = Color.Transparent
            };

            viewportHost.Controls.Add(glControl);
            viewportHost.Controls.Add(axisOverlay);
            contentPanel.Controls.Add(viewportHost);
            contentPanel.Controls.Add(colorBar);

            Controls.Add(contentPanel);
            Controls.Add(labelHint);
            Controls.Add(labelTitle);

            glControl.Load += GlControl_Load;
            glControl.Paint += GlControl_Paint;
            glControl.Resize += GlControl_Resize;
            glControl.MouseDown += GlControl_MouseDown;
            glControl.MouseUp += GlControl_MouseUp;
            glControl.MouseMove += GlControl_MouseMove;
            glControl.MouseWheel += GlControl_MouseWheel;
            glControl.DoubleClick += GlControl_DoubleClick;
            viewportHost.Resize += ViewportHost_Resize;
        }

        public void ShowMesh(SurfaceMesh mesh)
        {
            currentMesh = mesh;
            preparedMesh = PrepareMesh(mesh);
            labelTitle.Text = mesh != null && mesh.HasData
                ? string.IsNullOrWhiteSpace(mesh.Subtitle) ? mesh.Title : $"{mesh.Title} | {mesh.Subtitle}"
                : "暂无三维数据";

            if (mesh != null && mesh.HasData)
            {
                ResetView();
                colorBar.SetScale(mesh.ZLabel, mesh.MinZ, mesh.MaxZ);
                colorBar.Visible = true;
                colorBar.BringToFront();
            }
            else
            {
                colorBar.Visible = false;
            }

            UpdateOverlay();
            glControl.Invalidate();
        }

        public void ClearView(string message)
        {
            currentMesh = null;
            preparedMesh = null;
            labelTitle.Text = message;
            colorBar.Visible = false;
            UpdateOverlay();
            glControl.Invalidate();
        }

        public Bitmap CaptureBitmap(int width, int height)
        {
            if (!glReady)
                throw new InvalidOperationException("OpenGL 视图尚未初始化完成。");

            int renderWidth = Math.Max(1, width);
            int renderHeight = Math.Max(1, height);

            glControl.MakeCurrent();
            RenderScene(new Size(renderWidth, renderHeight));

            Bitmap bitmap = new Bitmap(renderWidth, renderHeight, System.Drawing.Imaging.PixelFormat.Format24bppRgb);
            BitmapData bitmapData = bitmap.LockBits(
                new Rectangle(0, 0, renderWidth, renderHeight),
                ImageLockMode.WriteOnly,
                System.Drawing.Imaging.PixelFormat.Format24bppRgb);

            try
            {
                GL.ReadPixels(0, 0, renderWidth, renderHeight, OpenTK.Graphics.OpenGL.PixelFormat.Bgr, PixelType.UnsignedByte, bitmapData.Scan0);
            }
            finally
            {
                bitmap.UnlockBits(bitmapData);
            }

            bitmap.RotateFlip(RotateFlipType.RotateNoneFlipY);
            return bitmap;
        }

        private void GlControl_Load(object sender, EventArgs e)
        {
            glControl.MakeCurrent();
            GL.ClearColor(Color.White);
            GL.Enable(EnableCap.DepthTest);
            GL.Enable(EnableCap.Blend);
            GL.BlendFunc(BlendingFactor.SrcAlpha, BlendingFactor.OneMinusSrcAlpha);
            GL.Enable(EnableCap.LineSmooth);
            GL.Hint(HintTarget.LineSmoothHint, HintMode.Nicest);
            GL.ShadeModel(ShadingModel.Smooth);
            glReady = true;
            ResetView();
        }

        private void GlControl_Resize(object sender, EventArgs e)
        {
            if (!glReady)
                return;

            glControl.MakeCurrent();
            GL.Viewport(0, 0, Math.Max(1, glControl.ClientSize.Width), Math.Max(1, glControl.ClientSize.Height));
            glControl.Invalidate();
            UpdateOverlay();
        }

        private void ViewportHost_Resize(object sender, EventArgs e)
        {
            UpdateOverlay();
        }

        private void GlControl_Paint(object sender, PaintEventArgs e)
        {
            if (!glReady)
                return;

            RenderScene(glControl.ClientSize);
            glControl.SwapBuffers();
            UpdateOverlay();
        }

        private void RenderScene(Size renderSize)
        {
            int width = Math.Max(1, renderSize.Width);
            int height = Math.Max(1, renderSize.Height);
            float aspect = width / (float)Math.Max(1, height);

            GL.Viewport(0, 0, width, height);
            GL.Clear(ClearBufferMask.ColorBufferBit | ClearBufferMask.DepthBufferBit);

            Matrix4 projection = Matrix4.CreatePerspectiveFieldOfView(
                MathHelper.DegreesToRadians(45f),
                aspect,
                0.1f,
                100f);
            GL.MatrixMode(MatrixMode.Projection);
            GL.LoadMatrix(ref projection);
            lastProjectionMatrix = projection;

            Matrix4 modelView =
                Matrix4.CreateTranslation(panX, panY, -zoomDistance) *
                Matrix4.CreateRotationX(MathHelper.DegreesToRadians(pitchDegrees)) *
                Matrix4.CreateRotationZ(MathHelper.DegreesToRadians(yawDegrees));
            GL.MatrixMode(MatrixMode.Modelview);
            GL.LoadMatrix(ref modelView);
            lastModelViewMatrix = modelView;
            lastViewportWidth = width;
            lastViewportHeight = height;

            if (preparedMesh == null || preparedMesh.Vertices.Count == 0)
                return;

            BuildAxisTicks();
            DrawReferencePlanes();
            DrawGridPlanes();
            DrawAxesWithTicks();
            DrawSurface();
        }

        private void DrawReferencePlanes()
        {
            float x0 = preparedMesh.MinX;
            float x1 = preparedMesh.MaxX;
            float y0 = preparedMesh.MinY;
            float y1 = preparedMesh.MaxY;
            float z0 = preparedMesh.BaseZ;
            float z1 = preparedMesh.MaxZ;

            GL.Enable(EnableCap.Blend);
            GL.Begin(PrimitiveType.Quads);

            GL.Color4(Color.FromArgb(22, 210, 218, 226));
            GL.Vertex3(x0, y0, z0);
            GL.Vertex3(x1, y0, z0);
            GL.Vertex3(x1, y1, z0);
            GL.Vertex3(x0, y1, z0);

            GL.Color4(Color.FromArgb(18, 205, 214, 223));
            GL.Vertex3(x0, y0, z0);
            GL.Vertex3(x1, y0, z0);
            GL.Vertex3(x1, y0, z1);
            GL.Vertex3(x0, y0, z1);

            GL.Color4(Color.FromArgb(16, 198, 208, 218));
            GL.Vertex3(x0, y0, z0);
            GL.Vertex3(x0, y1, z0);
            GL.Vertex3(x0, y1, z1);
            GL.Vertex3(x0, y0, z1);

            GL.End();
        }

        private void DrawGridPlanes()
        {
            GL.LineWidth(1f);
            GL.Color4(Color.FromArgb(186, 194, 204));
            GL.Begin(PrimitiveType.Lines);

            float x0 = preparedMesh.MinX;
            float x1 = preparedMesh.MaxX;
            float y0 = preparedMesh.MinY;
            float y1 = preparedMesh.MaxY;
            float z0 = preparedMesh.BaseZ;
            float z1 = preparedMesh.MaxZ;

            foreach (AxisTick xt in xTicks)
            {
                float x = Lerp(x0, x1, xt.Position);
                GL.Vertex3(x, y0, z0);
                GL.Vertex3(x, y1, z0);
                GL.Vertex3(x, y0, z0);
                GL.Vertex3(x, y0, z1);
            }

            foreach (AxisTick yt in yTicks)
            {
                float y = Lerp(y0, y1, yt.Position);
                GL.Vertex3(x0, y, z0);
                GL.Vertex3(x1, y, z0);
                GL.Vertex3(x0, y, z0);
                GL.Vertex3(x0, y, z1);
            }

            foreach (AxisTick zt in zTicks)
            {
                float z = Lerp(z0, z1, zt.Position);
                GL.Vertex3(x0, y0, z);
                GL.Vertex3(x1, y0, z);
                GL.Vertex3(x0, y0, z);
                GL.Vertex3(x0, y1, z);
            }

            GL.End();
        }

        private void DrawAxesWithTicks()
        {
            float x0 = preparedMesh.MinX;
            float x1 = preparedMesh.MaxX;
            float y0 = preparedMesh.MinY;
            float y1 = preparedMesh.MaxY;
            float z0 = preparedMesh.BaseZ;
            float z1 = preparedMesh.MaxZ;
            Color axisColor = Color.FromArgb(34, 43, 54);

            GL.LineWidth(2.4f);
            GL.Begin(PrimitiveType.Lines);
            GL.Color4(axisColor);
            GL.Vertex3(x0, y0, z0); GL.Vertex3(x1, y0, z0);
            GL.Vertex3(x0, y0, z0); GL.Vertex3(x0, y1, z0);
            GL.Vertex3(x0, y0, z0); GL.Vertex3(x0, y0, z1);
            GL.End();

            float xyTick = Math.Max(0.015f, Math.Min(0.06f, Math.Max(preparedMesh.XSpan, preparedMesh.YSpan) * 0.015f));
            float zTick = Math.Max(0.015f, Math.Min(0.06f, preparedMesh.ZSpan * 0.03f));

            GL.LineWidth(1.2f);
            GL.Begin(PrimitiveType.Lines);
            foreach (AxisTick tick in xTicks)
            {
                float x = Lerp(x0, x1, tick.Position);
                GL.Color4(axisColor);
                GL.Vertex3(x, y0, z0);
                GL.Vertex3(x, y0 - xyTick, z0);
            }

            foreach (AxisTick tick in yTicks)
            {
                float y = Lerp(y0, y1, tick.Position);
                GL.Color4(axisColor);
                GL.Vertex3(x0, y, z0);
                GL.Vertex3(x0 - xyTick, y, z0);
            }

            foreach (AxisTick tick in zTicks)
            {
                float z = Lerp(z0, z1, tick.Position);
                GL.Color4(axisColor);
                GL.Vertex3(x0, y0, z);
                GL.Vertex3(x0 - zTick, y0, z);
            }
            GL.End();
        }

        private void DrawSurface()
        {
            GL.Enable(EnableCap.PolygonOffsetFill);
            GL.PolygonOffset(1f, 1f);

            foreach (PreparedQuad quad in preparedMesh.Quads)
            {
                GL.Begin(PrimitiveType.Quads);
                DrawVertex(quad.V00);
                DrawVertex(quad.V10);
                DrawVertex(quad.V11);
                DrawVertex(quad.V01);
                GL.End();
            }

            GL.Disable(EnableCap.PolygonOffsetFill);

            if (preparedMesh.Quads.Count <= 1800)
            {
                GL.LineWidth(0.8f);
                GL.Color4(Color.FromArgb(52, 44, 62, 80));
                foreach (PreparedQuad quad in preparedMesh.Quads)
                {
                    GL.Begin(PrimitiveType.LineLoop);
                    GL.Vertex3(quad.V00.X, quad.V00.Y, quad.V00.Z);
                    GL.Vertex3(quad.V10.X, quad.V10.Y, quad.V10.Z);
                    GL.Vertex3(quad.V11.X, quad.V11.Y, quad.V11.Z);
                    GL.Vertex3(quad.V01.X, quad.V01.Y, quad.V01.Z);
                    GL.End();
                }
            }

            if (preparedMesh.Vertices.Count <= 900)
            {
                GL.PointSize(preparedMesh.Vertices.Count > 500 ? 2.8f : 4f);
                GL.Begin(PrimitiveType.Points);
                foreach (PreparedVertex vertex in preparedMesh.Vertices)
                {
                    Color color = MapColor(vertex.OriginalZ);
                    GL.Color4(Color.FromArgb(180, color));
                    GL.Vertex3(vertex.X, vertex.Y, vertex.Z);
                }
                GL.End();
            }
        }

        private void DrawVertex(PreparedVertex vertex)
        {
            Color color = MapColor(vertex.OriginalZ);
            GL.Color4(Color.FromArgb(238, color));
            GL.Vertex3(vertex.X, vertex.Y, vertex.Z);
        }

        private void UpdateOverlay()
        {
            List<OverlayPlacement> placements = BuildOverlayPlacements();
            axisOverlay.SetPlacements(placements);
            axisOverlay.BringToFront();
        }

        private List<OverlayPlacement> BuildOverlayPlacements()
        {
            var placements = new List<OverlayPlacement>();
            if (preparedMesh == null || !glReady || lastViewportWidth <= 0 || lastViewportHeight <= 0)
                return placements;

            float x0 = preparedMesh.MinX;
            float x1 = preparedMesh.MaxX;
            float y0 = preparedMesh.MinY;
            float y1 = preparedMesh.MaxY;
            float z0 = preparedMesh.BaseZ;
            float z1 = preparedMesh.MaxZ;

            Color axisColor = Color.FromArgb(34, 43, 54);

            foreach (AxisTick tick in xTicks)
                AddPlacementIfVisible(placements, new Vector3(Lerp(x0, x1, tick.Position), y0, z0), tick.Label, axisColor, false, 0f, 12f);

            foreach (AxisTick tick in yTicks)
                AddPlacementIfVisible(placements, new Vector3(x0, Lerp(y0, y1, tick.Position), z0), tick.Label, axisColor, false, -28f, 0f);

            foreach (AxisTick tick in zTicks)
                AddPlacementIfVisible(placements, new Vector3(x0, y0, Lerp(z0, z1, tick.Position)), tick.Label, axisColor, false, -28f, 0f);

            AddPlacementIfVisible(placements, new Vector3(x1, y0, z0), currentMesh?.XLabel ?? "X", axisColor, true, 22f, 12f);
            AddPlacementIfVisible(placements, new Vector3(x0, y1, z0), currentMesh?.YLabel ?? "Y", axisColor, true, -20f, -16f);
            AddPlacementIfVisible(placements, new Vector3(x0, y0, z1), currentMesh?.ZLabel ?? "B", axisColor, true, -18f, -16f);

            return placements;
        }

        private void AddPlacementIfVisible(List<OverlayPlacement> placements, Vector3 world, string text, Color color, bool bold, float dx, float dy)
        {
            if (!TryProjectWorldToScreen(world, out PointF screen))
                return;

            screen.X += dx;
            screen.Y += dy;

            if (screen.X < -40 || screen.X > viewportHost.ClientSize.Width + 40 || screen.Y < -30 || screen.Y > viewportHost.ClientSize.Height + 30)
                return;

            placements.Add(new OverlayPlacement
            {
                Text = text,
                Color = color,
                Bold = bold,
                Screen = screen,
                Visible = true
            });
        }

        private void BuildAxisTicks()
        {
            xTicks = CalculateTicks(preparedMesh.MinOriginalX, preparedMesh.MaxOriginalX)
                .Select(value => new AxisTick(value, MapOriginalToNormalized(value, preparedMesh.MinOriginalX, preparedMesh.MaxOriginalX), FormatTickLabel(value)))
                .ToList();
            yTicks = CalculateTicks(preparedMesh.MinOriginalY, preparedMesh.MaxOriginalY)
                .Select(value => new AxisTick(value, MapOriginalToNormalized(value, preparedMesh.MinOriginalY, preparedMesh.MaxOriginalY), FormatTickLabel(value)))
                .ToList();
            zTicks = CalculateTicks(preparedMesh.MinOriginalZ, preparedMesh.MaxOriginalZ)
                .Select(value => new AxisTick(value, MapOriginalToNormalized(value, preparedMesh.MinOriginalZ, preparedMesh.MaxOriginalZ), FormatTickLabel(value)))
                .ToList();
        }

        private static IReadOnlyList<double> CalculateTicks(double minValue, double maxValue, int targetTickCount = 6)
        {
            if (targetTickCount < 2)
                targetTickCount = 2;

            if (Math.Abs(maxValue - minValue) < 1e-12)
                return new[] { minValue };

            double min = Math.Min(minValue, maxValue);
            double max = Math.Max(minValue, maxValue);
            double range = NiceNumber(max - min, false);
            double step = NiceNumber(range / (targetTickCount - 1), true);

            double tickMin = Math.Ceiling(min / step) * step;
            double tickMax = Math.Floor(max / step) * step;
            var ticks = new List<double>();

            for (double value = tickMin; value <= tickMax + step * 0.5; value += step)
                ticks.Add(value);

            if (ticks.Count == 0)
            {
                ticks.Add(min);
                ticks.Add(max);
            }

            if (ticks[0] > min + 1e-10)
                ticks.Insert(0, min);

            if (ticks[ticks.Count - 1] < max - 1e-10)
                ticks.Add(max);

            return ticks;
        }

        private static double NiceNumber(double value, bool round)
        {
            if (value <= 0)
                return 1;

            double exponent = Math.Floor(Math.Log10(value));
            double fraction = value / Math.Pow(10, exponent);
            double niceFraction;

            if (round)
            {
                if (fraction < 1.5) niceFraction = 1;
                else if (fraction < 3) niceFraction = 2;
                else if (fraction < 7) niceFraction = 5;
                else niceFraction = 10;
            }
            else
            {
                if (fraction <= 1) niceFraction = 1;
                else if (fraction <= 2) niceFraction = 2;
                else if (fraction <= 5) niceFraction = 5;
                else niceFraction = 10;
            }

            return niceFraction * Math.Pow(10, exponent);
        }

        private static float MapOriginalToNormalized(double value, double minOriginal, double maxOriginal)
        {
            if (Math.Abs(maxOriginal - minOriginal) < 1e-12)
                return 0f;

            double normalized = (value - minOriginal) / (maxOriginal - minOriginal);
            normalized = Math.Max(0, Math.Min(1, normalized));
            return (float)normalized;
        }

        private static string FormatTickLabel(double value)
        {
            if (Math.Abs(value) < 1e-9)
                value = 0;

            double abs = Math.Abs(value);
            if (abs >= 1000)
                return value.ToString("F0", CultureInfo.InvariantCulture);

            if (abs >= 100)
                return value.ToString("F1", CultureInfo.InvariantCulture).TrimEnd('0').TrimEnd('.');

            return value.ToString("0.###", CultureInfo.InvariantCulture);
        }

        private static float Lerp(float a, float b, float t)
        {
            return a + (b - a) * t;
        }

        private bool TryProjectWorldToScreen(Vector3 world, out PointF screen)
        {
            screen = PointF.Empty;
            if (lastViewportWidth <= 0 || lastViewportHeight <= 0)
                return false;

            Vector4 clip = Vector4.Transform(new Vector4(world.X, world.Y, world.Z, 1f), lastModelViewMatrix);
            clip = Vector4.Transform(clip, lastProjectionMatrix);
            if (Math.Abs(clip.W) < 1e-6f)
                return false;

            Vector3 ndc = new Vector3(clip.X, clip.Y, clip.Z) / clip.W;
            if (ndc.Z < -1.1f || ndc.Z > 1.1f)
                return false;

            float x = (ndc.X * 0.5f + 0.5f) * lastViewportWidth;
            float y = (1f - (ndc.Y * 0.5f + 0.5f)) * lastViewportHeight;
            screen = new PointF(x, y);
            return true;
        }

        private Color MapColor(double zValue)
        {
            if (preparedMesh == null)
                return Color.FromArgb(40, 120, 215);

            return HeatmapColorMapper.GetHeatmapColor(zValue, preparedMesh.MinOriginalZ, preparedMesh.MaxOriginalZ);
        }

        private void GlControl_MouseDown(object sender, MouseEventArgs e)
        {
            lastMouseLocation = e.Location;
            isRotating = e.Button == MouseButtons.Left;
            isPanning = e.Button == MouseButtons.Right || e.Button == MouseButtons.Middle;
        }

        private void GlControl_MouseUp(object sender, MouseEventArgs e)
        {
            isRotating = false;
            isPanning = false;
        }

        private void GlControl_MouseMove(object sender, MouseEventArgs e)
        {
            int dx = e.X - lastMouseLocation.X;
            int dy = e.Y - lastMouseLocation.Y;

            if (isRotating)
            {
                yawDegrees += dx * 0.7f;
                pitchDegrees = Math.Max(-89f, Math.Min(89f, pitchDegrees + dy * 0.6f));
                glControl.Invalidate();
            }
            else if (isPanning)
            {
                panX += dx * 0.0045f;
                panY -= dy * 0.0045f;
                glControl.Invalidate();
            }

            lastMouseLocation = e.Location;
        }

        private void GlControl_MouseWheel(object sender, MouseEventArgs e)
        {
            zoomDistance -= e.Delta * 0.0015f;
            zoomDistance = Math.Max(1.6f, Math.Min(18f, zoomDistance));
            glControl.Invalidate();
        }

        private void GlControl_DoubleClick(object sender, EventArgs e)
        {
            ResetView();
            glControl.Invalidate();
        }

        private void ResetView()
        {
            yawDegrees = -42f;
            pitchDegrees = 28f;
            panX = 0f;
            panY = 0f;
            zoomDistance = preparedMesh == null ? 5.2f : preparedMesh.SuggestedZoomDistance;
        }

        private PreparedMesh PrepareMesh(SurfaceMesh mesh)
        {
            if (mesh == null || !mesh.HasData)
                return null;

            double xCenter = (mesh.MinX + mesh.MaxX) / 2.0;
            double yCenter = (mesh.MinY + mesh.MaxY) / 2.0;
            double zCenter = (mesh.MinZ + mesh.MaxZ) / 2.0;
            double horizontalSpan = Math.Max(mesh.XSpan, mesh.YSpan);
            double targetZSpan = horizontalSpan * 0.75;
            double zScale = targetZSpan / Math.Max(1e-9, mesh.ZSpan);
            zScale = Math.Max(0.25, Math.Min(6.0, zScale));
            double normalization = Math.Max(1e-6, horizontalSpan / 2.0);

            Dictionary<SurfaceVertex, PreparedVertex> vertexMap = mesh.Vertices.ToDictionary(
                vertex => vertex,
                vertex => new PreparedVertex(
                    (float)((vertex.X - xCenter) / normalization),
                    (float)((vertex.Y - yCenter) / normalization),
                    (float)(((vertex.Z - zCenter) * zScale) / normalization),
                    vertex.Z));

            PreparedMesh prepared = new PreparedMesh
            {
                MinOriginalZ = mesh.MinZ,
                MaxOriginalZ = mesh.MaxZ,
                MinOriginalX = mesh.MinX,
                MaxOriginalX = mesh.MaxX,
                MinOriginalY = mesh.MinY,
                MaxOriginalY = mesh.MaxY,
                SuggestedZoomDistance = mesh.Quads.Count > 0 ? 4.8f : 5.6f
            };

            prepared.Vertices.AddRange(vertexMap.Values);
            prepared.Quads.AddRange(mesh.Quads.Select(quad => new PreparedQuad(
                vertexMap[quad.V00],
                vertexMap[quad.V10],
                vertexMap[quad.V11],
                vertexMap[quad.V01])));
            prepared.BaseZ = prepared.Vertices.Count == 0 ? -1f : prepared.Vertices.Min(vertex => vertex.Z) - 0.12f;

            return prepared;
        }

        private sealed class PreparedMesh
        {
            public List<PreparedVertex> Vertices { get; } = new List<PreparedVertex>();
            public List<PreparedQuad> Quads { get; } = new List<PreparedQuad>();
            public float BaseZ { get; set; }
            public float MinX => Vertices.Count == 0 ? -1f : Vertices.Min(v => v.X);
            public float MaxX => Vertices.Count == 0 ? 1f : Vertices.Max(v => v.X);
            public float MinY => Vertices.Count == 0 ? -1f : Vertices.Min(v => v.Y);
            public float MaxY => Vertices.Count == 0 ? 1f : Vertices.Max(v => v.Y);
            public float MinZ => Vertices.Count == 0 ? 0f : Vertices.Min(v => v.Z);
            public float MaxZ => Vertices.Count == 0 ? 1f : Vertices.Max(v => v.Z);
            public float XSpan => Math.Max(1e-6f, MaxX - MinX);
            public float YSpan => Math.Max(1e-6f, MaxY - MinY);
            public float ZSpan => Math.Max(1e-6f, MaxZ - MinZ);
            public double MinOriginalX { get; set; }
            public double MaxOriginalX { get; set; }
            public double MinOriginalY { get; set; }
            public double MaxOriginalY { get; set; }
            public double MinOriginalZ { get; set; }
            public double MaxOriginalZ { get; set; }
            public double OriginalZSpan => Math.Max(1e-9, MaxOriginalZ - MinOriginalZ);
            public float SuggestedZoomDistance { get; set; }
        }

        private sealed class PreparedVertex
        {
            public PreparedVertex(float x, float y, float z, double originalZ)
            {
                X = x;
                Y = y;
                Z = z;
                OriginalZ = originalZ;
            }

            public float X { get; }
            public float Y { get; }
            public float Z { get; }
            public double OriginalZ { get; }
        }

        private sealed class PreparedQuad
        {
            public PreparedQuad(PreparedVertex v00, PreparedVertex v10, PreparedVertex v11, PreparedVertex v01)
            {
                V00 = v00;
                V10 = v10;
                V11 = v11;
                V01 = v01;
            }

            public PreparedVertex V00 { get; }
            public PreparedVertex V10 { get; }
            public PreparedVertex V11 { get; }
            public PreparedVertex V01 { get; }
        }

        private sealed class AxisTick
        {
            public AxisTick(double value, float position, string label)
            {
                Value = value;
                Position = position;
                Label = label;
            }

            public double Value { get; }
            public float Position { get; }
            public string Label { get; }
        }

        private sealed class OverlayPlacement
        {
            public string Text { get; set; } = string.Empty;
            public PointF Screen { get; set; }
            public Color Color { get; set; }
            public bool Bold { get; set; }
            public bool Visible { get; set; }
        }

        private sealed class AxisOverlayControl : System.Windows.Forms.Control
        {
            private const int WmNCHitTest = 0x84;
            private const int HtTransparent = -1;
            private IReadOnlyList<OverlayPlacement> placements = Array.Empty<OverlayPlacement>();

            public AxisOverlayControl()
            {
                SetStyle(ControlStyles.AllPaintingInWmPaint |
                    ControlStyles.OptimizedDoubleBuffer |
                    ControlStyles.ResizeRedraw |
                    ControlStyles.UserPaint |
                    ControlStyles.SupportsTransparentBackColor, true);
                TabStop = false;
            }

            public void SetPlacements(IReadOnlyList<OverlayPlacement> newPlacements)
            {
                placements = newPlacements ?? Array.Empty<OverlayPlacement>();
                Invalidate();
            }

            protected override void OnPaint(PaintEventArgs e)
            {
                base.OnPaint(e);

                e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
                e.Graphics.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit;

                foreach (OverlayPlacement placement in placements)
                {
                    if (!placement.Visible)
                    {
                        continue;
                    }

                    using (Font font = placement.Bold
                        ? new Font("Microsoft YaHei UI", 8.8f, FontStyle.Bold, GraphicsUnit.Point)
                        : new Font("Microsoft YaHei UI", 8.0f, FontStyle.Regular, GraphicsUnit.Point))
                    {
                        Size textSize = TextRenderer.MeasureText(
                            e.Graphics,
                            placement.Text,
                            font,
                            new Size(int.MaxValue, int.MaxValue),
                            TextFormatFlags.NoPadding);

                        Rectangle rect = new Rectangle(
                            (int)Math.Round(placement.Screen.X - textSize.Width / 2f),
                            (int)Math.Round(placement.Screen.Y - textSize.Height / 2f),
                            textSize.Width + 6,
                            textSize.Height + 4);

                        rect.X = Math.Max(0, Math.Min(Math.Max(0, ClientSize.Width - rect.Width), rect.X));
                        rect.Y = Math.Max(0, Math.Min(Math.Max(0, ClientSize.Height - rect.Height), rect.Y));

                        using (GraphicsPath path = CreateRoundedRectangle(rect, 4))
                        using (SolidBrush backgroundBrush = new SolidBrush(placement.Bold
                            ? Color.FromArgb(244, 255, 255, 255)
                            : Color.FromArgb(232, 255, 255, 255)))
                        using (Pen borderPen = new Pen(placement.Bold
                            ? Color.FromArgb(180, 180, 188, 198)
                            : Color.FromArgb(140, 200, 208, 216)))
                        using (SolidBrush textBrush = new SolidBrush(placement.Color))
                        {
                            e.Graphics.FillPath(backgroundBrush, path);
                            e.Graphics.DrawPath(borderPen, path);
                            e.Graphics.DrawString(placement.Text, font, textBrush, rect.Left + 3, rect.Top + 1);
                        }
                    }
                }
            }

            private static GraphicsPath CreateRoundedRectangle(Rectangle rect, int radius)
            {
                int diameter = radius * 2;
                GraphicsPath path = new GraphicsPath();

                if (rect.Width <= diameter || rect.Height <= diameter)
                {
                    path.AddRectangle(rect);
                    return path;
                }

                path.AddArc(rect.Left, rect.Top, diameter, diameter, 180, 90);
                path.AddArc(rect.Right - diameter, rect.Top, diameter, diameter, 270, 90);
                path.AddArc(rect.Right - diameter, rect.Bottom - diameter, diameter, diameter, 0, 90);
                path.AddArc(rect.Left, rect.Bottom - diameter, diameter, diameter, 90, 90);
                path.CloseFigure();
                return path;
            }

            protected override void WndProc(ref Message m)
            {
                if (m.Msg == WmNCHitTest)
                {
                    m.Result = (IntPtr)HtTransparent;
                    return;
                }

                base.WndProc(ref m);
            }
        }

        private sealed class SurfaceColorBarControl : System.Windows.Forms.Control
        {
            private string axisTitle = "B";
            private double minValue;
            private double maxValue = 1;

            public SurfaceColorBarControl()
            {
                SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw | ControlStyles.UserPaint, true);
                BackColor = Color.FromArgb(246, 248, 250);
                MinimumSize = new Size(120, 160);
            }

            public void SetScale(string title, double min, double max)
            {
                axisTitle = string.IsNullOrWhiteSpace(title) ? "B" : title;
                minValue = min;
                maxValue = max;
                Invalidate();
            }

            protected override void OnPaint(PaintEventArgs e)
            {
                base.OnPaint(e);

                Graphics g = e.Graphics;
                g.Clear(BackColor);
                using (var borderPen = new Pen(Color.FromArgb(210, 217, 224)))
                {
                    g.DrawRectangle(borderPen, 0, 0, Width - 1, Height - 1);
                }

                int barX = 24;
                int barWidth = 26;
                int barY = 46;
                int barHeight = Math.Max(140, Height - 120);

                using (var titleFont = new Font("Microsoft YaHei UI", 8.8f, FontStyle.Bold))
                using (var tickFont = new Font("Microsoft YaHei UI", 8.2f, FontStyle.Regular))
                {
                    g.DrawString("Color Bar", titleFont, Brushes.Black, 16, 10);
                    g.DrawString(axisTitle, titleFont, Brushes.Black, 16, 28);

                    for (int i = 0; i < barHeight; i++)
                    {
                        double normalized = 1.0 - i / (double)Math.Max(1, barHeight - 1);
                    using (var pen = new Pen(HeatmapColorMapper.GetColorFromNormalized(normalized)))
                        {
                            g.DrawLine(pen, barX, barY + i, barX + barWidth, barY + i);
                        }
                    }

                    g.DrawRectangle(Pens.Black, new Rectangle(barX, barY, barWidth, barHeight));

                    foreach (double tick in CalculateTicks(minValue, maxValue, 5))
                    {
                        float t = MapOriginalToNormalized(tick, minValue, maxValue);
                        float y = barY + (1f - t) * (barHeight - 1);
                        g.DrawLine(Pens.Black, barX + barWidth + 4, y, barX + barWidth + 12, y);
                        g.DrawString(FormatTickLabel(tick), tickFont, Brushes.Black, barX + barWidth + 16, y - 7);
                    }

                    g.DrawString($"Max: {FormatTickLabel(maxValue)}", tickFont, Brushes.Black, 16, barY + barHeight + 10);
                    g.DrawString($"Min: {FormatTickLabel(minValue)}", tickFont, Brushes.Black, 16, barY + barHeight + 28);
                }
            }
        }
    }
}
