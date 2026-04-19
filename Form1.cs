using grbloxy.Properties;
using System;
using System.CodeDom.Compiler;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.IO.Ports;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Globalization;
using ScottPlot.WinForms;



namespace grbloxy
{
    public partial class Form1 : Form
    {
        private bool enableRealTimeDrawing = true;

        private sealed class ZLayerViewItem
        {
            public int LayerIndex { get; set; }
            public double ZValue { get; set; }

            public override string ToString()
            {
                return $"第 {LayerIndex + 1} 层  Z={ZValue:F2} mm";
            }
        }

        private sealed class LayerSurfaceCache
        {
            public int LayerIndex { get; set; }
            public double ZValue { get; set; }
            public List<Control.PointData> Points { get; set; }
        }

        private enum CurveLayerDisplayMode
        {
            SingleLayer,
            AllLayers,
            AverageOnly,
            AllLayersWithAverage
        }

        private sealed class CurveDisplayModeItem
        {
            public CurveLayerDisplayMode Mode { get; set; }
            public string Text { get; set; }

            public override string ToString()
            {
                return Text;
            }
        }

        private enum PlotViewMode
        {
            RotatableSingleLayerHeatmap3D,
            StaticSingleLayerHeatmap,
            RotatableAverageAlongZHeatmap3D,
            StaticAverageAlongZHeatmap,
            XVoltageCurve,
            YVoltageCurve,
            ZVoltageCurve
        }

        private sealed class PlotModeItem
        {
            public PlotViewMode Mode { get; set; }
            public string Text { get; set; }

            public override string ToString()
            {
                return Text;
            }
        }

        private readonly SortedDictionary<double, LayerSurfaceCache> zLayerCache = new SortedDictionary<double, LayerSurfaceCache>();
        private PlotCurveCollectionData xCurveCollection = new PlotCurveCollectionData();
        private PlotCurveCollectionData yCurveCollection = new PlotCurveCollectionData();
        private PlotCurveData zCurveData = new PlotCurveData();
        private OpenTkSurfaceControl surfaceViewer;
        private FormsPlot formsPlot;
        private HeatmapViewControl heatmapViewer;
        private bool isUpdatingZLayerSelector = false;
        private bool isUpdatingPlotModeSelector = false;
        private bool isUpdatingCurveModeSelector = false;
        private bool isUpdatingCurveLayerSelector = false;
        private double? selectedZLayer = null;
        private double? selectedCurveZLayer = null;
        private Task<bool> stopAndReturnTask;
        private bool allowCloseAfterSafeStop;
        private bool isCloseSequenceRunning;
        private DateTime lastRealtimeCurveRefreshUtc = DateTime.MinValue;
        private const int RealtimeCurveRefreshIntervalMs = 60;
        private int latestRealtimeCurvePointCount = 0;
        private void Form1_Load(object sender, EventArgs e)
        {

            // 初始化编码下拉框选项
            comboBox4.Items.Add("ASCII");
            comboBox4.Items.Add("UTF-8");
            comboBox4.Items.Add("GB2312");
            comboBox4.Items.Add("GBK");
            comboBox4.Items.Add("Default");
            comboBox4.SelectedIndex = 0; // 设置默认选项

            // 显示模式选择（新增）
            comboBox5.Items.Add("文本显示");
            comboBox5.Items.Add("十六进制显示");
            comboBox5.Items.Add("二进制显示");
            comboBox5.SelectedIndex = 0;

            // 默认波特率：串口1（控制）115200，串口2（采集/电压）115200
            try
            {
                if (comboBox2.Items.Count > 4) comboBox2.SelectedIndex = 4; else comboBox2.Text = "115200";
                comboBox8.Text = "115200";

                // 给串口控件设置相同的默认值，打开前用户仍可更改
                serialPort1.BaudRate = 115200;
                serialPort2.BaudRate = 115200;
            }
            catch { }

            comboBox4.Enabled = false;
            comboBox5.Enabled = false;
            InitializePlotModeSelector();
            InitializeCurveDisplayModeSelector();
            if (string.IsNullOrWhiteSpace(textBox7.Text)) textBox7.Text = "0";
            if (string.IsNullOrWhiteSpace(textBox14.Text)) textBox14.Text = "10";
            if (string.IsNullOrWhiteSpace(textBox8.Text)) textBox8.Text = "0";
            if (string.IsNullOrWhiteSpace(textBox15.Text)) textBox15.Text = "10";
            if (string.IsNullOrWhiteSpace(textBox13.Text)) textBox13.Text = "0";
            if (string.IsNullOrWhiteSpace(textBox16.Text)) textBox16.Text = "10";
            if (string.IsNullOrWhiteSpace(textBox9.Text)) textBox9.Text = "1";
            if (string.IsNullOrWhiteSpace(textBox10.Text)) textBox10.Text = "1";
            if (string.IsNullOrWhiteSpace(textBox11.Text)) textBox11.Text = "1";
            if (string.IsNullOrWhiteSpace(textBox17.Text)) textBox17.Text = Control.DefaultSampleCount.ToString(CultureInfo.InvariantCulture);
            ResetLayerView();
            UpdatePlotModeUi();
        }
        private void InitialSerialPorts(System.IO.Ports.SerialPort port1, System.IO.Ports.SerialPort port2)
        {
            // 控制串口保留异步接收，采集串口改为同步查询，避免响应被接收事件提前读走
            port1.DataReceived += Port_DataReceived;
        }

        public Control ctrl;
        public Form1()
        {
            InitializeComponent();
            this.Load += Form1_Load;
            this.FormClosing += Form1_FormClosing;
            CheckForIllegalCrossThreadCalls = false;
            //串口初始化,绑定串口数据接收事件
            InitialSerialPorts(serialPort1, serialPort2);
            //串口1配置（根据你的实际串口号/波特率修改）
            ctrl = new Control(); // 或 new Control()，根据你的构造函数
            ctrl.LogMessage = AppendSerialLog;
            EnableDoubleBuffer(panel1);

            ctrl.PointScanned += HandlePointScanned;
            ctrl.ZLayerCompleted += HandleZLayerCompleted;
            ctrl.ScanStopped += HandleScanStopped;
            InitializePlotHosts();
        }

        private void HandleScanStopped()
        {
            if (!IsHandleCreated || IsDisposed)
            {
                return;
            }

            try
            {
                BeginInvoke(new Action(() =>
                {
                    button6.Enabled = true;
                    button9.Enabled = true;
                }));
            }
            catch
            {
            }
        }

        private void HandlePointScanned(int completedPoints, Control.PointData latestPoint)
        {
            if (!IsHandleCreated || IsDisposed)
            {
                return;
            }

            if (InvokeRequired)
            {
                try
                {
                    BeginInvoke(new Action(() => HandlePointScanned(completedPoints, latestPoint)));
                }
                catch
                {
                }

                return;
            }

            if (!enableRealTimeDrawing || !IsCurvePlotMode(GetCurrentPlotMode()))
            {
                return;
            }

            bool shouldRefresh = completedPoints <= 3 ||
                completedPoints >= ctrl.TotalPoints ||
                (DateTime.UtcNow - lastRealtimeCurveRefreshUtc).TotalMilliseconds >= RealtimeCurveRefreshIntervalMs;

            if (!shouldRefresh)
            {
                return;
            }

            latestRealtimeCurvePointCount = Math.Max(latestRealtimeCurvePointCount, completedPoints);
            RefreshRealtimeCurvePlot(completedPoints);
        }

        private void InitializePlotModeSelector()
        {
            isUpdatingPlotModeSelector = true;
            try
            {
                comboBox11.Items.Clear();
                comboBox11.Items.Add(new PlotModeItem { Mode = PlotViewMode.RotatableSingleLayerHeatmap3D, Text = "可旋转三维热力图（单个 Z 层）" });
                comboBox11.Items.Add(new PlotModeItem { Mode = PlotViewMode.StaticSingleLayerHeatmap, Text = "原始风格静态热力图（单个 Z 层）" });
                comboBox11.Items.Add(new PlotModeItem { Mode = PlotViewMode.RotatableAverageAlongZHeatmap3D, Text = "可旋转三维热力图（沿 Z 平均）" });
                comboBox11.Items.Add(new PlotModeItem { Mode = PlotViewMode.StaticAverageAlongZHeatmap, Text = "静态热力图（沿 Z 平均）" });
                comboBox11.Items.Add(new PlotModeItem { Mode = PlotViewMode.XVoltageCurve, Text = "B - X 曲线图" });
                comboBox11.Items.Add(new PlotModeItem { Mode = PlotViewMode.YVoltageCurve, Text = "B - Y 曲线图" });
                comboBox11.Items.Add(new PlotModeItem { Mode = PlotViewMode.ZVoltageCurve, Text = "B - Z 曲线图" });
                comboBox11.SelectedIndex = 0;
            }
            finally
            {
                isUpdatingPlotModeSelector = false;
            }
        }

        private void InitializeCurveDisplayModeSelector()
        {
            isUpdatingCurveModeSelector = true;
            try
            {
                comboBox12.Items.Clear();
                comboBox12.Items.Add(new CurveDisplayModeItem { Mode = CurveLayerDisplayMode.SingleLayer, Text = "单层" });
                comboBox12.Items.Add(new CurveDisplayModeItem { Mode = CurveLayerDisplayMode.AllLayers, Text = "全部层" });
                comboBox12.Items.Add(new CurveDisplayModeItem { Mode = CurveLayerDisplayMode.AverageOnly, Text = "平均值" });
                comboBox12.Items.Add(new CurveDisplayModeItem { Mode = CurveLayerDisplayMode.AllLayersWithAverage, Text = "全部层 + 平均值" });
                comboBox12.SelectedIndex = 3;
            }
            finally
            {
                isUpdatingCurveModeSelector = false;
            }
        }

        private PlotViewMode GetCurrentPlotMode()
        {
            if (comboBox11.SelectedItem is PlotModeItem item)
            {
                return item.Mode;
            }

            return PlotViewMode.RotatableSingleLayerHeatmap3D;
        }

        private CurveLayerDisplayMode GetCurrentCurveDisplayMode()
        {
            if (comboBox12.SelectedItem is CurveDisplayModeItem item)
            {
                return item.Mode;
            }

            return CurveLayerDisplayMode.AllLayersWithAverage;
        }

        private void UpdatePlotModeUi()
        {
            PlotViewMode mode = GetCurrentPlotMode();
            bool showSurfaceZSelector = IsSingleLayerMode(mode);
            bool is3DMode = Is3DPlotMode(mode);
            bool isStaticHeatmapMode = IsStaticHeatmapMode(mode);
            bool isCurveMode = IsCurvePlotMode(mode);
            bool showCurveModeSelector = SupportsLayeredCurveMode(mode);
            bool showCurveLayerSelector = showCurveModeSelector && GetCurrentCurveDisplayMode() == CurveLayerDisplayMode.SingleLayer;

            label19.Visible = showSurfaceZSelector;
            comboBox10.Visible = showSurfaceZSelector;
            comboBox10.Enabled = showSurfaceZSelector && comboBox10.Items.Count > 0;
            label25.Visible = showCurveModeSelector;
            comboBox12.Visible = showCurveModeSelector;
            comboBox12.Enabled = showCurveModeSelector;
            label26.Visible = showCurveLayerSelector;
            comboBox13.Visible = showCurveLayerSelector;
            comboBox13.Enabled = showCurveLayerSelector && comboBox13.Items.Count > 0;
            panel1.Visible = is3DMode;
            panel2DHost.Visible = !is3DMode;
            if (formsPlot != null)
            {
                formsPlot.Visible = isCurveMode;
            }

            if (heatmapViewer != null)
            {
                heatmapViewer.Visible = isStaticHeatmapMode;
            }

            button5.Text = "刷新图像";

            switch (mode)
            {
                case PlotViewMode.RotatableSingleLayerHeatmap3D:
                    groupBox7.Text = "图像视图（OpenTK 可旋转三维热力图）";
                    break;
                case PlotViewMode.StaticSingleLayerHeatmap:
                    groupBox7.Text = "图像视图（原始风格静态热力图）";
                    break;
                case PlotViewMode.RotatableAverageAlongZHeatmap3D:
                    groupBox7.Text = "图像视图（OpenTK 沿 Z 平均三维热力图）";
                    break;
                case PlotViewMode.StaticAverageAlongZHeatmap:
                    groupBox7.Text = "图像视图（沿 Z 平均静态热力图）";
                    break;
                case PlotViewMode.XVoltageCurve:
                    groupBox7.Text = "图像视图（ScottPlot B - X）";
                    break;
                case PlotViewMode.YVoltageCurve:
                    groupBox7.Text = "图像视图（ScottPlot B - Y）";
                    break;
                case PlotViewMode.ZVoltageCurve:
                    groupBox7.Text = "图像视图（ScottPlot B - Z）";
                    break;
            }
        }

        private bool Is3DPlotMode(PlotViewMode mode)
        {
            return mode == PlotViewMode.RotatableSingleLayerHeatmap3D ||
                mode == PlotViewMode.RotatableAverageAlongZHeatmap3D;
        }

        private bool IsStaticHeatmapMode(PlotViewMode mode)
        {
            return mode == PlotViewMode.StaticSingleLayerHeatmap ||
                mode == PlotViewMode.StaticAverageAlongZHeatmap;
        }

        private bool IsCurvePlotMode(PlotViewMode mode)
        {
            return mode == PlotViewMode.XVoltageCurve ||
                mode == PlotViewMode.YVoltageCurve ||
                mode == PlotViewMode.ZVoltageCurve;
        }

        private bool SupportsLayeredCurveMode(PlotViewMode mode)
        {
            return mode == PlotViewMode.XVoltageCurve ||
                mode == PlotViewMode.YVoltageCurve;
        }

        private bool IsSingleLayerMode(PlotViewMode mode)
        {
            return mode == PlotViewMode.RotatableSingleLayerHeatmap3D ||
                mode == PlotViewMode.StaticSingleLayerHeatmap;
        }

        private void InitializePlotHosts()
        {
            if (surfaceViewer == null)
            {
                surfaceViewer = new OpenTkSurfaceControl
                {
                    Dock = DockStyle.Fill
                };
                panel1.Controls.Clear();
                panel1.Controls.Add(surfaceViewer);
            }

            if (formsPlot == null)
            {
                formsPlot = new FormsPlot
                {
                    Dock = DockStyle.Fill,
                    Visible = false
                };
                panel2DHost.Controls.Add(formsPlot);
            }

            if (heatmapViewer == null)
            {
                heatmapViewer = new HeatmapViewControl
                {
                    Dock = DockStyle.Fill,
                    Visible = false
                };
                panel2DHost.Controls.Add(heatmapViewer);
            }
        }

        private bool TryParseFlexibleDouble(string input, out double value)
        {
            string normalized = (input ?? string.Empty).Trim();
            return double.TryParse(normalized, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.CurrentCulture, out value) ||
                double.TryParse(normalized, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out value);
        }

        private double ParseRequiredDouble(TextBox textBox, string fieldName)
        {
            if (!TryParseFlexibleDouble(textBox.Text, out double value))
            {
                throw new FormatException($"{fieldName} 输入无效，请输入数字。");
            }

            return value;
        }

        private int ParseRequiredInt(TextBox textBox, string fieldName)
        {
            string normalized = (textBox.Text ?? string.Empty).Trim();
            if (!int.TryParse(normalized, NumberStyles.Integer, CultureInfo.CurrentCulture, out int value) &&
                !int.TryParse(normalized, NumberStyles.Integer, CultureInfo.InvariantCulture, out value))
            {
                throw new FormatException($"{fieldName} 输入无效，请输入整数。");
            }

            return value;
        }

        private void ValidateRange(string axisName, double minValue, double maxValue, double step)
        {
            if (step <= 0)
            {
                throw new InvalidOperationException($"{axisName} 轴分辨率必须大于 0。");
            }
        }
        
    
        //扫描串口按钮，按下进行扫描获取端口数据
        private void button1_Click(object sender, EventArgs e)
        {
            SearchAnAddSerialToComboBox(serialPort1, comboBox1);
            SearchAnAddSerialToComboBox(serialPort2, comboBox6);
        }

        private void SearchAnAddSerialToComboBox(SerialPort MyPort, ComboBox MyBox)
        {
            // 更快且不会阻塞UI的枚举串口方式：使用 SerialPort.GetPortNames()
            // 这个实现不会尝试打开每个端口，从而避免长时间阻塞或引发异常开销
            MyBox.Items.Clear();
            try
            {
                // 获取可用串口并按名称排序
                string[] ports = SerialPort.GetPortNames();
                Array.Sort(ports, StringComparer.OrdinalIgnoreCase);

                // 使用主线程更新UI
                if (this.InvokeRequired)
                {
                    this.Invoke(new Action(() =>
                    {
                        MyBox.Items.AddRange(ports);
                        if (ports.Length > 0) MyBox.Text = ports[0];
                    }));
                }
                else
                {
                    MyBox.Items.AddRange(ports);
                    if (ports.Length > 0) MyBox.Text = ports[0];
                }
            }
            catch (Exception ex)
            {
                // 遇到异常时不抛出，记录到日志框（如果有）并提示短时失败
                System.Diagnostics.Debug.WriteLine("枚举串口失败：" + ex.Message);
            }
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
        private void button2_Click(object sender, EventArgs e)
        {
            // 打开串口逻辑
            if (button2.Text == "打开串口")
            {
                try
                {
                    // 先检查两个串口是否都处于关闭状态，避免重复打开
                    if (serialPort1.IsOpen || serialPort2.IsOpen)
                    {
                        MessageBox.Show("已有串口处于打开状态，请先关闭");
                        return;
                    }

                    // 配置并打开串口1
                    serialPort1.PortName = comboBox1.Text;
                    serialPort1.BaudRate = Convert.ToInt32(comboBox2.Text);
                    ConfigureControlPort(serialPort1);
                    serialPort1.Open();

                    // 配置并打开串口2
                    serialPort2.PortName = comboBox6.Text;
                    serialPort2.BaudRate = Convert.ToInt32(comboBox8.Text);
                    ConfigureInstrumentPort(serialPort2);
                    serialPort2.Open();

                    if (!TryIdentifyMeasurementInstrument(serialPort2, out string instrumentId))
                    {
                        throw new InvalidOperationException("采集串口未识别到 DM3058/兼容 SCPI 仪器，请确认串口号和波特率。");
                    }

                    // 修正：正确显示两个串口信息，避免覆盖
                    comboBox9.Items.Clear(); // 清空原有项，防止重复添加
                    comboBox9.Items.Add(serialPort1.PortName);
                    comboBox9.Items.Add(serialPort2.PortName);
                    comboBox9.SelectedIndex = 0;
                    AppendSerialLog($"[系统] 采集设备已连接: {instrumentId}");

                    // 更新按钮状态和控件权限
                    button2.Text = "关闭串口";
                    comboBox1.Enabled = false;  // 打开后禁用串口配置
                    comboBox2.Enabled = false;
                    comboBox6.Enabled = false;  // 补充：串口2的配置也需要禁用
                    comboBox8.Enabled = false;
                    comboBox3.Enabled = true;
                    comboBox4.Enabled = true;
                    comboBox5.Enabled = true;
                    button4.Enabled = true;
                    button1.Enabled = false;
                }
                catch (Exception ex)
                {
                    // 关键修复：如果打开失败，立即关闭已打开的串口（防止部分打开）
                    if (serialPort1.IsOpen) serialPort1.Close();
                    if (serialPort2.IsOpen) serialPort2.Close();
                    MessageBox.Show($"串口打开失败：{ex.Message}"); // 显示具体错误，便于排查
                }
            }
            // 关闭串口逻辑
            else if (button2.Text == "关闭串口")
            {
                try
                {
                    // 关键修复：同时关闭两个串口，避免serialPort2残留打开状态
                    if (serialPort1.IsOpen) serialPort1.Close();
                    if (serialPort2.IsOpen) serialPort2.Close();

                    // 更新按钮和控件状态
                    button2.Text = "打开串口";
                    comboBox1.Enabled = true;
                    comboBox2.Enabled = true;
                    comboBox6.Enabled = true;  // 补充：恢复串口2配置权限
                    comboBox8.Enabled = true;
                    comboBox3.Enabled = true;
                    comboBox4.Enabled = false;
                    comboBox5.Enabled = false;
                    button4.Enabled = false;
                    button1.Enabled = true;

                    // 清空串口显示框
                    comboBox9.Items.Clear();
                    comboBox9.Text = "";
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"串口关闭失败：{ex.Message}");
                }
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            textBox1.Clear();
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void comboBox4_SelectedIndexChanged(object sender, EventArgs e)
        {
            switch (comboBox4.Text)
            {
                case "ASCII":
                    serialPort1.Encoding = Encoding.ASCII;
                    break;
                case "UTF-8":
                    serialPort1.Encoding = Encoding.UTF8;
                    break;
                case "GB2312":
                    serialPort1.Encoding = Encoding.GetEncoding("GB2312");
                    break;
                case "GBK":
                    serialPort1.Encoding = Encoding.GetEncoding("GBK");
                    break;
                case "Default":
                    serialPort1.Encoding = Encoding.Default;
                    break;
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            if (ctrl != null && ctrl.IsScanRunning && comboBox9.Text == serialPort2.PortName)
            {
                MessageBox.Show("扫描采集中暂时不能手动操作采集串口");
                return;
            }

            if(comboBox9.Text == serialPort1.PortName)
            {
                choseTosendText(serialPort1, textBox2.Text);
            }
            else if(comboBox9.Text == serialPort2.PortName)
            {
                choseTosendText(serialPort2, textBox2.Text);
            }
            else
            {
                MessageBox.Show("请选择要发送的串口");
                return;
            }
        }
        private void choseTosendText(SerialPort resCOM, string str)
        {
            if (!resCOM.IsOpen)
            {
                MessageBox.Show("请先打开串口");
                return;
            }
            if (string.IsNullOrEmpty(str))
            {

                MessageBox.Show("发送内容不能为空");
                return;
            }
            try
            {
                if (ReferenceEquals(resCOM, serialPort2))
                {
                    string response = SendInstrumentCommand(resCOM, str);
                    AppendSerialLog("<-" + str);
                    if (!string.IsNullOrWhiteSpace(response))
                    {
                        AppendSerialLog($"[{resCOM.PortName}]: {response}");
                    }
                    return;
                }

                byte[] data = resCOM.Encoding.GetBytes(str + "\r\n");
                resCOM.Write(data, 0, data.Length);
                AppendSerialLog("<-" + str);
            }
            catch (Exception ex)
            {
                MessageBox.Show("发送失败: " + ex.Message);
            }
        }

        private void ConfigureInstrumentPort(SerialPort port)
        {
            port.DataBits = 8;
            port.Parity = Parity.None;
            port.StopBits = StopBits.One;
            port.Handshake = Handshake.None;
            port.ReadTimeout = 1000;
            port.WriteTimeout = 1000;
            port.NewLine = "\n";
            port.Encoding = Encoding.ASCII;
        }

        private void ConfigureControlPort(SerialPort port)
        {
            port.DataBits = 8;
            port.Parity = Parity.None;
            port.StopBits = StopBits.One;
            port.Handshake = Handshake.None;
            port.ReadTimeout = 1000;
            port.WriteTimeout = 1000;
            port.NewLine = "\n";
            port.Encoding = Encoding.ASCII;
        }

        private bool TryIdentifyMeasurementInstrument(SerialPort port, out string instrumentId)
        {
            instrumentId = string.Empty;

            try
            {
                instrumentId = SendInstrumentCommand(port, "*IDN?");
                return !string.IsNullOrWhiteSpace(instrumentId) &&
                    (instrumentId.IndexOf("RIGOL", StringComparison.OrdinalIgnoreCase) >= 0 ||
                     instrumentId.IndexOf("DM3058", StringComparison.OrdinalIgnoreCase) >= 0);
            }
            catch
            {
                instrumentId = string.Empty;
                return false;
            }
        }

        private string SendInstrumentCommand(SerialPort port, string command)
        {
            if (port == null || !port.IsOpen)
            {
                throw new InvalidOperationException("采集串口未打开");
            }

            string normalized = (command ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(normalized))
            {
                throw new InvalidOperationException("发送内容不能为空");
            }

            port.DiscardInBuffer();
            port.Write(normalized + "\r\n");

            if (!normalized.EndsWith("?", StringComparison.Ordinal))
            {
                return string.Empty;
            }

            string response = port.ReadLine();
            return response?.Trim() ?? string.Empty;
        }

        private void AppendSerialLog(string message)
        {
            if (string.IsNullOrWhiteSpace(message))
            {
                return;
            }

            string line = message + "\r\n";
            if (textBox1.InvokeRequired)
            {
                textBox1.Invoke(new Action(() =>
                {
                    textBox1.AppendText(line);
                    textBox1.ScrollToCaret();
                }));
            }
            else
            {
                textBox1.AppendText(line);
                textBox1.ScrollToCaret();
            }
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox7_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox8_TextChanged(object sender, EventArgs e)
        {

        }

        private void button9_Click(object sender, EventArgs e)
        {
            if (!serialPort1.IsOpen)
            {
                MessageBox.Show("请先通过串口连接设备"); return;
            }
            else if (serialPort1.IsOpen)
            {
                double minX = ParseRequiredDouble(textBox7, "X 起点");
                double maxX = ParseRequiredDouble(textBox14, "X 终点");
                double minY = ParseRequiredDouble(textBox8, "Y 起点");
                double maxY = ParseRequiredDouble(textBox15, "Y 终点");
                double minZ = ParseRequiredDouble(textBox13, "Z 起点");
                double maxZ = ParseRequiredDouble(textBox16, "Z 终点");
                double stepX = ParseRequiredDouble(textBox9, "X 分辨率");
                double stepY = ParseRequiredDouble(textBox10, "Y 分辨率");
                double stepZ = ParseRequiredDouble(textBox11, "Z 分辨率");
                double textBoxSetSpeed = ParseRequiredDouble(textBox12, "速度");
                int sampleCount = ParseRequiredInt(textBox17, "单点采样次数");

                ValidateRange("X", minX, maxX, stepX);
                ValidateRange("Y", minY, maxY, stepY);
                ValidateRange("Z", minZ, maxZ, stepZ);
                if (sampleCount <= 0)
                {
                    throw new InvalidOperationException("单点采样次数必须大于 0。");
                }

                if (sampleCount > Control.MaxSampleCount)
                {
                    throw new InvalidOperationException($"单点采样次数不能超过 {Control.MaxSampleCount}。");
                }

                int estimatedPoints = ctrl.EstimateTotalPoints(minX, maxX, minY, maxY, minZ, maxZ, stepX, stepY, stepZ);
                if (estimatedPoints <= 1)
                {
                    MessageBox.Show(
                        $"当前参数只会采集 {estimatedPoints} 个点。\r\n" +
                        $"X范围={minX} ~ {maxX} mm，Y范围={minY} ~ {maxY} mm，Z范围={minZ} ~ {maxZ} mm。\r\n" +
                        $"X分辨率={stepX} mm，Y分辨率={stepY} mm，Z分辨率={stepZ} mm。\r\n" +
                        $"请增大扫描范围，或减小分辨率后再开始。",
                        "点数过少",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                    return;
                }

                AppendSerialLog(
                    $"[系统] 本次预计采集 {estimatedPoints} 个点 | " +
                    $"X={minX:F4}~{maxX:F4} mm, Y={minY:F4}~{maxY:F4} mm, Z={minZ:F4}~{maxZ:F4} mm | " +
                    $"步长=({stepX:F4}, {stepY:F4}, {stepZ:F4}) mm | 单点采样={sampleCount} 次");
                ResetLayerView();
                ctrl.GrblControl(serialPort1, serialPort2, minX, maxX, minY, maxY, minZ, maxZ, stepX, stepY, stepZ, textBoxSetSpeed, sampleCount);
           
            }
        }
        
        private void groupBox3_Enter(object sender, EventArgs e)
        {

        }

        private void groupBox4_Enter(object sender, EventArgs e)
        {

        }

        private void label11_Click(object sender, EventArgs e)
        {

        }

        private void textBox12_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox11_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox10_TextChanged(object sender, EventArgs e)
        {

        }

        private void groupBox6_Enter(object sender, EventArgs e)
        {

        }

        private void comboBox10_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (isUpdatingZLayerSelector)
            {
                return;
            }

            if (comboBox10.SelectedItem is ZLayerViewItem selectedLayer)
            {
                selectedZLayer = NormalizeLayerZ(selectedLayer.ZValue);
            }
            else
            {
                selectedZLayer = null;
            }

            DrawSelectedLayer();
        }

        private void comboBox11_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (isUpdatingPlotModeSelector)
            {
                return;
            }

            UpdatePlotModeUi();
            DrawSelectedLayer();
        }

        private void comboBox12_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (isUpdatingCurveModeSelector)
            {
                return;
            }

            UpdatePlotModeUi();
            RefreshPlotDisplay();
        }

        private void comboBox13_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (isUpdatingCurveLayerSelector)
            {
                return;
            }

            if (comboBox13.SelectedItem is ZLayerViewItem selectedLayer)
            {
                selectedCurveZLayer = NormalizeLayerZ(selectedLayer.ZValue);
            }
            else
            {
                selectedCurveZLayer = null;
            }

            RefreshPlotDisplay();
        }

        private void groupBox7_Enter(object sender, EventArgs e)
        {

        }

        private void comboBox9_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void comboBox7_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void comboBox9_SelectedIndexChanged_1(object sender, EventArgs e)
        {

        }

        private void button6_Click(object sender, EventArgs e)
        {
            if (button2.Text == "打开串口")
            {

                try
                {
                    serialPort2.PortName = comboBox1.Text;
                    serialPort2.BaudRate = Convert.ToInt32(comboBox2.Text);
                    serialPort2.Open();
                    button2.Text = "关闭串口";


                    comboBox1.Enabled = false;
                    comboBox2.Enabled = false;
                    comboBox3.Enabled = false;
                    comboBox4.Enabled = true;
                    comboBox5.Enabled = true;
                    button1.Enabled = false;
                    button4.Enabled = true;

                }
                catch
                {
                    MessageBox.Show("串口打开失败");
                }
            }
            else if (button2.Text == "关闭串口")
            {

                try
                {

                    serialPort2.Close();
                    button2.Text = "打开串口";

                    comboBox1.Enabled = true;
                    comboBox2.Enabled = true;
                    comboBox3.Enabled = true;
                    comboBox4.Enabled = false;
                    comboBox5.Enabled = false;
                    button4.Enabled = false;
                    button1.Enabled = true;
                }
                catch
                {
                    MessageBox.Show("串口无法关闭");
                }
            }
        }

        private void comboBox6_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void label15_Click(object sender, EventArgs e)
        {

        }

        private void showTextInTextBox(string displayMode, TextBox mytextbox, string portname, byte[] buffer)
        {
            string content = "";
            switch (displayMode)
            {
                case "文本显示":
                    content = Encoding.ASCII.GetString(buffer); // 或用port.Encoding
                    break;
                case "十六进制显示":
                    content = BitConverter.ToString(buffer).Replace("-", " ");
                    break;
                case "二进制显示":
                    StringBuilder binaryBuilder = new StringBuilder();
                    foreach (byte b in buffer)
                        binaryBuilder.Append(Convert.ToString(b, 2).PadLeft(8, '0') + " ");
                    content = binaryBuilder.ToString().Trim();
                    break;
            }
            string displayText = $"[{portname}]: {content}";
            if (mytextbox.InvokeRequired)
            {
                mytextbox.Invoke(new Action(() => mytextbox.AppendText(displayText + "\n")));
            }
            else
            {
                mytextbox.AppendText(displayText + "\n");
            }
        }

        private void label10_Click(object sender, EventArgs e)
        {

        }

        private void button5_Click(object sender, EventArgs e)
        {
            DrawSelectedLayer();
        }

        private void Form1_Load_1(object sender, EventArgs e)
        {

        }

        private void RefreshPlotDisplay()
        {
            InitializePlotHosts();
            UpdatePlotModeUi();
            if (IsCurvePlotMode(GetCurrentPlotMode()))
            {
                int completedPoints = Math.Max(ctrl?.CurrentScannedPoints ?? 0, latestRealtimeCurvePointCount);
                RefreshCurveViewsFromCurrentData(completedPoints);
            }

            switch (GetCurrentPlotMode())
            {
                case PlotViewMode.RotatableSingleLayerHeatmap3D:
                    ShowSingleLayerSurface3D();
                    break;
                case PlotViewMode.StaticSingleLayerHeatmap:
                    ShowSingleLayerStaticHeatmap();
                    break;
                case PlotViewMode.RotatableAverageAlongZHeatmap3D:
                    ShowAverageAlongZSurface3D();
                    break;
                case PlotViewMode.StaticAverageAlongZHeatmap:
                    ShowAverageAlongZStaticHeatmap();
                    break;
                case PlotViewMode.XVoltageCurve:
                    ShowLayeredCurvePlot(xCurveCollection, "B - X 曲线图");
                    break;
                case PlotViewMode.YVoltageCurve:
                    ShowLayeredCurvePlot(yCurveCollection, "B - Y 曲线图");
                    break;
                case PlotViewMode.ZVoltageCurve:
                    ShowSingleCurvePlot(zCurveData);
                    break;
            }
        }

        private void ShowSingleLayerSurface3D()
        {
            LayerSurfaceCache layerCache = GetSelectedLayerCache();
            if (layerCache == null || layerCache.Points == null || layerCache.Points.Count == 0)
            {
                surfaceViewer?.ClearView("暂无三维数据，请先完成至少一层扫描");
                return;
            }

            SurfaceMesh mesh = PlotDataBuilder.BuildSurfaceMesh(
                layerCache.Points,
                "单个 Z 层三维图",
                $"第 {layerCache.LayerIndex + 1} 层 | Z={layerCache.ZValue:F4} mm | 点数={layerCache.Points.Count}");
            surfaceViewer?.ShowMesh(mesh);
        }

        private void ShowSingleLayerStaticHeatmap()
        {
            LayerSurfaceCache layerCache = GetSelectedLayerCache();
            if (layerCache == null || layerCache.Points == null || layerCache.Points.Count == 0)
            {
                heatmapViewer?.ClearView("暂无热力图数据，请先完成至少一层扫描");
                return;
            }

            HeatmapGridData grid = PlotDataBuilder.BuildHeatmapGrid(
                layerCache.Points,
                "原始风格静态热力图",
                $"第 {layerCache.LayerIndex + 1} 层 | Z={layerCache.ZValue:F4} mm | 点数={layerCache.Points.Count}");
            heatmapViewer?.ShowHeatmap(grid);
            heatmapViewer?.BringToFront();
        }

        private void ShowAverageAlongZSurface3D()
        {
            List<Control.PointData> allPoints = GetAllCompletedPoints();
            if (allPoints.Count == 0)
            {
                surfaceViewer?.ClearView("暂无三维数据，请先完成至少一层扫描");
                return;
            }

            SurfaceMesh mesh = PlotDataBuilder.BuildAverageAlongZSurface(allPoints);
            surfaceViewer?.ShowMesh(mesh);
        }

        private void ShowAverageAlongZStaticHeatmap()
        {
            List<Control.PointData> allPoints = GetAllCompletedPoints();
            if (allPoints.Count == 0)
            {
                heatmapViewer?.ClearView("暂无热力图数据，请先完成至少一层扫描");
                return;
            }

            HeatmapGridData grid = PlotDataBuilder.BuildAverageAlongZHeatmap(allPoints);
            heatmapViewer?.ShowHeatmap(grid);
            heatmapViewer?.BringToFront();
        }

        private void ShowLayeredCurvePlot(PlotCurveCollectionData curveCollection, string emptyTitle)
        {
            InitializePlotHosts();
            formsPlot.Visible = true;
            heatmapViewer.Visible = false;
            formsPlot.BringToFront();
            formsPlot.Plot.Clear();

            if (curveCollection == null || !curveCollection.HasData)
            {
                formsPlot.Plot.Title(emptyTitle);
                formsPlot.Refresh();
                return;
            }

            List<PlotCurveSeries> seriesToRender = BuildVisibleCurveSeries(curveCollection);
            if (seriesToRender.Count == 0)
            {
                formsPlot.Plot.Title(curveCollection.Title);
                formsPlot.Plot.XLabel(curveCollection.XLabel);
                formsPlot.Plot.YLabel(curveCollection.YLabel);
                formsPlot.Refresh();
                return;
            }

            System.Drawing.Color[] palette = GetCurvePalette();
            for (int index = 0; index < seriesToRender.Count; index++)
            {
                PlotCurveSeries series = seriesToRender[index];
                bool isAverage = string.Equals(series.LegendText, "平均值", StringComparison.Ordinal);
                AddCurveSeriesToPlot(series, palette[index % palette.Length], isAverage);
            }

            formsPlot.Plot.Title(curveCollection.Title);
            formsPlot.Plot.XLabel(curveCollection.XLabel);
            formsPlot.Plot.YLabel(curveCollection.YLabel);
            formsPlot.Plot.ShowLegend(ScottPlot.Alignment.UpperRight);
            formsPlot.Plot.Axes.AutoScale();
            formsPlot.Refresh();
        }

        private void ShowSingleCurvePlot(PlotCurveData curveData)
        {
            InitializePlotHosts();
            formsPlot.Visible = true;
            heatmapViewer.Visible = false;
            formsPlot.BringToFront();
            formsPlot.Plot.Clear();

            if (curveData == null || curveData.Points == null || curveData.Points.Count == 0)
            {
                formsPlot.Refresh();
                return;
            }

            AddCurveSeriesToPlot(
                new PlotCurveSeries
                {
                    LegendText = string.Empty,
                    Points = curveData.Points
                },
                System.Drawing.Color.FromArgb(30, 90, 170),
                false);

            formsPlot.Plot.Title(curveData.Title);
            formsPlot.Plot.XLabel(curveData.XLabel);
            formsPlot.Plot.YLabel(curveData.YLabel);
            formsPlot.Plot.HideLegend();
            formsPlot.Plot.Axes.AutoScale();
            formsPlot.Refresh();
        }

        private List<PlotCurveSeries> BuildVisibleCurveSeries(PlotCurveCollectionData curveCollection)
        {
            var visibleSeries = new List<PlotCurveSeries>();
            CurveLayerDisplayMode displayMode = GetCurrentCurveDisplayMode();

            switch (displayMode)
            {
                case CurveLayerDisplayMode.SingleLayer:
                    PlotCurveSeries singleLayerSeries = GetSelectedCurveLayerSeries(curveCollection);
                    if (singleLayerSeries != null && singleLayerSeries.HasData)
                    {
                        visibleSeries.Add(singleLayerSeries);
                    }
                    break;
                case CurveLayerDisplayMode.AllLayers:
                    visibleSeries.AddRange(curveCollection.LayerCurves.Where(series => series != null && series.HasData));
                    break;
                case CurveLayerDisplayMode.AverageOnly:
                    if (curveCollection.AverageCurve != null && curveCollection.AverageCurve.HasData)
                    {
                        visibleSeries.Add(curveCollection.AverageCurve);
                    }
                    break;
                case CurveLayerDisplayMode.AllLayersWithAverage:
                    visibleSeries.AddRange(curveCollection.LayerCurves.Where(series => series != null && series.HasData));
                    if (curveCollection.AverageCurve != null && curveCollection.AverageCurve.HasData)
                    {
                        visibleSeries.Add(curveCollection.AverageCurve);
                    }
                    break;
            }

            return visibleSeries;
        }

        private PlotCurveSeries GetSelectedCurveLayerSeries(PlotCurveCollectionData curveCollection)
        {
            if (curveCollection?.LayerCurves == null || curveCollection.LayerCurves.Count == 0)
            {
                return null;
            }

            double targetZ = selectedCurveZLayer ?? curveCollection.LayerCurves.First().ZValue.GetValueOrDefault();
            PlotCurveSeries matchedSeries = curveCollection.LayerCurves.FirstOrDefault(series =>
                series != null &&
                series.ZValue.HasValue &&
                NormalizeLayerZ(series.ZValue.Value) == NormalizeLayerZ(targetZ));

            return matchedSeries ?? curveCollection.LayerCurves.FirstOrDefault(series => series != null && series.HasData);
        }

        private void AddCurveSeriesToPlot(PlotCurveSeries series, System.Drawing.Color color, bool isAverage)
        {
            if (series == null || !series.HasData)
            {
                return;
            }

            double[] xs = series.Points.OrderBy(point => point.AxisValue).Select(point => point.AxisValue).ToArray();
            double[] ys = series.Points.OrderBy(point => point.AxisValue).Select(point => point.B).ToArray();
            var scatter = formsPlot.Plot.Add.Scatter(xs, ys);
            scatter.LegendText = series.LegendText;
            scatter.LineWidth = isAverage ? 3.5f : 2f;
            scatter.MarkerSize = isAverage ? 0f : 4f;
            scatter.Color = ScottPlot.Color.FromColor(color);
        }

        private System.Drawing.Color[] GetCurvePalette()
        {
            return new[]
            {
                System.Drawing.Color.FromArgb(31, 119, 180),
                System.Drawing.Color.FromArgb(255, 127, 14),
                System.Drawing.Color.FromArgb(44, 160, 44),
                System.Drawing.Color.FromArgb(214, 39, 40),
                System.Drawing.Color.FromArgb(148, 103, 189),
                System.Drawing.Color.FromArgb(140, 86, 75),
                System.Drawing.Color.FromArgb(227, 119, 194),
                System.Drawing.Color.FromArgb(127, 127, 127)
            };
        }

        private void Clear2DPlot()
        {
            if (formsPlot == null)
            {
                return;
            }

            formsPlot.Plot.Clear();
            formsPlot.Refresh();
        }

        private void HandleZLayerCompleted(int layerIndex, double zValue, Control.PointData[] layerData)
        {
            if (InvokeRequired)
            {
                BeginInvoke(new Action(() => HandleZLayerCompleted(layerIndex, zValue, layerData)));
                return;
            }

            double normalizedZ = NormalizeLayerZ(zValue);
            List<Control.PointData> validPoints = (layerData ?? Array.Empty<Control.PointData>())
                .Where(point => point.Order >= 0)
                .OrderBy(point => point.Order)
                .ToList();

            if (validPoints.Count == 0)
            {
                return;
            }

            zLayerCache[normalizedZ] = BuildLayerSurfaceCache(layerIndex, normalizedZ, validPoints);
            RefreshComputedViews();
            RefreshZLayerSelector();
            RefreshCurveLayerSelector();
            SelectLayer(normalizedZ);

            if (enableRealTimeDrawing)
            {
                RefreshPlotDisplay();
            }
        }

        private LayerSurfaceCache BuildLayerSurfaceCache(int layerIndex, double zValue, List<Control.PointData> validData)
        {
            return new LayerSurfaceCache
            {
                LayerIndex = layerIndex,
                ZValue = zValue,
                Points = validData
            };
        }

        private void RefreshComputedViews()
        {
            RefreshCurveViewsFromCurrentData(ctrl?.CurrentScannedPoints ?? 0);
        }

        private List<Control.PointData> GetAllCompletedPoints()
        {
            return zLayerCache.Values
                .OrderBy(layer => layer.LayerIndex)
                .SelectMany(layer => layer.Points ?? Enumerable.Empty<Control.PointData>())
                .Where(point => point.Order >= 0)
                .OrderBy(point => point.Order)
                .ToList();
        }

        private List<Control.PointData> GetCurrentScannedPoints(int completedPoints)
        {
            if (ctrl?.DataArray == null || ctrl.DataArray.Length == 0 || completedPoints <= 0)
            {
                return new List<Control.PointData>();
            }

            int safeCount = Math.Min(completedPoints, ctrl.DataArray.Length);
            return ctrl.DataArray
                .Take(safeCount)
                .Where(point => point.Order >= 0)
                .OrderBy(point => point.Order)
                .ToList();
        }

        private List<PlotCurveLayerSource> BuildCurveLayerSources()
        {
            return zLayerCache.Values
                .OrderBy(layer => layer.LayerIndex)
                .Select(layer => new PlotCurveLayerSource
                {
                    LayerIndex = layer.LayerIndex,
                    ZValue = layer.ZValue,
                    Points = (layer.Points ?? new List<Control.PointData>())
                        .Where(point => point.Order >= 0)
                        .OrderBy(point => point.Order)
                        .ToList()
                })
                .ToList();
        }

        private List<PlotCurveLayerSource> BuildCurveLayerSources(IEnumerable<Control.PointData> scannedPoints)
        {
            return (scannedPoints ?? Enumerable.Empty<Control.PointData>())
                .Where(point => point.Order >= 0)
                .GroupBy(point => NormalizeLayerZ(point.Z))
                .OrderBy(group => group.Key)
                .Select((group, index) => new PlotCurveLayerSource
                {
                    LayerIndex = index,
                    ZValue = group.First().Z,
                    Points = group.OrderBy(point => point.Order).ToList()
                })
                .ToList();
        }

        private void RefreshCurveViewsFromCurrentData(int completedPoints)
        {
            List<Control.PointData> scannedPoints = GetCurrentScannedPoints(completedPoints);
            List<PlotCurveLayerSource> layerSources = BuildCurveLayerSources(scannedPoints);
            xCurveCollection = PlotDataBuilder.BuildLayeredCurve(layerSources, point => point.X, "X", "B - X 曲线图");
            yCurveCollection = PlotDataBuilder.BuildLayeredCurve(layerSources, point => point.Y, "Y", "B - Y 曲线图");
            zCurveData = PlotDataBuilder.BuildAverageCurve(scannedPoints, point => point.Z, "Z", "B - Z 曲线图");
        }

        private void RefreshRealtimeCurvePlot(int completedPoints)
        {
            RefreshCurveViewsFromCurrentData(completedPoints);
            lastRealtimeCurveRefreshUtc = DateTime.UtcNow;
            RefreshPlotDisplay();
        }

        private void ResetLayerView()
        {
            zLayerCache.Clear();
            xCurveCollection = new PlotCurveCollectionData();
            yCurveCollection = new PlotCurveCollectionData();
            zCurveData = new PlotCurveData();
            selectedZLayer = null;
            selectedCurveZLayer = null;
            latestRealtimeCurvePointCount = 0;
            lastRealtimeCurveRefreshUtc = DateTime.MinValue;
            isUpdatingZLayerSelector = true;
            try
            {
                comboBox10.BeginUpdate();
                comboBox10.Items.Clear();
                comboBox10.Text = string.Empty;
                comboBox10.Enabled = false;
            }
            finally
            {
                comboBox10.EndUpdate();
                isUpdatingZLayerSelector = false;
            }

            isUpdatingCurveLayerSelector = true;
            try
            {
                comboBox13.BeginUpdate();
                comboBox13.Items.Clear();
                comboBox13.Text = string.Empty;
                comboBox13.Enabled = false;
            }
            finally
            {
                comboBox13.EndUpdate();
                isUpdatingCurveLayerSelector = false;
            }

            UpdatePlotModeUi();
            surfaceViewer?.ClearView("暂无三维数据");
            heatmapViewer?.ClearView("暂无热力图数据");
            Clear2DPlot();
        }

        private void RefreshZLayerSelector()
        {
            isUpdatingZLayerSelector = true;
            try
            {
                comboBox10.BeginUpdate();
                comboBox10.Items.Clear();

                foreach (LayerSurfaceCache layerCache in zLayerCache.Values.OrderBy(layer => layer.LayerIndex))
                {
                    comboBox10.Items.Add(new ZLayerViewItem
                    {
                        LayerIndex = layerCache.LayerIndex,
                        ZValue = layerCache.ZValue
                    });
                }

                comboBox10.Enabled = comboBox10.Items.Count > 0;

                if (selectedZLayer.HasValue)
                {
                    SelectLayerInCombo(selectedZLayer.Value);
                }
                else if (comboBox10.Items.Count > 0)
                {
                    comboBox10.SelectedIndex = comboBox10.Items.Count - 1;
                }
            }
            finally
            {
                comboBox10.EndUpdate();
                isUpdatingZLayerSelector = false;
            }

            UpdatePlotModeUi();
        }

        private void RefreshCurveLayerSelector()
        {
            isUpdatingCurveLayerSelector = true;
            try
            {
                comboBox13.BeginUpdate();
                comboBox13.Items.Clear();

                foreach (LayerSurfaceCache layerCache in zLayerCache.Values.OrderBy(layer => layer.LayerIndex))
                {
                    comboBox13.Items.Add(new ZLayerViewItem
                    {
                        LayerIndex = layerCache.LayerIndex,
                        ZValue = layerCache.ZValue
                    });
                }

                comboBox13.Enabled = comboBox13.Items.Count > 0;

                if (comboBox13.Items.Count == 0)
                {
                    selectedCurveZLayer = null;
                    return;
                }

                ZLayerViewItem selectedItem = comboBox13.Items
                    .Cast<ZLayerViewItem>()
                    .FirstOrDefault(item => selectedCurveZLayer.HasValue &&
                        Math.Abs(NormalizeLayerZ(item.ZValue) - NormalizeLayerZ(selectedCurveZLayer.Value)) < 0.000001);

                if (selectedItem == null)
                {
                    selectedItem = comboBox13.Items.Cast<ZLayerViewItem>().First();
                    selectedCurveZLayer = NormalizeLayerZ(selectedItem.ZValue);
                }

                comboBox13.SelectedItem = selectedItem;
            }
            finally
            {
                comboBox13.EndUpdate();
                isUpdatingCurveLayerSelector = false;
            }

            UpdatePlotModeUi();
        }

        private void SelectLayer(double zValue)
        {
            selectedZLayer = NormalizeLayerZ(zValue);
            SelectLayerInCombo(selectedZLayer.Value);
            DrawSelectedLayer();
        }

        private void SelectLayerInCombo(double zValue)
        {
            for (int index = 0; index < comboBox10.Items.Count; index++)
            {
                if (comboBox10.Items[index] is ZLayerViewItem item &&
                    Math.Abs(NormalizeLayerZ(item.ZValue) - NormalizeLayerZ(zValue)) < 0.000001)
                {
                    comboBox10.SelectedIndex = index;
                    return;
                }
            }
        }

        private void DrawSelectedLayer()
        {
            RefreshPlotDisplay();
        }

        private LayerSurfaceCache GetSelectedLayerCache()
        {
            if (zLayerCache.Count == 0)
            {
                return null;
            }

            double layerKey = selectedZLayer.HasValue
                ? NormalizeLayerZ(selectedZLayer.Value)
                : zLayerCache.Keys.Last();

            if (zLayerCache.TryGetValue(layerKey, out LayerSurfaceCache selectedLayer))
            {
                return selectedLayer;
            }

            return zLayerCache.Values.Last();
        }

        private double NormalizeLayerZ(double zValue)
        {
            return Math.Round(zValue, 6);
        }

        // 切换实时绘图开关
        public void ToggleRealTimeDrawing(bool enable)
        {
            enableRealTimeDrawing = enable;
        }

        private bool HasAnyPlotData()
        {
            return GetSelectedLayerCache() != null ||
                (xCurveCollection?.HasData ?? false) ||
                (yCurveCollection?.HasData ?? false) ||
                (zCurveData?.Points?.Count > 0);
        }

        private void btnDraw_Click(object sender, EventArgs e)
        {
            if (HasAnyPlotData())
            {
                DrawSelectedLayer();
            }
            else
            {
                MessageBox.Show("当前还没有可供绘图的数据");
            }
        }

        private void groupBox7_Enter_1(object sender, EventArgs e)
        {

        }
        private async void button6_Click_1(object sender, EventArgs e)
        {
            bool success = await StopAndReturnHomeFromUiAsync("停止按钮");
            if (success)
            {
                MessageBox.Show("设备已停止并完成复位。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private async Task<bool> StopAndReturnHomeFromUiAsync(string triggerSource, bool showFailureDialog = true)
        {
            if (stopAndReturnTask != null && !stopAndReturnTask.IsCompleted)
            {
                return await stopAndReturnTask;
            }

            stopAndReturnTask = StopAndReturnHomeInternalAsync(triggerSource, showFailureDialog);
            try
            {
                return await stopAndReturnTask;
            }
            finally
            {
                stopAndReturnTask = null;
            }
        }

        private async Task<bool> StopAndReturnHomeInternalAsync(string triggerSource, bool showFailureDialog)
        {
            button6.Enabled = false;
            button9.Enabled = false;

            try
            {
                AppendSerialLog($"[系统] {triggerSource} 触发停止并复位流程");

                if (!serialPort1.IsOpen)
                {
                    AppendSerialLog("[系统] 控制串口未打开，跳过自动复位流程");
                    if (showFailureDialog && triggerSource == "停止按钮")
                    {
                        MessageBox.Show("控制串口未打开，无法执行自动复位。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }

                    return false;
                }

                bool success = await ctrl.StopAndReturnHomeAsync();
                AppendSerialLog(success
                    ? "[系统] 停止并复位流程完成"
                    : "[系统] 停止流程已结束，但自动复位未完全成功");

                if (!success && showFailureDialog)
                {
                    MessageBox.Show(
                        "自动复位未完全成功，请检查设备状态后再继续。",
                        "自动复位提示",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                }

                return success;
            }
            catch (Exception ex)
            {
                AppendSerialLog($"[系统] 停止并复位异常：{ex.Message}");
                if (showFailureDialog)
                {
                    MessageBox.Show(
                        $"停止并复位失败：{ex.Message}",
                        "错误",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                }

                return false;
            }
            finally
            {
                button6.Enabled = true;
                button9.Enabled = true;
            }
        }

        private async void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (allowCloseAfterSafeStop)
            {
                return;
            }

            if (isCloseSequenceRunning)
            {
                e.Cancel = true;
                return;
            }

            bool shouldAttemptReturnHome = serialPort1.IsOpen;
            if (!shouldAttemptReturnHome)
            {
                return;
            }

            e.Cancel = true;
            isCloseSequenceRunning = true;

            bool success = false;
            try
            {
                success = await StopAndReturnHomeFromUiAsync("关闭窗口", false);
            }
            finally
            {
                isCloseSequenceRunning = false;
            }

            if (!success)
            {
                DialogResult forceClose = MessageBox.Show(
                    "自动复位未完全成功，是否仍然强制退出程序？",
                    "关闭确认",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Warning);

                if (forceClose != DialogResult.Yes)
                {
                    return;
                }
            }

            allowCloseAfterSafeStop = true;
            CloseOpenPortsSilently();
            Close();
        }

        private void CloseOpenPortsSilently()
        {
            try
            {
                if (serialPort1.IsOpen)
                {
                    serialPort1.Close();
                }
            }
            catch
            {
            }

            try
            {
                if (serialPort2.IsOpen)
                {
                    serialPort2.Close();
                }
            }
            catch
            {
            }
        }

        // 当panel1大小改变时重绘扫描数据
        private void panel1_Resize(object sender, EventArgs e)
        {
            try
            {
                surfaceViewer?.Invalidate();
                formsPlot?.Refresh();
            }
            catch { }
        }

        private void EnableDoubleBuffer(System.Windows.Forms.Control target)
        {
            try
            {
                typeof(System.Windows.Forms.Control)
                    .GetProperty("DoubleBuffered", BindingFlags.Instance | BindingFlags.NonPublic)
                    ?.SetValue(target, true, null);
            }
            catch
            {
            }
        }

        private void Form1_Resize(object sender, EventArgs e)
        {
            try
            {
                surfaceViewer?.Invalidate();
                formsPlot?.Refresh();
            }
            catch
            {
            }
        }

        private void ArrangeMainLayout()
        {
            if (!IsHandleCreated)
            {
                return;
            }

            int gap = 12;
            int outer = 12;
            int topHeight = 310;
            int leftAreaWidth = Math.Max(760, ClientSize.Width / 2);
            int leftColumnWidth = 320;
            int operationWidth = 150;
            int centerWidth = Math.Max(320, leftAreaWidth - leftColumnWidth - operationWidth - gap * 2);
            int chartLeft = outer + leftAreaWidth + gap;
            int chartWidth = Math.Max(420, ClientSize.Width - chartLeft - outer);
            int bottomTop = outer + topHeight + gap;
            int bottomHeight = Math.Max(250, ClientSize.Height - bottomTop - outer);

            groupBox1.SetBounds(outer, outer, leftColumnWidth, topHeight);
            groupBox3.SetBounds(groupBox1.Right + gap, outer, centerWidth, topHeight);
            groupBox8.SetBounds(groupBox3.Right + gap, outer, operationWidth, 210);

            int progressTop = groupBox8.Bottom + gap;
            int progressWidth = Math.Max(operationWidth, groupBox8.Right - groupBox3.Left);
            groupBox6.SetBounds(groupBox3.Left, progressTop, progressWidth, Math.Max(80, topHeight - (progressTop - outer)));

            groupBox2.SetBounds(outer, bottomTop, leftAreaWidth, bottomHeight);
            groupBox7.SetBounds(chartLeft, outer, chartWidth, ClientSize.Height - outer * 2);

            ArrangeLogControls();
            ArrangeChartControls();
        }

        private void ArrangeLogControls()
        {
            int innerGap = 12;
            int topMargin = 26;
            int rowTop = 28;
            int rowHeight = 24;
            int labelTop = rowTop + 4;
            int bottomRowTop = Math.Max(84, groupBox2.ClientSize.Height - 48);
            int rightButtonWidth = 88;
            int leftComboWidth = 120;
            int textLeft = 250;
            int textWidth = Math.Max(160, groupBox2.ClientSize.Width - textLeft - rightButtonWidth - innerGap * 2);
            int textHeight = Math.Max(120, bottomRowTop - (rowTop + rowHeight + innerGap) - innerGap);

            comboBox5.Location = new Point(26, rowTop);
            label6.Location = new Point(23, topMargin);
            comboBox4.Location = new Point(160, rowTop);
            label5.Location = new Point(157, topMargin);
            button3.Location = new Point(groupBox2.ClientSize.Width - rightButtonWidth - innerGap, rowTop);

            textBox1.SetBounds(innerGap, rowTop + rowHeight + innerGap, groupBox2.ClientSize.Width - innerGap * 2, textHeight);

            label16.Location = new Point(23, bottomRowTop + 4);
            comboBox9.SetBounds(91, bottomRowTop, leftComboWidth, rowHeight);
            label17.Location = new Point(218, bottomRowTop + 4);
            textBox2.SetBounds(textLeft, bottomRowTop, textWidth, rowHeight);
            button4.Location = new Point(groupBox2.ClientSize.Width - rightButtonWidth - innerGap, bottomRowTop);
        }

        private void ArrangeChartControls()
        {
            int innerGap = 12;
            label21.Location = new Point(innerGap, 28);
            comboBox11.Location = new Point(label21.Right + 8, 24);
            label19.Location = new Point(Math.Max(comboBox11.Right + 28, groupBox7.ClientSize.Width - 190), 28);
            comboBox10.Location = new Point(Math.Max(label19.Right + 8, groupBox7.ClientSize.Width - 135), 24);
            panel1.SetBounds(innerGap, 56, Math.Max(200, groupBox7.ClientSize.Width - innerGap * 2),
                Math.Max(180, groupBox7.ClientSize.Height - 56 - 46));
            panel2DHost.SetBounds(panel1.Left, panel1.Top, panel1.Width, panel1.Height);
            button5.Location = new Point(innerGap, groupBox7.ClientSize.Height - button5.Height - 10);
        }

        private void button7_Click(object sender, EventArgs e)
        {
            using (SaveFileDialog saveFileDialog1 = new SaveFileDialog())
            {
                saveFileDialog1.Filter = "CSV文件 (*.csv)|*.csv|文本文件 (*.txt)|*.txt|所有文件 (*.*)|*.*";
                saveFileDialog1.FilterIndex = 1;
                saveFileDialog1.Title = "保存点位数据";
                saveFileDialog1.FileName = $"点位数据_{DateTime.Now:yyyyMMddHHmmss}";
                saveFileDialog1.RestoreDirectory = true;

                if (saveFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        string dataFilePath = saveFileDialog1.FileName;
                        string imageFilePath = Path.ChangeExtension(dataFilePath, ".png");

                        SavePointDataToFile(dataFilePath);
                        SavePlotImageToFile(imageFilePath);

                        MessageBox.Show($"数据保存成功！\r\n点数据：{dataFilePath}\r\n图像：{imageFilePath}",
                            "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    catch (IOException ex)
                    {
                        MessageBox.Show($"文件保存失败：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"未知错误：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }
        private void SavePointDataToFile(string filePath)
        {
            Control.PointData[] _pointDataList = ctrl.DataArray;
            if (_pointDataList == null || _pointDataList.Length == 0)
            {
                throw new InvalidOperationException("没有可保存的点位数据");
            }

            using (StreamWriter writer = new StreamWriter(filePath, false, System.Text.Encoding.UTF8))
            {
                writer.WriteLine("序号,X坐标,Y坐标,Z坐标,U,B");
                foreach (var point in _pointDataList)
                {
                    if (point.Order != -1)
                    {
                        writer.WriteLine($"{point.Order},{point.X:F4},{point.Y:F4},{point.Z:F4},{point.Voltage:F6},{point.B:F6}");
                    }
                }
            }
        }

        private void SavePlotImageToFile(string filePath)
        {
            if (!HasAnyPlotData())
            {
                throw new InvalidOperationException("当前没有可保存的图像数据");
            }

            int width = Math.Max(1200, panel1.ClientSize.Width);
            int height = Math.Max(800, Math.Max(panel1.ClientSize.Height, panel2DHost.ClientSize.Height));

            if (Is3DPlotMode(GetCurrentPlotMode()))
            {
                using (Bitmap bitmap = surfaceViewer.CaptureBitmap(width, height))
                {
                    bitmap.Save(filePath, ImageFormat.Png);
                }
            }
            else if (IsStaticHeatmapMode(GetCurrentPlotMode()))
            {
                using (Bitmap bitmap = heatmapViewer.CaptureBitmap(width, height))
                {
                    bitmap.Save(filePath, ImageFormat.Png);
                }
            }
            else
            {
                formsPlot.Plot.SavePng(filePath, width, height);
            }
        }

        private void textBox13_TextChanged(object sender, EventArgs e)
        {

        }


        private void saveFileDialog1_FileOk(object sender, CancelEventArgs e)
        {

        }
    }
}
