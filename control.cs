using System;
using System.Collections;
using System.Collections.Generic;
using System.Data.Common;
using System.Drawing.Text;
using System.IO;
using System.IO.Ports;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;
using static System.Windows.Forms.LinkLabel;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using System.Globalization;

namespace grbloxy
{
    public partial class Form1 : Form  // 假设你的类名是这个
    {
        // 添加这行 - 取消令牌源
        //private CancellationTokenSource _cancellationTokenSource;
        // 串口1的数据接收处理
        private StringBuilder serialBuffer = new StringBuilder();
        private void Port_DataReceived(object sender, SerialDataReceivedEventArgs e)
        {
            // 1. 安全转换串口对象并校验状态
            SerialPort port = sender as SerialPort;
            if (port == null || !port.IsOpen)
            {
                return; // 串口无效/已关闭，直接返回
            }

            try
            {
                // 2. 校验可读字节数，避免空读取
                int bytesToRead = port.BytesToRead;
                if (bytesToRead <= 0)
                {
                    return; // 无数据可读，直接返回
                }

                // 3. 读取串口数据（安全读取）
                byte[] buffer = new byte[bytesToRead];
                int actualReadBytes = port.Read(buffer, 0, bytesToRead);
                if (actualReadBytes <= 0)
                {
                    return; // 未读取到有效数据
                }

                // 4. 转换为字符串并追加到缓存
                string received = port.Encoding.GetString(buffer, 0, actualReadBytes); // 仅转换实际读取的字节
                serialBuffer.Append(received);

                // 5. 按换行符分割数据（保留原有分割逻辑）
                string[] lines = serialBuffer.ToString()
                    .Split(new[] { "\r\n", "\n" }, StringSplitOptions.None);

                // 6. 关键：校验lines数组有效性，避免越界
                if (lines == null || lines.Length == 0)
                {
                    serialBuffer.Clear(); // 异常情况清空缓存
                    return;
                }

                // 7. 更新LatestLines供Scan方法调用
                ctrl.LatestLines = lines;

                // 8. 遍历有效行（除最后一项）并显示（安全遍历）
                if (lines.Length > 1) // 确保有可显示的完整行
                {
                    for (int i = 0; i < lines.Length - 1; i++)
                    {
                        string line = lines[i];
                        if (string.IsNullOrWhiteSpace(line))
                        {
                            continue; // 跳过空行，避免无效显示
                        }

                        // 拼接显示文本（保留端口号前缀）
                        string displayText = $"[{port.PortName}]: {line}";

                        // 9. 跨线程安全更新UI（增加异常捕获）
                        try
                        {
                            if (textBox1.InvokeRequired)
                            {
                                textBox1.Invoke(new Action(() =>
                                {
                                    textBox1.AppendText(displayText + "\r\n");
                                    // 可选：自动滚动到最新行
                                    textBox1.ScrollToCaret();
                                }));
                            }
                            else
                            {
                                textBox1.AppendText(displayText + "\r\n");
                                textBox1.ScrollToCaret();
                            }
                        }
                        catch (InvalidOperationException)
                        {
                            // 窗体已关闭/控件已释放时捕获异常，不影响主线程
                        }
                    }
                }

                // 10. 安全更新缓存（保留最后一项不完整数据）
                serialBuffer.Clear();
                // 仅当最后一项非空时追加，避免缓存空字符串
                if (!string.IsNullOrWhiteSpace(lines[lines.Length - 1]))
                {
                    serialBuffer.Append(lines[lines.Length - 1]);
                }
            }
            catch (ArgumentOutOfRangeException ex)
            {
                // 捕获索引越界异常，记录日志但不崩溃
                System.Diagnostics.Debug.WriteLine($"串口数据解析越界：{ex.Message}");
                serialBuffer.Clear(); // 异常时清空缓存，避免后续持续报错
            }
            catch (Exception ex)
            {
                // 捕获其他串口/IO异常
                System.Diagnostics.Debug.WriteLine($"串口接收异常：{ex.Message}");
                serialBuffer.Clear();
            }
        }


    }

    public class Control
    {
        // Cancellation token source to allow cooperative cancellation of the scan task
        private CancellationTokenSource _cts;
        private Task _scanTask = Task.CompletedTask;
        private readonly object _stopAndReturnSync = new object();
        private Task<bool> _stopAndReturnTask;

        // Indicates whether a scan is currently running
        public bool IsScanRunning { get; private set; }
        public bool IsStopAndReturnInProgress => _stopAndReturnTask != null && !_stopAndReturnTask.IsCompleted;

        // Event fired when scanning stops (either normally or by cancellation)
        public event Action ScanStopped;
        // 可配置的保持命令和间隔（用于在落笔时持续施加力/位置保持）
        // 默认使用更慢的进给率以便提供更大的保持力矩（设备相关，可在运行时调整）
        public string HoldCommand { get; set; } = "$J=G90 Z0.0 F500";
        public int HoldIntervalMs { get; set; } = 300;
        // 可配置的降笔/抬笔命令（默认值可根据机器调整）
        public string DropCommand { get; set; } = "$J=G90 Z0.0 F1000";
        public string RaiseCommand { get; set; } = "$J=G91 Z4.0 F1000";
        // 在 Control 类中添加以下公共属性
        public PointData[] DataArray => dataArray;
        public int TotalPoints => scanSettings?.TotalPoints ?? 0;
        public bool HasData => scanSettings != null && dataArray != null;
        public string[] LatestLines { get; set; }
        public int CurrentScannedPoints => scanSettings?.NextPosition ?? 0;
        
        // 获取每行的点数
        public int GetPointsPerRow()
        {
            // X_points is double; cast to int safely and add 1
            if (scanSettings == null) return 0;
            return scanSettings.XPointCount;
        }

        public int GetPointsPerLayer()
        {
            if (scanSettings == null) return 0;
            return scanSettings.XPointCount * scanSettings.YPointCount;
        }
        
        // 添加实时绘图事件：在单点数据写入完成后触发
        public event Action<int, PointData> PointScanned;
        public event Action<int, double, PointData[]> ZLayerCompleted;
        public Action<string> LogMessage { get; set; }
        public const double VoltageToBCoefficient = 7.3973e-5;
        private const int MaxAllowedScanPoints = 500000;
        private const int CoordinateDecimals = 4;

        public static double ConvertVoltageToB(double voltage)
        {
            return voltage / VoltageToBCoefficient;
        }

        public static double ConvertBToVoltage(double bValue)
        {
            return bValue * VoltageToBCoefficient;
        }

        private static double NormalizeCoordinate(double value)
        {
            return Math.Round(value, CoordinateDecimals);
        }

        private static string FormatCoordinate(double value)
        {
            return NormalizeCoordinate(value).ToString($"F{CoordinateDecimals}", CultureInfo.InvariantCulture);
        }
        // 声明结构体
        public struct PointData
        {
            public int Order;
            public double X;
            public double Y;
            public double Z;
            public double Voltage;
            public double B => ConvertVoltageToB(Voltage);
                public PointData(int order, double x, double y, double z, double voltage)
                {
                    Order = order;
                    X = x;
                    Y = y;
                    Z = z;
                    Voltage = voltage;
            }
        }
        // 数据数组；在生成路径时按需要分配大小以避免索引越界/溢出
        PointData[] dataArray = new PointData[0];
        public class ScanSettings
        {
            public int NextPosition { get; set; }
            public int TotalPoints { get; set; }
            public int XPointCount { get; set; }
            public int YPointCount { get; set; }
            public int ZPointCount { get; set; }
            public ScanSettings(int nextPosition, int totalPoints, int xPointCount, int yPointCount, int zPointCount)
            {
                NextPosition = nextPosition;
                TotalPoints = totalPoints;
                XPointCount = xPointCount;
                YPointCount = yPointCount;
                ZPointCount = zPointCount;
            }
        }
        ScanSettings scanSettings;
        // 声明数组（假设长度为N）

        enum GrblState
        {
            Run,
            Idle,
            Alarm,
            Hold,
            Jog,
            Door,
            Check,
            Home,
            Sleep,
            Unknown
        }
        GrblState MachineState;
        private GrblState _lastObservedGrblState = GrblState.Unknown;

        private SerialPort SerialPort_control, SerialPort_data;
        // 定义一个结构体来封装扫描参数
        struct ScanParameter
        {
            // 自动属性（推荐：封装性更好）
            public double MinX { get; set; }
            public double MaxX { get; set; }
            public double MinY { get; set; }
            public double MaxY { get; set; }
            public double MinZ { get; set; }
            public double MaxZ { get; set; }
            public double StepX { get; set; }
            public double StepY { get; set; }
            public double StepZ { get; set; }
            public double Speed { get; set; }
            public int SampleCount { get; set; }
        

            // 可选：自定义构造函数（初始化参数）
             public ScanParameter(double minX, double maxX, double minY, double maxY, double minZ, double maxZ, double stepX, double stepY, double stepZ, double speed, int sampleCount)
            {
                MinX = minX;
                MaxX = maxX;
                MinY = minY;
                MaxY = maxY;
                MinZ = minZ;
                MaxZ = maxZ;
                StepX = stepX;
                StepY = stepY;
                StepZ = stepZ;
                Speed = speed;
                SampleCount = sampleCount;

            }

        }
        private ScanParameter scanParam;// 定义一个成员变量来保存扫描参数

        private List<double> BuildAxisPositions(double minValue, double maxValue, double step, string axisName)
        {
            if (step <= 0)
            {
                throw new InvalidOperationException($"{axisName} 轴步长必须大于 0。");
            }

            if (Math.Abs(minValue - maxValue) < 1e-8)
            {
                return new List<double> { NormalizeCoordinate(minValue) };
            }

            double tolerance = Math.Max(1e-8, Math.Abs(step) * 1e-6);
            var positions = new List<double>();
            bool descending = minValue > maxValue;
            double signedStep = descending ? -step : step;
            double current = minValue;
            int guard = 0;

            while ((descending && current > maxValue + tolerance) ||
                   (!descending && current < maxValue - tolerance))
            {
                positions.Add(NormalizeCoordinate(current));
                current += signedStep;
                guard++;

                if (guard > MaxAllowedScanPoints)
                {
                    throw new InvalidOperationException($"{axisName} 轴步长过小，生成点数过多，请增大步长。");
                }
            }

            if (positions.Count == 0 || Math.Abs(positions[positions.Count - 1] - maxValue) > tolerance)
            {
                positions.Add(NormalizeCoordinate(maxValue));
            }
            else
            {
                positions[positions.Count - 1] = NormalizeCoordinate(maxValue);
            }

            return positions;
        }

        public int EstimateTotalPoints(double minX, double maxX, double minY, double maxY, double minZ, double maxZ, double stepX, double stepY, double stepZ)
        {
            List<double> xPositions = BuildAxisPositions(minX, maxX, stepX, "X");
            List<double> yPositions = BuildAxisPositions(minY, maxY, stepY, "Y");
            List<double> zPositions = BuildAxisPositions(minZ, maxZ, stepZ, "Z");

            long totalPoints = (long)xPositions.Count * yPositions.Count * zPositions.Count;

            if (totalPoints > MaxAllowedScanPoints)
            {
                throw new InvalidOperationException($"计算得到的点位数量为 {totalPoints}，已超过上限 {MaxAllowedScanPoints}，请增大步长或缩小扫描范围。");
            }

            return (int)totalPoints;
        }

        public void GrblControl(SerialPort Port_contrl, SerialPort Port_data,
            double minX, double maxX, double minY, double maxY, double minZ, double maxZ,
            double stepX, double stepY, double stepZ, double speed, int sampleCount)
        {
            // 如果已有扫描在运行，先请求停止
            if (IsScanRunning)
            {
                try { _cts?.Cancel(); } catch { }
            }

            SerialPort_control = Port_contrl;
            SerialPort_data = Port_data;
            scanParam = new ScanParameter(minX, maxX, minY, maxY, minZ, maxZ, stepX, stepY, stepZ, speed, sampleCount);

            // 创建新的取消令牌源并启动扫描任务（后台运行）
            _cts = new CancellationTokenSource();
            CancellationToken ct = _cts.Token;
            IsScanRunning = true;
            _scanTask = Task.Run(async () =>
            {
                try
                {
                    await this.ExecuteAsync(SerialPort_control, SerialPort_data, scanParam, ct);
                }
                catch (OperationCanceledException) { /* 取消正常 */ }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine("扫描任务异常：" + ex.Message);
                    MessageBox.Show($"扫描任务异常：{ex.Message}", "扫描失败",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                finally
                {
                    IsScanRunning = false;
                    ScanStopped?.Invoke();
                }
            });
        }

        public Task<bool> StopAndReturnHomeAsync()
        {
            lock (_stopAndReturnSync)
            {
                if (_stopAndReturnTask != null && !_stopAndReturnTask.IsCompleted)
                {
                    return _stopAndReturnTask;
                }

                _stopAndReturnTask = StopAndReturnHomeCoreAsync();
                return _stopAndReturnTask;
            }
        }

        private async Task<bool> StopAndReturnHomeCoreAsync()
        {
            try
            {
                bool wasScanning = IsScanRunning;
                if (wasScanning)
                {
                    LogMessage?.Invoke("[系统] 收到停止请求，准备安全停止并回到扫描起点");
                    StopScan();
                    await WaitForScanTaskToFinishAsync();
                }

                if (SerialPort_control == null || !SerialPort_control.IsOpen)
                {
                    LogMessage?.Invoke("[控制] 控制串口未打开，无法执行自动复位");
                    return false;
                }

                await WaitForIdleOrTimeout("停止后等待设备空闲", 8000);
                return await ReturnToOriginCoreAsync(SerialPort_control, scanParam, "停止/关闭流程");
            }
            finally
            {
                lock (_stopAndReturnSync)
                {
                    _stopAndReturnTask = null;
                }
            }
        }
        
        private void SendCode(SerialPort port, string CODE) 
        {
            byte[] data = Encoding.ASCII.GetBytes(CODE+"\r\n");
            port.Write(data, 0, data.Length);
            LogMessage?.Invoke($"[控制] {CODE}");
        }
        //检查机器运行状态，返回值为布尔型变量，1为运行
        

        private static GrblState ParseGrblStateFromText(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
            {
                return GrblState.Unknown;
            }

            if (text.IndexOf("Run", StringComparison.OrdinalIgnoreCase) >= 0) return GrblState.Run;
            if (text.IndexOf("Idle", StringComparison.OrdinalIgnoreCase) >= 0) return GrblState.Idle;
            if (text.IndexOf("Alarm", StringComparison.OrdinalIgnoreCase) >= 0) return GrblState.Alarm;
            if (text.IndexOf("Hold", StringComparison.OrdinalIgnoreCase) >= 0) return GrblState.Hold;
            if (text.IndexOf("Jog", StringComparison.OrdinalIgnoreCase) >= 0) return GrblState.Jog;
            if (text.IndexOf("Door", StringComparison.OrdinalIgnoreCase) >= 0) return GrblState.Door;
            if (text.IndexOf("Check", StringComparison.OrdinalIgnoreCase) >= 0) return GrblState.Check;
            if (text.IndexOf("Home", StringComparison.OrdinalIgnoreCase) >= 0) return GrblState.Home;
            if (text.IndexOf("Sleep", StringComparison.OrdinalIgnoreCase) >= 0) return GrblState.Sleep;

            return GrblState.Unknown;
        }

        private async Task<GrblState> IsRunState()
        {
            try
            {
                if (!SerialPort_control.IsOpen)
                {
                    Console.WriteLine("串口未打开");
                    return GrblState.Unknown;
                }

                LatestLines = Array.Empty<string>();
                SendCode(SerialPort_control, "?");

                // 最多等待 1.2s，优先解析异步缓存；若缓存未更新，再尝试直接读串口缓冲
                for (int i = 0; i < 12; i++)
                {
                    await Task.Delay(100);

                    string[] response = LatestLines;
                    if (response != null && response.Length > 0)
                    {
                        foreach (string line in response)
                        {
                            GrblState parsed = ParseGrblStateFromText(line);
                            if (parsed != GrblState.Unknown)
                            {
                                _lastObservedGrblState = parsed;
                                return parsed;
                            }
                        }
                    }

                    string directText = string.Empty;
                    try
                    {
                        directText = SerialPort_control.ReadExisting();
                    }
                    catch
                    {
                        // 忽略直接读取失败，继续走缓存解析
                    }

                    GrblState directParsed = ParseGrblStateFromText(directText);
                    if (directParsed != GrblState.Unknown)
                    {
                        _lastObservedGrblState = directParsed;
                        return directParsed;
                    }
                }

                // 若本次未读到，返回最近一次可用状态，减少“误判 Unknown”带来的停滞
                return _lastObservedGrblState;
            }
            catch (Exception ex)
            {
                Console.WriteLine("查询GRBL状态失败：" + ex.Message);
                return _lastObservedGrblState;
            }
        }
        //生成一个扫描坐标路径，输入为扫描参数结构体，输出为坐标列表
        private void Creatpath(ScanParameter ScanParameterStructure)
        {
            List<double> xPositions = BuildAxisPositions(ScanParameterStructure.MinX, ScanParameterStructure.MaxX, ScanParameterStructure.StepX, "X");
            List<double> yPositions = BuildAxisPositions(ScanParameterStructure.MinY, ScanParameterStructure.MaxY, ScanParameterStructure.StepY, "Y");
            List<double> zPositions = BuildAxisPositions(ScanParameterStructure.MinZ, ScanParameterStructure.MaxZ, ScanParameterStructure.StepZ, "Z");

            long totalPointsLong = (long)xPositions.Count * yPositions.Count * zPositions.Count;
            if (totalPointsLong > MaxAllowedScanPoints)
            {
                throw new InvalidOperationException($"计算得到的点位数量为 {totalPointsLong}，已超过上限 {MaxAllowedScanPoints}。");
            }
            int totalPoints = (int)totalPointsLong; // 安全转换

            // 根据实际点数分配或调整数组大小，避免后续索引越界
            dataArray = new PointData[totalPoints];

            scanSettings = new ScanSettings(0, totalPoints, xPositions.Count, yPositions.Count, zPositions.Count);
            LogMessage?.Invoke(
                $"[路径] 轴点数: X={xPositions.Count}, Y={yPositions.Count}, Z={zPositions.Count}, Total={totalPoints} | " +
                $"X:{xPositions.First():F4}->{xPositions.Last():F4}, " +
                $"Y:{yPositions.First():F4}->{yPositions.Last():F4}, " +
                $"Z:{zPositions.First():F4}->{zPositions.Last():F4}");

            // 先按 Z 层扫描，再在每个 Z 层内进行 XY 蛇形扫描
            int i = 0;
            for (int zp = 0; zp < zPositions.Count; zp++)
            {
                for (int yp = 0; yp < yPositions.Count; yp++)
                {
                    // 使用蛇形扫描：偶数行从左到右，奇数行从右到左
                    if (yp % 2 == 0)
                    {
                        for (int xp = 0; xp < xPositions.Count; xp++)
                        {
                            dataArray[i].X = xPositions[xp];
                            dataArray[i].Y = yPositions[yp];
                            dataArray[i].Z = zPositions[zp];
                            dataArray[i].Order = i;
                            i++;
                        }
                    }
                    else
                    {
                        for (int xp = xPositions.Count - 1; xp >= 0; xp--)
                        {
                            dataArray[i].X = xPositions[xp];
                            dataArray[i].Y = yPositions[yp];
                            dataArray[i].Z = zPositions[zp];
                            dataArray[i].Order = i;
                            i++;
                        }
                    }
                }
            }
        }

        private async Task<bool> ReturnToOriginCoreAsync(SerialPort port, ScanParameter scanParameterStructure, string trigger)
        {
            if (port == null || !port.IsOpen)
            {
                LogMessage?.Invoke($"[复位] 控制串口不可用，无法执行自动复位：{trigger}");
                return false;
            }

            // 复位策略：先回到扫描范围上边界的 Z（避免贴面横移），再回 XY 原点，最后回 Z=0（扫描起点）。
            double safeReturnZ = Math.Max(scanParameterStructure.MinZ, scanParameterStructure.MaxZ);
            double resolvedSpeed = scanParameterStructure.Speed > 0 ? scanParameterStructure.Speed : 1000d;
            string speed = resolvedSpeed.ToString("F0", CultureInfo.InvariantCulture);
            LogMessage?.Invoke($"[复位] 触发来源={trigger} | 回位目标: safeZ={safeReturnZ:F4}, XY=(0.0000,0.0000), Z=0.0000");

            SendCode(port, $"$J=G90 Z{FormatCoordinate(safeReturnZ)} F{speed}");
            bool safeZReady = await WaitForIdleOrTimeout("复位到安全 Z");

            SendCode(port, $"$J=G90 X0.0000 Y0.0000 F{speed}");
            bool xyReady = await WaitForIdleOrTimeout("复位到 XY 原点");

            SendCode(port, $"$J=G90 Z0.0000 F{speed}");
            bool zReady = await WaitForIdleOrTimeout("复位到 Z 原点");
            bool success = safeZReady && xyReady && zReady;
            LogMessage?.Invoke(success
                ? "[控制] 已自动返回扫描起点(0,0,0)"
                : "[控制] 自动回位已执行，但存在阶段超时，请检查设备状态");
            return success;
        }

        private async Task WaitForScanTaskToFinishAsync(int timeoutMs = 15000)
        {
            Task activeScanTask = _scanTask;
            if (activeScanTask == null || activeScanTask.IsCompleted)
            {
                return;
            }

            Task completedTask = await Task.WhenAny(activeScanTask, Task.Delay(timeoutMs));
            if (completedTask != activeScanTask)
            {
                LogMessage?.Invoke($"[系统] 等待扫描任务停止超时（{timeoutMs} ms）");
                return;
            }

            try
            {
                await activeScanTask;
            }
            catch (OperationCanceledException)
            {
            }
        }

        private async Task<bool> WaitForIdleOrTimeout(string stage, int timeoutMs = 12000)
        {
            int elapsed = 0;
            while (elapsed < timeoutMs)
            {
                GrblState state = await IsRunState();
                if (state == GrblState.Idle)
                {
                    return true;
                }

                await Task.Delay(100);
                elapsed += 100;
            }

            LogMessage?.Invoke($"[控制] 等待设备 Idle 超时: {stage}（{timeoutMs} ms）");
            return false;
        }

        private async Task MoveToNextPoint(ScanSettings stru, SerialPort Port_contrl, ScanParameter ScanParameterStructure)
        {
            // 优先移动到目标 Z，再移动到目标 XY
            if (stru.NextPosition >= stru.TotalPoints)
            {
                MessageBox.Show("扫描完成");
                return;
            }

            PointData target = dataArray[stru.NextPosition];
            ValidateTargetInRange(target, ScanParameterStructure);
            LogMessage?.Invoke($"[路径] 点#{stru.NextPosition + 1:D4}/{stru.TotalPoints:D4} -> X={target.X:F4}, Y={target.Y:F4}, Z={target.Z:F4}");
            double previousZ = stru.NextPosition > 0 ? dataArray[stru.NextPosition - 1].Z : double.NaN;

            if (double.IsNaN(previousZ) || Math.Abs(previousZ - target.Z) > 0.0001)
            {
                SendCode(Port_contrl, $"$J=G90 Z{FormatCoordinate(target.Z)} F{ScanParameterStructure.Speed.ToString("F0", CultureInfo.InvariantCulture)}");
                await Task.Delay(800);
            }

            SendCode(Port_contrl,
                $"$J=G90 X{FormatCoordinate(target.X)} Y{FormatCoordinate(target.Y)} F{ScanParameterStructure.Speed.ToString("F0", CultureInfo.InvariantCulture)}");
            await Task.Delay(800);
        }

        private static bool IsWithinRange(double value, double a, double b, double tolerance = 1e-6)
        {
            double min = Math.Min(a, b) - tolerance;
            double max = Math.Max(a, b) + tolerance;
            return value >= min && value <= max;
        }

        private void ValidateTargetInRange(PointData target, ScanParameter scanParameterStructure)
        {
            if (!IsWithinRange(target.X, scanParameterStructure.MinX, scanParameterStructure.MaxX) ||
                !IsWithinRange(target.Y, scanParameterStructure.MinY, scanParameterStructure.MaxY) ||
                !IsWithinRange(target.Z, scanParameterStructure.MinZ, scanParameterStructure.MaxZ))
            {
                throw new InvalidOperationException(
                    $"目标点超出扫描范围: X={target.X:F4},Y={target.Y:F4},Z={target.Z:F4} " +
                    $"| 范围 X[{scanParameterStructure.MinX:F4},{scanParameterStructure.MaxX:F4}] " +
                    $"Y[{scanParameterStructure.MinY:F4},{scanParameterStructure.MaxY:F4}] " +
                    $"Z[{scanParameterStructure.MinZ:F4},{scanParameterStructure.MaxZ:F4}]");
            }
        }

        private PointData[] GetLayerSnapshot(int completedPoints)
        {
            int pointsPerLayer = GetPointsPerLayer();
            if (pointsPerLayer <= 0 || completedPoints <= 0)
            {
                return Array.Empty<PointData>();
            }

            int layerIndex = (completedPoints - 1) / pointsPerLayer;
            int startIndex = layerIndex * pointsPerLayer;
            int count = Math.Min(pointsPerLayer, Math.Max(0, dataArray.Length - startIndex));

            if (count <= 0)
            {
                return Array.Empty<PointData>();
            }

            PointData[] snapshot = new PointData[count];
            Array.Copy(dataArray, startIndex, snapshot, 0, count);
            return snapshot;
        }

        private void RaiseLayerCompletedIfNeeded(int completedPoints)
        {
            int pointsPerLayer = GetPointsPerLayer();
            if (pointsPerLayer <= 0 || completedPoints <= 0)
            {
                return;
            }

            bool layerCompleted = completedPoints % pointsPerLayer == 0 || completedPoints >= TotalPoints;
            if (!layerCompleted)
            {
                return;
            }

            PointData[] layerSnapshot = GetLayerSnapshot(completedPoints);
            if (layerSnapshot.Length == 0)
            {
                return;
            }

            int layerIndex = (completedPoints - 1) / pointsPerLayer;
            double zValue = layerSnapshot[0].Z;
            ZLayerCompleted?.Invoke(layerIndex, zValue, layerSnapshot);
        }
        // 简化后的滤波配置（保留核心参数，降低复杂度）
        private readonly double VOLTAGE_UPPER_LIMIT = 10.0;    // 电压上限
        private readonly double VOLTAGE_LOWER_LIMIT = -10.0;   // 电压下限
        public const int DefaultSampleCount = 50;
        public const int MaxSampleCount = 1000;
        private readonly double OUTLIER_THRESHOLD = 3.0;       // 3σ异常值剔除阈值
        private readonly int SCAN_RETRY_TIMES = 2;             // 整体重试次数（简化）
        private readonly object instrumentPortLock = new object();
  

        // 数据项结构体（补充定义）
        public struct DataItem
        {
            public double Voltage;
            // 可添加其他字段
        }

        private async Task Scan(int presentPosition, SerialPort Port_data)
        {
            double finalVoltage = -1.0; // 默认无效值
            bool isDataValid = false;
            int retryCount = 0;
            int measureTimes = Math.Max(1, Math.Min(MaxSampleCount, scanParam.SampleCount));

            // 核心重试逻辑
            while (!isDataValid && retryCount < SCAN_RETRY_TIMES)
            {
                try
                {
                    // 1. 获取多次测量的原始数据
                    var multiMeasureVoltages = new List<double>();
                    for (int measureIdx = 0; measureIdx < measureTimes; measureIdx++)
                    {
                        double singleVoltage = await SingleMeasure(Port_data);
                        // 基础范围校验
                        if (singleVoltage >= VOLTAGE_LOWER_LIMIT && singleVoltage <= VOLTAGE_UPPER_LIMIT)
                        {
                            multiMeasureVoltages.Add(singleVoltage);
                        }
                        await Task.Delay(30); // 测量间隔
                    }

                    // 2. 简化的滤波逻辑（适配小数据）
                    if (multiMeasureVoltages.Count > 0)
                    {
                        var filtered = FilterOutliers(multiMeasureVoltages);
                        finalVoltage = filtered.Count > 0 ? filtered.Average() : multiMeasureVoltages.Average();
                        isDataValid = true; // 只要有数据就视为有效

                        PointData currentPoint = dataArray[presentPosition];
                        LogMessage?.Invoke(
                            $"[采集] 点#{presentPosition + 1:D4} " +
                            $"X={currentPoint.X:F2} Y={currentPoint.Y:F2} Z={currentPoint.Z:F2} " +
                            $"平均电压={finalVoltage:F6} V " +
                            $"(有效样本 {filtered.Count}/{multiMeasureVoltages.Count})");
                    }
                    else
                    {
                        // 无有效测量值，重试
                        retryCount++;
                        await Task.Delay(20);
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"扫描位置{presentPosition}时出错：{ex.Message}");
                    retryCount++;
                    await Task.Delay(20);
                }
            }

            // 赋值最终结果
            dataArray[presentPosition].Voltage = finalVoltage;
            PointScanned?.Invoke(presentPosition + 1, dataArray[presentPosition]);
            ScanCompleteState = isDataValid ? CompleteOrNotState.YES : CompleteOrNotState.NO;
        }

        /// <summary>
        /// 单次测量：改为使用 DM3058/SCPI 查询直流电压
        /// </summary>
        private async Task<double> SingleMeasure(SerialPort Port_data)
        {
            try
            {
                await Task.Delay(20);
                return QueryDcVoltage(Port_data);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"单次测量失败: {ex.Message}");
                return -1.0;
            }
        }

        private List<double> FilterOutliers(List<double> values)
        {
            if (values == null || values.Count < 2)
            {
                return values ?? new List<double>();
            }

            double mean = values.Average();
            double variance = values.Sum(v => Math.Pow(v - mean, 2)) / Math.Max(1, values.Count - 1);
            double std = Math.Sqrt(variance);

            if (std <= double.Epsilon)
            {
                return new List<double>(values);
            }

            return values
                .Where(v => Math.Abs(v - mean) <= OUTLIER_THRESHOLD * std)
                .ToList();
        }

        /// <summary>
        /// 电压解析：兼容直接数值响应和简单文本响应
        /// </summary>
        public double ParseVoltageData(string input)
        {
            if (string.IsNullOrWhiteSpace(input))
                return -1;

            try
            {
                string normalized = input.Trim();
                if (double.TryParse(normalized, NumberStyles.Float, CultureInfo.InvariantCulture, out double directVoltage))
                {
                    return IsVoltageInRange(directVoltage) ? directVoltage : -1;
                }

                string[] parts = normalized.Split(',');
                foreach (string part in parts)
                {
                    if (double.TryParse(part.Trim(), NumberStyles.Float, CultureInfo.InvariantCulture, out double parsedValue))
                    {
                        double voltage = Math.Abs(parsedValue) > VOLTAGE_UPPER_LIMIT ? parsedValue / 1000.0 : parsedValue;
                        return IsVoltageInRange(voltage) ? voltage : -1;
                    }
                }
            }
            catch
            {
                // 静默失败，返回无效值
            }
            return -1;
        }

        private double QueryDcVoltage(SerialPort port)
        {
            string response = QueryInstrumentLine(port, ":MEAS:VOLT:DC?");
            return ParseVoltageData(response);
        }

        private string QueryInstrumentLine(SerialPort port, string command)
        {
            if (port == null || !port.IsOpen)
            {
                throw new InvalidOperationException("采集串口未打开");
            }

            lock (instrumentPortLock)
            {
                port.ReadTimeout = 1000;
                port.WriteTimeout = 1000;
                port.NewLine = "\n";
                port.DiscardInBuffer();
                port.Write((command ?? string.Empty).Trim() + "\r\n");
                string response = port.ReadLine();
                return response?.Trim() ?? string.Empty;
            }
        }

        private bool IsVoltageInRange(double voltage)
        {
            return voltage >= VOLTAGE_LOWER_LIMIT && voltage <= VOLTAGE_UPPER_LIMIT;
        }
        public string SaveSimplePointData(System.Windows.Forms.TextBox logTextBox = null)
        {
            try
            {
                // 1. 核心：创建默认目录+生成不重复文件名
                string saveDir = Path.Combine(Application.StartupPath, "ScanData");
                Directory.CreateDirectory(saveDir); // 无需判断是否存在，CreateDirectory会自动处理
                string fileName = $"ScanData_{DateTime.Now:yyyyMMdd_HHmmss}.csv";
                string fullPath = Path.Combine(saveDir, fileName);

                // 2. 核心：写入数据（仅保存有效数据）
                int saveCount = 0;
                using (StreamWriter writer = new StreamWriter(fullPath))
                {
                    writer.WriteLine("Order,X,Y,Z,U,B");
                    foreach (PointData point in dataArray)
                    {
                        if (point.Order >= scanSettings.TotalPoints) break;
                        writer.WriteLine($"{point.Order},{point.X:F4},{point.Y:F4},{point.Z:F4},{point.Voltage:F6},{point.B:F6}");
                        saveCount++;
                    }
                }

                // 3. 反馈1：追加到文本框（跨线程安全）
                string successMsg = $"[{DateTime.Now:HH:mm:ss}] 保存成功：{fullPath} | 共{saveCount}条数据\r\n";
                if (logTextBox != null)
                {
                    if (logTextBox.InvokeRequired)
                        logTextBox.Invoke(new Action(() => logTextBox.AppendText(successMsg)));
                    else
                        logTextBox.AppendText(successMsg);
                }

                // 4. 反馈2：弹窗提示
                MessageBox.Show($"数据保存成功！\r\n路径：{fullPath}\r\n记录数：{scanSettings.TotalPoints}", "保存完成",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);

                return fullPath;
            }
            catch (Exception ex)
            {
                // 异常反馈
                string errorMsg = $"[{DateTime.Now:HH:mm:ss}] 保存失败：{ex.Message}\r\n";
                if (logTextBox != null)
                {
                    if (logTextBox.InvokeRequired)
                        logTextBox.Invoke(new Action(() => logTextBox.AppendText(errorMsg)));
                    else
                        logTextBox.AppendText(errorMsg);
                }
                MessageBox.Show($"保存失败：{ex.Message}", "错误",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return string.Empty;
            }
        }
        enum CompleteOrNotState
        {
            YES,
            NO,
        }
        CompleteOrNotState ScanCompleteState, MoveCompleteState;

        enum PrensentWorkState
        {
            WAIT,
            IDLE,
            SCAN,
            RUN,
            WORKFINISHIED,
            ALLFINISHED,
            EXTI
        }
        PrensentWorkState prensentWorkState;

        public void StopScan()
        {
            // 请求取消后台扫描任务
            try
            {
                _cts?.Cancel();
            }
            catch { }
            // 同步设置状态机为退出状态，ExecuteAsync 会检测取消并返回
            prensentWorkState = PrensentWorkState.EXTI;
        }
        private async Task ExecuteAsync(SerialPort Port_contrl, SerialPort Port_data, ScanParameter ScanParameterStructure, CancellationToken ct)
        {
            for (int i = 0; i < dataArray.Length; i++)
            {
                dataArray[i].Order = -1; // 初始化数据数组
                dataArray[i].X = 0;
                dataArray[i].Y = 0;
                dataArray[i].Voltage = 0;
            }

            SendCode(Port_contrl, "G92 X0 Y0 Z0"); // 设置当前坐标为原点
            LogMessage?.Invoke(
                $"[系统] 扫描开始，当前物理位置映射为工作坐标原点(0,0,0) | " +
                $"范围 X={ScanParameterStructure.MinX:F4}~{ScanParameterStructure.MaxX:F4}, " +
                $"Y={ScanParameterStructure.MinY:F4}~{ScanParameterStructure.MaxY:F4}, " +
                $"Z={ScanParameterStructure.MinZ:F4}~{ScanParameterStructure.MaxZ:F4}");
            Creatpath(ScanParameterStructure); // 生成扫描路径
            prensentWorkState = PrensentWorkState.SCAN;
            ScanCompleteState = CompleteOrNotState.NO;

            while (true)
            {
                ct.ThrowIfCancellationRequested();
                if(scanSettings.NextPosition == 0)
                {
                    prensentWorkState = PrensentWorkState.SCAN;
                }
                // 根据当前状态执行相应的操作
                switch (prensentWorkState)
                {
                    case PrensentWorkState.WAIT:
                        int delaytime = 0;
                        int pointsPerRow = Math.Max(1, scanSettings.XPointCount);
                        if (scanSettings.NextPosition > 0 && scanSettings.NextPosition % pointsPerRow == 0)
                        {
                            delaytime = (int)(Math.Abs(ScanParameterStructure.MaxX - ScanParameterStructure.MinX) / ScanParameterStructure.Speed * 1000) + 200;

                        }
                        else
                        {
                            double axisStep = Math.Max(ScanParameterStructure.StepX,
                                Math.Max(ScanParameterStructure.StepY, ScanParameterStructure.StepZ));
                            delaytime = (int)(axisStep / ScanParameterStructure.Speed * 1000);
                        }
                        await Task.Delay(delaytime, ct);

                        GrblState grbstate = await IsRunState();
                        if (grbstate == GrblState.Idle)
                        {
                            await Task.Delay(100, ct);
                            prensentWorkState = PrensentWorkState.SCAN;
                        }
                        else if (grbstate == GrblState.Unknown)
                        {
                            LogMessage?.Invoke("[控制] 未收到 GRBL 状态响应，请检查控制串口和控制器状态");
                            await Task.Delay(100, ct);
                        }
                        else
                        {
                            await Task.Delay(50, ct);
                        }
                        break;
                    case PrensentWorkState.IDLE:
                        break;
                    case PrensentWorkState.RUN:
                        ct.ThrowIfCancellationRequested();
                        await MoveToNextPoint(scanSettings, Port_contrl, ScanParameterStructure);
                        prensentWorkState = PrensentWorkState.WAIT;
                        break;
                    case PrensentWorkState.SCAN:
                        if (ScanCompleteState == CompleteOrNotState.NO)
                        {
                            int presentposition = scanSettings.NextPosition;
                            ct.ThrowIfCancellationRequested();
                            await Scan(presentposition, Port_data);
                        }
                        else if (ScanCompleteState == CompleteOrNotState.YES)
                        {
                            ScanCompleteState = CompleteOrNotState.NO;
                            scanSettings.NextPosition++;

                            RaiseLayerCompletedIfNeeded(scanSettings.NextPosition);

                            prensentWorkState = PrensentWorkState.RUN;
                        }
                        if (scanSettings.NextPosition >= scanSettings.TotalPoints)
                        {
                            LogMessage?.Invoke($"[系统] 扫描完成，总点数={scanSettings.TotalPoints}，开始执行回位");
                            MessageBox.Show($"数据扫描完成\r\n记录数：{scanSettings.TotalPoints}");
                        prensentWorkState = PrensentWorkState.ALLFINISHED;
                        }
                        break;
                    case PrensentWorkState.ALLFINISHED:
                        await ReturnToOriginCoreAsync(Port_contrl, ScanParameterStructure, "扫描自然结束");
                        scanSettings.NextPosition = 0;
                        //弹出窗口指定路径生成文件
                        
                       //SaveSimplePointData();
                        return; // 结束异步方法
                    case PrensentWorkState.EXTI:
                        scanSettings.NextPosition = 0;
                        ct.ThrowIfCancellationRequested();
                        return; // 结束异步方法
                }
            }

        }
    }
    
}
