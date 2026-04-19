using System;
using System.Collections.Generic;
using System.IO.Ports;
using System.Text;
using System.Threading;

namespace grbloxy
{
    /// <summary>
    /// 串口数据处理器 - 自动存储接收到的数据
    /// </summary>
    public class SerialDataHandler : IDisposable
    {
        // 串口对象
        private SerialPort _port;

        // 数据存储
        private readonly object _lockObject = new object();
        private List<byte> _rawDataBuffer = new List<byte>();        // 原始字节数据
        private StringBuilder _textBuffer = new StringBuilder();      // 文本数据
        private Queue<byte[]> _dataPackets = new Queue<byte[]>();     // 数据包队列

        // 配置选项
        public Encoding Encoding { get; set; } = Encoding.ASCII;
        public bool AutoClearAfterRead { get; set; } = false;
        public int MaxBufferSize { get; set; } = 1024 * 1024; // 1MB 限制

        // 事件
        public event EventHandler<DataReceivedEventArgs> DataReceived;
        public event EventHandler<ErrorEventArgs> ErrorOccurred;

        // 属性
        public bool IsRunning { get; private set; }
        public int DataCount => _rawDataBuffer.Count;
        public int PacketCount => _dataPackets.Count;

        /// <summary>
        /// 构造函数
        /// </summary>
        public SerialDataHandler(SerialPort port)
        {
            _port = port ?? throw new ArgumentNullException(nameof(port));
        }

        /// <summary>
        /// 开始监听数据
        /// </summary>
        public void Start()
        {
            if (!IsRunning)
            {
                _port.DataReceived += Port_DataReceived;
                IsRunning = true;
                OnDataReceived(new DataReceivedEventArgs("开始监听串口数据", DataType.System));
            }
        }

        /// <summary>
        /// 停止监听
        /// </summary>
        public void Stop()
        {
            if (IsRunning)
            {
                _port.DataReceived -= Port_DataReceived;
                IsRunning = false;
                OnDataReceived(new DataReceivedEventArgs("停止监听串口数据", DataType.System));
            }
        }

        /// <summary>
        /// 串口数据接收事件
        /// </summary>
        private void Port_DataReceived(object sender, SerialDataReceivedEventArgs e)
        {
            try
            {
                SerialPort port = (SerialPort)sender;

                // 读取所有可用数据
                int bytesToRead = port.BytesToRead;
                if (bytesToRead <= 0) return;

                byte[] buffer = new byte[bytesToRead];
                int bytesRead = port.Read(buffer, 0, bytesToRead);

                if (bytesRead > 0)
                {
                    // 截取实际读取的数据
                    byte[] data = new byte[bytesRead];
                    Array.Copy(buffer, data, bytesRead);

                    // 存储数据
                    StoreData(data);

                    // 触发事件
                    OnDataReceived(new DataReceivedEventArgs(data, Encoding));
                }
            }
            catch (Exception ex)
            {
                OnErrorOccurred(new ErrorEventArgs(ex));
            }
        }

        /// <summary>
        /// 存储数据到各个缓冲区
        /// </summary>
        private void StoreData(byte[] data)
        {
            lock (_lockObject)
            {
                // 存储到原始字节缓冲区
                _rawDataBuffer.AddRange(data);

                // 限制缓冲区大小
                if (_rawDataBuffer.Count > MaxBufferSize)
                {
                    _rawDataBuffer.RemoveRange(0, _rawDataBuffer.Count - MaxBufferSize);
                }

                // 存储到文本缓冲区
                string text = Encoding.GetString(data);
                _textBuffer.Append(text);

                // 限制文本缓冲区长度
                if (_textBuffer.Length > MaxBufferSize)
                {
                    _textBuffer.Remove(0, _textBuffer.Length - MaxBufferSize);
                }

                // 存储为数据包
                _dataPackets.Enqueue(data);

                // 限制队列大小
                while (_dataPackets.Count > 1000) // 最多保存1000个包
                {
                    _dataPackets.Dequeue();
                }
            }
        }

        #region 数据访问方法

        /// <summary>
        /// 获取所有原始数据（字节数组）
        /// </summary>
        public byte[] GetAllRawData()
        {
            lock (_lockObject)
            {
                return _rawDataBuffer.ToArray();
            }
        }

        /// <summary>
        /// 获取所有文本数据
        /// </summary>
        public string GetAllText()
        {
            lock (_lockObject)
            {
                return _textBuffer.ToString();
            }
        }

        /// <summary>
        /// 获取最近N个字节的数据
        /// </summary>
        public byte[] GetLastBytes(int count)
        {
            lock (_lockObject)
            {
                if (_rawDataBuffer.Count == 0) return new byte[0];

                int start = Math.Max(0, _rawDataBuffer.Count - count);
                int length = _rawDataBuffer.Count - start;

                byte[] result = new byte[length];
                _rawDataBuffer.CopyTo(start, result, 0, length);
                return result;
            }
        }

        /// <summary>
        /// 获取最近N个字符的文本
        /// </summary>
        public string GetLastText(int length)
        {
            lock (_lockObject)
            {
                if (_textBuffer.Length == 0) return "";

                int start = Math.Max(0, _textBuffer.Length - length);
                return _textBuffer.ToString(start, _textBuffer.Length - start);
            }
        }

        /// <summary>
        /// 获取所有数据包
        /// </summary>
        public List<byte[]> GetAllPackets()
        {
            lock (_lockObject)
            {
                return new List<byte[]>(_dataPackets);
            }
        }

        /// <summary>
        /// 获取下一个数据包（FIFO）
        /// </summary>
        public byte[] GetNextPacket()
        {
            lock (_lockObject)
            {
                return _dataPackets.Count > 0 ? _dataPackets.Dequeue() : null;
            }
        }

        /// <summary>
        /// 获取所有数据包并转换为文本
        /// </summary>
        public List<string> GetAllPacketsAsText()
        {
            lock (_lockObject)
            {
                var result = new List<string>();
                foreach (var packet in _dataPackets)
                {
                    result.Add(Encoding.GetString(packet));
                }
                return result;
            }
        }

        /// <summary>
        /// 查找包含特定文本的数据包
        /// </summary>
        public List<byte[]> FindPackets(string text)
        {
            lock (_lockObject)
            {
                var result = new List<byte[]>();
                byte[] searchBytes = Encoding.GetBytes(text);

                foreach (var packet in _dataPackets)
                {
                    if (ContainsBytes(packet, searchBytes))
                    {
                        result.Add(packet);
                    }
                }
                return result;
            }
        }

        /// <summary>
        /// 检查字节数组是否包含子数组
        /// </summary>
        private bool ContainsBytes(byte[] source, byte[] pattern)
        {
            if (pattern.Length > source.Length) return false;

            for (int i = 0; i <= source.Length - pattern.Length; i++)
            {
                bool found = true;
                for (int j = 0; j < pattern.Length; j++)
                {
                    if (source[i + j] != pattern[j])
                    {
                        found = false;
                        break;
                    }
                }
                if (found) return true;
            }
            return false;
        }

        /// <summary>
        /// 等待特定数据出现（超时可取消）
        /// </summary>
        public async Task<bool> WaitForDataAsync(string expectedText, int timeoutMs, CancellationToken cancellationToken = default)
        {
            byte[] expectedBytes = Encoding.GetBytes(expectedText);
            return await WaitForDataAsync(expectedBytes, timeoutMs, cancellationToken);
        }

        /// <summary>
        /// 等待特定数据出现（超时可取消）
        /// </summary>
        public async Task<bool> WaitForDataAsync(byte[] expectedBytes, int timeoutMs, CancellationToken cancellationToken = default)
        {
            using var cts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
            cts.CancelAfter(timeoutMs);

            try
            {
                while (!cts.Token.IsCancellationRequested)
                {
                    lock (_lockObject)
                    {
                        byte[] allData = _rawDataBuffer.ToArray();
                        if (ContainsBytes(allData, expectedBytes))
                        {
                            return true;
                        }
                    }

                    await Task.Delay(10, cts.Token);
                }
            }
            catch (OperationCanceledException)
            {
                // 超时或取消
            }

            return false;
        }

        #endregion

        #region 数据清除方法

        /// <summary>
        /// 清除所有数据
        /// </summary>
        public void ClearAll()
        {
            lock (_lockObject)
            {
                _rawDataBuffer.Clear();
                _textBuffer.Clear();
                _dataPackets.Clear();
            }
            OnDataReceived(new DataReceivedEventArgs("数据已清除", DataType.System));
        }

        /// <summary>
        /// 清除原始数据缓冲区
        /// </summary>
        public void ClearRawData()
        {
            lock (_lockObject)
            {
                _rawDataBuffer.Clear();
            }
        }

        /// <summary>
        /// 清除文本缓冲区
        /// </summary>
        public void ClearText()
        {
            lock (_lockObject)
            {
                _textBuffer.Clear();
            }
        }

        /// <summary>
        /// 清除数据包队列
        /// </summary>
        public void ClearPackets()
        {
            lock (_lockObject)
            {
                _dataPackets.Clear();
            }
        }

        /// <summary>
        /// 清除指定时间之前的数据（基于索引）
        /// </summary>
        public void ClearBefore(DateTime time)
        {
            // 由于我们没有时间戳，这里提供一个基于包数量的清除
            lock (_lockObject)
            {
                if (_dataPackets.Count > 100)
                {
                    int removeCount = _dataPackets.Count / 2;
                    for (int i = 0; i < removeCount; i++)
                    {
                        _dataPackets.Dequeue();
                    }
                }
            }
        }

        #endregion

        #region 事件触发

        protected virtual void OnDataReceived(DataReceivedEventArgs args)
        {
            DataReceived?.Invoke(this, args);
        }

        protected virtual void OnErrorOccurred(ErrorEventArgs args)
        {
            ErrorOccurred?.Invoke(this, args);
        }

        #endregion

        #region IDisposable

        public void Dispose()
        {
            Stop();
            ClearAll();
        }

        #endregion
    }

    /// <summary>
    /// 数据类型枚举
    /// </summary>
    public enum DataType
    {
        Normal,     // 普通数据
        Command,    // 命令
        Response,   // 响应
        Error,      // 错误
        System      // 系统消息
    }

    /// <summary>
    /// 数据接收事件参数
    /// </summary>
    public class DataReceivedEventArgs : EventArgs
    {
        public byte[] RawData { get; }
        public string Text { get; }
        public DataType DataType { get; }
        public DateTime Timestamp { get; }

        public DataReceivedEventArgs(byte[] data, Encoding encoding)
        {
            RawData = data;
            Text = encoding.GetString(data);
            DataType = DataType.Normal;
            Timestamp = DateTime.Now;
        }

        public DataReceivedEventArgs(string text, DataType type)
        {
            Text = text;
            RawData = Encoding.ASCII.GetBytes(text);
            DataType = type;
            Timestamp = DateTime.Now;
        }

        public string GetHexString()
        {
            return BitConverter.ToString(RawData).Replace("-", " ");
        }

        public string GetBinaryString()
        {
            StringBuilder sb = new StringBuilder();
            foreach (byte b in RawData)
            {
                sb.Append(Convert.ToString(b, 2).PadLeft(8, '0') + " ");
            }
            return sb.ToString().Trim();
        }
    }

    /// <summary>
    /// 错误事件参数
    /// </summary>
    public class ErrorEventArgs : EventArgs
    {
        public Exception Exception { get; }
        public string Message { get; }

        public ErrorEventArgs(Exception ex)
        {
            Exception = ex;
            Message = ex.Message;
        }
    }
}