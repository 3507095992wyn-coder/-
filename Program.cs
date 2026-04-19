using System;
using System.IO;
using System.Text;
using System.Threading;
using System.Windows.Forms;

namespace grbloxy
{
    internal static class Program
    {
        [STAThread]
        private static void Main()
        {
            Application.ThreadException += Application_ThreadException;
            AppDomain.CurrentDomain.UnhandledException += CurrentDomain_UnhandledException;

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Form1());
        }

        private static void Application_ThreadException(object sender, ThreadExceptionEventArgs e)
        {
            HandleUnhandledException(e.Exception, "UI 线程异常");
        }

        private static void CurrentDomain_UnhandledException(object sender, UnhandledExceptionEventArgs e)
        {
            HandleUnhandledException(e.ExceptionObject as Exception, "非 UI 线程异常");
        }

        private static void HandleUnhandledException(Exception exception, string title)
        {
            string details = BuildExceptionDetails(exception);
            string logPath = WriteCrashLog(details);

            MessageBox.Show(
                $"{title}\r\n\r\n{details}\r\n\r\n日志已保存到：\r\n{logPath}",
                "程序异常",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
        }

        private static string BuildExceptionDetails(Exception exception)
        {
            if (exception == null)
            {
                return "未捕获到具体异常对象。";
            }

            StringBuilder builder = new StringBuilder();
            int level = 0;
            Exception current = exception;

            while (current != null)
            {
                builder.AppendLine($"[{level}] {current.GetType().FullName}");
                builder.AppendLine(current.Message);

                if (!string.IsNullOrWhiteSpace(current.StackTrace))
                {
                    builder.AppendLine(current.StackTrace);
                }

                current = current.InnerException;
                level++;

                if (current != null)
                {
                    builder.AppendLine();
                    builder.AppendLine("---- Inner Exception ----");
                }
            }

            return builder.ToString();
        }

        private static string WriteCrashLog(string details)
        {
            string folder = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Logs");
            Directory.CreateDirectory(folder);

            string path = Path.Combine(folder, $"crash_{DateTime.Now:yyyyMMdd_HHmmss}.log");
            File.WriteAllText(path, details ?? string.Empty, Encoding.UTF8);
            return path;
        }
    }
}
