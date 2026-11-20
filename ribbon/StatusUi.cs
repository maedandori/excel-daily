using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Web.Script.Serialization;
using System.Windows.Forms;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;

namespace ExcelDailyRibbon
{
    // No modal windows: all operation feedback is written to Excel's
    // status bar. Keep this polling on Excel's UI thread (WinForms Timer),
    // never touch Excel COM from an Agent/background thread.
    public sealed class StatusUi : IExcelAddIn
    {
        private static Timer _timer;
        private static IRibbonUI _ribbon;
        private static string _directory;
        private static string _lastShown = "";
        private static bool _busy;
        private static bool _invalidatePending;
        private static DateTime _loadedUtc;
        private static DateTime _submittedUtc;
        private static string _lastCompletion = "";
        private static readonly JavaScriptSerializer Serializer = new JavaScriptSerializer();

        public static bool IsBusy
        {
            get
            {
                EnsureDirectory();
                if (File.Exists(Path.Combine(_directory, "excel-daily-command.json")) ||
                    File.Exists(Path.Combine(_directory, "excel-daily-busy.lock"))) return true;
                if (_submittedUtc == DateTime.MinValue) return false;
                AgentStatus latest = Read();
                if (latest != null && (latest.state == "success" || latest.state == "error"))
                {
                    try
                    {
                        if (DateTimeOffset.FromUnixTimeMilliseconds(latest.updated_ms).UtcDateTime >=
                            _submittedUtc)
                            return false;
                    }
                    catch { }
                }
                return (DateTime.UtcNow - _submittedUtc).TotalSeconds < 5;
            }
        }

        public void AutoOpen() { Initialize(); }

        public void AutoClose()
        {
            if (_timer != null) { _timer.Stop(); _timer.Dispose(); _timer = null; }
            try { ((dynamic)ExcelDnaUtil.Application).StatusBar = false; } catch { }
            _ribbon = null;
        }

        public static void SetRibbon(IRibbonUI ui)
        {
            _ribbon = ui;
            Initialize();
            _invalidatePending = true;
            Poll();
        }

        public static bool IsCommandEnabled(IRibbonControl _) { return !IsBusy; }

        private static void EnsureDirectory()
        {
            if (_directory == null)
                _directory = Path.GetDirectoryName(ExcelDnaUtil.XllPath) ?? ".";
        }

        private static void Initialize()
        {
            EnsureDirectory();
            if (_timer != null) return;
            _loadedUtc = DateTime.UtcNow;
            _busy = IsBusy;
            _timer = new Timer { Interval = 600 };
            _timer.Tick += (sender, args) => Poll();
            _timer.Start();
        }

        public static bool Guard()
        {
            if (!IsBusy) return true;
            Show("Excel Daily: 更新中 — các nút tạm khóa; không gửi yêu cầu trùng.");
            return false;
        }

        public static void Submitted(string fileName)
        {
            _submittedUtc = DateTime.UtcNow;
            _invalidatePending = true;
            Show("Excel Daily: 更新中 — đang cập nhật " + fileName);
            Poll();
        }

        public static void ShowError(Exception ex)
        {
            Show("Excel Daily: エラー — " + ex.Message);
            try
            {
                EnsureDirectory();
                File.AppendAllText(Path.Combine(_directory, "excel-daily-ribbon.log"),
                    DateTime.Now.ToString("s") + " ERROR " + ex + Environment.NewLine,
                    Encoding.UTF8);
            }
            catch { }
        }

        public static void Show(string message)
        {
            string display = (message ?? "").Replace("\r", " ").Replace("\n", " ").Trim();
            if (display.Length > 220) display = display.Substring(0, 217) + "...";
            try
            {
                dynamic app = ExcelDnaUtil.Application;
                app.StatusBar = display;
                _lastShown = display;
            }
            catch (COMException) { }
            catch (InvalidOperationException) { }
        }

        public sealed class AgentStatus
        {
            public string state { get; set; }
            public string message { get; set; }
            public long updated_ms { get; set; }
        }

        private static AgentStatus Read()
        {
            try
            {
                EnsureDirectory();
                string path = Path.Combine(_directory, "excel-daily-status.json");
                if (!File.Exists(path)) return null;
                return Serializer.Deserialize<AgentStatus>(File.ReadAllText(path, Encoding.UTF8));
            }
            catch (Exception) { return null; } // half-write/stale file: retry on next Tick
        }

        private static void Poll()
        {
            try
            {
                bool busy = IsBusy;
                if (_busy != busy)
                {
                    _busy = busy;
                    _invalidatePending = true;
                }
                if (_invalidatePending && _ribbon != null)
                {
                    try { _ribbon.Invalidate(); _invalidatePending = false; }
                    catch (COMException) { }
                }

                AgentStatus status = Read();
                if (status == null || String.IsNullOrWhiteSpace(status.message)) return;
                string statusMessage = "Excel Daily: " + status.message;
                if (busy)
                {
                    // Ignore the previous task's completed status while the
                    // next request is queued but not yet picked up by Agent.
                    if (status.state == "running" &&
                        !String.Equals(_lastShown, statusMessage, StringComparison.Ordinal))
                        Show(statusMessage);
                    return;
                }

                // Don't replay yesterday's status when Excel opens. Show only
                // this session's recent completion/error, exactly once.
                DateTime changed;
                try
                {
                    changed = DateTimeOffset.FromUnixTimeMilliseconds(status.updated_ms).UtcDateTime;
                }
                catch { return; }
                if (changed < _loadedUtc.AddSeconds(-1)) return;
                if (status.state != "success" && status.state != "error") return;
                string completedKey = status.updated_ms + ":" + status.state + ":" + status.message;
                if (_lastCompletion != completedKey)
                {
                    _lastCompletion = completedKey;
                    _submittedUtc = DateTime.MinValue;
                    _invalidatePending = true;
                    Show(statusMessage);
                }
            }
            catch (Exception) { /* Excel busy/shutting down: poll again. */ }
        }
    }
}
