using System;
using System.IO;
using System.Text;
using System.Windows.Forms;
using ExcelDna.Integration;

namespace ExcelDailyRibbon
{
    internal static class WorkbookAutoSave
    {
        private static Timer _timer;
        private static string _directory;
        private static bool _queued;
        private static int _generation;
        private static readonly AutoSaveSchedule Schedule = new AutoSaveSchedule();

        public static void Start(string directory)
        {
            if (_timer != null) return;
            _directory = directory;
            Schedule.Reset();
            _timer = new Timer { Interval = 5000 };
            _timer.Tick += (sender, args) => Tick();
            _timer.Start();
            Tick();
        }

        public static void Stop()
        {
            _generation++;
            if (_timer != null) { _timer.Stop(); _timer.Dispose(); _timer = null; }
            _queued = false;
            Schedule.Reset();
        }

        private static AutoSaveOptions ReadOptions()
        {
            return AutoSaveOptions.Parse(File.ReadAllText(Path.Combine(_directory, "settings.json"), Encoding.UTF8));
        }

        private static void Tick()
        {
            if (_queued || _timer == null) return;
            try
            {
                var options = ReadOptions();
                if (!Schedule.Due(options, DateTime.UtcNow) || StatusUi.IsBusy) return;
                _queued = true;
                int generation = _generation;
                // Excel-DNA executes macros only when Excel is able to accept one;
                // never use Excel COM from a background timer or while editing a cell.
                ExcelAsyncUtil.QueueAsMacro(() =>
                {
                    if (generation != _generation || _timer == null) return;
                    try { SaveDue(); }
                    catch (Exception ex) { StatusUi.ShowError(ex); Schedule.Completed(DateTime.UtcNow); }
                    finally { _queued = false; }
                });
            }
            catch { _queued = false; Schedule.Reset(); } // invalid/mid-write settings: never save
        }

        private static void SaveDue()
        {
            AutoSaveOptions options;
            try { options = ReadOptions(); }
            catch { Schedule.Reset(); return; }
            if (!Schedule.Due(options, DateTime.UtcNow) || StatusUi.IsBusy) return;
            dynamic app = ExcelDnaUtil.Application;
            if (!Convert.ToBoolean(app.Ready)) return;
            dynamic books = app.Workbooks;
            int count = Convert.ToInt32(books.Count);
            int saved = 0, failed = 0;
            for (int i = 1; i <= count; i++)
            {
                if (StatusUi.IsBusy) return;
                try
                {
                    dynamic book = books.Item(i);
                    string name = Convert.ToString(book.Name);
                    if (!options.Matches(name)) continue;
                    string path = Convert.ToString(book.FullName);
                    if (!options.ShouldSave(name, Convert.ToBoolean(book.Saved),
                        Convert.ToBoolean(book.ReadOnly), Convert.ToBoolean(book.MultiUserEditing),
                        Path.IsPathRooted(path) && File.Exists(path))) continue;
                    if (StatusUi.IsBusy) return;
                    // Normal Excel Save, not SaveAs. Preserve native conflict handling.
                    book.Save();
                    if (Convert.ToBoolean(book.Saved)) saved++;
                    else failed++;
                }
                catch (Exception ex) { failed++; StatusUi.ShowError(ex); }
            }
            Schedule.Completed(DateTime.UtcNow);
            if (failed > 0) StatusUi.Show("Excel Daily: 自動保存 " + saved + "件、失敗 " + failed + "件。");
            else if (saved > 0) StatusUi.Show("Excel Daily: 自動保存完了 " + saved + "件。");
        }
    }
}
