using System;
using System.IO;
using ExcelDailyRibbon;

internal static class Program
{
    private static int _checks;
    private static void Check(bool assertion, string message)
    {
        _checks++;
        if (!assertion) throw new Exception("FAILED: " + message);
    }

    private static void Throws(Action action, string message)
    {
        bool thrown = false;
        try { action(); }
        catch (InvalidOperationException) { thrown = true; }
        Check(thrown, message);
    }

    public static int Main()
    {
        try
        {
            Check(SheetButtonContract.Command("data") == "ExcelDaily_Refresh", "data callback stable");
            Check(SheetButtonContract.Command("time") == "ExcelDaily_RefreshTimes", "time callback stable");
            Check(SheetButtonContract.IsKind("ED_BTN_DATA_1234", null, "data"), "copied data shape without alt");
            Check(SheetButtonContract.IsKind("Shape 4", SheetButtonContract.DataTag, "data"), "copied shape with renamed name");
            Check(!SheetButtonContract.IsKind("ED_BTN_DATA_1234", null, "time"), "data cannot call time command");
            Check(SheetButtonContract.IsKind("ED_BTN_TIME_8888", null, "time"), "copied time button");
            Check(!SheetButtonContract.IsKind("Button 1", "untrusted", "data"), "unrelated shape ignored");
            Check(SheetButtonContract.Tag("data") != SheetButtonContract.Tag("time"), "distinct markers");
            string temp = Path.Combine(Path.GetTempPath(), "excel-daily-contract-" + Guid.NewGuid().ToString("N") + ".xlsx");
            try
            {
                File.WriteAllText(temp, "test");
                Check(String.Equals(SheetButtonContract.RequireSavedWorkbook(temp),
                    Path.GetFullPath(temp), StringComparison.OrdinalIgnoreCase), "full path preserved");
                Throws(() => SheetButtonContract.RequireSavedWorkbook("relative.xlsx"), "reject relative workbook");
                Throws(() => SheetButtonContract.RequireSavedWorkbook(null), "reject missing workbook");
            }
            finally { File.Delete(temp); }
            var options = AutoSaveOptions.Parse("{\"workbook\":{\"filename_pattern\":\"成形スケジュール*.xlsx\"}}");
            Check(options.Enabled && options.Minutes == 15, "legacy config defaults enabled 15 min");
            Check(options.ShouldSave("成形スケジュール20261125.xlsx", false, false, false, true), "dirty matching workbook");
            Check(!options.ShouldSave("other.xlsx", false, false, false, true), "unrelated workbook");
            Check(!options.ShouldSave("成形スケジュール.xlsx", true, false, false, true), "unchanged workbook");
            Check(!options.ShouldSave("成形スケジュール.xlsx", false, true, false, true), "read-only workbook");
            Check(!options.ShouldSave("成形スケジュール.xlsx", false, false, true, true), "shared workbook");
            Check(!options.ShouldSave("成形スケジュール.xlsx", false, false, false, false), "unsaved or missing file");
            Check(!options.Matches("成形スケジュール.xlsm"), "extension restricted");
            Check(!AutoSaveOptions.Parse("{\"workbook\":{\"filename_pattern\":\"*.xlsx\"}}").Matches("other.xlsx"), "no prefix rejected");
            Check(!AutoSaveOptions.Parse("{\"workbook\":{\"filename_pattern\":\"\"}}").Matches("other.xlsx"), "empty pattern rejected");
            var off = AutoSaveOptions.Parse("{\"auto_save\":{\"enabled\":false,\"interval_minutes\":30},\"workbook\":{\"filename_pattern\":\"成形*.xlsx\"}}");
            Check(!off.Matches("成形スケジュール.xlsx") && off.Minutes == 30, "disabled");
            var schedule = new AutoSaveSchedule();
            var start = new DateTime(2026, 11, 25, 7, 0, 0, DateTimeKind.Utc);
            Check(!schedule.Due(options, start), "not immediately");
            Check(!schedule.Due(options, start.AddMinutes(14)), "not early");
            Check(schedule.Due(options, start.AddMinutes(15)), "due at 15 minutes");
            Check(schedule.Due(options, start.AddMinutes(16)), "busy deferral retains due");
            schedule.Completed(start.AddMinutes(16));
            Check(!schedule.Due(options, start.AddMinutes(30)), "next cycle not early");
            Check(schedule.Due(options, start.AddMinutes(31)), "next cycle due");
            Check(!schedule.Due(off, start.AddMinutes(32)), "disabled resets schedule");
            Check(!schedule.Due(options, start.AddMinutes(33)), "re-enable starts new interval");
            Console.WriteLine("RIBBON_CONTRACT_TESTS_OK checks=" + _checks);
            return 0;
        }
        catch (Exception ex) { Console.Error.WriteLine(ex); return 1; }
    }
}
