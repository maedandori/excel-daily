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
            Console.WriteLine("RIBBON_CONTRACT_TESTS_OK checks=" + _checks);
            return 0;
        }
        catch (Exception ex) { Console.Error.WriteLine(ex); return 1; }
    }
}
