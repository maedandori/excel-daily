using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Web.Script.Serialization;
using System.Windows.Forms;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;

namespace ExcelDailyRibbon
{
    [ComVisible(true)]
    public class RibbonController : ExcelRibbon
    {
        public override string GetCustomUI(string ribbonId)
        {
            return @"<customUI xmlns='http://schemas.microsoft.com/office/2006/01/customui' onLoad='RibbonLoaded'>
              <ribbon><tabs>
                <tab id='excelDaily' label='データ更新'>
                  <group id='updates' label='MICS 予定 / 更新'>
                    <button id='refreshData' label='データ更新' size='large'
                      imageMso='RefreshAll' getEnabled='IsCommandEnabled' onAction='OnUpdate'/>
                    <button id='refreshTime' label='時間更新' size='large'
                      imageMso='RefreshAll' getEnabled='IsCommandEnabled' onAction='OnUpdate'/>
                    <button id='tomorrow' label='翌日予定' size='large'
                      imageMso='RefreshAll' getEnabled='IsCommandEnabled' onAction='OnUpdate'/>
                  </group>
                  <group id='sheetButtons' label='シート上のボタン'>
                    <button id='installButtons' label='シートにボタン追加'
                      getEnabled='IsCommandEnabled' onAction='OnUpdate'/>
                    <button id='repairButtons' label='ボタンを再設定'
                      imageMso='RefreshAll' getEnabled='IsCommandEnabled' onAction='OnUpdate'/>
                  </group>
                </tab>
              </tabs></ribbon>
            </customUI>";
        }

        public void RibbonLoaded(IRibbonUI ribbon) { StatusUi.SetRibbon(ribbon); }

        public bool IsCommandEnabled(IRibbonControl control) { return StatusUi.IsCommandEnabled(control); }

        public void OnUpdate(IRibbonControl control)
        {
            if (!StatusUi.Guard()) return;
            try
            {
                switch (control.Id)
                {
                    case "installButtons":
                        SheetButtons.InstallOrRepair(false);
                        return;
                    case "repairButtons":
                        SheetButtons.InstallOrRepair(true);
                        return;
                    case "refreshData":
                    {
                        var target = ActiveSheetTarget();
                        QueueAction("all", target.Item1, target.Item2);
                        return;
                    }
                    case "refreshTime":
                    {
                        var target = ActiveSheetTarget();
                        QueueAction("refresh_times", target.Item1, target.Item2);
                        return;
                    }
                    case "tomorrow":
                        QueueAction("tomorrow", ActiveWorkbookPath());
                        return;
                }
            }
            catch (Exception ex)
            {
                DisplayError(ex);
            }
        }

        // Name-only callback is embedded into a worksheet Shape's OnAction.
        // It stays portable across workbook copies and machines with this XLL.
        [ExcelCommand(Name = SheetButtonContract.DataCommand, Description = "Excel Daily refresh selected workbook")]
        public static void RefreshClickedSheet()
        {
            if (!StatusUi.Guard()) return;
            try {
                var target = ClickedShapeTarget("data");
                QueueAction("all", target.Item1, target.Item2);
            }
            catch (Exception ex) { DisplayError(ex); }
        }

        [ExcelCommand(Name = SheetButtonContract.TimeCommand, Description = "Excel Daily update pending times")]
        public static void RefreshClickedTimes()
        {
            if (!StatusUi.Guard()) return;
            try {
                var target = ClickedShapeTarget("time");
                QueueAction("refresh_times", target.Item1, target.Item2);
            }
            catch (Exception ex) { DisplayError(ex); }
        }

        private static string ActiveWorkbookPath()
        {
            dynamic app = ExcelDnaUtil.Application;
            dynamic book = app.ActiveWorkbook;
            if (book == null) throw new InvalidOperationException("更新するExcelファイルを開いてください。");
            return SheetButtonContract.RequireSavedWorkbook(Convert.ToString(book.FullName));
        }

        private static Tuple<string, string> ActiveSheetTarget()
        {
            dynamic app = ExcelDnaUtil.Application;
            dynamic sheet = app.ActiveSheet;
            if (sheet == null)
                throw new InvalidOperationException("更新するシートを選択してください。");
            string name = Convert.ToString(sheet.Name);
            if (String.IsNullOrWhiteSpace(name))
                throw new InvalidOperationException("シート名が取得できません。");
            string path = SheetButtonContract.RequireSavedWorkbook(Convert.ToString(sheet.Parent.FullName));
            return Tuple.Create(path, name);
        }

        private static Tuple<string, string> ClickedShapeTarget(string kind)
        {
            dynamic app = ExcelDnaUtil.Application;
            object caller = app.Caller;
            string name = caller as string;
            if (String.IsNullOrWhiteSpace(name))
                throw new InvalidOperationException("この操作はワークシート上のボタンから実行してください。");
            dynamic sheet = app.ActiveSheet;
            dynamic shape;
            try { shape = sheet.Shapes.Item(name); }
            catch (COMException)
            {
                throw new InvalidOperationException("現在のシートで対象ボタンが見つかりません。");
            }
            if (!SheetButtonContract.IsKind(
                Convert.ToString(shape.Name), Convert.ToString(shape.AlternativeText), kind))
                throw new InvalidOperationException("ボタンの種類が一致しません。「ボタンを再設定」を実行してください。");
            string path = SheetButtonContract.RequireSavedWorkbook(Convert.ToString(sheet.Parent.FullName));
            return Tuple.Create(path, Convert.ToString(sheet.Name));
        }

        internal static void QueueAction(string action, string workbookFullName, string worksheetName = null)
        {
            if (!StatusUi.Guard()) return;
            string dir = Path.GetDirectoryName(ExcelDnaUtil.XllPath);
            if (String.IsNullOrWhiteSpace(dir))
                throw new InvalidOperationException("アドインのインストール先が確認できません。");
            string agent = Path.Combine(dir, "excel-daily.exe");
            if (!File.Exists(agent) || !File.Exists(Path.Combine(dir, "settings.json")))
                throw new FileNotFoundException("Excel Daily本体またはsettings.jsonがありません。ZIP版をすべて展開してください。");
            // Older Agents might ignore the worksheet target and modify both
            // shifts. Reject mixed Agent/Add-in versions instead of guessing.
            string version = FileVersionInfo.GetVersionInfo(agent).FileVersion;
            Version parsed;
            if (!Version.TryParse(version, out parsed) || parsed.CompareTo(new Version(0, 3, 9)) < 0)
                throw new InvalidOperationException("Excel Daily本体はv0.3.9以降が必要です。");

            workbookFullName = SheetButtonContract.RequireSavedWorkbook(workbookFullName);
            string queue = Path.Combine(dir, "excel-daily-command.json");
            if (File.Exists(queue) || File.Exists(Path.Combine(dir, "excel-daily-busy.lock")))
                throw new InvalidOperationException("更新処理中です。完了するまでお待ちください。");

            // Use the JSON serializer to preserve Japanese / UNC paths and
            // avoid hand-built strings. The queue file is published atomically.
            var command = new Dictionary<string, string>
            {
                { "action", action }, { "workbook_full_name", workbookFullName }
            };
            if (!String.IsNullOrWhiteSpace(worksheetName))
            {
                if (worksheetName.Length > 31)
                    throw new InvalidOperationException("シート名が無効です。");
                command.Add("sheet_name", worksheetName);
            }
            var json = new JavaScriptSerializer().Serialize(command);
            string temp = Path.Combine(dir, "excel-daily-command-" +
                Guid.NewGuid().ToString("N") + ".tmp");
            try
            {
                File.WriteAllText(temp, json, new UTF8Encoding(false));
                File.Move(temp, queue);
            }
            finally { if (File.Exists(temp)) File.Delete(temp); }

            if (!IsAgentRunning(agent))
            {
                try
                {
                    Process.Start(new ProcessStartInfo(agent)
                    {
                        WorkingDirectory = dir,
                        UseShellExecute = false,
                        CreateNoWindow = true
                    });
                }
                catch
                {
                    // Never replay a stale command if starting the Agent failed.
                    try { File.Delete(queue); } catch { }
                    throw;
                }
            }
            StatusUi.Submitted(Path.GetFileName(workbookFullName));
        }

        internal static void DisplayError(Exception ex)
        {
            StatusUi.ShowError(ex);
        }

        private static bool IsAgentRunning(string executable)
        {
            foreach (Process process in Process.GetProcessesByName("excel-daily"))
            {
                try
                {
                    if (!process.HasExited &&
                        String.Equals(Path.GetFullPath(process.MainModule.FileName),
                            Path.GetFullPath(executable), StringComparison.OrdinalIgnoreCase))
                        return true;
                }
                catch (InvalidOperationException) { }
                catch (System.ComponentModel.Win32Exception) { }
                finally { process.Dispose(); }
            }
            return false;
        }
    }

    internal static class SheetButtons
    {
        // User may drag/resize the Shape after creation. The OnAction command
        // is independent of geometry and does not contain a local install path.
        private static dynamic Add(dynamic sheet, string kind, float left, float top)
        {
            dynamic shape = sheet.Shapes.AddShape(5, left, top, 140f, 34f);
            shape.Name = (kind == "data" ? SheetButtonContract.DataPrefix : SheetButtonContract.TimePrefix) +
                Guid.NewGuid().ToString("N").Substring(0, 8);
            shape.AlternativeText = SheetButtonContract.Tag(kind);
            shape.OnAction = SheetButtonContract.Command(kind);
            shape.TextFrame2.TextRange.Text = kind == "data" ? "データ更新" : "時間更新";
            try
            {
                shape.Placement = 2; // Move with cells, do not resize with cells
                shape.LockAspectRatio = 0;
                shape.Fill.ForeColor.RGB = kind == "data" ? 0x00529A48 : 0x008A7554;
                shape.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = 0x00FFFFFF;
                shape.TextFrame2.TextRange.Font.Size = 12;
            }
            catch (COMException) { /* Formatting is optional; callback isn't. */ }
            return shape;
        }

        public static void InstallOrRepair(bool repairOnly)
        {
            dynamic excel = ExcelDnaUtil.Application;
            dynamic sheet = excel.ActiveSheet;
            if (sheet == null || Convert.ToBoolean(sheet.ProtectDrawingObjects))
                throw new InvalidOperationException("シートが保護されているためボタンを編集できません。保護を解除してください。");
            int dataCount = 0;
            int timeCount = 0;
            dynamic shapes = sheet.Shapes;
            int existing = Convert.ToInt32(shapes.Count, CultureInfo.InvariantCulture);
            for (int i = 1; i <= existing; i++)
            {
                dynamic shape = shapes.Item(i);
                string name = Convert.ToString(shape.Name);
                string alt = Convert.ToString(shape.AlternativeText);
                if (SheetButtonContract.IsKind(name, alt, "data"))
                {
                    shape.OnAction = SheetButtonContract.DataCommand;
                    shape.AlternativeText = SheetButtonContract.DataTag;
                    dataCount++;
                }
                else if (SheetButtonContract.IsKind(name, alt, "time"))
                {
                    shape.OnAction = SheetButtonContract.TimeCommand;
                    shape.AlternativeText = SheetButtonContract.TimeTag;
                    timeCount++;
                }
            }
            if (!repairOnly)
            {
                dynamic cell = excel.ActiveCell;
                float left = Convert.ToSingle(cell.Left, CultureInfo.InvariantCulture);
                float top = Convert.ToSingle(cell.Top, CultureInfo.InvariantCulture);
                if (dataCount == 0) { Add(sheet, "data", left, top); dataCount++; }
                if (timeCount == 0) { Add(sheet, "time", left + 150f, top); timeCount++; }
            }
            if (dataCount + timeCount == 0)
                throw new InvalidOperationException("このシートに再設定できるExcel Dailyボタンがありません。");
            excel.StatusBar = "Excel Daily: ボタンの再設定が完了しました。データ更新: " +
                dataCount + "件、時刻更新: " + timeCount + "件。";
        }
    }
}
