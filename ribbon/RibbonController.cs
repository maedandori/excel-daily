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
            return @"<customUI xmlns='http://schemas.microsoft.com/office/2006/01/customui'>
              <ribbon><tabs>
                <tab id='excelDaily' label='EXCEL-DAILY'>
                  <group id='updates' label='MICS 予定 / 更新'>
                    <button id='refreshData' label='データ更新' size='large'
                      imageMso='RefreshAll' onAction='OnUpdate'/>
                    <button id='refreshTime' label='時間更新' size='large'
                      imageMso='RefreshAll' onAction='OnUpdate'/>
                    <button id='tomorrow' label='翌日予定' size='large'
                      imageMso='RefreshAll' onAction='OnUpdate'/>
                  </group>
                  <group id='sheetButtons' label='シート上のボタン'>
                    <button id='installButtons' label='シートにボタン追加'
                      imageMso='InsertShapes' onAction='OnUpdate'/>
                    <button id='repairButtons' label='ボタンを再設定'
                      imageMso='RefreshAll' onAction='OnUpdate'/>
                  </group>
                </tab>
              </tabs></ribbon>
            </customUI>";
        }

        public void OnUpdate(IRibbonControl control)
        {
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
                        QueueAction("all", ActiveWorkbookPath());
                        return;
                    case "refreshTime":
                        QueueAction("refresh_times", ActiveWorkbookPath());
                        return;
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
            try { QueueAction("all", ClickedShapeWorkbookPath("data")); }
            catch (Exception ex) { DisplayError(ex); }
        }

        [ExcelCommand(Name = SheetButtonContract.TimeCommand, Description = "Excel Daily update pending times")]
        public static void RefreshClickedTimes()
        {
            try { QueueAction("refresh_times", ClickedShapeWorkbookPath("time")); }
            catch (Exception ex) { DisplayError(ex); }
        }

        private static string ActiveWorkbookPath()
        {
            dynamic app = ExcelDnaUtil.Application;
            dynamic book = app.ActiveWorkbook;
            if (book == null) throw new InvalidOperationException("Hãy mở file Excel cần cập nhật.");
            return SheetButtonContract.RequireSavedWorkbook(Convert.ToString(book.FullName));
        }

        private static string ClickedShapeWorkbookPath(string kind)
        {
            dynamic app = ExcelDnaUtil.Application;
            object caller = app.Caller;
            string name = caller as string;
            if (String.IsNullOrWhiteSpace(name))
                throw new InvalidOperationException("Chức năng chỉ chạy từ nút Shape trong worksheet.");
            dynamic sheet = app.ActiveSheet;
            dynamic shape;
            try { shape = sheet.Shapes.Item(name); }
            catch (COMException)
            {
                throw new InvalidOperationException("Không tìm thấy nút trong sheet đang mở.");
            }
            if (!SheetButtonContract.IsKind(
                Convert.ToString(shape.Name), Convert.ToString(shape.AlternativeText), kind))
                throw new InvalidOperationException("Nút không thuộc loại cập nhật này. Hãy chọn ボタンを再設定.");
            return SheetButtonContract.RequireSavedWorkbook(Convert.ToString(sheet.Parent.FullName));
        }

        internal static void QueueAction(string action, string workbookFullName)
        {
            string dir = Path.GetDirectoryName(ExcelDnaUtil.XllPath);
            if (String.IsNullOrWhiteSpace(dir))
                throw new InvalidOperationException("Không xác định được thư mục Add-in.");
            string agent = Path.Combine(dir, "excel-daily.exe");
            if (!File.Exists(agent) || !File.Exists(Path.Combine(dir, "settings.json")))
                throw new FileNotFoundException("Thiếu Excel Daily Agent/settings.json. Cài bộ ZIP đầy đủ.");
            // Fail closed: the 0.3.5 agent would ignore the workbook target and
            // might update another open file. Never silently use that old agent.
            string version = FileVersionInfo.GetVersionInfo(agent).FileVersion;
            Version parsed;
            if (!Version.TryParse(version, out parsed) || parsed < new Version(0, 3, 6))
                throw new InvalidOperationException("Excel Daily Agent phải là v0.3.6 trở lên.");

            workbookFullName = SheetButtonContract.RequireSavedWorkbook(workbookFullName);
            string queue = Path.Combine(dir, "excel-daily-command.json");
            if (File.Exists(queue) || File.Exists(Path.Combine(dir, "excel-daily-busy.lock")))
                throw new InvalidOperationException("Excel Daily đang cập nhật hoặc đã có yêu cầu chờ xử lý.");

            // Use the JSON serializer to preserve Japanese / UNC paths and
            // avoid hand-built strings. The queue file is published atomically.
            var json = new JavaScriptSerializer().Serialize(new Dictionary<string, string>
            {
                { "action", action }, { "workbook_full_name", workbookFullName }
            });
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
            dynamic excel = ExcelDnaUtil.Application;
            excel.StatusBar = "Excel Daily: cập nhật " + Path.GetFileName(workbookFullName) +
                " đã được yêu cầu; kiểm tra excel-daily.log để biết kết quả.";
        }

        internal static void DisplayError(Exception ex)
        {
            MessageBox.Show(ex.Message, "Excel Daily", MessageBoxButtons.OK, MessageBoxIcon.Warning);
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
                throw new InvalidOperationException("Sheet bị khóa hình vẽ; hãy mở bảo vệ trước.");
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
                throw new InvalidOperationException("Sheet này chưa có nút Excel Daily để gắn lại.");
            excel.StatusBar = "Excel Daily: đã gắn lệnh cho " +
                dataCount + " nút dữ liệu và " + timeCount + " nút thời gian.";
        }
    }
}
