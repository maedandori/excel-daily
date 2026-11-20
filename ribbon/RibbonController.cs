using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;

namespace ExcelDailyRibbon
{
    // No VBA and no direct Turso credential. Existing Excel Daily Agent calls
    // the read-only Cloudflare /api/excel-daily/events endpoint over HTTPS.
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
                </tab>
              </tabs></ribbon>
            </customUI>";
        }

        public void OnUpdate(IRibbonControl control)
        {
            string action;
            switch (control.Id)
            {
                case "refreshData": action = "all"; break;
                case "refreshTime": action = "refresh_times"; break;
                case "tomorrow": action = "tomorrow"; break;
                default: return;
            }

            try
            {
                string dir = Path.GetDirectoryName(ExcelDnaUtil.XllPath);
                if (String.IsNullOrWhiteSpace(dir))
                    throw new InvalidOperationException("Không xác định được thư mục Add-in.");
                string settings = Path.Combine(dir, "settings.json");
                string agent = Path.Combine(dir, "excel-daily.exe");
                if (!File.Exists(agent) || !File.Exists(settings))
                    throw new FileNotFoundException(
                        "Hãy đặt Add-in cùng thư mục excel-daily.exe và settings.json.");

                string commandPath = Path.Combine(dir, "excel-daily-command.json");
                string busyPath = Path.Combine(dir, "excel-daily-busy.lock");
                if (File.Exists(commandPath) || File.Exists(busyPath))
                    throw new InvalidOperationException(
                        "Excel Daily đang cập nhật hoặc đã có yêu cầu chờ xử lý.");

                // Atomic publish: the Agent cannot consume a half-written JSON command.
                string temp = Path.Combine(dir, "excel-daily-command-" +
                    Guid.NewGuid().ToString("N") + ".tmp");
                try
                {
                    File.WriteAllText(temp, "{\"action\":\"" + action + "\"}",
                        new UTF8Encoding(false));
                    File.Move(temp, commandPath);
                }
                finally
                {
                    if (File.Exists(temp)) File.Delete(temp);
                }

                // Agent might not be running yet, e.g. immediately after Excel starts.
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
                        // Avoid replaying a queued request after an unrelated restart.
                        try { File.Delete(commandPath); } catch { }
                        throw;
                    }
                }
                dynamic excel = ExcelDnaUtil.Application;
                excel.StatusBar = "Excel Daily: yêu cầu " + control.Id +
                    " đã gửi. Kết quả ghi trong excel-daily.log.";
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Excel Daily - Ribbon",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private static bool IsAgentRunning(string executable)
        {
            foreach (Process process in Process.GetProcessesByName("excel-daily"))
            {
                try
                {
                    if (!process.HasExited &&
                        String.Equals(Path.GetFullPath(process.MainModule.FileName),
                            Path.GetFullPath(executable),
                            StringComparison.OrdinalIgnoreCase))
                        return true;
                }
                catch (InvalidOperationException) { }
                catch (System.ComponentModel.Win32Exception) { }
                finally { process.Dispose(); }
            }
            return false;
        }
    }
}
