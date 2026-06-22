using System;
using System.Runtime.InteropServices;

namespace TeklifAgent
{
    public static class ExcelRunner
    {
        public static bool IsExcelRunning()
        {
            return System.Diagnostics.Process.GetProcessesByName("EXCEL").Length > 0;
        }

        public static string RunRemoteModule(string moduleName, string param)
        {
            if (!IsExcelRunning())
                return "ERR:Excel çalışmıyor";

            object excelObj = null;
            try
            {
                excelObj = Marshal.GetActiveObject("Excel.Application");
                dynamic xl = excelObj;

                bool prevAlerts = xl.DisplayAlerts;
                bool prevScreen = xl.ScreenUpdating;
                xl.DisplayAlerts = false;
                xl.ScreenUpdating = false;

                try
                {
                    EnsureHostWorkbook(xl);

                    if (string.IsNullOrEmpty(param))
                        xl.Run("zInternet.RunRemoteCode", moduleName);
                    else
                        xl.Run("zInternet.RunRemoteCode", moduleName, param);

                    return "OK";
                }
                finally
                {
                    xl.DisplayAlerts = prevAlerts;
                    xl.ScreenUpdating = prevScreen;
                }
            }
            catch (Exception ex)
            {
                return "ERR:" + ex.Message;
            }
            finally
            {
                if (excelObj != null)
                    Marshal.ReleaseComObject(excelObj);
            }
        }

        private static void EnsureHostWorkbook(dynamic xl)
        {
            bool hasWb = false;
            foreach (dynamic wb in xl.Workbooks)
            {
                try
                {
                    if (!wb.IsAddin)
                    {
                        hasWb = true;
                        break;
                    }
                }
                catch { }
            }
            if (!hasWb)
                xl.Workbooks.Add();
        }
    }
}
