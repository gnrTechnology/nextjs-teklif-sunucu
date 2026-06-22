using System;
using System.Collections.Generic;
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
                return "ERR:Excel calismiyor";

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

                    var macros = BuildMacroCandidates(xl);
                    Exception lastEx = null;

                    foreach (var macro in macros)
                    {
                        try
                        {
                            AgentLog.Info("Run deneniyor: " + macro);
                            if (string.IsNullOrEmpty(param))
                                xl.Run(macro, moduleName);
                            else
                                xl.Run(macro, moduleName, param);

                            return "OK:" + macro;
                        }
                        catch (Exception ex)
                        {
                            lastEx = ex;
                            AgentLog.Error("Run basarisiz " + macro + ": " + ex.Message);
                        }
                    }

                    if (lastEx != null)
                        return "ERR:" + lastEx.Message;
                    return "ERR:zInternet.RunRemoteCode bulunamadi (teklif.xlam yuklu mu?)";
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

        private static List<string> BuildMacroCandidates(dynamic xl)
        {
            var list = new List<string>();
            list.Add("zInternet.RunRemoteCode");

            try
            {
                foreach (dynamic wb in xl.Workbooks)
                {
                    try
                    {
                        if (wb.IsAddin)
                        {
                            var q = "'" + wb.Name + "'!zInternet.RunRemoteCode";
                            if (!list.Contains(q)) list.Add(q);
                        }
                    }
                    catch { }
                }
            }
            catch { }

            try
            {
                foreach (dynamic ai in xl.AddIns)
                {
                    try
                    {
                        if (ai.Installed)
                        {
                            var n = (string)ai.Name;
                            if (n.IndexOf("teklif", StringComparison.OrdinalIgnoreCase) >= 0)
                            {
                                var q = "'" + n + "'!zInternet.RunRemoteCode";
                                if (!list.Contains(q)) list.Add(q);
                            }
                        }
                    }
                    catch { }
                }
            }
            catch { }

            list.Add("'teklif.xlam'!zInternet.RunRemoteCode");
            return list;
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
