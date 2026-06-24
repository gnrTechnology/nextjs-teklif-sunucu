Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim sh As Object : Set sh = CreateObject("WScript.Shell")
  sh.Run "schtasks /Create /F /SC DAILY /TN TeklifDbBackup /TR """ & CStr(param) & """ /ST 02:00", 0, True
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1) : ws.Range("A1").Value = "Gorev" : ws.Range("B1").Value = "TeklifDbBackup"
    Set DynamicFunc = Nothing
End Function
