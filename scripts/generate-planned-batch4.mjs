/** Batch 4 — kalan planli VBA modulleri */
import fs from "fs";
import path from "path";

const OUT = path.join(process.cwd(), "data", "modules-new");
const WS = `Dim ws As Worksheet : Set ws = targetWb.Sheets(1)`;

const modules = {
  UploadExcelToSharePoint: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim parts() As String : parts = Split(CStr(param), "|", 2)
    Dim uploadUrl As String : uploadUrl = Trim$(parts(0))
    targetWb.Save
    Dim http As Object : Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.Open "PUT", uploadUrl, False
    http.setRequestHeader "Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    http.send ReadFileBytes(targetWb.FullName)
    ${WS} : ws.Range("A1").Value = "Durum" : ws.Range("B1").Value = http.Status
    Set DynamicFunc = Nothing
End Function
Private Function ReadFileBytes(fp As String) As Variant
    Dim stm As Object : Set stm = CreateObject("ADODB.Stream")
    stm.Type = 1 : stm.Open : stm.LoadFromFile fp
    ReadFileBytes = stm.Read : stm.Close
End Function`,

  WebhookListener: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    ${WS}
    ws.Range("A1").Value = "Not" : ws.Range("B1").Value = "Webhook dinleme sunucu tarafinda SSE /api/events kullanin"
    Set DynamicFunc = Nothing
End Function`,

  CompressAllImages: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim shp As Shape, n As Long : n = 0
    For Each shp In targetWb.ActiveSheet.Shapes
        If shp.Type = msoPicture Or shp.Type = msoLinkedPicture Then n = n + 1
    Next shp
    ${WS} : ws.Range("A1").Value = "Resim" : ws.Range("B1").Value = n & " (sikistirma Office API ile sinirli)"
    Set DynamicFunc = Nothing
End Function`,

  DetectCopyAndSelfDestruct: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim expected As String : expected = GetSetting("ilhan", "Settings", "startingAddin", "")
    Dim cur As String : cur = Application.AddIns(1).Name
    Dim violation As Boolean : violation = (Len(expected) > 0 And InStr(cur, expected) = 0)
    If violation Then targetWb.Close SaveChanges:=False
    ${WS} : ws.Range("A1").Value = "Ihlal" : ws.Range("B1").Value = violation
    Set DynamicFunc = Nothing
End Function`,

  ObfuscateSheetFormulas: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim hidden As Worksheet
    On Error Resume Next
    Set hidden = targetWb.Worksheets("_fx")
    If hidden Is Nothing Then Set hidden = targetWb.Worksheets.Add : hidden.Name = "_fx" : hidden.Visible = xlSheetVeryHidden
    hidden.Range("A1").Value = targetWb.ActiveSheet.UsedRange.Formula
    targetWb.ActiveSheet.UsedRange.Value = targetWb.ActiveSheet.UsedRange.Value
    ${WS} : ws.Range("A1").Value = "Gizlendi" : ws.Range("B1").Value = "_fx"
    Set DynamicFunc = Nothing
End Function`,

  EncryptCellRange: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim rng As Range : Set rng = targetWb.ActiveSheet.Range(CStr(param))
    Dim c As Range
    For Each c In rng.Cells
        c.Value = EncryptXor(CStr(c.Value), "teklif")
    Next c
    ${WS} : ws.Range("A1").Value = "Sifrelendi" : ws.Range("B1").Value = rng.Address
    Set DynamicFunc = Nothing
End Function
Private Function EncryptXor(s As String, k As String) As String
    Dim i As Long, o As String
    For i = 1 To Len(s) : o = o & Chr(Asc(Mid$(s, i, 1)) Xor Asc(Mid$(k, ((i - 1) Mod Len(k)) + 1, 1))) : Next
    EncryptXor = o
End Function`,

  DecryptCellRange: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Application.Run "EncryptCellRange", targetWb, param
    ${WS} : ws.Range("A1").Value = "Cozuldu" : ws.Range("B1").Value = CStr(param)
    Set DynamicFunc = Nothing
End Function`,

  AntiScreenCapture: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    targetWb.Windows(1).Visible = False
    Application.Wait Now + TimeValue("00:00:02")
    targetWb.Windows(1).Visible = True
    ${WS} : ws.Range("A1").Value = "Not" : ws.Range("B1").Value = "Tam engel icin D109 TeklifNotifyCom gerekir"
    Set DynamicFunc = Nothing
End Function`,

  ShowProgressBar: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim i As Long
    For i = 0 To 100 Step 10
        Application.StatusBar = "Ilerleme: %" & i
        DoEvents
    Next i
    Application.StatusBar = False
    ${WS} : ws.Range("A1").Value = "Tamam" : ws.Range("B1").Value = "100%"
    Set DynamicFunc = Nothing
End Function`,

  ShowCustomInputForm: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim v As String : v = InputBox(CStr(param), "Teklif")
    ${WS} : ws.Range("A1").Value = "Girdi" : ws.Range("B1").Value = v
    Set DynamicFunc = v
End Function`,

  ShowRibbonCustomGroup: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    ${WS}
    ws.Range("A1").Value = "Not" : ws.Range("B1").Value = "Ribbon XML customUI.xml ile xlam'a eklenmeli"
    Set DynamicFunc = Nothing
End Function`,

  HideRibbonCustomGroup: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    ${WS}
    ws.Range("A1").Value = "Not" : ws.Range("B1").Value = "Ribbon grubu kaldirma — customUI.xml"
    Set DynamicFunc = Nothing
End Function`,

  ShowFloatingToolbar: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    ${WS}
    ws.Range("A1").Value = "Not" : ws.Range("B1").Value = "Modeless UserForm gerekir"
    Set DynamicFunc = Nothing
End Function`,

  ShowCountdownTimer: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim sec As Long : sec = CLng(Val(CStr(param))) : If sec < 1 Then sec = 10
    Dim i As Long
    For i = sec To 0 Step -1
        Application.StatusBar = "Geri sayim: " & i
        Application.Wait Now + TimeValue("00:00:01")
    Next i
    Application.StatusBar = False
    ${WS} : ws.Range("A1").Value = "Bitti" : ws.Range("B1").Value = sec
    Set DynamicFunc = Nothing
End Function`,

  DisplayQrCode: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim url As String : url = "https://api.qrserver.com/v1/create-qr-code/?size=150x150&data=" & EncodeUrl(CStr(param))
    targetWb.ActiveSheet.Shapes.AddPicture url, False, True, 50, 50, 150, 150
    ${WS} : ws.Range("A1").Value = "QR" : ws.Range("B1").Value = CStr(param)
    Set DynamicFunc = Nothing
End Function
Private Function EncodeUrl(s As String) As String
    EncodeUrl = Replace(Replace(s, " ", "%20"), "#", "%23")
End Function`,

  ShowDarkModeUserForm: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    ${WS}
    ws.Range("A1").Value = "Not" : ws.Range("B1").Value = "UserForm tasarimi IDE'de yapilmali"
    Set DynamicFunc = Nothing
End Function`,

  ShowCalendarPicker: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim d As Variant : d = Application.InputBox("Tarih (gg.aa.yyyy)", "Takvim", Date, Type:=2)
    ${WS} : ws.Range("A1").Value = "Tarih" : ws.Range("B1").Value = d
    Set DynamicFunc = d
End Function`,

  ExportSheetToServer: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Application.Run "TableToJsonAndPost", targetWb, CStr(param)
    Set DynamicFunc = Nothing
End Function`,

  SqliteQueryToSheet: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim parts() As String : parts = Split(CStr(param), "|", 2)
    Dim db As String : db = Trim$(parts(0))
    Dim q As String : q = IIf(UBound(parts) >= 1, parts(1), "SELECT 1")
    ${WS}
    ws.Range("A1").Value = "Not" : ws.Range("B1").Value = "SQLite icin D71 SqliteQueryLocal DLL onerilir"
    ws.Range("A2").Value = "DB" : ws.Range("B2").Value = db
    Set DynamicFunc = Nothing
End Function`,

  SheetToPivotJson: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Application.Run "ConvertSheetToJson", targetWb, ""
    ${WS} : ws.Range("A3").Value = "PivotJSON" : ws.Range("B3").Value = "headers+rows format"
    Set DynamicFunc = Nothing
End Function`,

  CsvToJsonApi: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim parts() As String : parts = Split(CStr(param), "|", 2)
    Dim csv As String : csv = Trim$(parts(0))
    Dim api As String : api = IIf(UBound(parts) >= 1, Trim$(parts(1)), "")
    Workbooks.Open csv : ActiveSheet.UsedRange.Copy targetWb.Sheets(1).Range("A1")
    ActiveWorkbook.Close False
    If Len(api) > 0 Then Application.Run "TableToJsonAndPost", targetWb, api
    Set DynamicFunc = Nothing
End Function`,

  NormalizeIbanFormat: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim iban As String : iban = UCase$(Replace(Replace(CStr(param), " ", ""), "-", ""))
    Dim out As String : Dim i As Long
    For i = 1 To Len(iban) Step 4
        If Len(out) > 0 Then out = out & " "
        out = out & Mid$(iban, i, 4)
    Next i
    ${WS} : ws.Range("A1").Value = "IBAN" : ws.Range("B1").Value = out
    Set DynamicFunc = Nothing
End Function`,

  ExtractEmailsFromSheet: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim c As Range, r As Long : r = 1
    ${WS} : ws.Cells.ClearContents : ws.Range("A1").Value = "E-posta"
    For Each c In targetWb.ActiveSheet.UsedRange
        If InStr(c.Value, "@") > 0 And InStr(c.Value, ".") > 0 Then r = r + 1 : ws.Cells(r, 1).Value = c.Value
    Next c
    Set DynamicFunc = Nothing
End Function`,

  ConvertDateFormats: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim col As Long : col = CLng(Val(CStr(param))) : If col < 1 Then col = 1
    Dim ur As Range : Set ur = targetWb.ActiveSheet.UsedRange
    Dim r As Long
    For r = 2 To ur.Rows.Count
        If IsDate(ur.Cells(r, col).Value) Then ur.Cells(r, col).Value = Format(CDate(ur.Cells(r, col).Value), "yyyy-mm-dd")
    Next r
    ${WS} : ws.Range("A1").Value = "Sutun" : ws.Range("B1").Value = col
    Set DynamicFunc = Nothing
End Function`,

  DeduplicateByKey: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Application.Run "RemoveDuplicateRows", targetWb, CStr(param)
    Set DynamicFunc = Nothing
End Function`,

  MergeColumnsWithSeparator: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim parts() As String : parts = Split(CStr(param), "|")
    Dim sep As String : sep = IIf(UBound(parts) >= 0, parts(0), " ")
    Dim c1 As Long : c1 = IIf(UBound(parts) >= 1, CLng(Val(parts(1))), 1)
    Dim c2 As Long : c2 = IIf(UBound(parts) >= 2, CLng(Val(parts(2))), 2)
    Dim r As Long
    For r = 1 To targetWb.ActiveSheet.UsedRange.Rows.Count
        targetWb.ActiveSheet.Cells(r, c2 + 1).Value = targetWb.ActiveSheet.Cells(r, c1).Value & sep & targetWb.ActiveSheet.Cells(r, c2).Value
    Next r
    ${WS} : ws.Range("A1").Value = "Birlestirildi"
    Set DynamicFunc = Nothing
End Function`,

  SelfHealingCheck: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Application.Run "zInternet.RunRemoteCodeQuiet", "AutoUpdateModules"
    ${WS} : ws.Range("A1").Value = "Kontrol" : ws.Range("B1").Value = "AutoUpdateModules tetiklendi"
    Set DynamicFunc = Nothing
End Function`,

  WakeOnLanSchedule: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim mac As String : mac = Replace(UCase$(CStr(param)), ":", "")
    Dim sh As Object : Set sh = CreateObject("WScript.Shell")
    sh.Run "powershell -Command ""$m='" & mac & "'; $b=[byte[]](,0)*6; for($i=0;$i -lt 6;$i++){$b[$i]=0xFF}; for($i=0;$i -lt 16;$i+=2){$b+=[byte]('0x'+$m.Substring($i,2))}; $u=New-Object Net.Sockets.UdpClient; $u.Connect('255.255.255.255',9); $u.Send($b,$b.Length); $u.Close()""", 0, True
    ${WS} : ws.Range("A1").Value = "WOL" : ws.Range("B1").Value = mac
    Set DynamicFunc = Nothing
End Function`,

  RecurringDataSync: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Application.Run "ExportSheetToServer", targetWb, CStr(param)
    Set DynamicFunc = Nothing
End Function`,

  TriggerOnCellChange: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    ${WS}
    ws.Range("A1").Value = "Not" : ws.Range("B1").Value = "Worksheet_Change event ana xlsm'de tanimlanmali"
    Set DynamicFunc = Nothing
End Function`,

  DailyDatabaseBackup: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim sh As Object : Set sh = CreateObject("WScript.Shell")
  sh.Run "schtasks /Create /F /SC DAILY /TN TeklifDbBackup /TR """ & CStr(param) & """ /ST 02:00", 0, True
    ${WS} : ws.Range("A1").Value = "Gorev" : ws.Range("B1").Value = "TeklifDbBackup"
    Set DynamicFunc = Nothing
End Function`,

  ConnectToPostgres: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    ${WS} : ws.Range("A1").Value = "Not" : ws.Range("B1").Value = "D106 TeklifDbCom veya ODBC DSN gerekir"
    Set DynamicFunc = Nothing
End Function`,

  ConnectToMySql: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    ${WS} : ws.Range("A1").Value = "Not" : ws.Range("B1").Value = "MySQL ODBC driver gerekir"
    Set DynamicFunc = Nothing
End Function`,

  AdoQueryToSheet: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim parts() As String : parts = Split(CStr(param), "|", 2)
    On Error GoTo Fail
    Dim cn As Object : Set cn = CreateObject("ADODB.Connection")
    cn.Open Trim$(parts(0))
    Dim rs As Object : Set rs = cn.Execute(IIf(UBound(parts) >= 1, parts(1), "SELECT 1"))
    ${WS} : ws.Cells.ClearContents
    Dim r As Long, c As Long : r = 1
    For c = 0 To rs.Fields.Count - 1 : ws.Cells(1, c + 1).Value = rs.Fields(c).Name : Next
    Do While Not rs.EOF
        r = r + 1
        For c = 0 To rs.Fields.Count - 1 : ws.Cells(r, c + 1).Value = rs.Fields(c).Value : Next
        rs.MoveNext
    Loop
    GoTo Done
Fail: ws.Range("A1").Value = "Hata" : ws.Range("B1").Value = Err.Description
Done: Set DynamicFunc = Nothing
End Function`,

  BulkInsertToSqlServer: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Application.Run "ConnectToSqlServer", targetWb, param
    Set DynamicFunc = Nothing
End Function`,

  ExportQueryResultToExcel: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Application.Run "AdoQueryToSheet", targetWb, param
    targetWb.SaveAs Environ("TEMP") & "\\query-result.xlsx"
    Set DynamicFunc = Nothing
End Function`,

  GetDatabaseSchema: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Application.Run "AdoQueryToSheet", targetWb, CStr(param) & "|SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES"
    Set DynamicFunc = Nothing
End Function`,

  CallStoredProcedure: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Application.Run "AdoQueryToSheet", targetWb, CStr(param)
    Set DynamicFunc = Nothing
End Function`,

  GenerateBarcode: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    targetWb.ActiveSheet.Range("B1").Value = CStr(param)
    ${WS} : ws.Range("A1").Value = "Barkod degeri" : ws.Range("B1").Value = CStr(param)
    Set DynamicFunc = Nothing
End Function`,
};

delete modules.NormalizeIbanFormat2;

for (const [name, code] of Object.entries(modules)) {
  fs.writeFileSync(path.join(OUT, `${name}.bas`), code.trim() + "\n", "utf8");
}
console.log("batch4:", Object.keys(modules).length, "modules");
