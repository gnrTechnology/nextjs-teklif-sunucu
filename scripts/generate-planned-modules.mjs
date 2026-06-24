/**
 * Batch 2+ planli modulleri data/modules-new/ altina yazar.
 * Calistir: node scripts/generate-planned-modules.mjs && node scripts/upsert-modules-new.mjs
 */
import fs from "fs";
import path from "path";

const OUT = path.join(process.cwd(), "data", "modules-new");
fs.mkdirSync(OUT, { recursive: true });

const WS = `Dim ws As Worksheet : Set ws = targetWb.Sheets(1)`;

const modules = {
  BackupVbaSettings: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim outPath As String
    outPath = Trim$(CStr(param))
    If Len(outPath) = 0 Then outPath = Environ("LOCALAPPDATA") & "\\TeklifAgent\\vba-settings-backup.json"
    Dim sections() As String : sections = Split("ilhan,scngnr,sercan", ",")
    Dim keys() As String : keys = Split("mac,mdip,TBveren,teklifYolu,startingAddin,ihlalDosyaYolu,ihlalDosyaAdi,apiBaseUrl", ",")
    Dim json As String : json = "{"
    Dim si As Long, ki As Long, first As Boolean : first = True
    For si = LBound(sections) To UBound(sections)
        For ki = LBound(keys) To UBound(keys)
            If Not first Then json = json & ","
            json = json & Chr(34) & sections(si) & "." & keys(ki) & Chr(34) & ":"
            json = json & Chr(34) & JsonEsc(GetSetting(sections(si), "Settings", keys(ki), "")) & Chr(34)
            first = False
        Next ki
    Next si
    json = json & "}"
    WriteText outPath, json
    ${WS}
    ws.Cells.ClearContents
    ws.Range("A1").Value = "Yedek dosyasi" : ws.Range("B1").Value = outPath
    ws.Range("A2").Value = "Boyut" : ws.Range("B2").Value = Len(json)
    ws.Columns.AutoFit
    Set DynamicFunc = Nothing
End Function
Private Function JsonEsc(s As String) As String
    JsonEsc = Replace(Replace(CStr(s), Chr(92), Chr(92) & Chr(92)), Chr(34), Chr(92) & Chr(34))
End Function
Private Sub WriteText(p As String, t As String)
    Dim f As Integer : f = FreeFile
    Dim dir As String : dir = Left$(p, InStrRev(p, Chr(92)) - 1)
    Dim fso As Object : Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(dir) Then fso.CreateFolder dir
    Open p For Output As #f : Print #f, t; : Close #f
End Sub`,

  RestoreVbaSettings: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim inPath As String : inPath = Trim$(CStr(param))
    If Len(inPath) = 0 Then inPath = Environ("LOCALAPPDATA") & "\\TeklifAgent\\vba-settings-backup.json"
    Dim fso As Object : Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(inPath) Then GoTo Done
    Dim ts As Object : Set ts = fso.OpenTextFile(inPath, 1)
    Dim json As String : json = ts.ReadAll : ts.Close
    Dim n As Long : n = 0
    Dim pos As Long : pos = 1
    Do
        pos = InStr(pos, json, Chr(34))
        If pos = 0 Then Exit Do
        Dim kEnd As Long : kEnd = InStr(pos + 1, json, Chr(34))
        If kEnd = 0 Then Exit Do
        Dim key As String : key = Mid$(json, pos + 1, kEnd - pos - 1)
        Dim vStart As Long : vStart = InStr(kEnd, json, Chr(34))
        If vStart = 0 Then Exit Do
        vStart = vStart + 1
        Dim vEnd As Long : vEnd = InStr(vStart, json, Chr(34))
        If vEnd = 0 Then Exit Do
        Dim val As String : val = Mid$(json, vStart, vEnd - vStart)
        Dim dot As Long : dot = InStr(key, ".")
        If dot > 0 Then
            SaveSetting Left$(key, dot - 1), "Settings", Mid$(key, dot + 1), val
            n = n + 1
        End If
        pos = vEnd + 1
    Loop
Done:
    ${WS}
    ws.Cells.ClearContents
    ws.Range("A1").Value = "Geri yuklenen" : ws.Range("B1").Value = n
    ws.Range("A2").Value = "Kaynak" : ws.Range("B2").Value = inPath
    ws.Columns.AutoFit
    Set DynamicFunc = Nothing
End Function`,

  MonitorRegistryChange: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim regPath As String : regPath = Trim$(CStr(param))
    If Len(regPath) = 0 Then regPath = "HKCU\\Software\\ilhan\\Settings"
    Dim snapFile As String : snapFile = Environ("LOCALAPPDATA") & "\\TeklifAgent\\reg-snap.txt"
    Dim sh As Object : Set sh = CreateObject("WScript.Shell")
    Dim cur As String : cur = ""
    On Error Resume Next
    cur = CStr(sh.RegRead(regPath))
    On Error GoTo 0
    Dim fso As Object : Set fso = CreateObject("Scripting.FileSystemObject")
    Dim old As String : old = ""
    If fso.FileExists(snapFile) Then
        Dim ts As Object : Set ts = fso.OpenTextFile(snapFile, 1)
        old = ts.ReadAll : ts.Close
    End If
    Dim changed As Boolean : changed = (old <> cur)
    If changed Or Len(old) = 0 Then
        Dim tw As Object : Set tw = fso.OpenTextFile(snapFile, 2, True)
        tw.Write cur : tw.Close
    End If
    ${WS}
    ws.Cells.ClearContents
    ws.Range("A1").Value = "Yol" : ws.Range("B1").Value = regPath
    ws.Range("A2").Value = "Degisti" : ws.Range("B2").Value = IIf(changed And Len(old) > 0, "EVET", "HAYIR")
    ws.Range("A3").Value = "Deger" : ws.Range("B3").Value = Left$(cur, 32000)
    ws.Columns.AutoFit
    Set DynamicFunc = Nothing
End Function`,

  CompareRegistrySnapshot: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim parts() As String
    If InStr(CStr(param), "|") > 0 Then parts = Split(CStr(param), "|", 2) Else ReDim parts(0 To 1) : parts(0) = CStr(param) : parts(1) = ""
    Dim fileA As String : fileA = Trim$(parts(0))
    Dim fileB As String : fileB = Trim$(parts(1))
    If Len(fileA) = 0 Then fileA = Environ("LOCALAPPDATA") & "\\TeklifAgent\\reg-snap-a.txt"
    If Len(fileB) = 0 Then fileB = Environ("LOCALAPPDATA") & "\\TeklifAgent\\reg-snap-b.txt"
  Dim fso As Object : Set fso = CreateObject("Scripting.FileSystemObject")
    Dim a As String, b As String
    If fso.FileExists(fileA) Then a = fso.OpenTextFile(fileA, 1).ReadAll
    If fso.FileExists(fileB) Then b = fso.OpenTextFile(fileB, 1).ReadAll
    ${WS}
    ws.Cells.ClearContents
    ws.Range("A1").Value = "Dosya A" : ws.Range("B1").Value = fileA
    ws.Range("A2").Value = "Dosya B" : ws.Range("B2").Value = fileB
    ws.Range("A3").Value = "Ayni" : ws.Range("B3").Value = (a = b)
    ws.Range("A4").Value = "Fark" : ws.Range("B4").Value = IIf(a = b, "-", "Farkli icerik")
    ws.Columns.AutoFit
    Set DynamicFunc = Nothing
End Function`,

  BasicAuthGet: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim url As String, user As String, pass As String
    ParseAuthParam CStr(param), url, user, pass
    Dim http As Object : Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.Open "GET", url, False
    http.setRequestHeader "Authorization", "Basic " & Base64Encode(user & ":" & pass)
    http.setTimeouts 5000, 10000, 30000, 30000
    http.send
    ${WS}
    ws.Cells.ClearContents
    ws.Range("A1").Value = "URL" : ws.Range("B1").Value = url
    ws.Range("A2").Value = "Durum" : ws.Range("B2").Value = http.Status
    ws.Range("A3").Value = "Yanit" : ws.Range("B3").Value = Left$(http.responseText, 32000)
    ws.Columns.AutoFit
    Set DynamicFunc = Nothing
End Function
Private Sub ParseAuthParam(p As String, ByRef url As String, ByRef user As String, ByRef pass As String)
    Dim parts() As String : parts = Split(p, "|")
    url = Trim$(parts(0)) : user = "" : pass = ""
    If UBound(parts) >= 1 Then user = Trim$(parts(1))
    If UBound(parts) >= 2 Then pass = Trim$(parts(2))
End Sub
Private Function Base64Encode(s As String) As String
    Dim dm As Object : Set dm = CreateObject("MSXML2.DOMDocument")
    Dim el As Object : Set el = dm.createElement("b64")
    el.DataType = "bin.base64"
    el.nodeTypedValue = StrConv(s, vbFromUnicode)
    Base64Encode = Replace(Replace(el.Text, vbCr, ""), vbLf, "")
End Function`,

  BearerTokenGet: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim url As String, token As String
    Dim parts() As String : parts = Split(CStr(param), "|", 2)
    url = Trim$(parts(0))
    token = IIf(UBound(parts) >= 1, Trim$(parts(1)), "")
    Dim http As Object : Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.Open "GET", url, False
    If Len(token) > 0 Then http.setRequestHeader "Authorization", "Bearer " & token
    http.setTimeouts 5000, 10000, 30000, 30000
    http.send
    ${WS}
    ws.Cells.ClearContents
    ws.Range("A1").Value = "URL" : ws.Range("B1").Value = url
    ws.Range("A2").Value = "Durum" : ws.Range("B2").Value = http.Status
    ws.Range("A3").Value = "Yanit" : ws.Range("B3").Value = Left$(http.responseText, 32000)
    ws.Columns.AutoFit
    Set DynamicFunc = Nothing
End Function`,

  SubmitFormData: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim url As String, body As String
    Dim parts() As String : parts = Split(CStr(param), "|", 2)
    url = Trim$(parts(0)) : body = IIf(UBound(parts) >= 1, parts(1), "a=1")
    Dim http As Object : Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.Open "POST", url, False
    http.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    http.send body
    ${WS}
    ws.Cells.ClearContents
    ws.Range("A1").Value = "URL" : ws.Range("B1").Value = url
    ws.Range("A2").Value = "Durum" : ws.Range("B2").Value = http.Status
    ws.Range("A3").Value = "Yanit" : ws.Range("B3").Value = Left$(http.responseText, 32000)
    ws.Columns.AutoFit
    Set DynamicFunc = Nothing
End Function`,

  GetRedirectedUrl: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim url As String : url = Trim$(CStr(param))
    Dim http As Object : Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.Open "GET", url, False
    http.setTimeouts 5000, 10000, 30000, 30000
    http.Option(6) = True
    http.send
    ${WS}
    ws.Cells.ClearContents
    ws.Range("A1").Value = "Baslangic" : ws.Range("B1").Value = url
    ws.Range("A2").Value = "Son URL" : ws.Range("B2").Value = http.getResponseHeader("Location")
    If Len(ws.Range("B2").Value) = 0 Then ws.Range("B2").Value = url
    ws.Range("A3").Value = "Durum" : ws.Range("B3").Value = http.Status
    ws.Columns.AutoFit
    Set DynamicFunc = Nothing
End Function`,

  WebScrapeSimple: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim url As String, tag As String
    Dim parts() As String : parts = Split(CStr(param), "|", 2)
    url = Trim$(parts(0)) : tag = IIf(UBound(parts) >= 1, LCase$(Trim$(parts(1))), "title")
    Dim http As Object : Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.Open "GET", url, False
    http.send
    Dim html As String : html = http.responseText
    Dim found As String : found = ExtractTag(html, tag)
    ${WS}
    ws.Cells.ClearContents
    ws.Range("A1").Value = "URL" : ws.Range("B1").Value = url
    ws.Range("A2").Value = "Etiket" : ws.Range("B2").Value = tag
    ws.Range("A3").Value = "Icerik" : ws.Range("B3").Value = Left$(found, 32000)
    ws.Columns.AutoFit
    Set DynamicFunc = Nothing
End Function
Private Function ExtractTag(html As String, tag As String) As String
    Dim o As String : o = "<" & tag
    Dim p1 As Long : p1 = InStr(1, html, o, vbTextCompare)
    If p1 = 0 Then Exit Function
    p1 = InStr(p1, html, ">") + 1
    Dim p2 As Long : p2 = InStr(p1, html, "</" & tag, vbTextCompare)
    If p2 > p1 Then ExtractTag = Trim$(Mid$(html, p1, p2 - p1))
End Function`,

  GetGoldPrice: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim url As String
    url = "https://api.exchangerate-api.com/v4/latest/USD"
    Dim http As Object : Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.Open "GET", url, False
    http.send
    Dim resp As String : resp = http.responseText
    Dim tryRate As Double : tryRate = JsonNum(resp, "TRY")
    ${WS}
    ws.Cells.ClearContents
    ws.Range("A1").Value = "1 USD (TRY)" : ws.Range("B1").Value = tryRate
    ws.Range("A2").Value = "Tahmini gram altin (TRY)" : ws.Range("B2").Value = Round((tryRate / 31.1035) * 0.95, 2)
    ws.Range("A3").Value = "Not" : ws.Range("B3").Value = "Yaklasik deger — resmi altin API icin param ile URL verin"
    ws.Columns.AutoFit
    Set DynamicFunc = Nothing
End Function
Private Function JsonNum(json As String, key As String) As Double
    Dim sk As String : sk = Chr(34) & key & Chr(34) & ":"
    Dim p As Long : p = InStr(json, sk)
    If p = 0 Then Exit Function
    p = p + Len(sk)
    JsonNum = CDbl(Val(Mid$(json, p, 12)))
End Function`,

  SendSlackMessage: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim webhook As String, msg As String
    Dim parts() As String : parts = Split(CStr(param), "|", 2)
    webhook = Trim$(parts(0)) : msg = IIf(UBound(parts) >= 1, parts(1), "Teklif bildirimi")
    Dim body As String : body = "{""text"":""" & JsonEsc(msg) & """}"
    Dim http As Object : Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.Open "POST", webhook, False
    http.setRequestHeader "Content-Type", "application/json"
    http.send body
    ${WS}
    ws.Range("A1").Value = "Durum" : ws.Range("B1").Value = http.Status
    ws.Range("A2").Value = "Yanit" : ws.Range("B2").Value = http.responseText
    ws.Columns.AutoFit
    Set DynamicFunc = Nothing
End Function
Private Function JsonEsc(s As String) As String
    JsonEsc = Replace(Replace(CStr(s), Chr(92), Chr(92) & Chr(92)), Chr(34), Chr(92) & Chr(34))
End Function`,

  SendTeamsMessage: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim webhook As String, msg As String
    Dim parts() As String : parts = Split(CStr(param), "|", 2)
    webhook = Trim$(parts(0)) : msg = IIf(UBound(parts) >= 1, parts(1), "Teklif bildirimi")
    Dim body As String
    body = "{""type"":""message"",""attachments"":[{""contentType"":""application/vnd.microsoft.card.adaptive"",""content"":{""type"":""AdaptiveCard"",""body"":[{""type"":""TextBlock"",""text"":""" & JsonEsc(msg) & """}]}}]}"
    Dim http As Object : Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.Open "POST", webhook, False
    http.setRequestHeader "Content-Type", "application/json"
    http.send body
    ${WS}
    ws.Range("A1").Value = "Durum" : ws.Range("B1").Value = http.Status
    ws.Columns.AutoFit
    Set DynamicFunc = Nothing
End Function
Private Function JsonEsc(s As String) As String
    JsonEsc = Replace(Replace(CStr(s), Chr(92), Chr(92) & Chr(92)), Chr(34), Chr(92) & Chr(34))
End Function`,

  FetchCurrencyHistory: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim base As String : base = Trim$(CStr(param)) : If Len(base) = 0 Then base = "USD"
    Dim url As String : url = "https://api.exchangerate-api.com/v4/latest/" & base
    Dim http As Object : Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.Open "GET", url, False : http.send
    Dim resp As String : resp = http.responseText
    ${WS}
    ws.Cells.ClearContents
    ws.Range("A1").Value = "Para" : ws.Range("B1").Value = "Kur (1 " & base & ")"
    Dim cur() As String : cur = Split("TRY,EUR,GBP,CHF,JPY", ",")
    Dim i As Long, r As Long : r = 2
    For i = LBound(cur) To UBound(cur)
        ws.Cells(r, 1).Value = cur(i)
        ws.Cells(r, 2).Value = JsonNum(resp, cur(i))
        r = r + 1
    Next i
    ws.Range(r, 1).Value = "Tarih" : ws.Range(r, 2).Value = Now
    ws.Columns.AutoFit
    Set DynamicFunc = Nothing
End Function
Private Function JsonNum(json As String, key As String) As Double
    Dim sk As String : sk = Chr(34) & key & Chr(34) & ":"
    Dim p As Long : p = InStr(json, sk)
    If p = 0 Then Exit Function
    JsonNum = CDbl(Val(Mid$(json, p + Len(sk), 12)))
End Function`,

  OAuthGetToken: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim tokenUrl As String, cid As String, secret As String, scope As String
    Dim parts() As String : parts = Split(CStr(param), "|")
    tokenUrl = Trim$(parts(0))
    cid = IIf(UBound(parts) >= 1, Trim$(parts(1)), "")
    secret = IIf(UBound(parts) >= 2, Trim$(parts(2)), "")
    scope = IIf(UBound(parts) >= 3, Trim$(parts(3)), "")
    Dim body As String
    body = "grant_type=client_credentials&client_id=" & cid & "&client_secret=" & secret
    If Len(scope) > 0 Then body = body & "&scope=" & scope
    Dim http As Object : Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.Open "POST", tokenUrl, False
    http.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    http.send body
    ${WS}
    ws.Range("A1").Value = "Durum" : ws.Range("B1").Value = http.Status
    ws.Range("A2").Value = "Token" : ws.Range("B2").Value = Left$(http.responseText, 32000)
    ws.Columns.AutoFit
    Set DynamicFunc = Nothing
End Function`,

  GetLatestModuleVersion: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim modName As String : modName = Trim$(CStr(param))
    Dim baseUrl As String
    baseUrl = GetSetting("ilhan", "Settings", "apiBaseUrl", "http://localhost:3000/api/")
    If Right(baseUrl, 1) <> "/" Then baseUrl = baseUrl & "/"
    Dim body As String : body = "{""methodName"":""" & modName & """}"
    Dim http As Object : Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.Open "POST", baseUrl & "module/", False
    http.setRequestHeader "Content-Type", "application/json"
    http.send body
    ${WS}
    ws.Range("A1").Value = "Modul" : ws.Range("B1").Value = modName
    ws.Range("A2").Value = "Durum" : ws.Range("B2").Value = http.Status
    ws.Range("A3").Value = "Yanit" : ws.Range("B3").Value = Left$(http.responseText, 2000)
    ws.Columns.AutoFit
    Set DynamicFunc = Nothing
End Function`,

  CheckForUpdate: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim localVer As String, modName As String
    Dim parts() As String : parts = Split(CStr(param), "|", 2)
    modName = Trim$(parts(0))
    localVer = IIf(UBound(parts) >= 1, Trim$(parts(1)), "1")
    Application.Run "GetLatestModuleVersion", targetWb, modName
    ${WS}
    ws.Range("A4").Value = "Yerel" : ws.Range("B4").Value = localVer
    ws.Range("A5").Value = "Guncelleme" : ws.Range("B5").Value = "Sunucu yanitini karsilastirin"
    Set DynamicFunc = Nothing
End Function`,

  WatchFolderForNewFile: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim folderPath As String, timeoutSec As Long
    Dim parts() As String : parts = Split(CStr(param), "|", 2)
    folderPath = Trim$(parts(0))
    timeoutSec = IIf(UBound(parts) >= 1, CLng(Val(parts(1))), 60)
    If Right(folderPath, 1) <> Chr(92) Then folderPath = folderPath & Chr(92)
    Dim fso As Object : Set fso = CreateObject("Scripting.FileSystemObject")
    Dim start As Double : start = Timer
    Dim found As String : found = ""
    Do While Timer - start < timeoutSec
        Dim fn As String : fn = Dir(folderPath & "*.*")
        Do While Len(fn) > 0
            If fn <> "." And fn <> ".." Then found = folderPath & fn : Exit Do
            fn = Dir()
        Loop
        If Len(found) > 0 Then Exit Do
        Application.Wait Now + TimeValue("00:00:01")
    Loop
    ${WS}
    ws.Range("A1").Value = "Klasor" : ws.Range("B1").Value = folderPath
    ws.Range("A2").Value = "Bulunan" : ws.Range("B2").Value = IIf(Len(found) > 0, found, "(yok)")
    ws.Columns.AutoFit
    Set DynamicFunc = Nothing
End Function`,

  ConvertFileToBase64: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim fp As String : fp = Trim$(CStr(param))
    Dim stm As Object : Set stm = CreateObject("ADODB.Stream")
    stm.Type = 1 : stm.Open
    stm.LoadFromFile fp
    Dim bytes() As Byte : bytes = stm.Read
    stm.Close
    Dim dm As Object : Set dm = CreateObject("MSXML2.DOMDocument")
    Dim el As Object : Set el = dm.createElement("b64")
    el.DataType = "bin.base64" : el.nodeTypedValue = bytes
    ${WS}
    ws.Range("A1").Value = "Dosya" : ws.Range("B1").Value = fp
    ws.Range("A2").Value = "Base64" : ws.Range("B2").Value = Left$(el.Text, 32000)
    ws.Columns.AutoFit
    Set DynamicFunc = Nothing
End Function`,

  ConvertBase64ToFile: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim b64 As String, outPath As String
    Dim parts() As String : parts = Split(CStr(param), "|", 2)
    b64 = parts(0) : outPath = IIf(UBound(parts) >= 1, Trim$(parts(1)), Environ("TEMP") & "\\decoded.bin")
    Dim dm As Object : Set dm = CreateObject("MSXML2.DOMDocument")
    Dim el As Object : Set el = dm.createElement("b64")
    el.DataType = "bin.base64" : el.Text = b64
    Dim stm As Object : Set stm = CreateObject("ADODB.Stream")
    stm.Type = 1 : stm.Open : stm.Write el.nodeTypedValue
    stm.SaveToFile outPath, 2 : stm.Close
    ${WS}
    ws.Range("A1").Value = "Cikti" : ws.Range("B1").Value = outPath
    ws.Columns.AutoFit
    Set DynamicFunc = Nothing
End Function`,

  SaveAllWorkbooks: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim n As Long : n = 0
    Dim wb As Workbook
    For Each wb In Application.Workbooks
        If Not wb.IsAddin Then wb.Save : n = n + 1
    Next wb
    ${WS}
    ws.Range("A1").Value = "Kaydedilen" : ws.Range("B1").Value = n
    Set DynamicFunc = Nothing
End Function`,

  ExportSheetAsPdf: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim outPdf As String : outPdf = Trim$(CStr(param))
    If Len(outPdf) = 0 Then outPdf = Environ("TEMP") & "\\export.pdf"
    targetWb.ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:=outPdf
    ${WS}
    ws.Range("A1").Value = "PDF" : ws.Range("B1").Value = outPdf
    Set DynamicFunc = Nothing
End Function`,

  AutoFitAllColumns: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim sh As Worksheet
    For Each sh In targetWb.Worksheets
        sh.Columns.AutoFit
    Next sh
    ${WS}
    ws.Range("A1").Value = "Sayfa" : ws.Range("B1").Value = targetWb.Worksheets.Count
    Set DynamicFunc = Nothing
End Function`,

  RefreshAllPivotTables: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim sh As Worksheet, pt As PivotTable, n As Long : n = 0
    For Each sh In targetWb.Worksheets
        For Each pt In sh.PivotTables
            pt.RefreshTable : n = n + 1
        Next pt
    Next sh
    ${WS}
    ws.Range("A1").Value = "Yenilenen pivot" : ws.Range("B1").Value = n
    Set DynamicFunc = Nothing
End Function`,

  RefreshAllConnections: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim cn As WorkbookConnection
    For Each cn In targetWb.Connections
        On Error Resume Next : cn.Refresh : On Error GoTo 0
    Next cn
    ${WS}
    ws.Range("A1").Value = "Baglanti" : ws.Range("B1").Value = targetWb.Connections.Count
    Set DynamicFunc = Nothing
End Function`,

  ConvertFormulasToValues: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim addr As String : addr = Trim$(CStr(param))
    If Len(addr) = 0 Then addr = targetWb.ActiveSheet.UsedRange.Address
    targetWb.ActiveSheet.Range(addr).Value = targetWb.ActiveSheet.Range(addr).Value
    ${WS}
    ws.Range("A1").Value = "Aralik" : ws.Range("B1").Value = addr
    Set DynamicFunc = Nothing
End Function`,

  RemoveDuplicateRows: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim col As Long : col = CLng(Val(CStr(param)))
    If col < 1 Then col = 1
    Dim ur As Range : Set ur = targetWb.ActiveSheet.UsedRange
    Dim before As Long : before = ur.Rows.Count
    ur.RemoveDuplicates Columns:=col, Header:=xlYes
    Dim after As Long : after = targetWb.ActiveSheet.UsedRange.Rows.Count
    ${WS}
    ws.Range("A1").Value = "Silinen" : ws.Range("B1").Value = before - after
    Set DynamicFunc = Nothing
End Function`,

  BulkFindAndReplace: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim oldT As String, newT As String
    Dim parts() As String : parts = Split(CStr(param), "|", 2)
    oldT = parts(0) : newT = IIf(UBound(parts) >= 1, parts(1), "")
    Dim sh As Worksheet, c As Range, n As Long : n = 0
    For Each sh In targetWb.Worksheets
        For Each c In sh.UsedRange
            If InStr(1, CStr(c.Value), oldT, vbTextCompare) > 0 Then
                c.Value = Replace(CStr(c.Value), oldT, newT, , , vbTextCompare) : n = n + 1
            End If
        Next c
    Next sh
    ${WS}
    ws.Range("A1").Value = "Degisen hucre" : ws.Range("B1").Value = n
    Set DynamicFunc = Nothing
End Function`,

  FreezeFirstRowAndColumn: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    With targetWb.ActiveWindow
        .FreezePanes = False
        .SplitColumn = 1
        .SplitRow = 1
        .FreezePanes = True
    End With
    ${WS}
    ws.Range("A1").Value = "Durum" : ws.Range("B1").Value = "Donduruldu"
    Set DynamicFunc = Nothing
End Function`,

  CreateSummarySheet: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim sumName As String : sumName = "Ozet"
    On Error Resume Next
    Application.DisplayAlerts = False
    targetWb.Worksheets(sumName).Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    Dim wsSum As Worksheet : Set wsSum = targetWb.Worksheets.Add(Before:=targetWb.Sheets(1))
    wsSum.Name = sumName
    wsSum.Range("A1").Value = "Sayfa" : wsSum.Range("B1").Value = "A1"
    Dim sh As Worksheet, r As Long : r = 2
    For Each sh In targetWb.Worksheets
        If sh.Name <> sumName Then
            wsSum.Cells(r, 1).Value = sh.Name
            wsSum.Cells(r, 2).Value = sh.Range("A1").Value
            r = r + 1
        End If
    Next sh
    Set DynamicFunc = Nothing
End Function`,

  ProtectAllSheets: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim pwd As String : pwd = CStr(param)
    Dim sh As Worksheet
    For Each sh In targetWb.Worksheets
        sh.Protect Password:=pwd
    Next sh
    ${WS}
    ws.Range("A1").Value = "Korunan" : ws.Range("B1").Value = targetWb.Worksheets.Count
    Set DynamicFunc = Nothing
End Function`,

  UnprotectAllSheets: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim pwd As String : pwd = CStr(param)
    Dim sh As Worksheet
    For Each sh In targetWb.Worksheets
        On Error Resume Next : sh.Unprotect Password:=pwd : On Error GoTo 0
    Next sh
    ${WS}
    ws.Range("A1").Value = "Acilan" : ws.Range("B1").Value = targetWb.Worksheets.Count
    Set DynamicFunc = Nothing
End Function`,

  SortSheetsByName: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim i As Long, j As Long
    For i = 1 To targetWb.Worksheets.Count - 1
        For j = i + 1 To targetWb.Worksheets.Count
            If targetWb.Worksheets(j).Name < targetWb.Worksheets(i).Name Then
                targetWb.Worksheets(j).Move Before:=targetWb.Worksheets(i)
            End If
        Next j
    Next i
    ${WS}
    ws.Range("A1").Value = "Siralandi" : ws.Range("B1").Value = Now
    Set DynamicFunc = Nothing
End Function`,

  CopySheetToNewWorkbook: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim newPath As String : newPath = Trim$(CStr(param))
    If Len(newPath) = 0 Then newPath = Environ("TEMP") & "\\sheet-copy.xlsx"
    targetWb.ActiveSheet.Copy
    ActiveWorkbook.SaveAs Filename:=newPath
    ActiveWorkbook.Close False
    ${WS}
    ws.Range("A1").Value = "Dosya" : ws.Range("B1").Value = newPath
    Set DynamicFunc = Nothing
End Function`,

  AutoNumberRows: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim col As Long : col = CLng(Val(CStr(param))) : If col < 1 Then col = 1
    Dim ur As Range : Set ur = targetWb.ActiveSheet.UsedRange
    Dim r As Long
    For r = 1 To ur.Rows.Count
        targetWb.ActiveSheet.Cells(ur.Row + r - 1, col).Value = r
    Next r
    ${WS}
    ws.Range("A1").Value = "Satir" : ws.Range("B1").Value = ur.Rows.Count
    Set DynamicFunc = Nothing
End Function`,

  CheckAdminRights: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim sh As Object : Set sh = CreateObject("Shell.Application")
    Dim isAdmin As Boolean : isAdmin = sh.IsUserAnAdmin
    ${WS}
    ws.Range("A1").Value = "Admin" : ws.Range("B1").Value = isAdmin
    Set DynamicFunc = Nothing
End Function`,

  GenerateHardwareId: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim cpu As String, mac As String
    On Error Resume Next
    Dim wmi As Object : Set wmi = GetObject("winmgmts:\\.\root\cimv2")
    Dim col As Object : Set col = wmi.ExecQuery("SELECT ProcessorId FROM Win32_Processor")
    Dim o As Object : For Each o In col : cpu = o.ProcessorId : Exit For : Next
    Set col = wmi.ExecQuery("SELECT MACAddress FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled=True")
    For Each o In col : If Not IsNull(o.MACAddress) Then mac = o.MACAddress : Exit For
    Dim raw As String : raw = UCase$(cpu & mac)
    Dim i As Long, hx As String : hx = ""
    For i = 1 To Len(raw)
        hx = hx & Right$("0" & Hex(Asc(Mid$(raw, i, 1)) Mod 256), 2)
    Next i
    ${WS}
    ws.Range("A1").Value = "HWID" : ws.Range("B1").Value = Left$(hx, 32)
    Set DynamicFunc = Nothing
End Function`,

  DetectVirtualMachine: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim vm As Boolean : vm = False
    On Error Resume Next
    Dim wmi As Object : Set wmi = GetObject("winmgmts:\\.\root\cimv2")
    Dim col As Object : Set col = wmi.ExecQuery("SELECT Model,Manufacturer FROM Win32_ComputerSystem")
    Dim o As Object
    For Each o In col
        Dim s As String : s = UCase$(o.Model & " " & o.Manufacturer)
        If InStr(s, "VMWARE") > 0 Or InStr(s, "VIRTUAL") > 0 Or InStr(s, "HYPER-V") > 0 Or InStr(s, "QEMU") > 0 Then vm = True
    Next
    ${WS}
    ws.Range("A1").Value = "VM" : ws.Range("B1").Value = vm
    Set DynamicFunc = Nothing
End Function`,

  ValidateTCKimlik: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim tc As String : tc = Replace(Trim$(CStr(param)), " ", "")
    Dim ok As Boolean : ok = False
    If Len(tc) = 11 And IsNumeric(tc) Then
        If Left$(tc, 1) <> "0" Then
            Dim d(1 To 11) As Long, i As Long
            For i = 1 To 11 : d(i) = CLng(Mid$(tc, i, 1)) : Next
            Dim s1 As Long, s2 As Long
            For i = 1 To 9 Step 2 : s1 = s1 + d(i) : Next
            For i = 2 To 8 Step 2 : s2 = s2 + d(i) : Next
            ok = ((s1 * 7 - s2) Mod 10 = d(10)) And (((s1 + s2 + d(10)) Mod 10) = d(11))
        End If
    End If
    ${WS}
    ws.Range("A1").Value = "TC" : ws.Range("B1").Value = tc
    ws.Range("A2").Value = "Gecerli" : ws.Range("B2").Value = ok
    Set DynamicFunc = Nothing
End Function`,

  NormalizePhoneNumbers: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim col As Long : col = CLng(Val(CStr(param))) : If col < 1 Then col = 1
    Dim ur As Range : Set ur = targetWb.ActiveSheet.UsedRange
    Dim r As Long, n As Long : n = 0
    For r = 1 To ur.Rows.Count
        Dim v As String : v = Replace(Replace(Replace(CStr(ur.Cells(r, col).Value), " ", ""), "-", ""), "(", "")
        v = Replace(v, ")", "")
        If Len(v) >= 10 Then
            If Left$(v, 1) = "0" Then v = Mid$(v, 2)
            If Left$(v, 2) <> "90" Then v = "90" & v
            ur.Cells(r, col).Value = "+90 " & Mid$(v, 3, 3) & " " & Mid$(v, 6, 3) & " " & Mid$(v, 9, 2) & " " & Mid$(v, 11, 2)
            n = n + 1
        End If
    Next r
    ${WS}
    ws.Range("A1").Value = "Normalize" : ws.Range("B1").Value = n
    Set DynamicFunc = Nothing
End Function`,

  FormatCurrencyColumn: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim col As Long : col = CLng(Val(CStr(param))) : If col < 1 Then col = 1
    targetWb.ActiveSheet.Columns(col).NumberFormat = "#,##0.00 ""₺"""
    ${WS}
    ws.Range("A1").Value = "Sutun" : ws.Range("B1").Value = col
    Set DynamicFunc = Nothing
End Function`,

  ConvertSheetToJson: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim ur As Range : Set ur = targetWb.ActiveSheet.UsedRange
    Dim json As String : json = "["
    Dim r As Long, c As Long
    For r = 1 To ur.Rows.Count
        If r > 1 Then json = json & ","
        json = json & "{"
        For c = 1 To ur.Columns.Count
            If c > 1 Then json = json & ","
            json = json & Chr(34) & "c" & c & Chr(34) & ":" & Chr(34) & Replace(CStr(ur.Cells(r, c).Value), Chr(34), "'") & Chr(34)
        Next c
        json = json & "}"
    Next r
    json = json & "]"
    ${WS}
    ws.Range("A1").Value = "JSON" : ws.Range("B1").Value = Left$(json, 32000)
    Set DynamicFunc = Nothing
End Function`,

  ImportJsonToSheet: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim json As String : json = CStr(param)
    ${WS}
    ws.Cells.ClearContents
    ws.Range("A1").Value = "JSON icerik (basit)"
    ws.Range("B1").Value = Left$(json, 32000)
    ws.Columns.AutoFit
    Set DynamicFunc = Nothing
End Function`,

  CleanHtmlToText: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim html As String : html = CStr(param)
    Dim i As Long, inside As Boolean, ch As String, out As String
    inside = False
    For i = 1 To Len(html)
        ch = Mid$(html, i, 1)
        If ch = "<" Then inside = True
        If Not inside Then out = out & ch
        If ch = ">" Then inside = False
    Next i
    out = Replace(Replace(out, "&nbsp;", " "), "&amp;", "&")
    ${WS}
    ws.Range("A1").Value = "Metin" : ws.Range("B1").Value = Left$(Trim$(out), 32000)
    Set DynamicFunc = Nothing
End Function`,

  CheckFileIntegrity: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim fp As String, expected As String
    Dim parts() As String : parts = Split(CStr(param), "|", 2)
    fp = Trim$(parts(0)) : expected = IIf(UBound(parts) >= 1, LCase$(Trim$(parts(1))), "")
    Application.Run "GetFileHashMd5", targetWb, fp
    ${WS}
    Dim actual As String : actual = LCase$(Trim$(CStr(ws.Range("B2").Value)))
    ws.Range("A4").Value = "Beklenen" : ws.Range("B4").Value = expected
    ws.Range("A5").Value = "Eslesme" : ws.Range("B5").Value = (actual = expected)
    Set DynamicFunc = Nothing
End Function`,

  SyncFolderToServer: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim folderPath As String, apiUrl As String
    Dim parts() As String : parts = Split(CStr(param), "|", 2)
    folderPath = Trim$(parts(0))
    apiUrl = IIf(UBound(parts) >= 1, Trim$(parts(1)), GetSetting("ilhan", "Settings", "apiBaseUrl", "") & "module-output/")
    Dim fso As Object : Set fso = CreateObject("Scripting.FileSystemObject")
    Dim folder As Object : Set folder = fso.GetFolder(folderPath)
    Dim f As Object, n As Long : n = 0
    For Each f In folder.Files
        n = n + 1
    Next f
    ${WS}
    ws.Range("A1").Value = "Klasor" : ws.Range("B1").Value = folderPath
    ws.Range("A2").Value = "Dosya sayisi" : ws.Range("B2").Value = n
    ws.Range("A3").Value = "API" : ws.Range("B3").Value = apiUrl
    Set DynamicFunc = Nothing
End Function`,

  ConvertPdfToText: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    ${WS}
    ws.Range("A1").Value = "Not"
    ws.Range("B1").Value = "PDF metin cikarma icin Adobe Acrobat COM veya D82 PdfExtractText DLL gerekir"
    Set DynamicFunc = Nothing
End Function`,

  ShowYesNoCancelDialog: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim msg As String : msg = CStr(param)
    Dim ans As Long : ans = MsgBox(msg, vbYesNoCancel + vbQuestion, "Teklif")
    Dim result As String
    If ans = vbYes Then result = "Yes" ElseIf ans = vbNo Then result = "No" Else result = "Cancel"
    ${WS}
    ws.Range("A1").Value = "Secim" : ws.Range("B1").Value = result
    Set DynamicFunc = result
End Function`,

  ExportAllSheetsAsPdf: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim outDir As String : outDir = Trim$(CStr(param))
    If Len(outDir) = 0 Then outDir = Environ("TEMP") & "\\pdf-export\\"
    Dim fso As Object : Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(outDir) Then fso.CreateFolder outDir
    Dim sh As Worksheet, n As Long : n = 0
    For Each sh In targetWb.Worksheets
        sh.ExportAsFixedFormat xlTypePDF, outDir & sh.Name & ".pdf"
        n = n + 1
    Next sh
    ${WS}
    ws.Range("A1").Value = "PDF sayisi" : ws.Range("B1").Value = n
    ws.Range("A2").Value = "Klasor" : ws.Range("B2").Value = outDir
    Set DynamicFunc = Nothing
End Function`,

  ImportSheetFromFile: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim src As String : src = Trim$(CStr(param))
    Dim wb As Workbook : Set wb = Workbooks.Open(src, ReadOnly:=True)
    wb.Sheets(1).Copy After:=targetWb.Sheets(targetWb.Sheets.Count)
    wb.Close False
    ${WS}
    ws.Range("A1").Value = "Kaynak" : ws.Range("B1").Value = src
    Set DynamicFunc = Nothing
End Function`,

  BatchRenameSheets: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim json As String : json = CStr(param)
    Dim pos As Long, oldN As String, newN As String, n As Long : n = 0
    pos = 1
    Do
        pos = InStr(pos, json, Chr(34))
        If pos = 0 Then Exit Do
        oldN = JsonPair(json, pos, newN)
        If Len(oldN) = 0 Then Exit Do
        On Error Resume Next
        targetWb.Worksheets(oldN).Name = newN
        If Err.Number = 0 Then n = n + 1
        Err.Clear
    Loop
    ${WS}
    ws.Range("A1").Value = "Yeniden ad" : ws.Range("B1").Value = n
    Set DynamicFunc = Nothing
End Function
Private Function JsonPair(j As String, start As Long, ByRef v As String) As String
    Dim p As Long : p = InStr(start, j, Chr(34)) : If p = 0 Then Exit Function
    p = p + 1 : Dim p2 As Long : p2 = InStr(p, j, Chr(34))
    JsonPair = Mid$(j, p, p2 - p)
    p = InStr(p2, j, ":") : p = InStr(p, j, Chr(34)) + 1
    p2 = InStr(p, j, Chr(34)) : v = Mid$(j, p, p2 - p)
End Function`,

  MergeMultipleFiles: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim list As String : list = CStr(param)
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Cells.ClearContents
    Dim r As Long : r = 1
    Dim f As Variant
    For Each f In Split(list, ";")
        Dim wb As Workbook : Set wb = Workbooks.Open(Trim$(CStr(f)), ReadOnly:=True)
        wb.Sheets(1).UsedRange.Copy ws.Cells(r, 1)
        r = r + wb.Sheets(1).UsedRange.Rows.Count
        wb.Close False
    Next f
    ${WS}
    ws.Range("A1").Value = "Birlestirildi" : ws.Range("B1").Value = r
    Set DynamicFunc = Nothing
End Function`,

  SplitSheetByColumn: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim col As Long : col = CLng(Val(CStr(param))) : If col < 1 Then col = 1
    Dim outDir As String : outDir = Environ("TEMP") & "\\split\\"
    Dim fso As Object : Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(outDir) Then fso.CreateFolder outDir
    Dim dict As Object : Set dict = CreateObject("Scripting.Dictionary")
    Dim ur As Range : Set ur = targetWb.ActiveSheet.UsedRange
    Dim r As Long
    For r = 2 To ur.Rows.Count
        Dim k As String : k = CStr(ur.Cells(r, col).Value)
        If Not dict.Exists(k) Then dict.Add k, 1
    Next r
    ${WS}
    ws.Range("A1").Value = "Grup" : ws.Range("B1").Value = dict.Count
    Set DynamicFunc = Nothing
End Function`,

  AddWatermarkToSheet: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim txt As String : txt = CStr(param) : If Len(txt) = 0 Then txt = "TEKLIF"
    With targetWb.ActiveSheet.PageSetup
        .CenterHeader = txt
    End With
    ${WS}
    ws.Range("A1").Value = "Filigran" : ws.Range("B1").Value = txt
    Set DynamicFunc = Nothing
End Function`,

  SendWorkbookByEmail: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    ${WS}
    Dim parts() As String : parts = Split(CStr(param), "|")
    Dim toAddr As String : toAddr = Trim$(parts(0))
    Dim subj As String : subj = IIf(UBound(parts) >= 1, parts(1), "Teklif")
    On Error GoTo Fail
    Dim ol As Object : Set ol = CreateObject("Outlook.Application")
    Dim mail As Object : Set mail = ol.CreateItem(0)
    mail.To = toAddr : mail.Subject = subj
    targetWb.Save
    mail.Attachments.Add targetWb.FullName
    mail.Send
    ws.Range("A1").Value = "Gonderildi" : ws.Range("B1").Value = toAddr
    GoTo Done
Fail:
    ws.Range("A1").Value = "Hata" : ws.Range("B1").Value = Err.Description
Done:
    Set DynamicFunc = Nothing
End Function`,

  TableToJsonAndPost: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim apiUrl As String : apiUrl = Trim$(CStr(param))
    Application.Run "ConvertSheetToJson", targetWb, ""
    Dim json As String : json = CStr(targetWb.Sheets(1).Range("B1").Value)
    Dim http As Object : Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.Open "POST", apiUrl, False
    http.setRequestHeader "Content-Type", "application/json"
    http.send json
    targetWb.Sheets(1).Range("A2").Value = "POST" : targetWb.Sheets(1).Range("B2").Value = http.Status
    Set DynamicFunc = Nothing
End Function`,

  CreateNamedRangeFromSelection: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim nm As String : nm = Trim$(CStr(param))
    targetWb.Names.Add Name:=nm, RefersTo:=targetWb.ActiveSheet.Selection
    ${WS}
    ws.Range("A1").Value = "Ad" : ws.Range("B1").Value = nm
    Set DynamicFunc = Nothing
End Function`,

  GeneratePivotFromData: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim addr As String : addr = Trim$(CStr(param))
    If Len(addr) = 0 Then addr = targetWb.ActiveSheet.UsedRange.Address
    Dim pc As PivotCache
    Set pc = targetWb.PivotCaches.Create(xlDatabase, targetWb.ActiveSheet.Range(addr))
    Dim pt As PivotTable
    Set pt = pc.CreatePivotTable(targetWb.ActiveSheet.Range("H1"), "PivotAuto")
    ${WS}
    ws.Range("A1").Value = "Pivot" : ws.Range("B1").Value = pt.Name
    Set DynamicFunc = Nothing
End Function`,

  ApplyConditionalFormatting: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim rng As String : rng = Trim$(CStr(param))
    If Len(rng) = 0 Then rng = targetWb.ActiveSheet.UsedRange.Address
    With targetWb.ActiveSheet.Range(rng).FormatConditions.Add(Type:=xlCellValue, Operator:=xlGreater, Formula1:="0")
        .Interior.Color = RGB(198, 239, 206)
    End With
    ${WS}
    ws.Range("A1").Value = "Aralik" : ws.Range("B1").Value = rng
    Set DynamicFunc = Nothing
End Function`,

  InsertChartFromData: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim addr As String : addr = Trim$(CStr(param))
    If Len(addr) = 0 Then addr = targetWb.ActiveSheet.UsedRange.Address
    targetWb.ActiveSheet.Shapes.AddChart2 227, xlColumnClustered, 300, 10, 400, 250
    targetWb.ActiveSheet.ChartObjects(1).Chart.SetSourceData targetWb.ActiveSheet.Range(addr)
    ${WS}
    ws.Range("A1").Value = "Grafik" : ws.Range("B1").Value = "OK"
    Set DynamicFunc = Nothing
End Function`,

  ConvertSheetToHtmlTable: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim outPath As String : outPath = Trim$(CStr(param))
    If Len(outPath) = 0 Then outPath = Environ("TEMP") & "\\sheet.html"
    targetWb.ActiveSheet.SaveAs outPath, xlHtml
    ${WS}
    ws.Range("A1").Value = "HTML" : ws.Range("B1").Value = outPath
    Set DynamicFunc = Nothing
End Function`,

  InsertRowAboveSelected: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim n As Long : n = CLng(Val(CStr(param))) : If n < 1 Then n = 1
    Dim r As Long
    For r = 1 To n
        targetWb.ActiveSheet.Rows(targetWb.ActiveSheet.Selection.Row).Insert
    Next r
    ${WS}
    ws.Range("A1").Value = "Eklenen" : ws.Range("B1").Value = n
    Set DynamicFunc = Nothing
End Function`,

  CheckLicenseStatus: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim lic As String : lic = GetSetting("ilhan", "Settings", "license", "")
    ${WS}
    ws.Range("A1").Value = "Lisans" : ws.Range("B1").Value = lic
    ws.Range("A2").Value = "Aktif" : ws.Range("B2").Value = (LCase$(lic) = "true" Or lic = "1")
    Set DynamicFunc = Nothing
End Function`,

  ValidateMacWithServer: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim mac As String : mac = GetSetting("ilhan", "Settings", "mac", "")
    Dim baseUrl As String : baseUrl = GetSetting("ilhan", "Settings", "apiBaseUrl", "http://localhost:3000/api/")
    If Right(baseUrl, 1) <> "/" Then baseUrl = baseUrl & "/"
    Dim http As Object : Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.Open "GET", baseUrl & "license/" & mac & "/", False
    http.send
    ${WS}
    ws.Range("A1").Value = "MAC" : ws.Range("B1").Value = mac
    ws.Range("A2").Value = "Durum" : ws.Range("B2").Value = http.Status
    ws.Range("A3").Value = "Yanit" : ws.Range("B3").Value = Left$(http.responseText, 2000)
    Set DynamicFunc = Nothing
End Function`,

  EncryptTextXor: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim parts() As String : parts = Split(CStr(param), "|", 2)
    Dim plain As String : plain = parts(0)
    Dim key As String : key = IIf(UBound(parts) >= 1, parts(1), "teklif")
    Dim i As Long, out As String
    For i = 1 To Len(plain)
        out = out & Chr(Asc(Mid$(plain, i, 1)) Xor Asc(Mid$(key, ((i - 1) Mod Len(key)) + 1, 1)))
    Next i
    ${WS}
    ws.Range("A1").Value = "Sifreli" : ws.Range("B1").Value = Base64Encode(out)
    Set DynamicFunc = Nothing
End Function
Private Function Base64Encode(s As String) As String
    Dim dm As Object : Set dm = CreateObject("MSXML2.DOMDocument")
    Dim el As Object : Set el = dm.createElement("b64")
    el.DataType = "bin.base64" : el.nodeTypedValue = StrConv(s, vbFromUnicode)
    Base64Encode = Replace(Replace(el.Text, vbCr, ""), vbLf, "")
End Function`,

  DecryptTextXor: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim parts() As String : parts = Split(CStr(param), "|", 2)
    Dim b64 As String : b64 = parts(0)
    Dim key As String : key = IIf(UBound(parts) >= 1, parts(1), "teklif")
    Dim bin As String : bin = Base64Decode(b64)
    Dim i As Long, out As String
    For i = 1 To Len(bin)
        out = out & Chr(Asc(Mid$(bin, i, 1)) Xor Asc(Mid$(key, ((i - 1) Mod Len(key)) + 1, 1)))
    Next i
    ${WS}
    ws.Range("A1").Value = "Cozuldu" : ws.Range("B1").Value = out
    Set DynamicFunc = Nothing
End Function
Private Function Base64Decode(s As String) As String
    Dim dm As Object : Set dm = CreateObject("MSXML2.DOMDocument")
    Dim el As Object : Set el = dm.createElement("b64")
    el.DataType = "bin.base64" : el.Text = s
    Base64Decode = StrConv(el.nodeTypedValue, vbUnicode)
End Function`,

  AuditLogAction: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim detail As String : detail = CStr(param)
    Dim mac As String : mac = GetSetting("ilhan", "Settings", "mac", "")
    Dim baseUrl As String : baseUrl = GetSetting("ilhan", "Settings", "apiBaseUrl", "http://localhost:3000/api/")
    If Right(baseUrl, 1) <> "/" Then baseUrl = baseUrl & "/"
    Dim body As String
    body = "{""eventType"":""audit"",""macAdresi"":""" & mac & """,""detail"":""" & Replace(detail, Chr(34), "'") & """}"
    Dim http As Object : Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.Open "POST", baseUrl & "activity/", False
    http.setRequestHeader "Content-Type", "application/json"
    http.send body
    ${WS}
    ws.Range("A1").Value = "Log" : ws.Range("B1").Value = http.Status
    Set DynamicFunc = Nothing
End Function`,

  CheckDebuggerAttached: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim attached As Boolean : attached = False
    On Error Resume Next
    attached = (Application.VBE.ActiveVBProject.Name <> "")
    On Error GoTo 0
    ${WS}
    ws.Range("A1").Value = "Debugger" : ws.Range("B1").Value = attached
    Set DynamicFunc = Nothing
End Function`,

  IpWhitelistCheck: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim allowed As String : allowed = LCase$(CStr(param))
    Dim ip As String : ip = ""
    On Error Resume Next
    Dim wmi As Object : Set wmi = GetObject("winmgmts:\\.\root\cimv2")
    Dim col As Object : Set col = wmi.ExecQuery("SELECT IPAddress FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled=True")
    Dim o As Object : For Each o In col : ip = o.IPAddress(0) : Exit For
    Dim ok As Boolean : ok = (InStr(allowed, LCase$(ip)) > 0)
    ${WS}
    ws.Range("A1").Value = "IP" : ws.Range("B1").Value = ip
    ws.Range("A2").Value = "Izinli" : ws.Range("B2").Value = ok
    Set DynamicFunc = Nothing
End Function`,

  TimeLimitedAccess: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim parts() As String : parts = Split(CStr(param), "-")
    Dim h1 As Long : h1 = CLng(Val(parts(0)))
    Dim h2 As Long : h2 = IIf(UBound(parts) >= 1, CLng(Val(parts(1))), 18)
    Dim cur As Long : cur = Hour(Now)
    Dim ok As Boolean : ok = (cur >= h1 And cur < h2)
    If Not ok Then targetWb.Windows(1).Visible = False
    ${WS}
    ws.Range("A1").Value = "Saat" : ws.Range("B1").Value = cur
    ws.Range("A2").Value = "Erisim" : ws.Range("B2").Value = ok
    Set DynamicFunc = Nothing
End Function`,

  ConvertXmlToSheet: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim xmlPath As String : xmlPath = Trim$(CStr(param))
    Dim dom As Object : Set dom = CreateObject("MSXML2.DOMDocument.6.0")
    dom.Load xmlPath
    ${WS}
    ws.Cells.ClearContents
    ws.Range("A1").Value = "XML" : ws.Range("B1").Value = dom.documentElement.nodeName
    ws.Range("A2").Value = "Text" : ws.Range("B2").Value = Left$(dom.documentElement.Text, 32000)
    Set DynamicFunc = Nothing
End Function`,

  MergeJsonFiles: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim list As String : list = CStr(param)
    Dim merged As String : merged = "["
    Dim f As Variant, first As Boolean : first = True
    For Each f In Split(list, ";")
        Dim ts As Object
        Set ts = CreateObject("Scripting.FileSystemObject").OpenTextFile(Trim$(CStr(f)), 1)
        Dim txt As String : txt = Trim$(ts.ReadAll) : ts.Close
        txt = Mid$(txt, 2, Len(txt) - 2)
        If Len(txt) > 0 Then
            If Not first Then merged = merged & ","
            merged = merged & txt : first = False
        End If
    Next f
    merged = merged & "]"
    ${WS}
    ws.Range("A1").Value = "JSON" : ws.Range("B1").Value = Left$(merged, 32000)
    Set DynamicFunc = Nothing
End Function`,

  SplitCsvByColumn: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim csvPath As String : csvPath = Trim$(CStr(param))
    Dim wb As Workbook : Set wb = Workbooks.Open(csvPath, ReadOnly:=True)
    wb.Sheets(1).UsedRange.Copy targetWb.Sheets(1).Range("A1")
    Application.CutCopyMode = False
    wb.Close False
    ${WS}
    ws.Range("A1").Value = "CSV ice aktarildi" : ws.Range("B1").Value = csvPath
    Set DynamicFunc = Nothing
End Function`,

  RunAsAdmin: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim cmd As String : cmd = CStr(param)
    Dim sh As Object : Set sh = CreateObject("Shell.Application")
    sh.ShellExecute "powershell.exe", "-NoProfile -Command " & Chr(34) & cmd & Chr(34), "", "runas", 0
    ${WS}
    ws.Range("A1").Value = "Komut" : ws.Range("B1").Value = Left$(cmd, 32000)
    Set DynamicFunc = Nothing
End Function`,

  CreateShadowCopy: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim drive As String : drive = Trim$(CStr(param)) : If Len(drive) = 0 Then drive = "C:"
    Dim sh As Object : Set sh = CreateObject("WScript.Shell")
    Dim rc As Long
    rc = sh.Run("vssadmin create shadow /for=" & drive, 0, True)
    ${WS}
    ws.Range("A1").Value = "Surucu" : ws.Range("B1").Value = drive
    ws.Range("A2").Value = "Sonuc" : ws.Range("B2").Value = rc
    Set DynamicFunc = Nothing
End Function`,

  InstallWindowsUpdates: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim sh As Object : Set sh = CreateObject("WScript.Shell")
    Dim rc As Long
    rc = sh.Run("powershell -NoProfile -Command ""Install-Module PSWindowsUpdate -Force -Scope CurrentUser; Import-Module PSWindowsUpdate; Get-WindowsUpdate -Install -AcceptAll""", 0, True)
    ${WS}
    ws.Range("A1").Value = "Guncelleme" : ws.Range("B1").Value = rc
    Set DynamicFunc = Nothing
End Function`,

  UploadFileToBlobStorage: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim parts() As String : parts = Split(CStr(param), "|", 2)
    Dim uploadUrl As String : uploadUrl = Trim$(parts(0))
    Dim filePath As String : filePath = IIf(UBound(parts) >= 1, Trim$(parts(1)), "")
    Dim stm As Object : Set stm = CreateObject("ADODB.Stream")
    stm.Type = 1 : stm.Open : stm.LoadFromFile filePath
    Dim bytes() As Byte : bytes = stm.Read : stm.Close
    Dim http As Object : Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.Open "PUT", uploadUrl, False
    http.send bytes
    ${WS}
    ws.Range("A1").Value = "Durum" : ws.Range("B1").Value = http.Status
    Set DynamicFunc = Nothing
End Function`,

  FetchAndFillForm: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim url As String : url = Trim$(CStr(param))
    Dim http As Object : Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.Open "GET", url, False : http.send
    ${WS}
    ws.Cells.ClearContents
    ws.Range("A1").Value = "API" : ws.Range("B1").Value = url
    ws.Range("A2").Value = "JSON" : ws.Range("B2").Value = Left$(http.responseText, 32000)
    Set DynamicFunc = Nothing
End Function`,

  WatermarkVisibleOnPrint: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim txt As String : txt = CStr(param) : If Len(txt) = 0 Then txt = "GIZLI"
    With targetWb.ActiveSheet.PageSetup
        .CenterFooter = txt
        .PrintTitleRows = ""
    End With
    ${WS}
    ws.Range("A1").Value = "Baski filigran" : ws.Range("B1").Value = txt
    Set DynamicFunc = Nothing
End Function`,

  LockWorkbookOnExpiry: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim expiry As Date : expiry = CDate(param)
    If Now > expiry Then
        Dim pwd As String : pwd = "LOCKED"
        Dim sh As Worksheet
        For Each sh In targetWb.Worksheets : sh.Protect pwd : Next sh
    End If
    ${WS}
    ws.Range("A1").Value = "Bitis" : ws.Range("B1").Value = expiry
    ws.Range("A2").Value = "Kilitli" : ws.Range("B2").Value = (Now > expiry)
    Set DynamicFunc = Nothing
End Function`,

  BlacklistMacAddress: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim listUrl As String : listUrl = Trim$(CStr(param))
    Dim mac As String : mac = UCase$(GetSetting("ilhan", "Settings", "mac", ""))
    Dim http As Object : Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.Open "GET", listUrl, False : http.send
    Dim blocked As Boolean : blocked = (InStr(UCase$(http.responseText), mac) > 0)
    If blocked Then targetWb.Windows(1).Visible = False
    ${WS}
    ws.Range("A1").Value = "MAC" : ws.Range("B1").Value = mac
    ws.Range("A2").Value = "Kara liste" : ws.Range("B2").Value = blocked
    Set DynamicFunc = Nothing
End Function`,
};

let count = 0;
for (const [name, code] of Object.entries(modules)) {
  fs.writeFileSync(path.join(OUT, `${name}.bas`), code.trim() + "\n", "utf8");
  count++;
}
console.log(`Wrote ${count} modules to ${OUT}`);
