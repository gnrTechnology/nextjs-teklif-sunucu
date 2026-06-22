/**
 * generate-batch3-modules.mjs
 * Registry, Internet/HTTP, PS eklentileri, UI/Bildirim, Uzman modüller
 * node scripts/generate-batch3-modules.mjs
 */
import { writeFileSync, readFileSync, mkdirSync } from "fs";
import { join, dirname } from "path";
import { fileURLToPath } from "url";

const __dirname = dirname(fileURLToPath(import.meta.url));
const ROOT = join(__dirname, "..");
const SRC  = join(ROOT, "data", "modules-source");
mkdirSync(SRC, { recursive: true });

/* ─── ortak VBA yardımcılar ─────────────────────────── */
const HELPERS = `
Private Function RunPS(cmd As String) As String
    Dim tmp As String : tmp = Environ("TEMP") & "\\ps_out_" & CLng(Timer*1000) & ".txt"
    Dim wsh As Object : Set wsh = CreateObject("WScript.Shell")
    wsh.Run "powershell -NonInteractive -NoProfile -Command " & Chr(34) & cmd & " | Out-File -Encoding UTF8 '" & tmp & "'" & Chr(34), 0, True
    Dim fso As Object : Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(tmp) Then
        Dim f As Object : Set f = fso.OpenTextFile(tmp, 1, False, -1)
        RunPS = f.ReadAll : f.Close : fso.DeleteFile tmp
    End If
End Function

Private Function EscJ(s As String) As String
    s = Replace(s, "\\", "\\\\") : s = Replace(s, Chr(34), "\\""")
    s = Replace(s, Chr(10), "\\n") : s = Replace(s, Chr(13), "")
    EscJ = s
End Function

Private Function ExtractJsonVal(json As String, key As String) As String
    Dim sk As String : sk = Chr(34) & key & Chr(34) & ":"
    Dim p1 As Long : p1 = InStr(1, json, sk, vbTextCompare)
    If p1 = 0 Then Exit Function
    p1 = p1 + Len(sk)
    Do While Mid(json, p1, 1) = " " : p1 = p1 + 1 : Loop
    If Mid(json, p1, 1) = Chr(34) Then
        p1 = p1 + 1 : Dim p2 As Long : p2 = InStr(p1, json, Chr(34))
        If p2 > p1 Then ExtractJsonVal = Mid(json, p1, p2 - p1)
    Else
        Dim p3 As Long : p3 = p1
        Do While p3 <= Len(json)
            If InStr(",}] " & Chr(13) & Chr(10), Mid(json, p3, 1)) > 0 Then Exit Do
            p3 = p3 + 1
        Loop
        ExtractJsonVal = Trim(Mid(json, p1, p3 - p1))
    End If
End Function

Private Sub WriteResult(ws As Worksheet, key As String, val As String, row As Long)
    ws.Cells(row, 1).Value = key
    ws.Cells(row, 2).Value = val
End Sub
`;

/* ─── modül tanımları ─────────────────────────────────── */
const modules = [

  /* ══════════════════════════════════════════════════════
     DONANIM EKLEMELERİ (#26-30)
  ══════════════════════════════════════════════════════ */
  {
    methodName: "GetCpuUsage",
    description: "Anlık CPU kullanım yüzdesi (WMI LoadPercentage)",
    category: "donanim",
    code: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim wmi As Object : Set wmi = GetObject("winmgmts:\\\\.\\root\\cimv2")
    Dim col As Object, obj As Object
    Set col = wmi.ExecQuery("SELECT LoadPercentage FROM Win32_Processor")
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Cells.ClearContents
    ws.Range("A1").Value = "Çekirdek" : ws.Range("B1").Value = "CPU %"
    ws.Range("A1:B1").Font.Bold = True
    Dim r As Long : r = 2
    For Each obj In col
        ws.Cells(r, 1).Value = "CPU " & (r - 1)
        ws.Cells(r, 2).Value = obj.LoadPercentage & " %"
        r = r + 1
    Next
    ws.Columns.AutoFit
    Set DynamicFunc = Nothing
End Function`
  },
  {
    methodName: "GetVirtualMemoryInfo",
    description: "Sayfa dosyası boyutu ve kullanımı",
    category: "donanim",
    code: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim wmi As Object : Set wmi = GetObject("winmgmts:\\\\.\\root\\cimv2")
    Dim col As Object, obj As Object
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Cells.ClearContents
    ws.Range("A1").Value = "Özellik" : ws.Range("B1").Value = "Değer"
    ws.Range("A1:B1").Font.Bold = True
    Dim r As Long : r = 2
    Set col = wmi.ExecQuery("SELECT TotalVirtualMemorySize, FreeVirtualMemory, TotalVisibleMemorySize, FreePhysicalMemory FROM Win32_OperatingSystem")
    For Each obj In col
        ws.Cells(r,1).Value = "Toplam Sanal Bellek"     : ws.Cells(r,2).Value = Format(CLng(obj.TotalVirtualMemorySize)/1024,"#,##0") & " MB" : r=r+1
        ws.Cells(r,1).Value = "Boş Sanal Bellek"        : ws.Cells(r,2).Value = Format(CLng(obj.FreeVirtualMemory)/1024,"#,##0") & " MB"     : r=r+1
        ws.Cells(r,1).Value = "Toplam Fiziksel Bellek"  : ws.Cells(r,2).Value = Format(CLng(obj.TotalVisibleMemorySize)/1024,"#,##0") & " MB" : r=r+1
        ws.Cells(r,1).Value = "Boş Fiziksel Bellek"     : ws.Cells(r,2).Value = Format(CLng(obj.FreePhysicalMemory)/1024,"#,##0") & " MB"    : r=r+1
    Next
    ws.Columns.AutoFit
    Set DynamicFunc = Nothing
End Function`
  },
  {
    methodName: "GetHardwareSerial",
    description: "Bilgisayar seri numarası (Win32_ComputerSystemProduct)",
    category: "donanim",
    code: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim wmi As Object : Set wmi = GetObject("winmgmts:\\\\.\\root\\cimv2")
    Dim col As Object, obj As Object
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Cells.ClearContents
    ws.Range("A1").Value = "Özellik" : ws.Range("B1").Value = "Değer"
    ws.Range("A1:B1").Font.Bold = True
    Dim r As Long : r = 2
    Set col = wmi.ExecQuery("SELECT Name, IdentifyingNumber, Vendor, Version, UUID FROM Win32_ComputerSystemProduct")
    For Each obj In col
        ws.Cells(r,1).Value = "Ürün Adı"      : ws.Cells(r,2).Value = obj.Name             : r=r+1
        ws.Cells(r,1).Value = "Seri No"        : ws.Cells(r,2).Value = obj.IdentifyingNumber : r=r+1
        ws.Cells(r,1).Value = "Üretici"        : ws.Cells(r,2).Value = obj.Vendor           : r=r+1
        ws.Cells(r,1).Value = "Sürüm"          : ws.Cells(r,2).Value = obj.Version          : r=r+1
        ws.Cells(r,1).Value = "UUID"            : ws.Cells(r,2).Value = obj.UUID             : r=r+1
    Next
    ws.Columns.AutoFit
    Set DynamicFunc = Nothing
End Function`
  },

  /* ══════════════════════════════════════════════════════
     KAYIT DEFTERİ (#31-48)
  ══════════════════════════════════════════════════════ */
  {
    methodName: "ReadRegistryValue",
    description: "param=HKCU\\App\\Key ile registry değeri okur, hücreye yazar",
    category: "registry",
    code: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim path As String : path = Trim(CStr(param))
    Dim wsh As Object : Set wsh = CreateObject("WScript.Shell")
    On Error Resume Next
    Dim val As String : val = wsh.RegRead(path)
    Dim errTxt As String : If Err.Number <> 0 Then errTxt = "HATA: " & Err.Description
    On Error GoTo 0
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Range("A1").Value = "Yol"   : ws.Range("B1").Value = path
    ws.Range("A2").Value = "Değer" : ws.Range("B2").Value = IIf(errTxt <> "", errTxt, val)
    ws.Columns.AutoFit
    Set DynamicFunc = Nothing
End Function`
  },
  {
    methodName: "WriteRegistryValue",
    description: 'param={"path":"HKCU\\\\App\\\\Key","name":"KeyName","value":"Val","type":"REG_SZ"} ile yazar',
    category: "registry",
    code: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim p As String : p = CStr(param)
    Dim regPath As String : regPath = ExtractJsonVal(p, "path")
    Dim name    As String : name    = ExtractJsonVal(p, "name")
    Dim val     As String : val     = ExtractJsonVal(p, "value")
    Dim regType As String : regType = ExtractJsonVal(p, "type")
    If regType = "" Then regType = "REG_SZ"
    Dim fullKey As String : fullKey = regPath & "\\" & name
    Dim wsh As Object : Set wsh = CreateObject("WScript.Shell")
    On Error Resume Next
    wsh.RegWrite fullKey, val, regType
    Dim ok As Boolean : ok = (Err.Number = 0)
    On Error GoTo 0
    targetWb.Sheets(1).Range("A1").Value = IIf(ok, "✅ Yazıldı: " & fullKey, "❌ Hata")
    Set DynamicFunc = Nothing
End Function
${HELPERS}`
  },
  {
    methodName: "DeleteRegistryKey",
    description: "Belirtilen registry anahtarını ve alt anahtarlarını siler",
    category: "registry",
    code: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim path As String : path = Trim(CStr(param))
    Dim wsh As Object : Set wsh = CreateObject("WScript.Shell")
    On Error Resume Next
    wsh.RegDelete path
    Dim ok As Boolean : ok = (Err.Number = 0)
    On Error GoTo 0
    targetWb.Sheets(1).Range("A1").Value = IIf(ok, "✅ Silindi: " & path, "❌ Hata silme")
    Set DynamicFunc = Nothing
End Function`
  },
  {
    methodName: "CheckRegistryKeyExists",
    description: "Belirtilen registry anahtarının var olup olmadığını kontrol eder",
    category: "registry",
    code: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim path As String : path = Trim(CStr(param))
    Dim wsh As Object : Set wsh = CreateObject("WScript.Shell")
    Dim exists As Boolean : exists = False
    On Error Resume Next
    Dim v As String : v = wsh.RegRead(path) : exists = (Err.Number = 0)
    On Error GoTo 0
    targetWb.Sheets(1).Range("A1").Value = IIf(exists, "✅ Mevcut: " & path, "❌ Bulunamadı: " & path)
    Set DynamicFunc = Nothing
End Function`
  },
  {
    methodName: "GetAllVbaSettings",
    description: "ilhan/scngnr/sercan VBA bölümlerini sayfaya döker",
    category: "registry",
    code: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim sections() As String : sections = Split("ilhan,scngnr,sercan", ",")
    Dim keys() As String : keys = Split("mac,mdip,TBveren,teklifYolu,startingAddin,ihlalDosyaYolu,ihlalDosyaAdi,apiBaseUrl", ",")
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Cells.ClearContents
    ws.Range("A1").Value = "Bölüm" : ws.Range("B1").Value = "Anahtar" : ws.Range("C1").Value = "Değer"
    ws.Range("A1:C1").Font.Bold = True
    Dim r As Long : r = 2
    Dim sec As Variant, k As Variant
    For Each sec In sections
        For Each k In keys
            Dim v As String : v = GetSetting(CStr(sec), "Settings", CStr(k), "(yok)")
            ws.Cells(r,1).Value = sec : ws.Cells(r,2).Value = k : ws.Cells(r,3).Value = v : r=r+1
        Next k
    Next sec
    ws.Columns.AutoFit
    Set DynamicFunc = Nothing
End Function`
  },
  {
    methodName: "GetStartupPrograms",
    description: "HKCU/HKLM Run anahtarlarındaki başlangıç programlarını listeler",
    category: "registry",
    code: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Cells.ClearContents
    ws.Range("A1").Value = "Konum" : ws.Range("B1").Value = "Ad" : ws.Range("C1").Value = "Komut"
    ws.Range("A1:C1").Font.Bold = True
    Dim r As Long : r = 2
    Dim paths(3) As String
    paths(0) = "HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\Run\\"
    paths(1) = "HKLM\\Software\\Microsoft\\Windows\\CurrentVersion\\Run\\"
    paths(2) = "HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\RunOnce\\"
    paths(3) = "HKLM\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\RunOnce\\"
    Dim wsh As Object : Set wsh = CreateObject("WScript.Shell")
    Dim i As Long
    For i = 0 To 3
        On Error Resume Next
        Dim val As String : val = wsh.RegRead(paths(i))
        On Error GoTo 0
    Next i
    ' PowerShell ile al
    Dim out As String
    out = RunPS("Get-ItemProperty 'HKLM:\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Run' | Select-Object -Property * | Format-List")
    ws.Cells(r,1).Value = "HKLM Run" : ws.Cells(r,2).Value = "..." : ws.Cells(r,3).Value = Left(out, 200) : r=r+1
    out = RunPS("Get-ItemProperty 'HKCU:\\Software\\Microsoft\\Windows\\CurrentVersion\\Run' | Select-Object -Property * | Format-List")
    ws.Cells(r,1).Value = "HKCU Run" : ws.Cells(r,2).Value = "..." : ws.Cells(r,3).Value = Left(out, 200) : r=r+1
    ws.Columns.AutoFit
    Set DynamicFunc = Nothing
End Function
${HELPERS}`
  },
  {
    methodName: "SetRunOnceCommand",
    description: "HKCU RunOnce'a başlangıç komutu ekler",
    category: "registry",
    code: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    ' param: {"name":"TaskName","command":"cmd /c ..."}
    Dim p As String : p = CStr(param)
    Dim name As String : name = ExtractJsonVal(p, "name")
    Dim cmd  As String : cmd  = ExtractJsonVal(p, "command")
    If name = "" Or cmd = "" Then
        targetWb.Sheets(1).Range("A1").Value = "Hata: name ve command zorunludur"
        Set DynamicFunc = Nothing : Exit Function
    End If
    Dim wsh As Object : Set wsh = CreateObject("WScript.Shell")
    wsh.RegWrite "HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\RunOnce\\" & name, cmd, "REG_SZ"
    targetWb.Sheets(1).Range("A1").Value = "✅ RunOnce eklendi: " & name
    Set DynamicFunc = Nothing
End Function
${HELPERS}`
  },
  {
    methodName: "AddStartupProgram",
    description: "HKCU Run anahtarına başlangıç programı ekler",
    category: "registry",
    code: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    ' param: {"name":"AppName","command":"C:\\path\\app.exe"}
    Dim p As String : p = CStr(param)
    Dim name As String : name = ExtractJsonVal(p, "name")
    Dim cmd  As String : cmd  = ExtractJsonVal(p, "command")
    If name = "" Or cmd = "" Then
        targetWb.Sheets(1).Range("A1").Value = "Hata: name ve command zorunludur"
        Set DynamicFunc = Nothing : Exit Function
    End If
    Dim wsh As Object : Set wsh = CreateObject("WScript.Shell")
    wsh.RegWrite "HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\Run\\" & name, cmd, "REG_SZ"
    targetWb.Sheets(1).Range("A1").Value = "✅ Başlangıç programı eklendi: " & name
    Set DynamicFunc = Nothing
End Function
${HELPERS}`
  },
  {
    methodName: "RemoveStartupProgram",
    description: "HKCU Run başlangıç programını kaldırır",
    category: "registry",
    code: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim name As String : name = Trim(CStr(param))
    Dim wsh As Object : Set wsh = CreateObject("WScript.Shell")
    On Error Resume Next
    wsh.RegDelete "HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\Run\\" & name
    Dim ok As Boolean : ok = (Err.Number = 0)
    On Error GoTo 0
    targetWb.Sheets(1).Range("A1").Value = IIf(ok, "✅ Kaldırıldı: " & name, "❌ Bulunamadı")
    Set DynamicFunc = Nothing
End Function`
  },
  {
    methodName: "GetInstalledAppPaths",
    description: "HKLM App Paths'tan uygulama yollarını listeler",
    category: "registry",
    code: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Cells.ClearContents
    ws.Range("A1").Value = "Uygulama" : ws.Range("B1").Value = "Yol"
    ws.Range("A1:B1").Font.Bold = True
    Dim out As String
    out = RunPS("Get-ChildItem 'HKLM:\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\App Paths' | ForEach-Object { [pscustomobject]@{App=$_.PSChildName; Path=(Get-ItemProperty $_.PSPath).'(default)'} } | ConvertTo-Csv -NoTypeInformation")
    Dim lines() As String : lines = Split(out, Chr(10))
    Dim r As Long : r = 2
    Dim i As Long
    For i = 1 To UBound(lines)
        Dim ln As String : ln = Trim(lines(i))
        If Len(ln) > 2 Then
            ln = Replace(ln, Chr(34), "")
            Dim parts() As String : parts = Split(ln, ",")
            If UBound(parts) >= 1 Then
                ws.Cells(r,1).Value = Trim(parts(0))
                ws.Cells(r,2).Value = Trim(parts(1))
                r = r + 1
            End If
        End If
    Next i
    ws.Columns.AutoFit
    Set DynamicFunc = Nothing
End Function
${HELPERS}`
  },

  /* ══════════════════════════════════════════════════════
     İNTERNET / HTTP (#81-110)
  ══════════════════════════════════════════════════════ */
  {
    methodName: "HttpGetJson",
    description: "URL'den JSON indirir, parse eder, sayfaya yazar",
    category: "internet",
    code: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim url As String : url = Trim(CStr(param))
    Dim http As Object : Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.Open "GET", url, False
    http.setTimeouts 5000, 10000, 30000, 30000
    http.send
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Cells.ClearContents
    If http.Status = 200 Then
        ws.Range("A1").Value = "URL"     : ws.Range("B1").Value = url
        ws.Range("A2").Value = "Durum"   : ws.Range("B2").Value = http.Status
        ws.Range("A3").Value = "Yanıt"   : ws.Range("B3").Value = Left(http.responseText, 32000)
        ws.Range("A3").Font.Bold = True
    Else
        ws.Range("A1").Value = "Hata: HTTP " & http.Status
    End If
    ws.Columns.AutoFit
    Set DynamicFunc = Nothing
End Function`
  },
  {
    methodName: "HttpPostJson",
    description: 'param={"url":"...","body":{"k":"v"}} - JSON body ile POST atar',
    category: "internet",
    code: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim p As String : p = CStr(param)
    Dim url As String : url = ExtractJsonVal(p, "url")
    Dim body As String : body = ExtractBodyJson(p)
    If url = "" Then url = p
    Dim http As Object : Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.Open "POST", url, False
    http.setRequestHeader "Content-Type", "application/json"
    http.setTimeouts 5000, 10000, 30000, 30000
    http.send body
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Cells.ClearContents
    ws.Range("A1").Value = "Durum"  : ws.Range("B1").Value = http.Status
    ws.Range("A2").Value = "Yanıt"  : ws.Range("B2").Value = Left(http.responseText, 32000)
    ws.Columns.AutoFit
    Set DynamicFunc = Nothing
End Function

Private Function ExtractBodyJson(json As String) As String
    Dim sk As String : sk = Chr(34) & "body" & Chr(34) & ":"
    Dim p1 As Long : p1 = InStr(json, sk)
    If p1 = 0 Then ExtractBodyJson = "{}" : Exit Function
    p1 = p1 + Len(sk)
    Do While Mid(json, p1, 1) = " " : p1 = p1 + 1 : Loop
    If Mid(json, p1, 1) <> "{" Then ExtractBodyJson = "{}" : Exit Function
    Dim depth As Long : depth = 0 : Dim p2 As Long : p2 = p1
    Do While p2 <= Len(json)
        If Mid(json, p2, 1) = "{" Then depth = depth + 1
        If Mid(json, p2, 1) = "}" Then depth = depth - 1 : If depth = 0 Then ExtractBodyJson = Mid(json, p1, p2-p1+1) : Exit Function
        p2 = p2 + 1
    Loop
    ExtractBodyJson = "{}"
End Function
${HELPERS}`
  },
  {
    methodName: "HttpDownloadFile",
    description: "Dosyayı ADODB.Stream ile binary olarak indirir",
    category: "internet",
    code: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    ' param: {"url":"https://...","savePath":"C:\\path\\file.ext"}
    Dim p As String : p = CStr(param)
    Dim url  As String : url      = ExtractJsonVal(p, "url")
    Dim dest As String : dest     = ExtractJsonVal(p, "savePath")
    If url = "" Then url = p
    If dest = "" Then dest = Environ("TEMP") & "\\downloaded_" & CLng(Timer*1000)
    Dim http As Object : Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.Open "GET", url, False
    http.setTimeouts 5000, 10000, 60000, 60000
    http.send
    If http.Status = 200 Then
        Dim st As Object : Set st = CreateObject("ADODB.Stream")
        st.Type = 1 : st.Open : st.Write http.responseBody
        st.SaveToFile dest, 2 : st.Close
        targetWb.Sheets(1).Range("A1").Value = "✅ İndirildi: " & dest
    Else
        targetWb.Sheets(1).Range("A1").Value = "❌ HTTP " & http.Status
    End If
    Set DynamicFunc = Nothing
End Function
${HELPERS}`
  },
  {
    methodName: "HttpGetText",
    description: "URL'den düz metin yanıt alır, hücreye yazar",
    category: "internet",
    code: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim url As String : url = Trim(CStr(param))
    Dim http As Object : Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.Open "GET", url, False
    http.setTimeouts 5000, 10000, 30000, 30000
    http.send
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Cells.ClearContents
    If http.Status = 200 Then
        ws.Range("A1").Value = http.responseText
    Else
        ws.Range("A1").Value = "HTTP " & http.Status
    End If
    Set DynamicFunc = Nothing
End Function`
  },
  {
    methodName: "CheckUrlReachable",
    description: "URL'ye HEAD isteği atarak erişilebilirliği test eder",
    category: "internet",
    code: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim url As String : url = Trim(CStr(param))
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Cells.ClearContents
    On Error Resume Next
    Dim http As Object : Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.Open "HEAD", url, False
    http.setTimeouts 3000, 5000, 10000, 10000
    http.send
    Dim status As Long : status = http.Status
    On Error GoTo 0
    ws.Range("A1").Value = "URL"    : ws.Range("B1").Value = url
    ws.Range("A2").Value = "Durum"  : ws.Range("B2").Value = IIf(status >= 200 And status < 400, "✅ Erişilebilir (HTTP " & status & ")", "❌ Erişilemiyor (" & status & ")")
    ws.Columns.AutoFit
    Set DynamicFunc = Nothing
End Function`
  },
  {
    methodName: "CheckInternetConnection",
    description: "İnternet bağlantısını test eder",
    category: "internet",
    code: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Cells.ClearContents
    On Error Resume Next
    Dim http As Object : Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.Open "GET", "http://www.msftconnecttest.com/connecttest.txt", False
    http.setTimeouts 3000, 5000, 8000, 8000
    http.send
    Dim ok As Boolean : ok = (http.Status = 200)
    On Error GoTo 0
    ws.Range("A1").Value = "İnternet Bağlantısı"
    ws.Range("B1").Value = IIf(ok, "✅ Bağlı", "❌ Bağlı Değil")
    ' Genel gecikme (ping-benzeri)
    Dim t1 As Double : t1 = Timer
    On Error Resume Next
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.Open "GET", "https://api.ipify.org", False
    http.setTimeouts 3000, 5000, 8000, 8000
    http.send
    Dim t2 As Double : t2 = Timer
    On Error GoTo 0
    ws.Range("A2").Value = "Yanıt Süresi" : ws.Range("B2").Value = Format((t2-t1)*1000,"0") & " ms"
    ws.Columns.AutoFit
    Set DynamicFunc = Nothing
End Function`
  },
  {
    methodName: "PingHost",
    description: "WMI Win32_PingStatus ile host'a ping atar, ms döner",
    category: "internet",
    code: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim host As String : host = Trim(CStr(param))
    If host = "" Then host = "8.8.8.8"
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Cells.ClearContents
    ws.Range("A1").Value = "Host" : ws.Range("B1").Value = host
    ws.Range("A1:B1").Font.Bold = True
    Dim wmi As Object : Set wmi = GetObject("winmgmts:\\\\.\\root\\cimv2")
    Dim col As Object, obj As Object
    Dim r As Long : r = 2
    Set col = wmi.ExecQuery("SELECT * FROM Win32_PingStatus WHERE Address='" & host & "'")
    For Each obj In col
        ws.Cells(r,1).Value = "Durum"       : ws.Cells(r,2).Value = IIf(obj.StatusCode=0,"✅ Başarılı","❌ Başarısız (" & obj.StatusCode & ")") : r=r+1
        ws.Cells(r,1).Value = "Yanıt ms"    : ws.Cells(r,2).Value = obj.ResponseTime & " ms" : r=r+1
        ws.Cells(r,1).Value = "TTL"         : ws.Cells(r,2).Value = obj.TimeToLive : r=r+1
        ws.Cells(r,1).Value = "Paket Boyutu": ws.Cells(r,2).Value = obj.BufferSize & " byte" : r=r+1
    Next
    ws.Columns.AutoFit
    Set DynamicFunc = Nothing
End Function`
  },
  {
    methodName: "GetExchangeRate",
    description: "TCMB API'den anlık döviz kurlarını çeker (USD, EUR, GBP)",
    category: "internet",
    code: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Cells.ClearContents
    ws.Range("A1").Value = "Döviz" : ws.Range("B1").Value = "Alış" : ws.Range("C1").Value = "Satış"
    ws.Range("A1:C1").Font.Bold = True
    ' exchangerate-api (ücretsiz, kayıt gerektirmeyen)
    On Error Resume Next
    Dim http As Object : Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.Open "GET", "https://api.exchangerate-api.com/v4/latest/TRY", False
    http.setTimeouts 5000, 10000, 20000, 20000
    http.send
    If http.Status = 200 Then
        Dim resp As String : resp = http.responseText
        ' Basit JSON parse - rates içinden USD, EUR, GBP
        Dim currencies() As String : currencies = Split("USD,EUR,GBP,CHF,JPY", ",")
        Dim r As Long : r = 2
        Dim c As Variant
        For Each c In currencies
            Dim key As String : key = Chr(34) & c & Chr(34) & ":"
            Dim p1 As Long : p1 = InStr(resp, key)
            If p1 > 0 Then
                p1 = p1 + Len(key)
                Dim p2 As Long : p2 = p1
                Do While InStr(",}", Mid(resp,p2,1)) = 0 : p2=p2+1 : Loop
                Dim rate As Double : rate = CDbl(Trim(Mid(resp,p1,p2-p1)))
                ws.Cells(r,1).Value = c
                ws.Cells(r,2).Value = Format(1/rate,"0.0000") & " TL"
                r = r + 1
            End If
        Next c
        ws.Cells(r,1).Value = "Güncelleme" : ws.Cells(r,2).Value = Now()
    Else
        ws.Range("A2").Value = "Hata: HTTP " & http.Status
    End If
    On Error GoTo 0
    ws.Columns.AutoFit
    Set DynamicFunc = Nothing
End Function`
  },
  {
    methodName: "SendTelegramMessage",
    description: 'param={"token":"BOT_TOKEN","chatId":"CHAT_ID","text":"Mesaj"} - Telegram Bot API ile mesaj gönderir',
    category: "internet",
    code: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim p As String : p = CStr(param)
    Dim token  As String : token  = ExtractJsonVal(p, "token")
    Dim chatId As String : chatId = ExtractJsonVal(p, "chatId")
    Dim text   As String : text   = ExtractJsonVal(p, "text")
    If token = "" Or chatId = "" Then
        targetWb.Sheets(1).Range("A1").Value = "Hata: token ve chatId zorunludur"
        Set DynamicFunc = Nothing : Exit Function
    End If
    Dim url As String : url = "https://api.telegram.org/bot" & token & "/sendMessage"
    Dim body As String : body = "{" & Chr(34) & "chat_id" & Chr(34) & ":" & Chr(34) & chatId & Chr(34) & "," & Chr(34) & "text" & Chr(34) & ":" & Chr(34) & EscJ(text) & Chr(34) & "}"
    On Error Resume Next
    Dim http As Object : Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.Open "POST", url, False
    http.setRequestHeader "Content-Type", "application/json"
    http.setTimeouts 5000, 10000, 15000, 15000
    http.send body
    Dim ok As Boolean : ok = (http.Status = 200)
    On Error GoTo 0
    targetWb.Sheets(1).Range("A1").Value = IIf(ok, "✅ Telegram mesajı gönderildi", "❌ Hata: " & http.Status)
    Set DynamicFunc = Nothing
End Function
${HELPERS}`
  },
  {
    methodName: "SendErrorReportToServer",
    description: "Hata bilgisini JSON ile sunucuya (/api/module-output) gönderir",
    category: "internet",
    code: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim errMsg As String : errMsg = Trim(CStr(param))
    Dim baseUrl As String : baseUrl = GetSetting("ilhan","Settings","apiBaseUrl","https://nextjs-teklif-sunucu.vercel.app")
    If Right(baseUrl,1) = "/" Then baseUrl = Left(baseUrl,Len(baseUrl)-1)
    Dim mac      As String : mac      = GetSetting("ilhan","Settings","mac","UNKNOWN")
    Dim hostname As String : hostname = Environ("COMPUTERNAME")
    Dim body As String
    body = "{" & Chr(34) & "mac" & Chr(34) & ":" & Chr(34) & EscJ(mac) & Chr(34) & "," & _
           Chr(34) & "moduleName" & Chr(34) & ":" & Chr(34) & "ErrorReport" & Chr(34) & "," & _
           Chr(34) & "hostname" & Chr(34) & ":" & Chr(34) & EscJ(hostname) & Chr(34) & "," & _
           Chr(34) & "output" & Chr(34) & ":{" & Chr(34) & "error" & Chr(34) & ":" & Chr(34) & EscJ(errMsg) & Chr(34) & "}}"
    On Error Resume Next
    Dim http As Object : Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.Open "POST", baseUrl & "/api/module-output", False
    http.setRequestHeader "Content-Type", "application/json"
    http.setTimeouts 5000,5000,15000,15000
    http.send body
    On Error GoTo 0
    targetWb.Sheets(1).Range("A1").Value = "✅ Hata raporu gönderildi"
    Set DynamicFunc = Nothing
End Function
${HELPERS}`
  },
  {
    methodName: "GetWeatherData",
    description: 'param={"city":"Istanbul","apiKey":"API_KEY"} - OpenWeatherMap hava durumu çeker',
    category: "internet",
    code: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim p As String : p = CStr(param)
    Dim city   As String : city   = ExtractJsonVal(p, "city")
    Dim apiKey As String : apiKey = ExtractJsonVal(p, "apiKey")
    If city = "" Then city = "Istanbul"
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Cells.ClearContents
    If apiKey = "" Then
        ws.Range("A1").Value = "Hata: OpenWeatherMap apiKey gerekli (openweathermap.org'dan ücretsiz alınır)"
        Set DynamicFunc = Nothing : Exit Function
    End If
    Dim url As String : url = "https://api.openweathermap.org/data/2.5/weather?q=" & city & "&appid=" & apiKey & "&units=metric&lang=tr"
    On Error Resume Next
    Dim http As Object : Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.Open "GET", url, False : http.setTimeouts 5000,10000,20000,20000 : http.send
    If http.Status = 200 Then
        Dim resp As String : resp = http.responseText
        ws.Range("A1").Value = "Şehir"        : ws.Range("B1").Value = city
        ws.Range("A2").Value = "Sıcaklık"     : ws.Range("B2").Value = ExtractJsonVal(ExtractJsonVal(resp,"main"),"temp") & " °C"
        ws.Range("A3").Value = "Nem"          : ws.Range("B3").Value = ExtractJsonVal(ExtractJsonVal(resp,"main"),"humidity") & " %"
        ws.Range("A4").Value = "Rüzgar"       : ws.Range("B4").Value = ExtractJsonVal(ExtractJsonVal(resp,"wind"),"speed") & " m/s"
        ws.Range("A5").Value = "Tarih"        : ws.Range("B5").Value = Now()
    Else
        ws.Range("A1").Value = "Hata: HTTP " & http.Status & " - " & Left(http.responseText,100)
    End If
    On Error GoTo 0
    ws.Columns.AutoFit
    Set DynamicFunc = Nothing
End Function
${HELPERS}`
  },
  {
    methodName: "DownloadAndOpenExcel",
    description: "URL'den Excel dosyasını indirir ve açar",
    category: "internet",
    code: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim url As String : url = Trim(CStr(param))
    Dim savePath As String : savePath = Environ("TEMP") & "\\downloaded_" & CLng(Timer*1000) & ".xlsx"
    On Error Resume Next
    Dim http As Object : Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.Open "GET", url, False : http.setTimeouts 5000,10000,60000,60000 : http.send
    If http.Status = 200 Then
        Dim st As Object : Set st = CreateObject("ADODB.Stream")
        st.Type = 1 : st.Open : st.Write http.responseBody
        st.SaveToFile savePath, 2 : st.Close
        Workbooks.Open savePath
        targetWb.Sheets(1).Range("A1").Value = "✅ Açıldı: " & savePath
    Else
        targetWb.Sheets(1).Range("A1").Value = "❌ HTTP " & http.Status
    End If
    On Error GoTo 0
    Set DynamicFunc = Nothing
End Function`
  },

  /* ══════════════════════════════════════════════════════
     POWERSHELL EKLEMELERİ
  ══════════════════════════════════════════════════════ */
  {
    methodName: "RunPsScript",
    description: ".ps1 dosyasını çalıştırır, çıktıyı hücreye yazar",
    category: "powershell",
    code: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim scriptPath As String : scriptPath = Trim(CStr(param))
    If Dir(scriptPath) = "" Then
        targetWb.Sheets(1).Range("A1").Value = "Hata: Dosya bulunamadı: " & scriptPath
        Set DynamicFunc = Nothing : Exit Function
    End If
    Dim out As String : out = RunPS("-ExecutionPolicy Bypass -File """ & scriptPath & """")
    targetWb.Sheets(1).Range("A1").Value = Left(Trim(out), 32000)
    Set DynamicFunc = Nothing
End Function
${HELPERS}`
  },
  {
    methodName: "GetPsOutputToSheet",
    description: "PowerShell çıktısını satır satır sayfaya yazar",
    category: "powershell",
    code: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim cmd As String : cmd = Trim(CStr(param))
    Dim out As String : out = RunPS(cmd)
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Cells.ClearContents
    Dim lines() As String : lines = Split(out, Chr(10))
    Dim r As Long : r = 1
    Dim i As Long
    For i = 0 To UBound(lines)
        Dim ln As String : ln = Trim(Replace(lines(i), Chr(13),""))
        If Len(ln) > 0 Then
            ws.Cells(r, 1).Value = ln
            r = r + 1
        End If
    Next i
    Set DynamicFunc = Nothing
End Function
${HELPERS}`
  },
  {
    methodName: "SetPsExecutionPolicy",
    description: "PowerShell ExecutionPolicy ayarlar (RemoteSigned, Unrestricted, vb.)",
    category: "powershell",
    code: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim policy As String : policy = Trim(CStr(param))
    If policy = "" Then policy = "RemoteSigned"
    Dim out As String : out = RunPS("Set-ExecutionPolicy " & policy & " -Scope CurrentUser -Force")
    Dim verify As String : verify = RunPS("Get-ExecutionPolicy -Scope CurrentUser")
    targetWb.Sheets(1).Range("A1").Value = "ExecutionPolicy: " & Trim(verify)
    Set DynamicFunc = Nothing
End Function
${HELPERS}`
  },
  {
    methodName: "GetWindowsUpdateList",
    description: "Bekleyen Windows güncellemelerini listeler",
    category: "powershell",
    code: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Cells.ClearContents
    ws.Range("A1").Value = "Güncelleme" : ws.Range("B1").Value = "KB" : ws.Range("C1").Value = "Boyut"
    ws.Range("A1:C1").Font.Bold = True
    ' Get-WindowsUpdate gerektirmeden temel kontrol
    Dim out As String
    out = RunPS("(New-Object -ComObject Microsoft.Update.Searcher).Search('IsInstalled=0 and Type=Software').Updates | Select-Object Title,@{N='KB';E={($_.KBArticleIDs -join ',')}},@{N='Size';E={[math]::Round($_.MaxDownloadSize/1MB,1)}} | ConvertTo-Csv -NoTypeInformation")
    Dim lines() As String : lines = Split(out, Chr(10))
    Dim r As Long : r = 2
    Dim i As Long
    For i = 1 To UBound(lines)
        Dim ln As String : ln = Trim(Replace(lines(i), Chr(13),""))
        If Len(ln) > 2 Then
            ln = Replace(ln, Chr(34), "")
            Dim parts() As String : parts = Split(ln, ",")
            If UBound(parts) >= 2 Then
                ws.Cells(r,1).Value = parts(0) : ws.Cells(r,2).Value = parts(1) : ws.Cells(r,3).Value = parts(2) & " MB"
                r = r + 1
            End If
        End If
    Next i
    If r = 2 Then ws.Range("A2").Value = "✅ Bekleyen güncelleme yok veya yetki gerekiyor"
    ws.Columns.AutoFit
    Set DynamicFunc = Nothing
End Function
${HELPERS}`
  },
  {
    methodName: "ResetNetworkAdapter",
    description: "Ağ adaptörünü devre dışı bırakıp yeniden etkinleştirir",
    category: "powershell",
    code: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim adapterName As String : adapterName = Trim(CStr(param))
    If adapterName = "" Then adapterName = "Ethernet"
    Dim out As String
    out = RunPS("$a=Get-NetAdapter -Name '" & adapterName & "' -ErrorAction SilentlyContinue; if($a){Disable-NetAdapter -Name '" & adapterName & "' -Confirm:$false; Start-Sleep 2; Enable-NetAdapter -Name '" & adapterName & "' -Confirm:$false; 'OK'} else {'Adaptör bulunamadı: " & adapterName & "'}")
    targetWb.Sheets(1).Range("A1").Value = "Adaptör: " & adapterName & " → " & Trim(out)
    Set DynamicFunc = Nothing
End Function
${HELPERS}`
  },
  {
    methodName: "SetStaticIp",
    description: 'param={"adapter":"Ethernet","ip":"192.168.1.100","mask":"255.255.255.0","gw":"192.168.1.1","dns":"8.8.8.8"} - statik IP atar',
    category: "powershell",
    code: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim p As String : p = CStr(param)
    Dim adapter As String : adapter = ExtractJsonVal(p,"adapter")
    Dim ip      As String : ip      = ExtractJsonVal(p,"ip")
    Dim mask    As String : mask    = ExtractJsonVal(p,"mask")
    Dim gw      As String : gw      = ExtractJsonVal(p,"gw")
    Dim dns     As String : dns     = ExtractJsonVal(p,"dns")
    If adapter="" Then adapter="Ethernet"
    If mask=""    Then mask="255.255.255.0"
    Dim prefix  As Long : prefix = 24
    ' mask'tan prefix hesapla basit
    Dim out As String
    out = RunPS("$a='" & adapter & "'; $i='" & ip & "'; $g='" & gw & "'; $d='" & dns & "'; netsh interface ip set address $a static $i " & mask & " $g; netsh interface ip set dns $a static $d; 'Tamamlandı'")
    targetWb.Sheets(1).Range("A1").Value = Trim(out)
    Set DynamicFunc = Nothing
End Function
${HELPERS}`
  },
  {
    methodName: "EnableRemoteDesktop",
    description: "RDP kaydını ve servisini aktif eder",
    category: "powershell",
    code: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim out As String
    out = RunPS("Set-ItemProperty -Path 'HKLM:\\System\\CurrentControlSet\\Control\\Terminal Server' -Name 'fDenyTSConnections' -Value 0; Enable-NetFirewallRule -DisplayGroup 'Remote Desktop'; 'RDP etkinleştirildi'")
    targetWb.Sheets(1).Range("A1").Value = Trim(out)
    Set DynamicFunc = Nothing
End Function
${HELPERS}`
  },
  {
    methodName: "GetInstalledDrivers",
    description: "Yüklü sürücüleri (InfName, Sürüm, Tarih) listeler",
    category: "powershell",
    code: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Cells.ClearContents
    ws.Range("A1").Value = "Sürücü" : ws.Range("B1").Value = "Sağlayıcı" : ws.Range("C1").Value = "Sürüm" : ws.Range("D1").Value = "Tarih"
    ws.Range("A1:D1").Font.Bold = True
    Dim out As String
    out = RunPS("Get-WindowsDriver -Online | Select-Object Driver,ProviderName,Version,Date | ConvertTo-Csv -NoTypeInformation")
    Dim lines() As String : lines = Split(out, Chr(10))
    Dim r As Long : r = 2
    Dim i As Long
    For i = 1 To UBound(lines)
        Dim ln As String : ln = Trim(Replace(lines(i), Chr(13),""))
        If Len(ln) > 2 Then
            ln = Replace(ln, Chr(34), "")
            Dim parts() As String : parts = Split(ln, ",")
            If UBound(parts) >= 3 Then
                ws.Cells(r,1).Value=parts(0): ws.Cells(r,2).Value=parts(1): ws.Cells(r,3).Value=parts(2): ws.Cells(r,4).Value=parts(3): r=r+1
            End If
        End If
    Next i
    ws.Columns.AutoFit
    Set DynamicFunc = Nothing
End Function
${HELPERS}`
  },
  {
    methodName: "GetDiskHealthStatus",
    description: "SMART verisi ile disk sağlık durumunu kontrol eder",
    category: "powershell",
    code: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Cells.ClearContents
    ws.Range("A1").Value = "Disk" : ws.Range("B1").Value = "Boyut" : ws.Range("C1").Value = "Durum" : ws.Range("D1").Value = "Ortam Tipi"
    ws.Range("A1:D1").Font.Bold = True
    Dim out As String
    out = RunPS("Get-PhysicalDisk | Select-Object FriendlyName,Size,HealthStatus,MediaType | ConvertTo-Csv -NoTypeInformation")
    Dim lines() As String : lines = Split(out, Chr(10))
    Dim r As Long : r = 2
    Dim i As Long
    For i = 1 To UBound(lines)
        Dim ln As String : ln = Trim(Replace(lines(i), Chr(13),""))
        If Len(ln) > 2 Then
            ln = Replace(ln, Chr(34), "")
            Dim parts() As String : parts = Split(ln, ",")
            If UBound(parts) >= 3 Then
                Dim sizeGb As String
                On Error Resume Next : sizeGb = Format(CDbl(parts(1))/1073741824,"0.0") & " GB" : On Error GoTo 0
                ws.Cells(r,1).Value=parts(0): ws.Cells(r,2).Value=sizeGb: ws.Cells(r,3).Value=parts(2): ws.Cells(r,4).Value=parts(3): r=r+1
            End If
        End If
    Next i
    If r = 2 Then ws.Range("A2").Value = "Admin hakları gerekebilir"
    ws.Columns.AutoFit
    Set DynamicFunc = Nothing
End Function
${HELPERS}`
  },
  {
    methodName: "GetShadowCopies",
    description: "Volume Shadow Copy listesini döndürür",
    category: "powershell",
    code: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Cells.ClearContents
    ws.Range("A1").Value = "Sürücü" : ws.Range("B1").Value = "Oluşturma" : ws.Range("C1").Value = "Boyut" : ws.Range("D1").Value = "ID"
    ws.Range("A1:D1").Font.Bold = True
    Dim out As String
    out = RunPS("Get-WmiObject Win32_ShadowCopy | Select-Object VolumeName,InstallDate,@{N='SizeMB';E={[math]::Round($_.Count/1MB,0)}},ID | ConvertTo-Csv -NoTypeInformation")
    Dim lines() As String : lines = Split(out, Chr(10))
    Dim r As Long : r = 2
    Dim i As Long
    For i = 1 To UBound(lines)
        Dim ln As String : ln = Trim(Replace(lines(i), Chr(13),""))
        If Len(ln) > 2 Then
            ln = Replace(ln, Chr(34), "")
            Dim parts() As String : parts = Split(ln, ",")
            If UBound(parts) >= 3 Then
                ws.Cells(r,1).Value=parts(0): ws.Cells(r,2).Value=parts(1): ws.Cells(r,3).Value=parts(2) & " MB": ws.Cells(r,4).Value=parts(3): r=r+1
            End If
        End If
    Next i
    If r = 2 Then ws.Range("A2").Value = "Shadow copy bulunamadı veya yetki gerekiyor"
    ws.Columns.AutoFit
    Set DynamicFunc = Nothing
End Function
${HELPERS}`
  },
  {
    methodName: "AddHostsEntry",
    description: 'param={"ip":"1.2.3.4","host":"myserver.local"} - hosts dosyasına satır ekler (admin gerekir)',
    category: "powershell",
    code: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim p As String : p = CStr(param)
    Dim ip   As String : ip   = ExtractJsonVal(p,"ip")
    Dim host As String : host = ExtractJsonVal(p,"host")
    If ip="" Or host="" Then
        targetWb.Sheets(1).Range("A1").Value = "Hata: ip ve host zorunludur"
        Set DynamicFunc = Nothing : Exit Function
    End If
    Dim out As String
    out = RunPS("Add-Content -Path $env:windir\\System32\\drivers\\etc\\hosts -Value ('" & ip & "' + [char]9 + '" & host & "') -Force; 'OK'")
    targetWb.Sheets(1).Range("A1").Value = IIf(Trim(out)="OK","✅ Eklendi: " & ip & " → " & host,"❌ " & Trim(out))
    Set DynamicFunc = Nothing
End Function
${HELPERS}`
  },
  {
    methodName: "ConnectWifi",
    description: 'param={"ssid":"NetworkName","password":"pass"} - Wi-Fi ağına bağlanır',
    category: "powershell",
    code: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim p As String : p = CStr(param)
    Dim ssid As String : ssid = ExtractJsonVal(p,"ssid")
    Dim pass As String : pass = ExtractJsonVal(p,"password")
    If ssid = "" Then ssid = p
    Dim out As String
    If pass <> "" Then
        out = RunPS("$prof=@'<?xml version=""1.0""?><WLANProfile xmlns=""http://www.microsoft.com/networking/WLAN/profile/v1""><name>" & ssid & "</name><SSIDConfig><SSID><name>" & ssid & "</name></SSID></SSIDConfig><connectionType>ESS</connectionType><connectionMode>auto</connectionMode><MSM><security><authEncryption><authentication>WPA2PSK</authentication><encryption>AES</encryption></authEncryption><sharedKey><keyType>passPhrase</keyType><protected>false</protected><keyMaterial>" & pass & "</keyMaterial></sharedKey></security></MSM></WLANProfile>'@; $tmp=New-TemporaryFile; $prof|Out-File $tmp.FullName; netsh wlan add profile filename=$tmp.FullName; Remove-Item $tmp; netsh wlan connect name='" & ssid & "'")
    Else
        out = RunPS("netsh wlan connect name='" & ssid & "'")
    End If
    targetWb.Sheets(1).Range("A1").Value = Trim(out)
    Set DynamicFunc = Nothing
End Function
${HELPERS}`
  },

  /* ══════════════════════════════════════════════════════
     BİLDİRİM & UI (#191-208)
  ══════════════════════════════════════════════════════ */
  {
    methodName: "ShowToastNotification",
    description: "Windows 10/11 baloncuk bildirimi (BurntToast veya XML fallback)",
    category: "bildirim",
    code: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    ' param: {"title":"Başlık","message":"Mesaj"}  veya düz metin
    Dim p As String : p = CStr(param)
    Dim title As String : title = ExtractJsonVal(p,"title")
    Dim msg   As String : msg   = ExtractJsonVal(p,"message")
    If title = "" Then title = "Excel Bildirimi"
    If msg   = "" Then msg   = p
    Dim out As String
    out = RunPS("$xml=[Windows.UI.Notifications.ToastNotificationManager,Windows.UI.Notifications,ContentType=WindowsRuntime];" & _
        "[void]$xml;" & _
        "$t=[Windows.UI.Notifications.ToastTemplateType,Windows.UI.Notifications,ContentType=WindowsRuntime]::ToastText02;" & _
        "$c=[Windows.UI.Notifications.ToastNotificationManager]::GetTemplateContent($t);" & _
        "$c.GetElementsByTagName('text')[0].AppendChild($c.CreateTextNode('" & Replace(title,"'","\`'") & "'));" & _
        "$c.GetElementsByTagName('text')[1].AppendChild($c.CreateTextNode('" & Replace(msg,"'","\`'") & "'));" & _
        "$n=[Windows.UI.Notifications.ToastNotification,Windows.UI.Notifications,ContentType=WindowsRuntime]::new($c);" & _
        "[Windows.UI.Notifications.ToastNotificationManager]::CreateToastNotifier('Excel').Show($n);" & _
        "'OK'")
    If Trim(out) <> "OK" Then
        ' Fallback: WScript popup
        CreateObject("WScript.Shell").Popup msg, 3, title, 64
    End If
    Set DynamicFunc = Nothing
End Function
${HELPERS}`
  },
  {
    methodName: "ShowSystemTrayBalloon",
    description: "WScript.Shell PopUp baloncuğu gösterir (3 saniye)",
    category: "bildirim",
    code: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim p As String : p = CStr(param)
    Dim title As String : title = ExtractJsonVal(p,"title")
    Dim msg   As String : msg   = ExtractJsonVal(p,"message")
    Dim secs  As Long   : secs  = 3
    If title = "" Then title = "Bildirim"
    If msg   = "" Then msg   = p
    Dim secsStr As String : secsStr = ExtractJsonVal(p,"seconds")
    If IsNumeric(secsStr) Then secs = CLng(secsStr)
    CreateObject("WScript.Shell").Popup msg, secs, title, 64
    Set DynamicFunc = Nothing
End Function
${HELPERS}`
  },
  {
    methodName: "OpenUrlInBrowser",
    description: "Varsayılan tarayıcıda URL açar",
    category: "bildirim",
    code: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim url As String : url = Trim(CStr(param))
    If url = "" Then url = "https://www.google.com"
    CreateObject("WScript.Shell").Run url
    targetWb.Sheets(1).Range("A1").Value = "✅ Açıldı: " & url
    Set DynamicFunc = Nothing
End Function`
  },
  {
    methodName: "SetExcelTitleBar",
    description: "Excel başlık çubuğunu özelleştirir",
    category: "bildirim",
    code: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim title As String : title = Trim(CStr(param))
    If title = "" Then title = "Excel - " & targetWb.Name
    Application.Caption = title
    targetWb.Sheets(1).Range("A1").Value = "✅ Başlık: " & title
    Set DynamicFunc = Nothing
End Function`
  },
  {
    methodName: "ShowStatusBarProgress",
    description: "Durum çubuğunda % göstergesiyle işlem sırasında ilerleme gösterir",
    category: "bildirim",
    code: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    ' param: {"steps":5,"message":"İşleniyor..."}
    Dim p As String : p = CStr(param)
    Dim steps   As Long   : steps   = 5
    Dim msg     As String : msg     = "İşleniyor"
    Dim stepsS  As String : stepsS  = ExtractJsonVal(p,"steps")
    Dim msgS    As String : msgS    = ExtractJsonVal(p,"message")
    If IsNumeric(stepsS) Then steps = CLng(stepsS)
    If msgS <> "" Then msg = msgS
    Dim i As Long
    For i = 1 To steps
        Dim pct As Long : pct = CLng(i / steps * 100)
        Application.StatusBar = msg & " [" & String(pct\\5, "=") & String(20-(pct\\5), "-") & "] " & pct & "%"
        Application.Wait Now + TimeValue("00:00:01")
    Next i
    Application.StatusBar = False
    targetWb.Sheets(1).Range("A1").Value = "✅ Tamamlandı"
    Set DynamicFunc = Nothing
End Function
${HELPERS}`
  },
  {
    methodName: "AnimateStatusMessage",
    description: "Durum çubuğunda kayan yazı efekti gösterir",
    category: "bildirim",
    code: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim msg As String : msg = Trim(CStr(param))
    If msg = "" Then msg = "Excel Teklif Sistemi"
    Dim padded As String : padded = Space(40) & msg & Space(40)
    Dim i As Long
    For i = 1 To Len(msg) + 40
        Application.StatusBar = Mid(padded, i, 40)
        Application.Wait Now + TimeValue("00:00:00") + 0.0001
    Next i
    Application.StatusBar = False
    Set DynamicFunc = Nothing
End Function`
  },
  {
    methodName: "PlaySystemSound",
    description: "Windows sistem sesi çalar (Asterisk, Critical, Exclamation, Hand, Question)",
    category: "bildirim",
    code: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim soundType As String : soundType = Trim(CStr(param))
    If soundType = "" Then soundType = "Asterisk"
    Select Case LCase(soundType)
        Case "asterisk":     Application.EnableCancelKey = xlDisabled : Beep
        Case "critical":     MsgBox "", vbCritical + vbOKOnly, "": Exit Function
        Case "question":     MsgBox "", vbQuestion + vbOKOnly, "": Exit Function
        Case "exclamation":  MsgBox "", vbExclamation + vbOKOnly, "": Exit Function
        Case Else
            Dim out As String : out = RunPS("[System.Media.SystemSounds]::" & soundType & ".Play()")
    End Select
    Set DynamicFunc = Nothing
End Function
${HELPERS}`
  },
  {
    methodName: "FlashTaskbarIcon",
    description: "Excel görev çubuğu simgesini yanıp söndürür (dikkat çekme)",
    category: "bildirim",
    code: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    ' Windows API FlashWindow ile
    Declare PtrSafe Function FlashWindow Lib "user32" (ByVal hwnd As LongPtr, ByVal bInvert As Long) As Long
    Dim i As Long
    For i = 1 To 6
        FlashWindow Application.hwnd, 1
        Application.Wait Now + TimeValue("00:00:01")
    Next i
    FlashWindow Application.hwnd, 0
    Set DynamicFunc = Nothing
End Function`
  },

  /* ══════════════════════════════════════════════════════
     UZMAN EKLEME
  ══════════════════════════════════════════════════════ */
  {
    methodName: "SendKeystrokes",
    description: 'param={"keys":"^s","window":"Notepad"} - SendKeys ile tuş dizisi gönderir',
    category: "uzman",
    code: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim p As String : p = CStr(param)
    Dim keys   As String : keys   = ExtractJsonVal(p,"keys")
    Dim window As String : window = ExtractJsonVal(p,"window")
    If keys = "" Then keys = p
    Dim wsh As Object : Set wsh = CreateObject("WScript.Shell")
    If window <> "" Then
        wsh.AppActivate window
        Application.Wait Now + TimeValue("00:00:01")
    End If
    wsh.SendKeys keys
    targetWb.Sheets(1).Range("A1").Value = "✅ Tuş gönderildi: " & keys
    Set DynamicFunc = Nothing
End Function
${HELPERS}`
  },
  {
    methodName: "GetNetworkSpeed",
    description: "Ağ adaptörü bant genişliği ve anlık kullanımı (Mbps)",
    category: "donanim",
    code: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Cells.ClearContents
    ws.Range("A1").Value = "Adaptör" : ws.Range("B1").Value = "Hız (Mbps)" : ws.Range("C1").Value = "Durum"
    ws.Range("A1:C1").Font.Bold = True
    Dim out As String
    out = RunPS("Get-NetAdapter | Where-Object {$_.Status -eq 'Up'} | Select-Object Name,LinkSpeed,Status | ConvertTo-Csv -NoTypeInformation")
    Dim lines() As String : lines = Split(out, Chr(10))
    Dim r As Long : r = 2
    Dim i As Long
    For i = 1 To UBound(lines)
        Dim ln As String : ln = Trim(Replace(lines(i),Chr(13),""))
        If Len(ln) > 2 Then
            ln = Replace(ln,Chr(34),"")
            Dim parts() As String : parts = Split(ln,",")
            If UBound(parts) >= 2 Then
                ws.Cells(r,1).Value=parts(0): ws.Cells(r,2).Value=parts(1): ws.Cells(r,3).Value=parts(2): r=r+1
            End If
        End If
    Next i
    ws.Columns.AutoFit
    Set DynamicFunc = Nothing
End Function
${HELPERS}`
  },
];

/* ─── .bas dosyalarını yaz ───────────────────────────── */
let written = 0;
for (const m of modules) {
  const filePath = join(SRC, `${m.methodName}.bas`);
  writeFileSync(filePath, m.code, "utf8");
  written++;
}
console.log(`✅ ${written} .bas dosyası yazıldı`);

/* ─── modules.json güncelle ──────────────────────────── */
const modulesJsonPath = join(ROOT, "data", "modules.json");
const existing = JSON.parse(readFileSync(modulesJsonPath, "utf8"));
const existingMap = new Map(existing.map(m => [m.methodName, m]));

let added = 0, updated = 0;
for (const m of modules) {
  if (existingMap.has(m.methodName)) {
    const ex = existingMap.get(m.methodName);
    ex.code = m.code;
    ex.description = m.description;
    ex.category = m.category;
    updated++;
  } else {
    existing.push({ methodName: m.methodName, description: m.description, category: m.category, active: true, code: m.code });
    added++;
  }
}

writeFileSync(modulesJsonPath, JSON.stringify(existing, null, 2), "utf8");
console.log(`✅ modules.json: ${added} eklendi, ${updated} güncellendi. Toplam: ${existing.length}`);
