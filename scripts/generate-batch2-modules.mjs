/**
 * Batch 2: Dosya/Klasör + PowerShell + Zamanlanmış + Uzman Modüller
 * Her PS/CMD çağrısı gizli VBScript üzerinden çalışır (window yok, kullanıcı görmez).
 */
import fs from "fs";
import path from "path";
import { fileURLToPath } from "url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const ROOT      = path.join(__dirname, "..");
const SRC_DIR   = path.join(ROOT, "data", "modules-source");
const JSON_PATH = path.join(ROOT, "data", "modules.json");

/* ── Ortak yardımcı VBA bloğu ──────────────────────────────────────────
   Tüm modüllere eklenen helper fonksiyonları
   ──────────────────────────────────────────────────────────────────── */
const HELPERS = `
'=== ORTAK YARDIMCILAR ===

' Gizli PowerShell komutu çalıştırır, stdout'u temp dosyadan okur
Private Function RunPS(cmd As String) As String
    Dim psPath  As String : psPath  = Environ("TEMP") & "\\__rps.ps1"
    Dim outPath As String : outPath = Environ("TEMP") & "\\__rps_out.txt"
    Dim fNum As Integer : fNum = FreeFile
    Open psPath For Output As #fNum
        Print #fNum, cmd & " | Out-File -FilePath '" & outPath & "' -Encoding UTF8 -Force"
    Close #fNum
    Dim sh As Object : Set sh = CreateObject("WScript.Shell")
    sh.Run "powershell -NonInteractive -NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File """ & psPath & """", 0, True
    Dim result As String
    If Dir(outPath) <> "" Then
        fNum = FreeFile
        Open outPath For Input As #fNum
        Dim line As String
        Do While Not EOF(fNum) : Line Input #fNum, line : result = result & line & vbCrLf : Loop
        Close #fNum
        On Error Resume Next : Kill outPath : Kill psPath : On Error GoTo 0
    End If
    RunPS = Trim(result)
End Function

' Gizli cmd komutu çalıştırır (yönlendirme destekler)
Private Sub RunCmdHidden(cmd As String)
    Dim sh As Object : Set sh = CreateObject("WScript.Shell")
    sh.Run "cmd.exe /c " & cmd, 0, True
End Sub

' Gizli VBScript ile uzun süren işlem başlatır (async)
Private Sub RunVbsHidden(vbsCode As String)
    Dim vbsPath As String : vbsPath = Environ("TEMP") & "\\__async.vbs"
    Dim fNum As Integer : fNum = FreeFile
    Open vbsPath For Output As #fNum : Print #fNum, vbsCode : Close #fNum
    Dim sh As Object : Set sh = CreateObject("WScript.Shell")
    sh.Run "wscript.exe //B //Nologo """ & vbsPath & """", 0, False
End Sub

Private Function ExtractJsonValue(json As String, key As String) As String
    Dim sk As String, p1 As Long, p2 As Long
    sk = """" & key & """:"
    p1 = InStr(1, json, sk, vbTextCompare)
    If p1 = 0 Then Exit Function
    p1 = p1 + Len(sk)
    Do While Mid(json, p1, 1) = " " : p1 = p1 + 1 : Loop
    If Mid(json, p1, 1) = """" Then
        p1 = p1 + 1 : p2 = InStr(p1, json, """")
        If p2 > p1 Then ExtractJsonValue = Mid(json, p1, p2 - p1)
    Else
        p2 = p1
        Do While p2 <= Len(json)
            If InStr(",}]" & Chr(13) & Chr(10), Mid(json, p2, 1)) > 0 Then Exit Do
            p2 = p2 + 1
        Loop
        ExtractJsonValue = Trim(Mid(json, p1, p2 - p1))
    End If
End Function`;

/* ── Modül tanımları ─────────────────────────────────────────────────── */
const modules = [

// ══════════════════════════════════════════════════════════════
//  3. DOSYA / KLASÖR İŞLEMLERİ
// ══════════════════════════════════════════════════════════════

{ name:"CreateFolder", desc:"Parametre ile klasör oluşturur (recursive)", cat:"dosya",
code:`Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim folderPath As String : folderPath = Trim(CStr(param))
    If Len(folderPath) = 0 Then MsgBox "Klasör yolu belirtilmedi.", vbExclamation : GoTo Done
    Dim fso As Object : Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FolderExists(folderPath) Then
        MsgBox "Klasör zaten mevcut: " & folderPath, vbInformation
    Else
        fso.CreateFolder folderPath
        MsgBox "Klasör oluşturuldu: " & folderPath, vbInformation
    End If
Done:
    Set DynamicFunc = Nothing
End Function${HELPERS}`},

{ name:"DeleteFolder", desc:"Klasörü içeriğiyle birlikte siler (gizli VBS)", cat:"dosya",
code:`Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim folderPath As String : folderPath = Trim(CStr(param))
    If Len(folderPath) = 0 Then MsgBox "Klasör yolu belirtilmedi.", vbExclamation : GoTo Done
    If MsgBox("'" & folderPath & "' klasörü kalıcı silinecek. Emin misiniz?", vbYesNo + vbExclamation, "DeleteFolder") <> vbYes Then GoTo Done
    Dim vbs As String
    vbs = "Dim fso : Set fso = CreateObject(""Scripting.FileSystemObject"")" & vbCrLf & _
          "If fso.FolderExists(""" & folderPath & """) Then fso.DeleteFolder """ & folderPath & """, True"
    RunVbsHidden vbs
    MsgBox "Silme komutu gönderildi.", vbInformation
Done:
    Set DynamicFunc = Nothing
End Function${HELPERS}`},

{ name:"CopyFolder", desc:"Klasörü hedefe kopyalar — param: {src, dst}", cat:"dosya",
code:`Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim p As String : p = Trim(CStr(param))
    Dim src As String, dst As String
    src = ExtractJsonValue(p, "src") : dst = ExtractJsonValue(p, "dst")
    If Len(src) = 0 Or Len(dst) = 0 Then MsgBox "Param: {""src"":""...\\"",""dst"":""...\\""}",vbExclamation:GoTo Done
    Dim fso As Object : Set fso = CreateObject("Scripting.FileSystemObject")
    fso.CopyFolder src, dst
    MsgBox "Kopyalandı: " & src & " → " & dst, vbInformation
Done:
    Set DynamicFunc = Nothing
End Function${HELPERS}`},

{ name:"MoveFolder", desc:"Klasörü taşır — param: {src, dst}", cat:"dosya",
code:`Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim p As String : p = Trim(CStr(param))
    Dim src As String, dst As String
    src = ExtractJsonValue(p, "src") : dst = ExtractJsonValue(p, "dst")
    If Len(src) = 0 Or Len(dst) = 0 Then MsgBox "Param: {""src"":""..."",""dst"":""...""}",vbExclamation:GoTo Done
    Dim fso As Object : Set fso = CreateObject("Scripting.FileSystemObject")
    fso.MoveFolder src, dst
    MsgBox "Taşındı: " & src & " → " & dst, vbInformation
Done:
    Set DynamicFunc = Nothing
End Function${HELPERS}`},

{ name:"ListFolderContents", desc:"Ad, boyut, tarih, uzantı bilgisiyle listeler", cat:"dosya",
code:`Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim folderPath As String : folderPath = Trim(CStr(param))
    If Len(folderPath) = 0 Then folderPath = Environ("USERPROFILE") & "\\Desktop"
    Dim fso As Object : Set fso = CreateObject("Scripting.FileSystemObject")
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Cells.ClearContents
    ws.Range("A1:E1").Value = Array("Ad", "Uzantı", "Boyut (KB)", "Son Değiştirilme", "Tür")
    ws.Range("A1:E1").Font.Bold = True
    Dim r As Long : r = 2
    Dim folder As Object : Set folder = fso.GetFolder(folderPath)
    Dim item As Object
    For Each item In folder.SubFolders
        ws.Cells(r,1).Value = item.Name : ws.Cells(r,5).Value = "Klasör"
        ws.Cells(r,4).Value = Format(item.DateLastModified,"dd.mm.yyyy HH:MM")
        r = r + 1
    Next item
    For Each item In folder.Files
        ws.Cells(r,1).Value = item.Name
        ws.Cells(r,2).Value = fso.GetExtensionName(item.Path)
        ws.Cells(r,3).Value = Format(item.Size/1024,"0.0")
        ws.Cells(r,4).Value = Format(item.DateLastModified,"dd.mm.yyyy HH:MM")
        ws.Cells(r,5).Value = "Dosya"
        r = r + 1
    Next item
    ws.Columns.AutoFit
    MsgBox (r-2) & " öğe listelendi: " & folderPath, vbInformation
    Set DynamicFunc = Nothing
End Function${HELPERS}`},

{ name:"CopyFile", desc:"Dosya kopyalar — param: {src, dst, overwrite}", cat:"dosya",
code:`Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim p As String : p = Trim(CStr(param))
    Dim src As String : src = ExtractJsonValue(p,"src")
    Dim dst As String : dst = ExtractJsonValue(p,"dst")
    Dim ow  As Boolean : ow = (LCase(ExtractJsonValue(p,"overwrite")) <> "false")
    Dim fso As Object : Set fso = CreateObject("Scripting.FileSystemObject")
    fso.CopyFile src, dst, ow
    MsgBox "Kopyalandı: " & src, vbInformation
    Set DynamicFunc = Nothing
End Function${HELPERS}`},

{ name:"MoveFile", desc:"Dosya taşır — param: {src, dst}", cat:"dosya",
code:`Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim p As String : p = Trim(CStr(param))
    Dim fso As Object : Set fso = CreateObject("Scripting.FileSystemObject")
    fso.MoveFile ExtractJsonValue(p,"src"), ExtractJsonValue(p,"dst")
    MsgBox "Taşındı.", vbInformation
    Set DynamicFunc = Nothing
End Function${HELPERS}`},

{ name:"DeleteFile", desc:"Dosya siler (geri dönüşüm kutusu bypass) — param: yol", cat:"dosya",
code:`Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim fp As String : fp = Trim(CStr(param))
    If MsgBox("'" & fp & "' kalıcı silinecek?", vbYesNo+vbExclamation, "DeleteFile") <> vbYes Then GoTo Done
    Dim vbs As String
    vbs = "Dim fso : Set fso = CreateObject(""Scripting.FileSystemObject"")" & vbCrLf & _
          "If fso.FileExists(""" & fp & """) Then fso.DeleteFile """ & fp & """, True"
    RunVbsHidden vbs
    MsgBox "Silme komutu gönderildi.", vbInformation
Done:
    Set DynamicFunc = Nothing
End Function${HELPERS}`},

{ name:"RenameFile", desc:"Dosyayı yeniden adlandırır — param: {path, newName}", cat:"dosya",
code:`Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim p As String : p = Trim(CStr(param))
    Dim fp As String : fp = ExtractJsonValue(p,"path")
    Dim nn As String : nn = ExtractJsonValue(p,"newName")
    Dim fso As Object : Set fso = CreateObject("Scripting.FileSystemObject")
    Dim parent As String : parent = fso.GetParentFolderName(fp)
    fso.MoveFile fp, parent & "\\" & nn
    MsgBox "Yeniden adlandırıldı: " & nn, vbInformation
    Set DynamicFunc = Nothing
End Function${HELPERS}`},

{ name:"FileExists", desc:"Dosyanın varlığını kontrol eder — param: yol", cat:"dosya",
code:`Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim fp As String : fp = Trim(CStr(param))
    Dim fso As Object : Set fso = CreateObject("Scripting.FileSystemObject")
    Dim exists As Boolean : exists = fso.FileExists(fp)
    targetWb.Sheets(1).Range("A1").Value = "Dosya Yolu"
    targetWb.Sheets(1).Range("B1").Value = "Mevcut"
    targetWb.Sheets(1).Range("A2").Value = fp
    targetWb.Sheets(1).Range("B2").Value = IIf(exists,"Evet","Hayır")
    targetWb.Sheets(1).Columns.AutoFit
    MsgBox fp & Chr(10) & IIf(exists,"✔ Mevcut","✖ Bulunamadı"), vbInformation,"FileExists"
    Set DynamicFunc = Nothing
End Function${HELPERS}`},

{ name:"GetFileSize", desc:"Dosya boyutu (B/KB/MB/GB otomatik birim) — param: yol", cat:"dosya",
code:`Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim fp As String : fp = Trim(CStr(param))
    Dim fso As Object : Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(fp) Then MsgBox "Dosya bulunamadı.", vbExclamation : GoTo Done
    Dim sz As Double : sz = fso.GetFile(fp).Size
    Dim szStr As String
    If sz < 1024 Then szStr = sz & " B"
    ElseIf sz < 1048576 Then szStr = Format(sz/1024,"0.00") & " KB"
    ElseIf sz < 1073741824 Then szStr = Format(sz/1048576,"0.00") & " MB"
    Else szStr = Format(sz/1073741824,"0.00") & " GB"
    End If
    MsgBox fso.GetFileName(fp) & ": " & szStr, vbInformation,"GetFileSize"
Done:
    Set DynamicFunc = Nothing
End Function${HELPERS}`},

{ name:"GetFileHashMd5", desc:"MD5 hash hesaplar — param: yol (gizli PS)", cat:"dosya",
code:`Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim fp As String : fp = Trim(CStr(param))
    Dim hash As String
    hash = RunPS("(Get-FileHash -Path '" & fp & "' -Algorithm MD5).Hash")
    targetWb.Sheets(1).Range("A1").Value = "Dosya"
    targetWb.Sheets(1).Range("B1").Value = "MD5 Hash"
    targetWb.Sheets(1).Range("A2").Value = fp
    targetWb.Sheets(1).Range("B2").Value = hash
    targetWb.Sheets(1).Columns.AutoFit
    MsgBox "MD5: " & hash, vbInformation,"GetFileHashMd5"
    Set DynamicFunc = Nothing
End Function${HELPERS}`},

{ name:"ReadTextFile", desc:"UTF-8 metin dosyasını okur, hücreye yazar — param: yol", cat:"dosya",
code:`Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim fp As String : fp = Trim(CStr(param))
    Dim content As String
    content = RunPS("Get-Content -Path '" & fp & "' -Encoding UTF8 -Raw")
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Cells.ClearContents
    ws.Range("A1").Value = "İçerik (" & fp & ")"
    ws.Range("A1").Font.Bold = True
    Dim lines() As String : lines = Split(content, vbLf)
    Dim i As Long
    For i = 0 To UBound(lines)
        ws.Cells(i+2, 1).Value = Replace(lines(i), vbCr, "")
    Next i
    ws.Columns(1).AutoFit
    Set DynamicFunc = Nothing
End Function${HELPERS}`},

{ name:"WriteTextFile", desc:"Hücre içeriğini dosyaya yazar — param: {path, content}", cat:"dosya",
code:`Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim p As String : p = Trim(CStr(param))
    Dim fp      As String : fp      = ExtractJsonValue(p,"path")
    Dim content As String : content = ExtractJsonValue(p,"content")
    Dim fso As Object : Set fso = CreateObject("Scripting.FileSystemObject")
    Dim ts  As Object : Set ts = fso.CreateTextFile(fp, True, True) ' Unicode
    ts.Write content : ts.Close
    MsgBox "Dosya yazıldı: " & fp, vbInformation
    Set DynamicFunc = Nothing
End Function${HELPERS}`},

{ name:"AppendToTextFile", desc:"Metin dosyasına satır ekler — param: {path, line}", cat:"dosya",
code:`Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim p As String : p = Trim(CStr(param))
    Dim fp   As String : fp   = ExtractJsonValue(p,"path")
    Dim line As String : line = ExtractJsonValue(p,"line")
    Dim fNum As Integer : fNum = FreeFile
    Open fp For Append As #fNum
    Print #fNum, line
    Close #fNum
    MsgBox "Satır eklendi: " & fp, vbInformation
    Set DynamicFunc = Nothing
End Function${HELPERS}`},

{ name:"ReadCsvToSheet", desc:"CSV dosyasını aktif sayfaya aktarır — param: yol", cat:"dosya",
code:`Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim fp As String : fp = Trim(CStr(param))
    Dim fso As Object : Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(fp) Then MsgBox "Dosya bulunamadı.", vbExclamation : GoTo Done
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Cells.ClearContents
    Dim fNum As Integer : fNum = FreeFile
    Open fp For Input As #fNum
    Dim r As Long : r = 1
    Dim rowLine As String
    Do While Not EOF(fNum)
        Line Input #fNum, rowLine
        Dim cols() As String : cols = Split(rowLine, ",")
        Dim c As Integer
        For c = 0 To UBound(cols)
            ws.Cells(r, c+1).Value = Replace(cols(c), """", "")
        Next c
        r = r + 1
    Loop
    Close #fNum
    ws.Rows(1).Font.Bold = True
    ws.Columns.AutoFit
    MsgBox (r-1) & " satır aktarıldı.", vbInformation
Done:
    Set DynamicFunc = Nothing
End Function${HELPERS}`},

{ name:"WriteSheetToCsv", desc:"Aktif sayfayı CSV olarak dışa aktarır — param: yol", cat:"dosya",
code:`Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim fp As String : fp = Trim(CStr(param))
    If Len(fp) = 0 Then fp = Environ("USERPROFILE") & "\\Desktop\\export_" & Format(Now,"yyyymmdd_HHMMSS") & ".csv"
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    Dim lastRow As Long : lastRow = ws.Cells(ws.Rows.Count,1).End(xlUp).Row
    Dim lastCol As Integer : lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    Dim fNum As Integer : fNum = FreeFile
    Open fp For Output As #fNum
    Dim r As Long, c As Integer
    For r = 1 To lastRow
        Dim rowStr As String : rowStr = ""
        For c = 1 To lastCol
            Dim v As String : v = CStr(ws.Cells(r,c).Value)
            If InStr(v,",") > 0 Or InStr(v,"""") > 0 Then v = """" & Replace(v,"""","""""") & """"
            rowStr = rowStr & IIf(c>1,",","") & v
        Next c
        Print #fNum, rowStr
    Next r
    Close #fNum
    MsgBox "CSV kaydedildi: " & fp, vbInformation
    Set DynamicFunc = Nothing
End Function${HELPERS}`},

{ name:"ZipFolder", desc:"Klasörü ZIP'ler — param: {src, dst} (gizli PS)", cat:"dosya",
code:`Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim p As String : p = Trim(CStr(param))
    Dim src As String : src = ExtractJsonValue(p,"src")
    Dim dst As String : dst = ExtractJsonValue(p,"dst")
    If Len(dst) = 0 Then dst = src & "_" & Format(Now,"yyyymmdd") & ".zip"
    Dim psCmd As String
    psCmd = "Compress-Archive -Path '" & src & "' -DestinationPath '" & dst & "' -Force"
    RunPS psCmd
    MsgBox "ZIP oluşturuldu: " & dst, vbInformation,"ZipFolder"
    Set DynamicFunc = Nothing
End Function${HELPERS}`},

{ name:"UnzipToFolder", desc:"ZIP dosyasını hedef klasöre açar — param: {zip, dst} (gizli PS)", cat:"dosya",
code:`Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim p As String : p = Trim(CStr(param))
    Dim zipPath As String : zipPath = ExtractJsonValue(p,"zip")
    Dim dst     As String : dst     = ExtractJsonValue(p,"dst")
    Dim psCmd   As String
    psCmd = "Expand-Archive -Path '" & zipPath & "' -DestinationPath '" & dst & "' -Force"
    RunPS psCmd
    MsgBox "ZIP açıldı: " & dst, vbInformation,"UnzipToFolder"
    Set DynamicFunc = Nothing
End Function${HELPERS}`},

{ name:"SearchFilesByPattern", desc:"Pattern ile özyinelemeli dosya arama — param: {folder, pattern}", cat:"dosya",
code:`Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim p As String : p = Trim(CStr(param))
    Dim folder  As String : folder  = ExtractJsonValue(p,"folder")
    Dim pattern As String : pattern = ExtractJsonValue(p,"pattern")
    If Len(folder)  = 0 Then folder  = Environ("USERPROFILE")
    If Len(pattern) = 0 Then pattern = "*.*"
    Dim psCmd As String
    psCmd = "Get-ChildItem -Path '" & folder & "' -Filter '" & pattern & "' -Recurse -File | Select-Object FullName,Length,LastWriteTime | Format-Table -AutoSize"
    Dim result As String : result = RunPS(psCmd)
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Cells.ClearContents
    ws.Range("A1").Value = "Arama: " & folder & " / " & pattern
    ws.Range("A1").Font.Bold = True
    Dim lines() As String : lines = Split(result, vbLf)
    Dim i As Long
    For i = 0 To UBound(lines)
        ws.Cells(i+2,1).Value = Trim(Replace(lines(i), vbCr, ""))
    Next i
    ws.Columns.AutoFit
    Set DynamicFunc = Nothing
End Function${HELPERS}`},

{ name:"GetFolderSize", desc:"Alt klasörler dahil toplam disk kullanımı — param: yol (gizli PS)", cat:"dosya",
code:`Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim fp As String : fp = Trim(CStr(param))
    If Len(fp) = 0 Then fp = Environ("USERPROFILE") & "\\Documents"
    Dim psCmd As String
    psCmd = "$s=(Get-ChildItem -Path '" & fp & "' -Recurse -File -ErrorAction SilentlyContinue | Measure-Object -Property Length -Sum).Sum; if($s -lt 1MB){""$([math]::Round($s/1KB,2)) KB""}elseif($s -lt 1GB){""$([math]::Round($s/1MB,2)) MB""}else{""$([math]::Round($s/1GB,2)) GB""}"
    Dim result As String : result = RunPS(psCmd)
    MsgBox fp & Chr(10) & "Toplam: " & result, vbInformation,"GetFolderSize"
    Set DynamicFunc = Nothing
End Function${HELPERS}`},

{ name:"OpenFileWithDefaultApp", desc:"Varsayılan uygulamada açar — param: yol", cat:"dosya",
code:`Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim fp As String : fp = Trim(CStr(param))
    Dim sh As Object : Set sh = CreateObject("WScript.Shell")
    sh.Run """" & fp & """", 1, False
    Set DynamicFunc = Nothing
End Function${HELPERS}`},

{ name:"CleanTempFolder", desc:"%TEMP% klasörünü temizler (gizli PS)", cat:"dosya",
code:`Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    If MsgBox("%TEMP% klasörü temizlenecek. Devam?", vbYesNo+vbQuestion,"CleanTempFolder") <> vbYes Then GoTo Done
    Dim psCmd As String
    psCmd = "Remove-Item -Path $env:TEMP\\* -Recurse -Force -ErrorAction SilentlyContinue"
    RunPS psCmd
    MsgBox "Temp klasörü temizlendi.", vbInformation
Done:
    Set DynamicFunc = Nothing
End Function${HELPERS}`},

{ name:"BackupFileWithTimestamp", desc:"Dosyayı ad_YYYYMMDD_HHMMSS.bak kopyalar — param: yol", cat:"dosya",
code:`Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim fp As String : fp = Trim(CStr(param))
    Dim fso As Object : Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(fp) Then MsgBox "Dosya bulunamadı.", vbExclamation : GoTo Done
    Dim baseName As String : baseName = fso.GetBaseName(fp)
    Dim ext      As String : ext      = fso.GetExtensionName(fp)
    Dim folder   As String : folder   = fso.GetParentFolderName(fp)
    Dim bakPath  As String : bakPath  = folder & "\\" & baseName & "_" & Format(Now,"yyyymmdd_HHmmss") & "." & ext & ".bak"
    fso.CopyFile fp, bakPath
    MsgBox "Yedek: " & bakPath, vbInformation,"BackupFileWithTimestamp"
Done:
    Set DynamicFunc = Nothing
End Function${HELPERS}`},

{ name:"CompareFilesIdentical", desc:"İki dosyanın MD5 hash'ini karşılaştırır — param: {f1, f2} (gizli PS)", cat:"dosya",
code:`Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim p As String : p = Trim(CStr(param))
    Dim f1 As String : f1 = ExtractJsonValue(p,"f1")
    Dim f2 As String : f2 = ExtractJsonValue(p,"f2")
    Dim psCmd As String
    psCmd = "$h1=(Get-FileHash '" & f1 & "' -Algorithm MD5).Hash;$h2=(Get-FileHash '" & f2 & "' -Algorithm MD5).Hash;if($h1 -eq $h2){'AYNI: '+$h1}else{'FARKLI - F1:'+$h1+' F2:'+$h2}"
    MsgBox RunPS(psCmd), vbInformation,"CompareFilesIdentical"
    Set DynamicFunc = Nothing
End Function${HELPERS}`},

{ name:"ReplaceTextInFile", desc:"Metin dosyasında bul-değiştir — param: {path, find, replace} (gizli PS)", cat:"dosya",
code:`Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim p As String : p = Trim(CStr(param))
    Dim fp  As String : fp  = ExtractJsonValue(p,"path")
    Dim fnd As String : fnd = ExtractJsonValue(p,"find")
    Dim rep As String : rep = ExtractJsonValue(p,"replace")
    Dim psCmd As String
    psCmd = "(Get-Content '" & fp & "') -replace '" & fnd & "','" & rep & "' | Set-Content '" & fp & "'"
    RunPS psCmd
    MsgBox "Değiştirme yapıldı: " & fnd & " → " & rep, vbInformation
    Set DynamicFunc = Nothing
End Function${HELPERS}`},

{ name:"GetNewestFileInFolder", desc:"Son değiştirilen dosyayı bulur — param: yol (gizli PS)", cat:"dosya",
code:`Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim fp As String : fp = Trim(CStr(param))
    If Len(fp) = 0 Then fp = Environ("USERPROFILE") & "\\Downloads"
    Dim psCmd As String
    psCmd = "$f=Get-ChildItem -Path '" & fp & "' -File | Sort-Object LastWriteTime -Descending | Select-Object -First 1; ""$($f.FullName) | $($f.LastWriteTime.ToString('dd.MM.yyyy HH:mm:ss'))"""
    Dim result As String : result = RunPS(psCmd)
    MsgBox "En yeni dosya:" & Chr(10) & result, vbInformation,"GetNewestFileInFolder"
    Set DynamicFunc = Nothing
End Function${HELPERS}`},

{ name:"GetFileAttributes", desc:"Gizli/Salt-okunur/Sistem özniteliklerini okur — param: yol", cat:"dosya",
code:`Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim fp As String : fp = Trim(CStr(param))
    Dim fso As Object : Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(fp) Then MsgBox "Dosya bulunamadı.", vbExclamation : GoTo Done
    Dim attrs As Long : attrs = fso.GetFile(fp).Attributes
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Cells.ClearContents
    ws.Range("A1:B1").Value = Array("Öznitelik","Değer") : ws.Range("A1:B1").Font.Bold = True
    ws.Range("A2:B2").Value = Array("Normal",       IIf(attrs = 0,"Evet","Hayır"))
    ws.Range("A3:B3").Value = Array("Salt-Okunur",  IIf((attrs And 1) > 0,"Evet","Hayır"))
    ws.Range("A4:B4").Value = Array("Gizli",        IIf((attrs And 2) > 0,"Evet","Hayır"))
    ws.Range("A5:B5").Value = Array("Sistem",       IIf((attrs And 4) > 0,"Evet","Hayır"))
    ws.Range("A6:B6").Value = Array("Arşiv",        IIf((attrs And 32) > 0,"Evet","Hayır"))
    ws.Columns.AutoFit
Done:
    Set DynamicFunc = Nothing
End Function${HELPERS}`},

{ name:"SetFileAttribute", desc:"Dosya özniteliğini değiştirir — param: {path, attr, value} attr: hidden/readonly/system", cat:"dosya",
code:`Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim p As String : p = Trim(CStr(param))
    Dim fp   As String  : fp   = ExtractJsonValue(p,"path")
    Dim attr As String  : attr = LCase(ExtractJsonValue(p,"attr"))
    Dim val  As Boolean : val  = (LCase(ExtractJsonValue(p,"value")) <> "false")
    Dim psCmd As String
    If attr = "hidden"   Then psCmd = "(Get-Item '" & fp & "').Attributes = (Get-Item '" & fp & "').Attributes " & IIf(val,"bor [System.IO.FileAttributes]::Hidden","band -bnot [System.IO.FileAttributes]::Hidden")
    If attr = "readonly" Then psCmd = "(Get-Item '" & fp & "').IsReadOnly = $" & LCase(CStr(val))
    If Len(psCmd) = 0 Then MsgBox "attr: hidden|readonly|system", vbExclamation : GoTo Done
    RunPS psCmd
    MsgBox "Öznitelik güncellendi: " & attr & " = " & CStr(val), vbInformation
Done:
    Set DynamicFunc = Nothing
End Function${HELPERS}`},

// ══════════════════════════════════════════════════════════════
//  5. POWERSHELL / CMD (Tümü gizli VBScript → PS)
// ══════════════════════════════════════════════════════════════

{ name:"RunPsCommand", desc:"PS komutu çalıştırır, stdout'u sayfaya yazar — param: ps kodu", cat:"powershell",
code:`Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim psCmd As String : psCmd = Trim(CStr(param))
    Dim result As String : result = RunPS(psCmd)
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Cells.ClearContents
    ws.Range("A1").Value = "Komut: " & Left(psCmd,80)
    ws.Range("A1").Font.Bold = True
    Dim lines() As String : lines = Split(result, vbLf)
    Dim i As Long
    For i = 0 To UBound(lines)
        ws.Cells(i+2,1).Value = Trim(Replace(lines(i),vbCr,""))
    Next i
    ws.Columns.AutoFit
    Set DynamicFunc = Nothing
End Function${HELPERS}`},

{ name:"RunCmdCommand", desc:"cmd.exe komutu çalıştırır (gizli) — param: {cmd, wait}", cat:"powershell",
code:`Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim p As String : p = Trim(CStr(param))
    Dim cmd   As String  : cmd   = ExtractJsonValue(p,"cmd")
    Dim wait  As Boolean : wait  = (LCase(ExtractJsonValue(p,"wait")) <> "false")
    If Len(cmd) = 0 Then cmd = p
    Dim vbs As String
    vbs = "Dim wsh : Set wsh = CreateObject(""WScript.Shell"")" & vbCrLf & _
          "wsh.Run ""cmd.exe /c " & cmd & """, 0, " & LCase(CStr(wait))
    RunVbsHidden vbs
    MsgBox "Komut gönderildi: " & cmd, vbInformation
    Set DynamicFunc = Nothing
End Function${HELPERS}`},

{ name:"GetEventLogErrors", desc:"Son N uygulama/sistem hatasını listeler — param: {log, count} (gizli PS)", cat:"powershell",
code:`Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim p As String : p = Trim(CStr(param))
    Dim logName As String : logName = ExtractJsonValue(p,"log")
    Dim cnt     As Long   : cnt     = CLng(IIf(Len(ExtractJsonValue(p,"count"))>0, ExtractJsonValue(p,"count"), "20"))
    If Len(logName) = 0 Then logName = "Application"
    Dim psCmd As String
    psCmd = "Get-EventLog -LogName '" & logName & "' -EntryType Error,Warning -Newest " & cnt & " | Select-Object TimeGenerated,EntryType,Source,Message | Format-Table -AutoSize -Wrap"
    Dim result As String : result = RunPS(psCmd)
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Cells.ClearContents
    ws.Range("A1").Value = "Event Log: " & logName & " (Son " & cnt & " hata)"
    ws.Range("A1").Font.Bold = True
    Dim lines() As String : lines = Split(result, vbLf)
    Dim i As Long
    For i = 0 To UBound(lines)
        ws.Cells(i+2,1).Value = Trim(Replace(lines(i),vbCr,""))
    Next i
    ws.Columns.AutoFit
    Set DynamicFunc = Nothing
End Function${HELPERS}`},

{ name:"FlushDnsCache", desc:"ipconfig /flushdns çalıştırır (gizli)", cat:"powershell",
code:`Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim vbs As String
    vbs = "Dim wsh : Set wsh = CreateObject(""WScript.Shell"")" & vbCrLf & _
          "wsh.Run ""cmd.exe /c ipconfig /flushdns"", 0, True"
    RunVbsHidden vbs
    MsgBox "DNS önbelleği temizlendi.", vbInformation,"FlushDnsCache"
    Set DynamicFunc = Nothing
End Function${HELPERS}`},

{ name:"GetNetworkConfig", desc:"IP, DNS, Gateway, DHCP bilgilerini sayfaya yazar (gizli PS)", cat:"powershell",
code:`Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim psCmd As String
    psCmd = "Get-NetIPConfiguration | Select-Object InterfaceAlias,IPv4Address,IPv4DefaultGateway,DNSServer | Format-List"
    Dim result As String : result = RunPS(psCmd)
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Cells.ClearContents
    ws.Range("A1").Value = "Ağ Yapılandırması" : ws.Range("A1").Font.Bold = True
    Dim lines() As String : lines = Split(result, vbLf)
    Dim i As Long
    For i = 0 To UBound(lines)
        ws.Cells(i+2,1).Value = Trim(Replace(lines(i),vbCr,""))
    Next i
    ws.Columns.AutoFit
    Set DynamicFunc = Nothing
End Function${HELPERS}`},

{ name:"GetFirewallRules", desc:"Aktif güvenlik duvarı kurallarını listeler (gizli PS)", cat:"powershell",
code:`Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim p As String : p = Trim(CStr(param))
    Dim direction As String : direction = ExtractJsonValue(p,"direction")
    If Len(direction) = 0 Then direction = "Inbound"
    Dim psCmd As String
    psCmd = "Get-NetFirewallRule -Direction " & direction & " -Enabled True | Select-Object DisplayName,Action,Profile,Direction | Format-Table -AutoSize"
    Dim result As String : result = RunPS(psCmd)
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Cells.ClearContents
    ws.Range("A1").Value = "Güvenlik Duvarı Kuralları (" & direction & ")"
    ws.Range("A1").Font.Bold = True
    Dim lines() As String : lines = Split(result, vbLf)
    Dim i As Long
    For i = 0 To UBound(lines)
        ws.Cells(i+2,1).Value = Trim(Replace(lines(i),vbCr,""))
    Next i
    ws.Columns.AutoFit
    Set DynamicFunc = Nothing
End Function${HELPERS}`},

{ name:"AddFirewallRule", desc:"Firewall kuralı ekler — param: {name, port, protocol, direction} (gizli PS)", cat:"powershell",
code:`Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim p As String : p = Trim(CStr(param))
    Dim name      As String : name      = ExtractJsonValue(p,"name")
    Dim port      As String : port      = ExtractJsonValue(p,"port")
    Dim proto     As String : proto     = ExtractJsonValue(p,"protocol")
    Dim direction As String : direction = ExtractJsonValue(p,"direction")
    If Len(proto)     = 0 Then proto     = "TCP"
    If Len(direction) = 0 Then direction = "Inbound"
    Dim psCmd As String
    psCmd = "New-NetFirewallRule -DisplayName '" & name & "' -Direction " & direction & " -Protocol " & proto & " -LocalPort " & port & " -Action Allow"
    RunPS psCmd
    MsgBox "Kural eklendi: " & name & " (" & direction & "/" & proto & "/" & port & ")", vbInformation
    Set DynamicFunc = Nothing
End Function${HELPERS}`},

{ name:"RestartWindowsService", desc:"Servis adına göre yeniden başlatır — param: servis adı (gizli PS)", cat:"powershell",
code:`Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim svcName As String : svcName = Trim(CStr(param))
    If MsgBox("'" & svcName & "' servisi yeniden başlatılacak. Devam?", vbYesNo+vbQuestion,"RestartService") <> vbYes Then GoTo Done
    Dim psCmd As String
    psCmd = "Restart-Service -Name '" & svcName & "' -Force"
    RunPS psCmd
    MsgBox svcName & " servisi yeniden başlatıldı.", vbInformation
Done:
    Set DynamicFunc = Nothing
End Function${HELPERS}`},

{ name:"GetServiceStatus", desc:"Servis durumunu döndürür — param: servis adı (gizli PS)", cat:"powershell",
code:`Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim svcName As String : svcName = Trim(CStr(param))
    If Len(svcName) = 0 Then svcName = "wuauserv"
    Dim psCmd As String
    psCmd = "$s = Get-Service -Name '" & svcName & "' -ErrorAction SilentlyContinue; if($s){'$($s.DisplayName): $($s.Status)'}else{'Servis bulunamadı: " & svcName & "'}"
    MsgBox RunPS(psCmd), vbInformation,"GetServiceStatus"
    Set DynamicFunc = Nothing
End Function${HELPERS}`},

{ name:"KillProcessByName", desc:"İsme göre process sonlandırır (gizli PS) — param: process adı", cat:"powershell",
code:`Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim procName As String : procName = Trim(CStr(param))
    If MsgBox("'" & procName & "' process sonlandırılacak. Emin misiniz?", vbYesNo+vbExclamation,"KillProcess") <> vbYes Then GoTo Done
    Dim psCmd As String
    psCmd = "Stop-Process -Name '" & procName & "' -Force -ErrorAction SilentlyContinue"
    RunPS psCmd
    MsgBox procName & " sonlandırıldı.", vbInformation
Done:
    Set DynamicFunc = Nothing
End Function${HELPERS}`},

{ name:"GetEnvironmentVariables", desc:"Tüm env değişkenlerini sayfaya yazar (gizli PS)", cat:"powershell",
code:`Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim psCmd As String
    psCmd = "Get-ChildItem Env: | Select-Object Name,Value | Sort-Object Name | Format-Table -AutoSize"
    Dim result As String : result = RunPS(psCmd)
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Cells.ClearContents
    ws.Range("A1").Value = "Environment Variables" : ws.Range("A1").Font.Bold = True
    Dim lines() As String : lines = Split(result, vbLf)
    Dim i As Long
    For i = 0 To UBound(lines)
        ws.Cells(i+2,1).Value = Trim(Replace(lines(i),vbCr,""))
    Next i
    ws.Columns.AutoFit
    Set DynamicFunc = Nothing
End Function${HELPERS}`},

{ name:"SetEnvironmentVariable", desc:"Kullanıcı env değişkeni atar — param: {name, value} (gizli PS)", cat:"powershell",
code:`Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim p As String : p = Trim(CStr(param))
    Dim varName  As String : varName  = ExtractJsonValue(p,"name")
    Dim varValue As String : varValue = ExtractJsonValue(p,"value")
    Dim psCmd As String
    psCmd = "[System.Environment]::SetEnvironmentVariable('" & varName & "','" & varValue & "',[System.EnvironmentVariableTarget]::User)"
    RunPS psCmd
    MsgBox varName & " = " & varValue & " (Kullanıcı)", vbInformation
    Set DynamicFunc = Nothing
End Function${HELPERS}`},

{ name:"GetHostsFile", desc:"hosts dosyasını okur ve sayfaya yazar (gizli PS)", cat:"powershell",
code:`Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim result As String
    result = RunPS("Get-Content 'C:\\Windows\\System32\\drivers\\etc\\hosts'")
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Cells.ClearContents
    ws.Range("A1").Value = "Hosts Dosyası" : ws.Range("A1").Font.Bold = True
    Dim lines() As String : lines = Split(result, vbLf)
    Dim i As Long
    For i = 0 To UBound(lines)
        ws.Cells(i+2,1).Value = Trim(Replace(lines(i),vbCr,""))
    Next i
    ws.Columns.AutoFit
    Set DynamicFunc = Nothing
End Function${HELPERS}`},

{ name:"GetWifiProfiles", desc:"Kayıtlı Wi-Fi profilleri listeler (gizli PS)", cat:"powershell",
code:`Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim psCmd As String
    psCmd = "netsh wlan show profiles | Select-String 'All User Profile' | ForEach-Object { $n = ($_ -split ':')[1].Trim(); $k = (netsh wlan show profile name=$n key=clear | Select-String 'Key Content'); $p = if($k){($k -split ':')[1].Trim()}else{'(şifre yok)'}; ""$n | $p"" }"
    Dim result As String : result = RunPS(psCmd)
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Cells.ClearContents
    ws.Range("A1:B1").Value = Array("SSID","Şifre") : ws.Range("A1:B1").Font.Bold = True
    Dim lines() As String : lines = Split(result, vbLf)
    Dim i As Long, r As Long : r = 2
    For i = 0 To UBound(lines)
        Dim ln As String : ln = Trim(Replace(lines(i),vbCr,""))
        If InStr(ln," | ") > 0 Then
            Dim parts() As String : parts = Split(ln," | ")
            ws.Cells(r,1).Value = Trim(parts(0))
            ws.Cells(r,2).Value = Trim(parts(1))
            r = r + 1
        End If
    Next i
    ws.Columns.AutoFit
    Set DynamicFunc = Nothing
End Function${HELPERS}`},

{ name:"GetBitLockerStatus", desc:"BitLocker durumunu kontrol eder (gizli PS)", cat:"powershell",
code:`Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim psCmd As String
    psCmd = "Get-BitLockerVolume | Select-Object MountPoint,VolumeStatus,ProtectionStatus,EncryptionPercentage | Format-Table -AutoSize"
    Dim result As String : result = RunPS(psCmd)
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Cells.ClearContents
    ws.Range("A1").Value = "BitLocker Durumu" : ws.Range("A1").Font.Bold = True
    Dim lines() As String : lines = Split(result, vbLf)
    Dim i As Long
    For i = 0 To UBound(lines)
        ws.Cells(i+2,1).Value = Trim(Replace(lines(i),vbCr,""))
    Next i
    ws.Columns.AutoFit
    Set DynamicFunc = Nothing
End Function${HELPERS}`},

// ══════════════════════════════════════════════════════════════
//  10. ZAMANLANMIŞ & OTOMATİK (yeni)
// ══════════════════════════════════════════════════════════════

{ name:"ScheduleTaskWeekly", desc:"Haftanın belirli günü tekrarlayan görev — param: {taskName, command, dayOfWeek, startTime}", cat:"zamanlanmis",
code:`Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim p As String : p = Trim(CStr(param))
    Dim taskName  As String : taskName  = ExtractJsonValue(p,"taskName")
    Dim command   As String : command   = ExtractJsonValue(p,"command")
    Dim day       As String : day       = ExtractJsonValue(p,"dayOfWeek")
    Dim startTime As String : startTime = ExtractJsonValue(p,"startTime")
    If Len(startTime) = 0 Then startTime = "09:00"
    If Len(day)       = 0 Then day       = "MON"
    Dim psCmd As String
    psCmd = "schtasks /create /tn """ & taskName & """ /tr """ & command & """ /sc WEEKLY /d " & day & " /st " & startTime & " /f"
    RunPS psCmd
    MsgBox "Haftalık görev oluşturuldu: " & taskName, vbInformation
    Set DynamicFunc = Nothing
End Function${HELPERS}`},

{ name:"ScheduleTaskMonthly", desc:"Ayın belirli günü tekrarlayan görev — param: {taskName, command, day, startTime}", cat:"zamanlanmis",
code:`Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim p As String : p = Trim(CStr(param))
    Dim taskName  As String : taskName  = ExtractJsonValue(p,"taskName")
    Dim command   As String : command   = ExtractJsonValue(p,"command")
    Dim day       As String : day       = ExtractJsonValue(p,"day")
    Dim startTime As String : startTime = ExtractJsonValue(p,"startTime")
    If Len(startTime) = 0 Then startTime = "09:00"
    If Len(day)       = 0 Then day       = "1"
    Dim psCmd As String
    psCmd = "schtasks /create /tn """ & taskName & """ /tr """ & command & """ /sc MONTHLY /d " & day & " /st " & startTime & " /f"
    RunPS psCmd
    MsgBox "Aylık görev oluşturuldu: " & taskName, vbInformation
    Set DynamicFunc = Nothing
End Function${HELPERS}`},

{ name:"AutoArchiveOldRows", desc:"N günden eski satırları arşiv sayfasına taşır — param: {dateColumn, days, archiveSheet}", cat:"zamanlanmis",
code:`Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim p As String : p = Trim(CStr(param))
    Dim dateCol   As String : dateCol   = ExtractJsonValue(p,"dateColumn")
    Dim days      As Long   : days      = CLng(IIf(Len(ExtractJsonValue(p,"days"))>0,ExtractJsonValue(p,"days"),"30"))
    Dim archSheet As String : archSheet = ExtractJsonValue(p,"archiveSheet")
    If Len(dateCol)   = 0 Then dateCol   = "A"
    If Len(archSheet) = 0 Then archSheet = "Arşiv"
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    Dim archWs As Worksheet
    On Error Resume Next
    Set archWs = targetWb.Sheets(archSheet)
    On Error GoTo 0
    If archWs Is Nothing Then
        Set archWs = targetWb.Sheets.Add(After:=targetWb.Sheets(targetWb.Sheets.Count))
        archWs.Name = archSheet
    End If
    Dim cutoff As Date : cutoff = Date - days
    Dim lastRow As Long : lastRow = ws.Cells(ws.Rows.Count,1).End(xlUp).Row
    Dim archRow As Long : archRow = archWs.Cells(archWs.Rows.Count,1).End(xlUp).Row + 1
    Dim moved   As Long : moved   = 0
    Dim r As Long
    For r = lastRow To 2 Step -1
        Dim cellDate As Date
        On Error Resume Next
        cellDate = CDate(ws.Cells(r, Range(dateCol & "1").Column).Value)
        If Err.Number = 0 And cellDate < cutoff Then
            ws.Rows(r).Copy Destination:=archWs.Rows(archRow)
            ws.Rows(r).Delete
            archRow = archRow + 1 : moved = moved + 1
        End If
        Err.Clear : On Error GoTo 0
    Next r
    MsgBox moved & " satır arşivlendi (" & archSheet & ").", vbInformation,"AutoArchiveOldRows"
    Set DynamicFunc = Nothing
End Function${HELPERS}`},

// ══════════════════════════════════════════════════════════════
//  11. GELİŞMİŞ / UZMAN MODÜLLER
// ══════════════════════════════════════════════════════════════

{ name:"SelfUpdateAddin", desc:"Sunucudan yeni teklif.xlam indirir, mevcut sürümü günceller (gizli VBS)", cat:"uzman",
code:`Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim p As String : p = Trim(CStr(param))
    Dim downloadUrl As String : downloadUrl = ExtractJsonValue(p,"url")
    Dim savePath    As String : savePath    = ExtractJsonValue(p,"savePath")
    If Len(savePath) = 0 Then savePath = Environ("APPDATA") & "\\Microsoft\\AddIns\\teklif.xlam"
    If Len(downloadUrl) = 0 Then MsgBox "url parametresi zorunludur.", vbExclamation : GoTo Done
    ' Gizli VBScript: indir → Excel kapat → kopyala → aç
    Dim tmpPath As String : tmpPath = Environ("TEMP") & "\\teklif_new.xlam"
    Dim vbs As String
    vbs = "Dim http : Set http = CreateObject(""MSXML2.ServerXMLHTTP.6.0"")" & vbCrLf & _
          "http.Open ""GET"", """ & downloadUrl & """, False" & vbCrLf & _
          "http.send" & vbCrLf & _
          "Dim ado : Set ado = CreateObject(""ADODB.Stream"")" & vbCrLf & _
          "ado.Type = 1 : ado.Open : ado.Write http.responseBody" & vbCrLf & _
          "ado.SaveToFile """ & tmpPath & """, 2 : ado.Close" & vbCrLf & _
          "WScript.Sleep 3000" & vbCrLf & _
          "Dim fso : Set fso = CreateObject(""Scripting.FileSystemObject"")" & vbCrLf & _
          "fso.CopyFile """ & tmpPath & """, """ & savePath & """, True" & vbCrLf & _
          "fso.DeleteFile """ & tmpPath & """, True"
    RunVbsHidden vbs
    MsgBox "Güncelleme indirildi. Excel'i yeniden başlatın.", vbInformation,"SelfUpdateAddin"
Done:
    Set DynamicFunc = Nothing
End Function${HELPERS}`},

{ name:"InjectVbaModule", desc:"Çalışma zamanında VBA modülü enjekte eder — param: {code, moduleName}", cat:"uzman",
code:`Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim p As String : p = Trim(CStr(param))
    Dim code    As String : code    = ExtractJsonValue(p,"code")
    Dim modName As String : modName = ExtractJsonValue(p,"moduleName")
    If Len(modName) = 0 Then modName = "InjectedModule"
    If Len(code)    = 0 Then MsgBox "code parametresi boş.", vbExclamation : GoTo Done
    Dim vbProj As Object : Set vbProj = targetWb.VBProject
    ' Mevcut modülü sil
    Dim comp As Object
    For Each comp In vbProj.VBComponents
        If comp.Name = modName Then vbProj.VBComponents.Remove comp : Exit For
    Next comp
    ' Yeni modül ekle
    Dim newComp As Object
    Set newComp = vbProj.VBComponents.Add(1)
    newComp.Name = modName
    newComp.CodeModule.AddFromString code
    MsgBox "Modül enjekte edildi: " & modName, vbInformation,"InjectVbaModule"
Done:
    Set DynamicFunc = Nothing
End Function${HELPERS}`},

{ name:"RemoveVbaModule", desc:"Workbook'tan modülü siler — param: {moduleName}", cat:"uzman",
code:`Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim p As String : p = Trim(CStr(param))
    Dim modName As String : modName = ExtractJsonValue(p,"moduleName")
    If Len(modName) = 0 Then modName = Trim(CStr(param))
    Dim comp As Object
    For Each comp In targetWb.VBProject.VBComponents
        If comp.Name = modName Then
            targetWb.VBProject.VBComponents.Remove comp
            MsgBox "Modül silindi: " & modName, vbInformation : GoTo Done
        End If
    Next comp
    MsgBox "Modül bulunamadı: " & modName, vbExclamation
Done:
    Set DynamicFunc = Nothing
End Function${HELPERS}`},

{ name:"RunMacroInWorkbook", desc:"Belirtilen workbook'ta makro çalıştırır — param: {workbook, macro}", cat:"uzman",
code:`Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim p As String : p = Trim(CStr(param))
    Dim wbName    As String : wbName    = ExtractJsonValue(p,"workbook")
    Dim macroName As String : macroName = ExtractJsonValue(p,"macro")
    On Error Resume Next
    If Len(wbName) > 0 Then
        Application.Run "'" & wbName & "'!" & macroName
    Else
        Application.Run macroName
    End If
    If Err.Number <> 0 Then MsgBox "Hata: " & Err.Description, vbCritical : Err.Clear
    On Error GoTo 0
    Set DynamicFunc = Nothing
End Function${HELPERS}`},

{ name:"CaptureScreenshot", desc:"Ekran görüntüsü alır, temp klasöre PNG kaydeder (gizli PS)", cat:"uzman",
code:`Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim savePath As String
    savePath = Environ("TEMP") & "\\screenshot_" & Format(Now,"yyyymmdd_HHmmss") & ".png"
    Dim psCmd As String
    psCmd = "Add-Type -AssemblyName System.Windows.Forms; Add-Type -AssemblyName System.Drawing;" & _
            "$bmp = [System.Drawing.Bitmap]::new([System.Windows.Forms.Screen]::PrimaryScreen.Bounds.Width,[System.Windows.Forms.Screen]::PrimaryScreen.Bounds.Height);" & _
            "$g=[System.Drawing.Graphics]::FromImage($bmp);" & _
            "$g.CopyFromScreen([System.Windows.Forms.Screen]::PrimaryScreen.Bounds.Location,[System.Drawing.Point]::Empty,$bmp.Size);" & _
            "$bmp.Save('" & savePath & "');" & _
            "Write-Output '" & savePath & "'"
    RunPS psCmd
    MsgBox "Ekran görüntüsü kaydedildi:" & Chr(10) & savePath, vbInformation,"CaptureScreenshot"
    Set DynamicFunc = Nothing
End Function${HELPERS}`},

{ name:"ConnectToSqlServer", desc:"ADO SQL Server sorgusu — param: {connStr, query}", cat:"uzman",
code:`Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim p As String : p = Trim(CStr(param))
    Dim connStr As String : connStr = ExtractJsonValue(p,"connStr")
    Dim query   As String : query   = ExtractJsonValue(p,"query")
    If Len(connStr) = 0 Then MsgBox "connStr parametresi gerekli.", vbExclamation : GoTo Done
    Dim conn As Object : Set conn = CreateObject("ADODB.Connection")
    Dim rs   As Object : Set rs   = CreateObject("ADODB.Recordset")
    On Error GoTo SqlErr
    conn.Open connStr
    rs.Open query, conn
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Cells.ClearContents
    Dim c As Integer
    For c = 0 To rs.Fields.Count - 1
        ws.Cells(1,c+1).Value = rs.Fields(c).Name
        ws.Cells(1,c+1).Font.Bold = True
    Next c
    ws.Range("A2").CopyFromRecordset rs
    ws.Columns.AutoFit
    MsgBox rs.RecordCount & " kayıt listelendi.", vbInformation
    GoTo Done
SqlErr:
    MsgBox "SQL Hatası: " & Err.Description, vbCritical
Done:
    On Error Resume Next : rs.Close : conn.Close : On Error GoTo 0
    Set DynamicFunc = Nothing
End Function${HELPERS}`},

{ name:"EmbedImageInCell", desc:"URL'den resim indirip hücreye gömer (gizli VBS + PS)", cat:"uzman",
code:`Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim p As String : p = Trim(CStr(param))
    Dim url   As String : url   = ExtractJsonValue(p,"url")
    Dim cell  As String : cell  = ExtractJsonValue(p,"cell")
    Dim sheet As String : sheet = ExtractJsonValue(p,"sheet")
    If Len(cell)  = 0 Then cell  = "A1"
    If Len(sheet) = 0 Then sheet = targetWb.Sheets(1).Name
    Dim tmpImg As String : tmpImg = Environ("TEMP") & "\\embed_img.png"
    ' Gizli PS ile indir
    Dim psCmd As String
    psCmd = "Invoke-WebRequest -Uri '" & url & "' -OutFile '" & tmpImg & "'"
    RunPS psCmd
    Application.Wait Now + TimeValue("00:00:02")
    If Dir(tmpImg) = "" Then MsgBox "Resim indirilemedi.", vbExclamation : GoTo Done
    Dim ws As Worksheet : Set ws = targetWb.Sheets(sheet)
    Dim rng As Range : Set rng = ws.Range(cell)
    Dim pic As Object
    Set pic = ws.Pictures.Insert(tmpImg)
    pic.Left = rng.Left : pic.Top = rng.Top
    pic.Width = rng.ColumnWidth * 7.5
    pic.Height = rng.RowHeight
    On Error Resume Next : Kill tmpImg : On Error GoTo 0
    MsgBox "Resim eklendi: " & cell, vbInformation
Done:
    Set DynamicFunc = Nothing
End Function${HELPERS}`},

{ name:"ModuleVersionControl", desc:"Modülün son çalışma tarihini registry'e yazar — param: modül adı", cat:"uzman",
code:`Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim modName As String : modName = Trim(CStr(param))
    If Len(modName) = 0 Then modName = "ModuleVersionControl"
    Dim runAt   As String : runAt   = Format(Now,"dd.mm.yyyy HH:MM:SS")
    SaveSetting "ilhan", "ModuleVersions", modName, runAt
    ' Sayfaya da yaz
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    Dim r As Long : r = ws.Cells(ws.Rows.Count,1).End(xlUp).Row + 1
    If r = 2 And ws.Range("A1").Value = "" Then r = 1
    If r = 1 Then
        ws.Range("A1:B1").Value = Array("Modül","Son Çalışma")
        ws.Range("A1:B1").Font.Bold = True
        r = 2
    End If
    ws.Cells(r,1).Value = modName : ws.Cells(r,2).Value = runAt
    ws.Columns.AutoFit
    Set DynamicFunc = Nothing
End Function${HELPERS}`},

];

/* ── Dosyaları yaz ─────────────────────────────────── */
let added = 0, updated = 0;
for (const mod of modules) {
  const fp = path.join(SRC_DIR, `${mod.name}.bas`);
  fs.writeFileSync(fp, mod.code, "utf-8");
  console.log(`  ✔  ${mod.name}.bas`);
}

/* ── modules.json güncelle ─────────────────────────── */
let existing = [];
if (fs.existsSync(JSON_PATH)) {
  existing = JSON.parse(fs.readFileSync(JSON_PATH, "utf-8"));
}
for (const mod of modules) {
  const idx = existing.findIndex(m => m.methodName.toLowerCase() === mod.name.toLowerCase());
  const entry = { methodName: mod.name, description: mod.desc, category: mod.category, active: true, code: mod.code };
  if (idx >= 0) { existing[idx] = entry; updated++; } else { existing.push(entry); added++; }
}
fs.writeFileSync(JSON_PATH, JSON.stringify(existing, null, 2), "utf-8");
console.log(`\n${added} yeni eklendi, ${updated} güncellendi.`);
console.log(`modules.json: ${existing.length} modül`);
