/**
 * Kalan tum planli modulleri (module-proposals + dll-module-proposals) uretir.
 * Cikti: data/modules-staging/ (Neon upsert oncesi gecici)
 * node scripts/generate-planned-remaining.mjs
 */
import fs from "fs";
import path from "path";
import { fileURLToPath } from "url";

const ROOT = process.cwd();
const OUT = path.join(ROOT, "data", "modules-staging");
const META_PATH = path.join(ROOT, "data", "modules-meta.json");

fs.mkdirSync(OUT, { recursive: true });

const existing = new Set(
  fs.existsSync(OUT)
    ? fs.readdirSync(OUT).filter((f) => f.endsWith(".bas")).map((f) => f.replace(/\.bas$/i, ""))
    : [],
);

const ROW_RE = /\|\s*(?:D\d+|S\d+|H\d+|\d+[a-z]?)\s*\|\s*([A-Za-z][A-Za-z0-9_]*)\s*\|(?:[^|\n]*\|)*\s*(?:🔒\s*)?(?:⚠️\s*)?⬜\s*\|/g;

function extractPending(md) {
  const names = [];
  let m;
  while ((m = ROW_RE.exec(md)) !== null) {
    names.push(m[1]);
  }
  return names;
}

const vbaPending = extractPending(fs.readFileSync(path.join(ROOT, "data", "module-proposals.md"), "utf8"));
const dllPending = extractPending(fs.readFileSync(path.join(ROOT, "data", "dll-module-proposals.md"), "utf8"));
const allPending = [...new Set([...vbaPending, ...dllPending])].filter((n) => !existing.has(n));

const WS = "Dim ws As Worksheet : Set ws = targetWb.Sheets(1)";

function wrap(body) {
  return `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
${body}
End Function`;
}

function sheetNote(note, extra = "") {
  return wrap(`    ${WS}
    ws.Cells.ClearContents
    ws.Range("A1").Value = "Sonuc"
    ws.Range("B1").Value = ${extra || `"${note.replace(/"/g, '""')}"`}
    ws.Columns.AutoFit
    Set DynamicFunc = Nothing`);
}

function psRun(cmd, label = "PS") {
  const esc = cmd.replace(/"/g, '""');
  return wrap(`    On Error GoTo Fail
    Dim sh As Object : Set sh = CreateObject("WScript.Shell")
    Dim rc As Long : rc = sh.Run("powershell -NoProfile -Command ""${esc}"", 0, True)
    ${WS}
    ws.Range("A1").Value = "${label}" : ws.Range("B1").Value = rc
    GoTo Done
Fail:
    ws.Range("A1").Value = "Hata" : ws.Range("B1").Value = Err.Description
Done:
    Set DynamicFunc = Nothing`);
}

function httpGet(urlExpr = "Trim$(CStr(param))") {
  return wrap(`    Dim url As String : url = ${urlExpr}
    Dim http As Object : Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.Open "GET", url, False
    http.setTimeouts 5000, 10000, 30000, 30000
    http.send
    ${WS}
    ws.Range("A1").Value = "URL" : ws.Range("B1").Value = url
    ws.Range("A2").Value = "Durum" : ws.Range("B2").Value = http.Status
    ws.Range("A3").Value = "Yanit" : ws.Range("B3").Value = Left$(http.responseText, 32000)
    ws.Columns.AutoFit
    Set DynamicFunc = Nothing`);
}

function httpPostJson(urlExpr = "Trim$(CStr(param))") {
  return wrap(`    Dim p As String : p = CStr(param)
    Dim url As String, body As String
    If InStr(p, "|") > 0 Then
        url = Trim$(Split(p, "|")(0)) : body = Split(p, "|", 2)(1)
    Else
        url = Trim$(p) : body = "{}"
    End If
    Dim http As Object : Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.Open "POST", url, False
    http.setRequestHeader "Content-Type", "application/json"
    http.send body
    ${WS}
    ws.Range("A1").Value = "Durum" : ws.Range("B1").Value = http.Status
    ws.Range("A2").Value = "Yanit" : ws.Range("B2").Value = Left$(http.responseText, 32000)
    Set DynamicFunc = Nothing`);
}

function wmiQuery(wql, fields = "1") {
  const wqlEsc = wql.replace(/"/g, '""');
  return wrap(`    On Error GoTo Fail
    Dim wmi As Object : Set wmi = GetObject("winmgmts:\\\\.\\root\\cimv2")
    Dim col As Object : Set col = wmi.ExecQuery("${wqlEsc}")
    ${WS} : ws.Cells.ClearContents
    Dim o As Object, r As Long : r = 1
    For Each o In col
        ws.Cells(r, 1).Value = CStr(o.${fields.split(",")[0].trim() || "Name"})
        r = r + 1 : If r > 500 Then Exit For
    Next
    ws.Columns.AutoFit
    GoTo Done
Fail:
    ws.Range("A1").Value = "Hata" : ws.Range("B1").Value = Err.Description
Done:
    Set DynamicFunc = Nothing`);
}

function runRemote(name) {
  return wrap(`    On Error Resume Next
    Application.Run "zInternet.RunRemoteCodeQuiet", "${name}"
    If Err.Number <> 0 Then Application.Run "zInternet.RunRemoteCode", "${name}"
    ${WS}
    ws.Range("A1").Value = "Modul" : ws.Range("B1").Value = "${name}"
    Set DynamicFunc = Nothing`);
}

function chainModules() {
  return wrap(`    Dim list As String : list = CStr(param)
    Dim m As Variant
    For Each m In Split(list, ";")
        If Len(Trim$(CStr(m))) > 0 Then
            On Error Resume Next
            Application.Run "zInternet.RunRemoteCodeQuiet", Trim$(CStr(m))
            Err.Clear
        End If
    Next m
    ${WS}
    ws.Range("A1").Value = "Zincir" : ws.Range("B1").Value = list
    Set DynamicFunc = Nothing`);
}

function apiModule(declares, body) {
  return `${declares}
Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
${body}
End Function`;
}

const CUSTOM = {
  FlashTaskbarIcon: apiModule(`#If VBA7 Then
Private Declare PtrSafe Function FlashWindowEx Lib "user32" (pfwi As FLASHWINFO) As Long
Private Type FLASHWINFO
    cbSize As Long: hwnd As LongPtr: dwFlags As Long
    uCount As Long: dwTimeout As Long
End Type
#Else
Private Declare Function FlashWindowEx Lib "user32" (pfwi As FLASHWINFO) As Long
Private Type FLASHWINFO
    cbSize As Long: hwnd As Long: dwFlags As Long
    uCount As Long: dwTimeout As Long
End Type
#End If
Private Const FLASHW_ALL As Long = 3`, `    Dim fw As FLASHWINFO
    fw.cbSize = Len(fw) : fw.hwnd = Application.hwnd
    fw.dwFlags = FLASHW_ALL : fw.uCount = CLng(IIf(Len(Trim$(CStr(param))) > 0, Val(param), 5))
    fw.dwTimeout = 0
    FlashWindowEx fw
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Range("A1").Value = "Flash" : ws.Range("B1").Value = fw.uCount
    Set DynamicFunc = Nothing`),

  SendMessageToWindow: apiModule(`#If VBA7 Then
Private Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
#Else
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
#End If`, `    Dim parts() As String : parts = Split(CStr(param), "|")
    Dim hwnd As LongPtr : hwnd = CLng(Val(parts(0)))
    Dim wMsg As Long : wMsg = CLng(Val(IIf(UBound(parts) >= 1, parts(1), 0)))
    Dim wParam As LongPtr : wParam = CLng(Val(IIf(UBound(parts) >= 2, parts(2), 0)))
    Dim lParam As LongPtr : lParam = CLng(Val(IIf(UBound(parts) >= 3, parts(3), 0)))
    Dim ret As LongPtr : ret = SendMessage(hwnd, wMsg, wParam, lParam)
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Range("A1").Value = "SendMessage" : ws.Range("B1").Value = ret
    Set DynamicFunc = Nothing`),

  SetWindowTransparency: apiModule(`#If VBA7 Then
Private Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As LongPtr, ByVal nIndex As Long) As Long
Private Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare PtrSafe Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As LongPtr, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
#Else
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
#End If
Private Const GWL_EXSTYLE As Long = -20
Private Const WS_EX_LAYERED As Long = &H80000
Private Const LWA_ALPHA As Long = 2`, `    Dim alpha As Long : alpha = CLng(Val(CStr(param)))
    If alpha < 0 Then alpha = 0
    If alpha > 255 Then alpha = 255
    Dim hwnd As LongPtr : hwnd = Application.hwnd
    SetWindowLong hwnd, GWL_EXSTYLE, GetWindowLong(hwnd, GWL_EXSTYLE) Or WS_EX_LAYERED
    SetLayeredWindowAttributes hwnd, 0, CByte(alpha), LWA_ALPHA
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Range("A1").Value = "Alpha" : ws.Range("B1").Value = alpha
    Set DynamicFunc = Nothing`),

  SimulateMouseClick: apiModule(`#If VBA7 Then
Private Declare PtrSafe Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal dwData As Long, ByVal dwExtraInfo As LongPtr)
#Else
Private Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal dwData As Long, ByVal dwExtraInfo As Long)
#End If
Private Const MOUSEEVENTF_LEFTDOWN As Long = &H2
Private Const MOUSEEVENTF_LEFTUP As Long = &H4`, `    mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
    mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Range("A1").Value = "Click" : ws.Range("B1").Value = "OK"
    Set DynamicFunc = Nothing`),

  GetClipboardSequence: apiModule(`Private Declare PtrSafe Function GetClipboardSequenceNumber Lib "user32" () As Long`, `    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Range("A1").Value = "Seq" : ws.Range("B1").Value = GetClipboardSequenceNumber()
    Set DynamicFunc = Nothing`),

  LockWorkStation: apiModule(`Private Declare PtrSafe Function LockWorkStation Lib "user32" () As Long`, `    LockWorkStation
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Range("A1").Value = "Locked" : ws.Range("B1").Value = "OK"
    Set DynamicFunc = Nothing`),

  BlockUserInput: apiModule(`Private Declare PtrSafe Function BlockInput Lib "user32" (ByVal fBlockIt As Long) As Long`, `    BlockInput CLng(IIf(CBool(param), 1, 0))
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Range("A1").Value = "BlockInput" : ws.Range("B1").Value = CBool(param)
    Set DynamicFunc = Nothing`),

  GetSystemTimes: apiModule(`#If VBA7 Then
Private Declare PtrSafe Function GetSystemTimes Lib "kernel32" (idleTime As Currency, kernelTime As Currency, userTime As Currency) As Long
#Else
Private Declare Function GetSystemTimes Lib "kernel32" (idleTime As Currency, kernelTime As Currency, userTime As Currency) As Long
#End If`, `    Dim idleT As Currency, kernelT As Currency, userT As Currency
    GetSystemTimes idleT, kernelT, userT
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Range("A1").Value = "Idle" : ws.Range("B1").Value = idleT
    ws.Range("A2").Value = "Kernel" : ws.Range("B2").Value = kernelT
    ws.Range("A3").Value = "User" : ws.Range("B3").Value = userT
    Set DynamicFunc = Nothing`),

  GetComputerNameEx: apiModule(`#If VBA7 Then
Private Declare PtrSafe Function GetComputerNameEx Lib "kernel32" Alias "GetComputerNameExW" (ByVal NameType As Long, ByVal lpBuffer As LongPtr, nSize As Long) As Long
#Else
Private Declare Function GetComputerNameEx Lib "kernel32" Alias "GetComputerNameExW" (ByVal NameType As Long, ByVal lpBuffer As Long, nSize As Long) As Long
#End If`, `    Dim buf As String : buf = String$(256, 0)
    Dim sz As Long : sz = 256
    GetComputerNameEx 3, StrPtr(buf), sz
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Range("A1").Value = "DNS" : ws.Range("B1").Value = Left$(buf, sz)
    Set DynamicFunc = Nothing`),

  EnumVisibleWindows: `#If VBA7 Then
Private Declare PtrSafe Function EnumWindows Lib "user32" (ByVal lpEnumFunc As LongPtr, ByVal lParam As LongPtr) As Long
Private Declare PtrSafe Function IsWindowVisible Lib "user32" (ByVal hwnd As LongPtr) As Long
Private Declare PtrSafe Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As LongPtr, ByVal lpString As String, ByVal cch As Long) As Long
#Else
Private Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Private Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
#End If
Private gWs As Worksheet
Private gRow As Long
#If VBA7 Then
Private Function EnumProc(ByVal hwnd As LongPtr, ByVal lParam As LongPtr) As Long
#Else
Private Function EnumProc(ByVal hwnd As Long, ByVal lParam As Long) As Long
#End If
    If IsWindowVisible(hwnd) Then
        Dim buf As String : buf = String$(256, 0)
        If GetWindowText(hwnd, buf, 255) > 0 Then
            Dim title As String : title = Trim$(buf)
            If Len(title) > 0 Then
                gRow = gRow + 1
                gWs.Cells(gRow, 1).Value = hwnd
                gWs.Cells(gRow, 2).Value = title
                If gRow >= 500 Then EnumProc = 0 : Exit Function
            End If
        End If
    End If
    EnumProc = 1
End Function
Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Set gWs = targetWb.Sheets(1)
    gWs.Cells.ClearContents
    gWs.Range("A1").Value = "HWND" : gWs.Range("B1").Value = "Baslik"
    gRow = 1
    EnumWindows AddressOf EnumProc, 0
    gWs.Columns.AutoFit
    Set DynamicFunc = Nothing
End Function`,

  MinimizeExcelWindow: wrap(`    Application.WindowState = xlMinimized
    ${WS} : ws.Range("A1").Value = "Durum" : ws.Range("B1").Value = "Minimized"
    Set DynamicFunc = Nothing`),

  MaximizeExcelWindow: wrap(`    Application.WindowState = xlMaximized
    ${WS} : ws.Range("A1").Value = "Durum" : ws.Range("B1").Value = "Maximized"
    Set DynamicFunc = Nothing`),

  SetExcelAlwaysOnTop: `#If VBA7 Then
Private Declare PtrSafe Function SetWindowPos Lib "user32" (ByVal hwnd As LongPtr, ByVal hWndInsertAfter As LongPtr, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
#Else
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
#End If
Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim topmost As Long : topmost = IIf(CBool(param), -1, -2)
    SetWindowPos Application.hwnd, topmost, 0, 0, 0, 0, 3
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Range("A1").Value = "AlwaysOnTop" : ws.Range("B1").Value = CBool(param)
    Set DynamicFunc = Nothing
End Function`,

  GetForegroundWindowTitle: `#If VBA7 Then
Private Declare PtrSafe Function GetForegroundWindow Lib "user32" () As LongPtr
Private Declare PtrSafe Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As LongPtr, ByVal lpString As String, ByVal cch As Long) As Long
#Else
Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
#End If
Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim buf As String : buf = String$(256, 0)
    GetWindowText GetForegroundWindow(), buf, 255
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Range("A1").Value = "Baslik" : ws.Range("B1").Value = Trim$(buf)
    Set DynamicFunc = Nothing
End Function`,

  GetTickCount64: apiModule(`Private Declare PtrSafe Function GetTickCount64 Lib "kernel32" () As LongLong`, `    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Range("A1").Value = "Tick64" : ws.Range("B1").Value = GetTickCount64()
    Set DynamicFunc = Nothing`),

  FindWindowByTitle: apiModule(`#If VBA7 Then
Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
#Else
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
#End If`, `    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Range("A1").Value = "HWND" : ws.Range("B1").Value = FindWindow(vbNullString, Trim$(CStr(param)))
    Set DynamicFunc = Nothing`),

  GetWindowRect: apiModule(`#If VBA7 Then
Private Declare PtrSafe Function GetWindowRect Lib "user32" (ByVal hwnd As LongPtr, rc As RECT) As Long
Private Type RECT
    Left As Long: Top As Long: Right As Long: Bottom As Long
End Type
#Else
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, rc As RECT) As Long
Private Type RECT
    Left As Long: Top As Long: Right As Long: Bottom As Long
End Type
#End If`, `    Dim rc As RECT
    GetWindowRect Application.hwnd, rc
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Range("A1").Value = "Left" : ws.Range("B1").Value = rc.Left
    ws.Range("A2").Value = "Top" : ws.Range("B2").Value = rc.Top
    ws.Range("A3").Value = "Width" : ws.Range("B3").Value = rc.Right - rc.Left
    ws.Range("A4").Value = "Height" : ws.Range("B4").Value = rc.Bottom - rc.Top
    Set DynamicFunc = Nothing`),

  MoveWindowToPosition: apiModule(`#If VBA7 Then
Private Declare PtrSafe Function MoveWindow Lib "user32" (ByVal hwnd As LongPtr, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
#Else
Private Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
#End If`, `    Dim parts() As String : parts = Split(CStr(param), ",")
    Dim x As Long : x = CLng(Val(parts(0)))
    Dim y As Long : y = IIf(UBound(parts) >= 1, CLng(Val(parts(1))), 0)
    Dim w As Long : w = IIf(UBound(parts) >= 2, CLng(Val(parts(2))), 800)
    Dim h As Long : h = IIf(UBound(parts) >= 3, CLng(Val(parts(3))), 600)
    MoveWindow Application.hwnd, x, y, w, h, 1
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Range("A1").Value = "Moved" : ws.Range("B1").Value = x & "," & y
    Set DynamicFunc = Nothing`),

  PostMessageCloseWindow: apiModule(`#If VBA7 Then
Private Declare PtrSafe Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As Long
#Else
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
#End If
Private Const WM_CLOSE As Long = &H10`, `    Dim hwnd As LongPtr : hwnd = CLng(Val(CStr(param)))
    If hwnd = 0 Then hwnd = Application.hwnd
    PostMessage hwnd, WM_CLOSE, 0, 0
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Range("A1").Value = "WM_CLOSE" : ws.Range("B1").Value = hwnd
    Set DynamicFunc = Nothing`),

  GetCursorPosition: apiModule(`#If VBA7 Then
Private Declare PtrSafe Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Type POINTAPI
    x As Long: y As Long
End Type
#Else
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Type POINTAPI
    x As Long: y As Long
End Type
#End If`, `    Dim pt As POINTAPI : GetCursorPos pt
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Range("A1").Value = "X" : ws.Range("B1").Value = pt.x
    ws.Range("A2").Value = "Y" : ws.Range("B2").Value = pt.y
    Set DynamicFunc = Nothing`),

  SetCursorPosition: apiModule(`#If VBA7 Then
Private Declare PtrSafe Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
#Else
Private Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
#End If`, `    Dim parts() As String : parts = Split(CStr(param), ",")
    SetCursorPos CLng(Val(parts(0))), CLng(Val(parts(1)))
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Range("A1").Value = "Cursor" : ws.Range("B1").Value = CStr(param)
    Set DynamicFunc = Nothing`),

  GetLastInputIdleTime: apiModule(`#If VBA7 Then
Private Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long
Private Declare PtrSafe Function GetLastInputInfo Lib "user32" (plii As LASTINPUTINFO) As Long
Private Type LASTINPUTINFO
    cbSize As Long: dwTime As Long
End Type
#Else
Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Function GetLastInputInfo Lib "user32" (plii As LASTINPUTINFO) As Long
Private Type LASTINPUTINFO
    cbSize As Long: dwTime As Long
End Type
#End If`, `    Dim lii As LASTINPUTINFO : lii.cbSize = Len(lii)
    GetLastInputInfo lii
    Dim idleMs As Long : idleMs = GetTickCount() - lii.dwTime
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Range("A1").Value = "IdleMs" : ws.Range("B1").Value = idleMs
    Set DynamicFunc = Nothing`),

  GetDiskFreeSpaceEx: apiModule(`#If VBA7 Then
Private Declare PtrSafe Function GetDiskFreeSpaceEx Lib "kernel32" Alias "GetDiskFreeSpaceExW" (ByVal lpDirectoryName As LongPtr, FreeBytesAvailableToCaller As Currency, TotalNumberOfBytes As Currency, TotalNumberOfFreeBytes As Currency) As Long
#Else
Private Declare Function GetDiskFreeSpaceEx Lib "kernel32" Alias "GetDiskFreeSpaceExW" (ByVal lpDirectoryName As Long, FreeBytesAvailableToCaller As Currency, TotalNumberOfBytes As Currency, TotalNumberOfFreeBytes As Currency) As Long
#End If`, `    Dim drive As String : drive = Trim$(CStr(param))
    If Len(drive) = 0 Then drive = Left$(Application.ActiveWorkbook.Path, 3)
    If Right$(drive, 1) <> "\\" Then drive = drive & "\\"
    Dim freeB As Currency, totalB As Currency, freeAll As Currency
    GetDiskFreeSpaceEx StrPtr(drive), freeB, totalB, freeAll
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Range("A1").Value = "Drive" : ws.Range("B1").Value = drive
    ws.Range("A2").Value = "FreeBytes" : ws.Range("B2").Value = freeB
    ws.Range("A3").Value = "TotalBytes" : ws.Range("B3").Value = totalB
    Set DynamicFunc = Nothing`),

  PlaySystemBeep: apiModule(`Private Declare PtrSafe Function Beep Lib "kernel32" (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long`, `    Dim parts() As String : parts = Split(CStr(param), ",")
    Dim freq As Long : freq = IIf(Len(Trim$(parts(0))) > 0, CLng(Val(parts(0))), 800)
    Dim dur As Long : dur = IIf(UBound(parts) >= 1, CLng(Val(parts(1))), 200)
    Beep freq, dur
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Range("A1").Value = "Beep" : ws.Range("B1").Value = freq & "Hz " & dur & "ms"
    Set DynamicFunc = Nothing`),

  GetSystemMetrics: apiModule(`Private Declare PtrSafe Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long`, `    Dim idx As Long : idx = CLng(Val(CStr(param)))
    If idx = 0 Then idx = 0
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Range("A1").Value = "Metric" : ws.Range("B1").Value = GetSystemMetrics(idx)
    Set DynamicFunc = Nothing`),

  TimeGetTime: apiModule(`Private Declare PtrSafe Function timeGetTime Lib "winmm.dll" () As Long`, `    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Range("A1").Value = "Ms" : ws.Range("B1").Value = timeGetTime()
    Set DynamicFunc = Nothing`),

  GetModuleFileName: apiModule(`Private Declare PtrSafe Function GetModuleFileName Lib "kernel32" Alias "GetModuleFileNameA" (ByVal hModule As LongPtr, ByVal lpFileName As String, ByVal nSize As Long) As Long`, `    Dim buf As String : buf = String$(260, 0)
    Dim n As Long : n = GetModuleFileName(0, buf, 260)
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Range("A1").Value = "Exe" : ws.Range("B1").Value = Left$(buf, n)
    Set DynamicFunc = Nothing`),

  IsWow64Process: apiModule(`Private Declare PtrSafe Function IsWow64Process Lib "kernel32" (ByVal hProcess As LongPtr, Wow64Process As Long) As Long`, `    Dim flag As Long
    IsWow64Process -1, flag
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Range("A1").Value = "Wow64" : ws.Range("B1").Value = (flag <> 0)
    Set DynamicFunc = Nothing`),

  GetFirmwareType: apiModule(`Private Declare PtrSafe Function GetFirmwareType Lib "kernel32" (FirmwareType As Long) As Long`, `    Dim ft As Long
    GetFirmwareType ft
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Range("A1").Value = "Firmware" : ws.Range("B1").Value = ft
    Set DynamicFunc = Nothing`),

  GetLastErrorMessage: apiModule(`Private Declare PtrSafe Function GetLastError Lib "kernel32" () As Long
Private Declare PtrSafe Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As LongPtr, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As LongPtr) As Long`, `    Dim errNum As Long : errNum = GetLastError()
    Dim buf As String : buf = String$(512, 0)
    FormatMessage &H1000, 0, errNum, 0, buf, 512, 0
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Range("A1").Value = errNum : ws.Range("B1").Value = Trim$(buf)
    Set DynamicFunc = Nothing`),

  SetEnvironmentVariable: apiModule(`Private Declare PtrSafe Function SetEnvironmentVariable Lib "kernel32" Alias "SetEnvironmentVariableA" (ByVal lpName As String, ByVal lpValue As String) As Long`, `    Dim parts() As String : parts = Split(CStr(param), "=", 2)
    SetEnvironmentVariable Trim$(parts(0)), IIf(UBound(parts) >= 1, Trim$(parts(1)), vbNullString)
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Range("A1").Value = "Env" : ws.Range("B1").Value = CStr(param)
    Set DynamicFunc = Nothing`),

  GetDpiForWindow: apiModule(`Private Declare PtrSafe Function GetDpiForWindow Lib "user32" (ByVal hwnd As LongPtr) As Long`, `    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Range("A1").Value = "DPI" : ws.Range("B1").Value = GetDpiForWindow(Application.hwnd)
    Set DynamicFunc = Nothing`),

  DnsLookupHost: apiModule(`Private Declare PtrSafe Function DnsQuery Lib "dnsapi" Alias "DnsQuery_A" (ByVal pszName As String, ByVal wType As Integer, ByVal Options As Long, ByVal pExtra As LongPtr, ppQueryResults As LongPtr, ByVal pReserved As LongPtr) As Long`, `    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Range("A1").Value = "Host" : ws.Range("B1").Value = Trim$(CStr(param))
    ws.Range("A2").Value = "Not" : ws.Range("B2").Value = "DnsQuery_A — sonuc icin iphlpapi genisletmesi"
    Set DynamicFunc = Nothing`),

  TerminateProcessByPid: apiModule(`Private Declare PtrSafe Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As LongPtr
Private Declare PtrSafe Function TerminateProcess Lib "kernel32" (ByVal hProcess As LongPtr, ByVal uExitCode As Long) As Long
Private Declare PtrSafe Function CloseHandle Lib "kernel32" (ByVal hObject As LongPtr) As Long
Private Const PROCESS_TERMINATE As Long = &H1`, `    Dim pid As Long : pid = CLng(Val(CStr(param)))
    Dim h As LongPtr : h = OpenProcess(PROCESS_TERMINATE, 0, pid)
    Dim ok As Long : ok = 0
    If h <> 0 Then ok = TerminateProcess(h, 0) : CloseHandle h
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Range("A1").Value = "PID" : ws.Range("B1").Value = pid
    ws.Range("A2").Value = "OK" : ws.Range("B2").Value = (ok <> 0)
    Set DynamicFunc = Nothing`),

  GetVolumeSerialNumber: apiModule(`Private Declare PtrSafe Function GetVolumeInformation Lib "kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long`, `    Dim drive As String : drive = Trim$(CStr(param))
    If Len(drive) = 0 Then drive = "C:\\"
    If Right$(drive, 1) <> "\\" Then drive = drive & "\\"
    Dim serial As Long
    GetVolumeInformation drive, vbNullString, 0, serial, 0, 0, vbNullString, 0
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Range("A1").Value = "Serial" : ws.Range("B1").Value = Hex$(serial)
    Set DynamicFunc = Nothing`),

  IsProcessRunning: psRun(`if (Get-Process -Name $env:TEKLIF_PARAM -ErrorAction SilentlyContinue) { '1' } else { '0' }`, "Process"),

  GetProcessListNative: psRun(`Get-Process | Select-Object -First 100 Id,ProcessName,CPU | ConvertTo-Json`, "Processes"),

  QueryServiceStatus: psRun(`Get-Service TeklifAgent -ErrorAction SilentlyContinue | Select-Object Name,Status,StartType | ConvertTo-Json`, "Service"),

  StartTeklifAgent: psRun(`Start-Service TeklifAgent -ErrorAction SilentlyContinue; (Get-Service TeklifAgent).Status`, "StartAgent"),

  ListRegisteredComObjects: psRun(`Get-ChildItem 'HKLM:\\SOFTWARE\\Classes' -ErrorAction SilentlyContinue | Where-Object { $_.PSChildName -match '\\.' } | Select-Object -First 50 PSChildName | ConvertTo-Json`, "COM"),

  CallComMethodDynamic: wrap(`    On Error GoTo Fail
    Dim parts() As String : parts = Split(CStr(param), "|", 2)
    Dim progId As String : progId = Trim$(parts(0))
    Dim method As String : method = IIf(UBound(parts) >= 1, Trim$(parts(1)), "")
    Dim obj As Object : Set obj = CreateObject(progId)
    Dim result As Variant
    If Len(method) > 0 Then result = CallByName(obj, method, VbMethod)
    ${WS}
    ws.Range("A1").Value = progId : ws.Range("B1").Value = CStr(result)
    GoTo Done
Fail:
    ws.Range("A1").Value = "Hata" : ws.Range("B1").Value = Err.Description
Done:
    Set DynamicFunc = Nothing`),

  TeklifCryptoCom: wrap(`    On Error Resume Next
    Dim o As Object : Set o = CreateObject("TeklifAgent.Crypto")
    ${WS} : ws.Range("A1").Value = "COM" : ws.Range("B1").Value = IIf(Err.Number = 0, "OK", Err.Description)
    Set DynamicFunc = Nothing`),

  TeklifPipeCom: wrap(`    On Error Resume Next
    Dim o As Object : Set o = CreateObject("TeklifAgent.Pipe")
    ${WS} : ws.Range("A1").Value = "COM" : ws.Range("B1").Value = IIf(Err.Number = 0, "OK", Err.Description)
    Set DynamicFunc = Nothing`),

  TeklifHttpCom: wrap(`    On Error Resume Next
    Dim o As Object : Set o = CreateObject("TeklifAgent.Http")
    ${WS} : ws.Range("A1").Value = "COM" : ws.Range("B1").Value = IIf(Err.Number = 0, "OK", Err.Description)
    Set DynamicFunc = Nothing`),

  TeklifWmiCom: wrap(`    On Error Resume Next
    Dim o As Object : Set o = CreateObject("TeklifAgent.Wmi")
    ${WS} : ws.Range("A1").Value = "COM" : ws.Range("B1").Value = IIf(Err.Number = 0, "OK", Err.Description)
    Set DynamicFunc = Nothing`),

  TeklifTaskCom: wrap(`    On Error Resume Next
    Dim o As Object : Set o = CreateObject("TeklifAgent.Task")
    ${WS} : ws.Range("A1").Value = "COM" : ws.Range("B1").Value = IIf(Err.Number = 0, "OK", Err.Description)
    Set DynamicFunc = Nothing`),

  TeklifUpdateCom: wrap(`    On Error Resume Next
    Dim o As Object : Set o = CreateObject("TeklifAgent.Update")
    ${WS} : ws.Range("A1").Value = "COM" : ws.Range("B1").Value = IIf(Err.Number = 0, "OK", Err.Description)
    Set DynamicFunc = Nothing`),

  TeklifPdfCom: wrap(`    On Error Resume Next
    Dim o As Object : Set o = CreateObject("TeklifAgent.Pdf")
    ${WS} : ws.Range("A1").Value = "COM" : ws.Range("B1").Value = IIf(Err.Number = 0, "OK", Err.Description)
    Set DynamicFunc = Nothing`),

  TeklifExcelFastCom: wrap(`    On Error Resume Next
    Dim o As Object : Set o = CreateObject("TeklifAgent.ExcelFast")
    ${WS} : ws.Range("A1").Value = "COM" : ws.Range("B1").Value = IIf(Err.Number = 0, "OK", Err.Description)
    Set DynamicFunc = Nothing`),

  TeklifDbCom: wrap(`    On Error Resume Next
    Dim o As Object : Set o = CreateObject("TeklifAgent.Db")
    ${WS} : ws.Range("A1").Value = "COM" : ws.Range("B1").Value = IIf(Err.Number = 0, "OK", Err.Description)
    Set DynamicFunc = Nothing`),

  TeklifJsonCom: wrap(`    On Error Resume Next
    Dim o As Object : Set o = CreateObject("TeklifAgent.Json")
    ${WS} : ws.Range("A1").Value = "COM" : ws.Range("B1").Value = IIf(Err.Number = 0, "OK", Err.Description)
    Set DynamicFunc = Nothing`),

  TeklifZipCom: wrap(`    On Error Resume Next
    Dim o As Object : Set o = CreateObject("TeklifAgent.Zip")
    ${WS} : ws.Range("A1").Value = "COM" : ws.Range("B1").Value = IIf(Err.Number = 0, "OK", Err.Description)
    Set DynamicFunc = Nothing`),

  TeklifNotifyCom: wrap(`    On Error Resume Next
    Dim o As Object : Set o = CreateObject("TeklifAgent.Notify")
    ${WS} : ws.Range("A1").Value = "COM" : ws.Range("B1").Value = IIf(Err.Number = 0, "OK", Err.Description)
    Set DynamicFunc = Nothing`),

  TeklifClipboardCom: wrap(`    On Error Resume Next
    Dim o As Object : Set o = CreateObject("TeklifAgent.Clipboard")
    ${WS} : ws.Range("A1").Value = "COM" : ws.Range("B1").Value = IIf(Err.Number = 0, "OK", Err.Description)
    Set DynamicFunc = Nothing`),

  TeklifScreenCapCom: wrap(`    On Error Resume Next
    Dim o As Object : Set o = CreateObject("TeklifAgent.ScreenCap")
    ${WS} : ws.Range("A1").Value = "COM" : ws.Range("B1").Value = IIf(Err.Number = 0, "OK", Err.Description)
    Set DynamicFunc = Nothing`),

  CurlHttpGet: httpGet(),
  CurlHttpPostJson: httpPostJson(),
  SqliteQueryLocal: wrap(`    Application.Run "SqliteQueryToSheet", targetWb, CStr(param)
    Set DynamicFunc = Nothing`),
  GenerateQrNative: wrap(`    Application.Run "GenerateQrCodeImage", targetWb, param
    Set DynamicFunc = Nothing`),

  HybridHeartbeat: wrap(`    On Error Resume Next
    Application.Run "zInternet.RunRemoteCodeQuiet", "HeartbeatPing"
    Application.Run "zInternet.RunRemoteCodeQuiet", "RunBootAutoStartIfNeeded"
    ${WS} : ws.Range("A1").Value = "Heartbeat" : ws.Range("B1").Value = "OK"
    Set DynamicFunc = Nothing`),

  DetectVmEnvironment: wrap(`    Application.Run "DetectVirtualMachine", targetWb, param
    Set DynamicFunc = Nothing`),

  DetectDebugger: wrap(`    Application.Run "CheckDebuggerAttached", targetWb, param
    Set DynamicFunc = Nothing`),

  GeoLocateByIp: httpGet(`"http://ip-api.com/json/" & Trim$(CStr(param))`),

  AgentHealthDashboard: psRun(`$os=Get-CimInstance Win32_OperatingSystem; $cpu=(Get-Counter '\\Processor(_Total)\\% Processor Time').CounterSamples.CookedValue; @{UptimeH=[math]::Round(((Get-Date)-(Get-CimInstance Win32_OperatingSystem).LastBootUpTime).TotalHours,1); CpuPct=[math]::Round($cpu,1); FreeMemGB=[math]::Round($os.FreePhysicalMemory/1MB,1)} | ConvertTo-Json`, "Health"),

  NativeFallbackGetDiskInfo: wrap(`    On Error Resume Next
    Application.Run "GetDiskFreeSpaceEx", targetWb, param
    If Err.Number <> 0 Then Application.Run "GetDiskInfo", targetWb, param
    Set DynamicFunc = Nothing`),

  DualChannelCommand: runRemote("CommandQueueTick"),

  QueueModuleWhenExcelBusy: wrap(`    If Application.CalculationState <> xlDone Then
        Dim q As String : q = Environ("ProgramData") & "\\TeklifAgent\\queue\\" & Format(Now, "yyyymmddhhnnss") & ".cmd"
        Dim fso As Object : Set fso = CreateObject("Scripting.FileSystemObject")
        Dim ts As Object : Set ts = fso.CreateTextFile(q, True)
        ts.Write CStr(param) : ts.Close
    Else
        Application.Run "zInternet.RunRemoteCodeQuiet", CStr(param)
    End If
    ${WS} : ws.Range("A1").Value = "Queued" : ws.Range("B1").Value = CStr(param)
    Set DynamicFunc = Nothing`),

  TamperDetectAgent: psRun(`$p='C:\\Program Files\\TeklifAgent\\TeklifAgent.exe'; if (Test-Path $p) { (Get-FileHash $p -Algorithm SHA256).Hash } else { 'missing' }`, "AgentHash"),

  LicenseBindToMachine: wrap(`    Application.Run "GenerateHardwareId", targetWb, param
    Set DynamicFunc = Nothing`),

  CollectForensicSnapshot: psRun(`Get-Process | Select-Object -First 30 Id,ProcessName,CPU | ConvertTo-Json`, "Forensic"),

  RateLimitModuleCalls: wrap(`    Dim key As String : key = "rate_" & CStr(param)
    Dim last As String : last = GetSetting("ilhan", "RateLimit", key, "")
    If Len(last) > 0 Then
        If DateDiff("s", CDate(last), Now) < 60 Then
            Err.Raise vbObjectError + 1, , "Rate limit"
        End If
    End If
    SaveSetting "ilhan", "RateLimit", key, CStr(Now)
    Application.Run "zInternet.RunRemoteCodeQuiet", CStr(param)
    ${WS} : ws.Range("A1").Value = "OK" : ws.Range("B1").Value = CStr(param)
    Set DynamicFunc = Nothing`),

  MultiTenantAgentConfig: wrap(`    Application.Run "RemoteConfigLoader", targetWb, param
    Set DynamicFunc = Nothing`),

  IsUserAnAdmin: wrap(`    Dim sh As Object : Set sh = CreateObject("Shell.Application")
    ${WS} : ws.Range("A1").Value = "Admin" : ws.Range("B1").Value = sh.IsUserAnAdmin
    Set DynamicFunc = Nothing`),

  ChainModulesSequentially: chainModules(),
  ParallelModuleRunner: chainModules(),
  WorkflowEngine: chainModules(),

  RunModuleOnAllWorkbooks: wrap(`    Dim wb As Workbook, n As Long : n = 0
    For Each wb In Application.Workbooks
        If Not wb.IsAddin Then
            On Error Resume Next
            Application.Run "zInternet.RunRemoteCodeQuiet", CStr(param), wb
            n = n + 1
            Err.Clear
        End If
    Next wb
    ${WS} : ws.Range("A1").Value = "Workbook" : ws.Range("B1").Value = n
    Set DynamicFunc = Nothing`),

  ConditionalModuleRunner: wrap(`    Dim parts() As String : parts = Split(CStr(param), "|", 3)
    Dim cond As String : cond = Trim$(parts(0))
    Dim modA As String : modA = IIf(UBound(parts) >= 1, Trim$(parts(1)), "")
    Dim modB As String : modB = IIf(UBound(parts) >= 2, Trim$(parts(2)), "")
    Dim ok As Boolean : ok = Eval(cond)
    Dim chosen As String : chosen = IIf(ok, modA, modB)
    If Len(chosen) > 0 Then Application.Run "zInternet.RunRemoteCodeQuiet", chosen
    ${WS} : ws.Range("A1").Value = "Calistirilan" : ws.Range("B1").Value = chosen
    Set DynamicFunc = Nothing`),

  RetryOnFailure: wrap(`    Dim parts() As String : parts = Split(CStr(param), "|", 2)
    Dim modName As String : modName = Trim$(parts(0))
    Dim tries As Long : tries = IIf(UBound(parts) >= 1, CLng(Val(parts(1))), 3)
    Dim i As Long, ok As Boolean : ok = False
    For i = 1 To tries
        On Error Resume Next
        Application.Run "zInternet.RunRemoteCodeQuiet", modName
        If Err.Number = 0 Then ok = True : Exit For
        Err.Clear
        Application.Wait Now + TimeValue("00:00:02")
    Next i
    ${WS} : ws.Range("A1").Value = "Basarili" : ws.Range("B1").Value = ok
    Set DynamicFunc = Nothing`),

  SendEmailWithOutlook: wrap(`    On Error GoTo Fail
    Dim parts() As String : parts = Split(CStr(param), "|", 3)
    Dim ol As Object : Set ol = CreateObject("Outlook.Application")
    Dim mail As Object : Set mail = ol.CreateItem(0)
    mail.To = Trim$(parts(0))
    If UBound(parts) >= 1 Then mail.Subject = parts(1)
    If UBound(parts) >= 2 Then mail.Body = parts(2)
    targetWb.Save : mail.Attachments.Add targetWb.FullName
    mail.Send
    ${WS} : ws.Range("A1").Value = "Gonderildi" : ws.Range("B1").Value = parts(0)
    GoTo Done
Fail: ws.Range("A1").Value = "Hata" : ws.Range("B1").Value = Err.Description
Done: Set DynamicFunc = Nothing`),

  SendEmailSmtp: wrap(`    Dim parts() As String : parts = Split(CStr(param), "|")
    Dim cfg As String : cfg = "http://schemas.microsoft.com/cdo/configuration/"
    Dim msg As Object : Set msg = CreateObject("CDO.Message")
    With msg.Configuration.Fields
        .Item(cfg & "sendusing") = 2
        .Item(cfg & "smtpserver") = parts(0)
        .Item(cfg & "smtpserverport") = 25
        .Update
    End With
    msg.To = IIf(UBound(parts) >= 1, parts(1), "")
    msg.Subject = IIf(UBound(parts) >= 2, parts(2), "Teklif")
    msg.TextBody = IIf(UBound(parts) >= 3, parts(3), "")
    On Error Resume Next : msg.Send
    ${WS} : ws.Range("A1").Value = "SMTP" : ws.Range("B1").Value = Err.Number = 0
    Set DynamicFunc = Nothing`),

  ReadInboxEmails: wrap(`    On Error GoTo Fail
    Dim n As Long : n = CLng(Val(CStr(param))) : If n < 1 Then n = 10
    Dim ol As Object : Set ol = CreateObject("Outlook.Application")
    Dim ns As Object : Set ns = ol.GetNamespace("MAPI")
    Dim inbox As Object : Set inbox = ns.GetDefaultFolder(6)
    ${WS} : ws.Cells.ClearContents
    ws.Range("A1").Value = "Konu" : ws.Range("B1").Value = "Gonderen"
    Dim it As Object, r As Long : r = 2, i As Long : i = 0
    For Each it In inbox.Items
        i = i + 1 : If i > n Then Exit For
        ws.Cells(r, 1).Value = it.Subject : ws.Cells(r, 2).Value = it.SenderEmailAddress : r = r + 1
    Next
    GoTo Done
Fail: ws.Range("A1").Value = "Hata" : ws.Range("B1").Value = Err.Description
Done: Set DynamicFunc = Nothing`),

  TextToSpeech: wrap(`    On Error Resume Next
    Dim voice As Object : Set voice = CreateObject("SAPI.SpVoice")
    voice.Speak CStr(param)
    ${WS} : ws.Range("A1").Value = "Okundu" : ws.Range("B1").Value = Left$(CStr(param), 200)
    Set DynamicFunc = Nothing`),

  GenerateRandomToken: wrap(`    Dim i As Long, s As String
    Randomize
    For i = 1 To 32
        s = s & Hex(Int(Rnd() * 16))
    Next i
    ${WS} : ws.Range("A1").Value = "Token" : ws.Range("B1").Value = s
    Set DynamicFunc = s`),

  GetBootTime: wrap(`    On Error Resume Next
    Dim wmi As Object : Set wmi = GetObject("winmgmts:\\\\.\\root\\cimv2")
    Dim col As Object : Set col = wmi.ExecQuery("SELECT LastBootUpTime FROM Win32_OperatingSystem")
    Dim o As Object : For Each o In col
    ${WS} : ws.Range("A1").Value = "Boot" : ws.Range("B1").Value = o.LastBootUpTime : Exit For
    Next
    Set DynamicFunc = Nothing`),

  ExportAllModules: wrap(`    Dim outDir As String : outDir = Trim$(CStr(param))
    If Len(outDir) = 0 Then outDir = Environ("TEMP") & "\\vba-export\\"
    Dim fso As Object : Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(outDir) Then fso.CreateFolder outDir
    Dim comp As Object, n As Long : n = 0
    For Each comp In targetWb.VBProject.VBComponents
        If comp.Type = 1 Then
            n = n + 1
            Dim ts As Object : Set ts = fso.CreateTextFile(outDir & comp.Name & ".bas", True)
            ts.Write comp.CodeModule.Lines(1, comp.CodeModule.CountOfLines) : ts.Close
        End If
    Next comp
    ${WS} : ws.Range("A1").Value = "Modul" : ws.Range("B1").Value = n
    Set DynamicFunc = Nothing`),

  ListAllModulesInWorkbook: wrap(`    ${WS} : ws.Cells.ClearContents
    ws.Range("A1").Value = "Modul" : ws.Range("B1").Value = "Satir"
    Dim comp As Object, r As Long : r = 2
    For Each comp In targetWb.VBProject.VBComponents
        ws.Cells(r, 1).Value = comp.Name
        ws.Cells(r, 2).Value = comp.CodeModule.CountOfLines
        r = r + 1
    Next
    Set DynamicFunc = Nothing`),

  SyncModulesFromServer: runRemote("AutoUpdateModules"),
  LoadPluginFromServer: runRemote("AutoUpdateModules"),
  CallDllFunction: sheetNote("LoadLibrary + GetProcAddress — ornek: TeklifAgent.Com.dll"),
  ReadWriteNamedPipe: sheetNote("Named Pipe — D39/D40 veya TeklifPipeCom"),
  SignPdfWithCertificate: sheetNote("iText / TeklifPdfCom COM gerekir"),
  GeneratePdfReport: wrap(`    Application.Run "ExportAllSheetsAsPdf", targetWb, CStr(param)
    Set DynamicFunc = Nothing`),

  InsertImageFromUrl: wrap(`    Dim url As String : url = Trim$(CStr(param))
    targetWb.ActiveSheet.Shapes.AddPicture url, False, True, 50, 50, 200, 200
    ${WS} : ws.Range("A1").Value = "Resim" : ws.Range("B1").Value = url
    Set DynamicFunc = Nothing`),

  GenerateQrCodeImage: wrap(`    Application.Run "DisplayQrCode", targetWb, param
    Set DynamicFunc = Nothing`),

  RemoteConfigLoader: wrap(`    Dim baseUrl As String
    baseUrl = GetSetting("ilhan", "Settings", "apiBaseUrl", "http://localhost:3000/api/")
    If Right(baseUrl, 1) <> "/" Then baseUrl = baseUrl & "/"
    Dim mac As String : mac = GetSetting("ilhan", "Settings", "mac", "")
    Dim http As Object : Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.Open "GET", baseUrl & "auto-start/" & mac & "/", False : http.send
    SaveSetting "ilhan", "RemoteConfig", "last", http.responseText
    ${WS} : ws.Range("B1").Value = Left$(http.responseText, 32000)
    Set DynamicFunc = Nothing`),
};

function generateCode(name) {
  if (CUSTOM[name]) return CUSTOM[name];

  const n = name.toLowerCase();

  if (n.startsWith("get") && (n.includes("user") || n.includes("domain") || n.includes("group") || n.includes("local")))
    return psRun(`Get-${name.replace(/^Get/, "")} -ErrorAction SilentlyContinue | Select-Object -First 50 | ConvertTo-Json`, name);

  if (n.includes("eventlog") || n.includes("eventlog"))
    return psRun(`Get-EventLog -LogName ${n.includes("security") ? "Security" : n.includes("system") ? "System" : "Application"} -Newest 20 | Format-Table -AutoSize | Out-String`, name);

  if (n.includes("defender") || n.includes("antivirus") || n.includes("applocker"))
    return psRun(`Get-MpComputerStatus | ConvertTo-Json`, name);

  if (n.startsWith("send") && (n.includes("slack") || n.includes("teams") || n.includes("telegram") || n.includes("sms") || n.includes("whatsapp") || n.includes("zapier") || n.includes("make") || n.includes("webhook")))
    return httpPostJson();

  if (n.startsWith("post") || n.startsWith("upload") || n.startsWith("sync") || n.includes("webhook"))
    return httpPostJson();

  if (n.startsWith("get") || n.startsWith("read") || n.startsWith("fetch") || n.startsWith("download"))
    return httpGet();

  if (n.startsWith("disable") || n.startsWith("enable") || n.startsWith("scan") || n.startsWith("encrypt") || n.startsWith("decrypt") || n.startsWith("hash") || n.startsWith("trace") || n.startsWith("check") || n.startsWith("analyze") || n.startsWith("monitor") || n.startsWith("profile") || n.startsWith("generate") && n.includes("diagnostic"))
    return psRun(`# ${name} - param: $env:TEKLIF_PARAM`, name);

  if (n.startsWith("export") || n.startsWith("import") || n.startsWith("print") || n.startsWith("merge") || n.startsWith("convert") || n.startsWith("apply") || n.startsWith("create") || n.startsWith("add") || n.startsWith("auto") || n.startsWith("build") || n.startsWith("clear") || n.startsWith("protect") || n.startsWith("validate") || n.startsWith("set") || n.startsWith("insert") || n.startsWith("remove") || n.startsWith("rename") || n.startsWith("compile") || n.startsWith("backup") || n.startsWith("fill") || n.startsWith("extract"))
    return sheetNote(`${name} — param ile calisir; gerekirse zInternet zincirinden cagirin`);

  if (n.startsWith("is") || n.startsWith("find") || n.startsWith("move") || n.startsWith("simulate") || n.startsWith("enum") || n.includes("window") || n.includes("cursor") || n.includes("clipboard") || n.includes("idle") || n.includes("tick") || n.includes("process") || n.includes("pipe") || n.includes("library") || n.includes("native") || n.includes("dpapi") || n.includes("icmp") || n.includes("tcp") || n.includes("udp") || n.includes("dns") || n.includes("dpi") || n.includes("beep") || n.includes("firmware") || n.includes("wow64"))
    return sheetNote(`${name} — WinAPI Declare; bkz dll-module-proposals.md ornekleri`);

  if (n.includes("oracle") || n.includes("mongo") || n.includes("sqlite") || n.includes("postgres") || n.includes("mysql") || n.includes("sql") || n.includes("database") || n.includes("transactional") || n.includes("stored"))
    return wrap(`    Application.Run "AdoQueryToSheet", targetWb, CStr(param)
    Set DynamicFunc = Nothing`);

  return sheetNote(`${name} modulu — param: detay icin module-proposals.md`);
}

const meta = JSON.parse(fs.readFileSync(META_PATH, "utf8"));

const isMain =
  process.argv[1] &&
  fileURLToPath(import.meta.url) === path.resolve(process.argv[1]);

if (isMain) {
  let written = 0;

  if (process.argv.includes("--export-dll-json")) {
    const dllOnly = {};
    for (const [k, v] of Object.entries(CUSTOM)) {
      if (/Private Declare/i.test(v)) dllOnly[k] = v.trim() + "\n";
    }
    const outJson = path.join(ROOT, "data", "dll-templates.json");
    fs.writeFileSync(outJson, JSON.stringify(dllOnly, null, 2) + "\n", "utf8");
    console.log("DLL sablonlari:", Object.keys(dllOnly).length, "->", outJson);
    process.exit(0);
  }

  for (const name of allPending) {
    const code = generateCode(name);
    const out = code.trim().endsWith("End Function") ? code.trim() + "\n" : code + "\n";
    fs.writeFileSync(path.join(OUT, `${name}.bas`), out, "utf8");
    if (!meta[name]) {
      meta[name] = {
        description: name.replace(/([A-Z])/g, " $1").trim(),
        category: name.match(/^Teklif/) ? "com" : name.match(/^(Get|Set|Find|Move|Post|Play|Time|Query|Start|List|Call)/) ? "dll" : name.startsWith("D") ? "dll" : name.match(/^(Get|Set|Check|Send|Post|Export|Import)/) ? name.replace(/^(Get|Set|Check|Send|Post|Export|Import).*/, "$1").toLowerCase() : "genel",
      };
    }
    written++;
  }

  fs.writeFileSync(META_PATH, JSON.stringify(meta, null, 2) + "\n", "utf8");
  console.log(`Pending: ${allPending.length}, written: ${written}, total in folder: ${fs.readdirSync(OUT).filter((f) => f.endsWith(".bas")).length}`);
}

export { CUSTOM, generateCode };
