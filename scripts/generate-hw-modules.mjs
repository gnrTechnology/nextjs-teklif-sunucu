/**
 * Generates 25 hardware/computer-info VBA .bas files and updates modules.json
 */
import fs from "fs";
import path from "path";
import { fileURLToPath } from "url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const ROOT       = path.join(__dirname, "..");
const SRC_DIR    = path.join(ROOT, "data", "modules-source");
const JSON_PATH  = path.join(ROOT, "data", "modules.json");

/* ── Helper header / footer ─────────────────────────── */
function wmiHeader() {
  return `Private Function GetWMI(query As String) As Object
    Dim objWMI As Object
    Set objWMI = GetObject("winmgmts:\\\\.\root\cimv2")
    Set GetWMI = objWMI.ExecQuery(query)
End Function

Private Sub WriteSheet(targetWb As Workbook, headers() As String, rows As Collection)
    Dim ws As Worksheet
    Set ws = targetWb.Sheets(1)
    ws.Cells.ClearContents
    Dim c As Integer
    For c = 0 To UBound(headers)
        ws.Cells(1, c + 1).Value = headers(c)
        ws.Cells(1, c + 1).Font.Bold = True
    Next c
    Dim r As Long : r = 2
    Dim rowData As Variant
    For Each rowData In rows
        Dim cols() As String
        cols = Split(CStr(rowData), Chr(9))
        For c = 0 To UBound(cols)
            ws.Cells(r, c + 1).Value = cols(c)
        Next c
        r = r + 1
    Next rowData
    ws.Columns.AutoFit
End Sub`;
}

/* ── Module definitions ─────────────────────────────── */
const modules = [

// 1
{ name: "GetComputerName",
  desc: "Bilgisayar adını döndürür",
  category: "donanim",
  code: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim name As String : name = Environ("COMPUTERNAME")
    targetWb.Sheets(1).Range("A1").Value = "Bilgisayar Adı"
    targetWb.Sheets(1).Range("B1").Value = name
    targetWb.Sheets(1).Columns.AutoFit
    MsgBox "Bilgisayar Adı: " & name, vbInformation, "GetComputerName"
    Set DynamicFunc = Nothing
End Function` },

// 2
{ name: "GetWindowsVersion",
  desc: "Windows sürüm + build numarası (WMI)",
  category: "donanim",
  code: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim objWMI As Object, col As Object, obj As Object
    Set objWMI = GetObject("winmgmts:\\\\.\root\cimv2")
    Set col = objWMI.ExecQuery("SELECT Caption, Version, BuildNumber, OSArchitecture FROM Win32_OperatingSystem")
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Cells.ClearContents
    ws.Range("A1:D1").Value = Array("İşletim Sistemi", "Sürüm", "Build", "Mimari")
    ws.Range("A1:D1").Font.Bold = True
    Dim r As Long : r = 2
    For Each obj In col
        ws.Cells(r, 1).Value = obj.Caption
        ws.Cells(r, 2).Value = obj.Version
        ws.Cells(r, 3).Value = obj.BuildNumber
        ws.Cells(r, 4).Value = obj.OSArchitecture
        r = r + 1
    Next obj
    ws.Columns.AutoFit
    Set DynamicFunc = Nothing
End Function` },

// 3
{ name: "GetCpuInfo",
  desc: "CPU adı, çekirdek sayısı, maks frekans (GHz)",
  category: "donanim",
  code: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim objWMI As Object, col As Object, obj As Object
    Set objWMI = GetObject("winmgmts:\\\\.\root\cimv2")
    Set col = objWMI.ExecQuery("SELECT Name, NumberOfCores, NumberOfLogicalProcessors, MaxClockSpeed FROM Win32_Processor")
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Cells.ClearContents
    ws.Range("A1:D1").Value = Array("İşlemci", "Çekirdek", "Mantıksal CPU", "Maks Frekans (GHz)")
    ws.Range("A1:D1").Font.Bold = True
    Dim r As Long : r = 2
    For Each obj In col
        ws.Cells(r, 1).Value = obj.Name
        ws.Cells(r, 2).Value = obj.NumberOfCores
        ws.Cells(r, 3).Value = obj.NumberOfLogicalProcessors
        ws.Cells(r, 4).Value = Format(obj.MaxClockSpeed / 1000, "0.00")
        r = r + 1
    Next obj
    ws.Columns.AutoFit
    Set DynamicFunc = Nothing
End Function` },

// 4
{ name: "GetRamInfo",
  desc: "Toplam / Kullanılan / Boş RAM (GB)",
  category: "donanim",
  code: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim objWMI As Object, col As Object, obj As Object
    Set objWMI = GetObject("winmgmts:\\\\.\root\cimv2")
    Set col = objWMI.ExecQuery("SELECT TotalVisibleMemorySize, FreePhysicalMemory FROM Win32_OperatingSystem")
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Cells.ClearContents
    ws.Range("A1:C1").Value = Array("Toplam RAM (GB)", "Boş RAM (GB)", "Kullanılan RAM (GB)")
    ws.Range("A1:C1").Font.Bold = True
    For Each obj In col
        Dim total As Double : total = obj.TotalVisibleMemorySize / 1048576
        Dim free  As Double : free  = obj.FreePhysicalMemory / 1048576
        ws.Range("A2:C2").Value = Array(Format(total, "0.00"), Format(free, "0.00"), Format(total - free, "0.00"))
    Next obj
    ws.Columns.AutoFit
    Set DynamicFunc = Nothing
End Function` },

// 5
{ name: "GetDiskInfo",
  desc: "Tüm sürücüler: boyut, boş alan, dosya sistemi",
  category: "donanim",
  code: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim objWMI As Object, col As Object, obj As Object
    Set objWMI = GetObject("winmgmts:\\\\.\root\cimv2")
    Set col = objWMI.ExecQuery("SELECT DeviceID, VolumeName, Size, FreeSpace, FileSystem FROM Win32_LogicalDisk WHERE DriveType=3")
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Cells.ClearContents
    ws.Range("A1:E1").Value = Array("Sürücü", "Etiket", "Toplam (GB)", "Boş (GB)", "Dosya Sistemi")
    ws.Range("A1:E1").Font.Bold = True
    Dim r As Long : r = 2
    For Each obj In col
        ws.Cells(r, 1).Value = obj.DeviceID
        ws.Cells(r, 2).Value = obj.VolumeName
        ws.Cells(r, 3).Value = Format(obj.Size / 1073741824, "0.00")
        ws.Cells(r, 4).Value = Format(obj.FreeSpace / 1073741824, "0.00")
        ws.Cells(r, 5).Value = obj.FileSystem
        r = r + 1
    Next obj
    ws.Columns.AutoFit
    Set DynamicFunc = Nothing
End Function` },

// 6
{ name: "GetMacAddress",
  desc: "Aktif NIC'in MAC adresini döndürür",
  category: "donanim",
  code: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim objWMI As Object, col As Object, obj As Object
    Set objWMI = GetObject("winmgmts:\\\\.\root\cimv2")
    Set col = objWMI.ExecQuery("SELECT MACAddress, Description FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled=True")
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Cells.ClearContents
    ws.Range("A1:B1").Value = Array("MAC Adresi", "Adaptör")
    ws.Range("A1:B1").Font.Bold = True
    Dim r As Long : r = 2
    For Each obj In col
        ws.Cells(r, 1).Value = obj.MACAddress
        ws.Cells(r, 2).Value = obj.Description
        r = r + 1
    Next obj
    ws.Columns.AutoFit
    Set DynamicFunc = Nothing
End Function` },

// 7
{ name: "GetIpAddress",
  desc: "IPv4 + IPv6 yerel adresler",
  category: "donanim",
  code: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim objWMI As Object, col As Object, obj As Object
    Set objWMI = GetObject("winmgmts:\\\\.\root\cimv2")
    Set col = objWMI.ExecQuery("SELECT IPAddress, Description FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled=True")
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Cells.ClearContents
    ws.Range("A1:B1").Value = Array("IP Adresi", "Adaptör")
    ws.Range("A1:B1").Font.Bold = True
    Dim r As Long : r = 2
    For Each obj In col
        If Not IsNull(obj.IPAddress) Then
            Dim ip As Variant
            For Each ip In obj.IPAddress
                ws.Cells(r, 1).Value = ip
                ws.Cells(r, 2).Value = obj.Description
                r = r + 1
            Next ip
        End If
    Next obj
    ws.Columns.AutoFit
    Set DynamicFunc = Nothing
End Function` },

// 8
{ name: "GetPublicIp",
  desc: "api.ipify.org üzerinden dış IP",
  category: "donanim",
  code: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim http As Object
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    On Error Resume Next
    http.Open "GET", "https://api.ipify.org", False
    http.setTimeouts 5000, 10000, 15000, 15000
    http.send
    Dim ip As String
    If http.Status = 200 Then
        ip = Trim(http.responseText)
    Else
        ip = "Erişilemedi"
    End If
    On Error GoTo 0
    targetWb.Sheets(1).Range("A1").Value = "Dış IP Adresi"
    targetWb.Sheets(1).Range("B1").Value = ip
    targetWb.Sheets(1).Range("A1").Font.Bold = True
    targetWb.Sheets(1).Columns.AutoFit
    MsgBox "Dış IP: " & ip, vbInformation, "GetPublicIp"
    Set DynamicFunc = Nothing
End Function` },

// 9
{ name: "GetLoggedInUser",
  desc: "Aktif Windows kullanıcısı",
  category: "donanim",
  code: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Cells.ClearContents
    ws.Range("A1:B1").Value = Array("Alan", "Değer")
    ws.Range("A1:B1").Font.Bold = True
    ws.Range("A2:B2").Value = Array("Kullanıcı Adı", Environ("USERNAME"))
    ws.Range("A3:B3").Value = Array("Bilgisayar Adı", Environ("COMPUTERNAME"))
    ws.Range("A4:B4").Value = Array("Kullanıcı Profil Yolu", Environ("USERPROFILE"))
    ws.Range("A5:B5").Value = Array("AppData", Environ("APPDATA"))
    ws.Columns.AutoFit
    Set DynamicFunc = Nothing
End Function` },

// 10
{ name: "GetDomainName",
  desc: "Domain / Workgroup adı",
  category: "donanim",
  code: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim objWMI As Object, col As Object, obj As Object
    Set objWMI = GetObject("winmgmts:\\\\.\root\cimv2")
    Set col = objWMI.ExecQuery("SELECT Name, Domain, Workgroup, PartOfDomain FROM Win32_ComputerSystem")
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Cells.ClearContents
    ws.Range("A1:D1").Value = Array("Bilgisayar Adı", "Domain", "Workgroup", "Domain Üyesi mi?")
    ws.Range("A1:D1").Font.Bold = True
    Dim r As Long : r = 2
    For Each obj In col
        ws.Cells(r, 1).Value = obj.Name
        ws.Cells(r, 2).Value = obj.Domain
        ws.Cells(r, 3).Value = IIf(IsNull(obj.Workgroup), "-", obj.Workgroup)
        ws.Cells(r, 4).Value = IIf(obj.PartOfDomain, "Evet", "Hayır")
        r = r + 1
    Next obj
    ws.Columns.AutoFit
    Set DynamicFunc = Nothing
End Function` },

// 11
{ name: "GetBiosInfo",
  desc: "BIOS sürümü, tarih, üretici",
  category: "donanim",
  code: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim objWMI As Object, col As Object, obj As Object
    Set objWMI = GetObject("winmgmts:\\\\.\root\cimv2")
    Set col = objWMI.ExecQuery("SELECT Manufacturer, Name, Version, ReleaseDate FROM Win32_BIOS")
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Cells.ClearContents
    ws.Range("A1:D1").Value = Array("Üretici", "BIOS Adı", "Sürüm", "Tarih")
    ws.Range("A1:D1").Font.Bold = True
    Dim r As Long : r = 2
    For Each obj In col
        ws.Cells(r, 1).Value = obj.Manufacturer
        ws.Cells(r, 2).Value = obj.Name
        ws.Cells(r, 3).Value = obj.Version
        Dim rawDate As String : rawDate = CStr(obj.ReleaseDate)
        ws.Cells(r, 4).Value = Left(rawDate, 4) & "-" & Mid(rawDate, 5, 2) & "-" & Mid(rawDate, 7, 2)
        r = r + 1
    Next obj
    ws.Columns.AutoFit
    Set DynamicFunc = Nothing
End Function` },

// 12
{ name: "GetMotherboardInfo",
  desc: "Anakart modeli ve seri numarası",
  category: "donanim",
  code: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim objWMI As Object, col As Object, obj As Object
    Set objWMI = GetObject("winmgmts:\\\\.\root\cimv2")
    Set col = objWMI.ExecQuery("SELECT Manufacturer, Product, SerialNumber, Version FROM Win32_BaseBoard")
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Cells.ClearContents
    ws.Range("A1:D1").Value = Array("Üretici", "Model", "Seri No", "Sürüm")
    ws.Range("A1:D1").Font.Bold = True
    Dim r As Long : r = 2
    For Each obj In col
        ws.Cells(r, 1).Value = obj.Manufacturer
        ws.Cells(r, 2).Value = obj.Product
        ws.Cells(r, 3).Value = obj.SerialNumber
        ws.Cells(r, 4).Value = obj.Version
        r = r + 1
    Next obj
    ws.Columns.AutoFit
    Set DynamicFunc = Nothing
End Function` },

// 13
{ name: "GetGpuInfo",
  desc: "Ekran kartı adı, VRAM, sürücü versiyonu",
  category: "donanim",
  code: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim objWMI As Object, col As Object, obj As Object
    Set objWMI = GetObject("winmgmts:\\\\.\root\cimv2")
    Set col = objWMI.ExecQuery("SELECT Name, AdapterRAM, DriverVersion, VideoProcessor FROM Win32_VideoController")
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Cells.ClearContents
    ws.Range("A1:D1").Value = Array("Ekran Kartı", "VRAM (MB)", "Sürücü", "İşlemci")
    ws.Range("A1:D1").Font.Bold = True
    Dim r As Long : r = 2
    For Each obj In col
        ws.Cells(r, 1).Value = obj.Name
        On Error Resume Next
        ws.Cells(r, 2).Value = Format(obj.AdapterRAM / 1048576, "0")
        On Error GoTo 0
        ws.Cells(r, 3).Value = obj.DriverVersion
        ws.Cells(r, 4).Value = obj.VideoProcessor
        r = r + 1
    Next obj
    ws.Columns.AutoFit
    Set DynamicFunc = Nothing
End Function` },

// 14
{ name: "GetScreenResolution",
  desc: "Ekran genişliği × yüksekliği (piksel)",
  category: "donanim",
  code: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim objWMI As Object, col As Object, obj As Object
    Set objWMI = GetObject("winmgmts:\\\\.\root\cimv2")
    Set col = objWMI.ExecQuery("SELECT CurrentHorizontalResolution, CurrentVerticalResolution, Name FROM Win32_VideoController")
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Cells.ClearContents
    ws.Range("A1:C1").Value = Array("Ekran Kartı", "Genişlik (px)", "Yükseklik (px)")
    ws.Range("A1:C1").Font.Bold = True
    Dim r As Long : r = 2
    For Each obj In col
        On Error Resume Next
        If obj.CurrentHorizontalResolution > 0 Then
            ws.Cells(r, 1).Value = obj.Name
            ws.Cells(r, 2).Value = obj.CurrentHorizontalResolution
            ws.Cells(r, 3).Value = obj.CurrentVerticalResolution
            r = r + 1
        End If
        On Error GoTo 0
    Next obj
    ws.Columns.AutoFit
    Set DynamicFunc = Nothing
End Function` },

// 15
{ name: "GetAllNetworkAdapters",
  desc: "Tüm NIC'leri (IP, MAC, durum) sayfaya yazar",
  category: "donanim",
  code: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim objWMI As Object, col As Object, obj As Object
    Set objWMI = GetObject("winmgmts:\\\\.\root\cimv2")
    Set col = objWMI.ExecQuery("SELECT Description, MACAddress, IPAddress, IPEnabled, DHCPEnabled FROM Win32_NetworkAdapterConfiguration")
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Cells.ClearContents
    ws.Range("A1:E1").Value = Array("Adaptör", "MAC", "IP Adresi", "IP Etkin", "DHCP")
    ws.Range("A1:E1").Font.Bold = True
    Dim r As Long : r = 2
    For Each obj In col
        Dim ipStr As String : ipStr = ""
        If Not IsNull(obj.IPAddress) Then
            Dim ip As Variant
            For Each ip In obj.IPAddress
                ipStr = ipStr & ip & " "
            Next ip
        End If
        ws.Cells(r, 1).Value = obj.Description
        ws.Cells(r, 2).Value = IIf(IsNull(obj.MACAddress), "-", obj.MACAddress)
        ws.Cells(r, 3).Value = Trim(ipStr)
        ws.Cells(r, 4).Value = IIf(obj.IPEnabled, "Evet", "Hayır")
        ws.Cells(r, 5).Value = IIf(obj.DHCPEnabled, "Evet", "Hayır")
        r = r + 1
    Next obj
    ws.Columns.AutoFit
    Set DynamicFunc = Nothing
End Function` },

// 16
{ name: "GetSystemUptime",
  desc: "Son açılıştan itibaren gün/saat/dakika",
  category: "donanim",
  code: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim objWMI As Object, col As Object, obj As Object
    Set objWMI = GetObject("winmgmts:\\\\.\root\cimv2")
    Set col = objWMI.ExecQuery("SELECT LastBootUpTime FROM Win32_OperatingSystem")
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Cells.ClearContents
    ws.Range("A1:B1").Value = Array("Alan", "Değer")
    ws.Range("A1:B1").Font.Bold = True
    For Each obj In col
        Dim raw As String : raw = obj.LastBootUpTime
        Dim bootYear  As Integer : bootYear  = CInt(Left(raw, 4))
        Dim bootMonth As Integer : bootMonth = CInt(Mid(raw, 5, 2))
        Dim bootDay   As Integer : bootDay   = CInt(Mid(raw, 7, 2))
        Dim bootHour  As Integer : bootHour  = CInt(Mid(raw, 9, 2))
        Dim bootMin   As Integer : bootMin   = CInt(Mid(raw, 11, 2))
        Dim bootSec   As Integer : bootSec   = CInt(Mid(raw, 13, 2))
        Dim bootDT As Date
        bootDT = DateSerial(bootYear, bootMonth, bootDay) + TimeSerial(bootHour, bootMin, bootSec)
        Dim diff  As Double : diff  = Now - bootDT
        Dim days  As Long   : days  = Int(diff)
        Dim hours As Long   : hours = Int((diff - days) * 24)
        Dim mins  As Long   : mins  = Int(((diff - days) * 24 - hours) * 60)
        ws.Range("A2:B2").Value = Array("Son Açılış", Format(bootDT, "dd.mm.yyyy HH:MM:SS"))
        ws.Range("A3:B3").Value = Array("Çalışma Süresi", days & " gün " & hours & " saat " & mins & " dk")
    Next obj
    ws.Columns.AutoFit
    Set DynamicFunc = Nothing
End Function` },

// 17
{ name: "GetTimeZone",
  desc: "Saat dilimi adı ve UTC offseti",
  category: "donanim",
  code: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim objWMI As Object, col As Object, obj As Object
    Set objWMI = GetObject("winmgmts:\\\\.\root\cimv2")
    Set col = objWMI.ExecQuery("SELECT Caption, Bias, StandardName FROM Win32_TimeZone")
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Cells.ClearContents
    ws.Range("A1:C1").Value = Array("Saat Dilimi", "UTC Farkı (dk)", "Standart Ad")
    ws.Range("A1:C1").Font.Bold = True
    Dim r As Long : r = 2
    For Each obj In col
        ws.Cells(r, 1).Value = obj.Caption
        ws.Cells(r, 2).Value = obj.Bias
        ws.Cells(r, 3).Value = obj.StandardName
        r = r + 1
    Next obj
    ws.Columns.AutoFit
    Set DynamicFunc = Nothing
End Function` },

// 18
{ name: "GetInstalledSoftwareList",
  desc: "Tüm yüklü yazılım listesini sayfaya döker",
  category: "donanim",
  code: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim wsh As Object : Set wsh = CreateObject("WScript.Shell")
    Dim objWMI As Object, col As Object, obj As Object
    Set objWMI = GetObject("winmgmts:\\\\.\root\cimv2")
    Set col = objWMI.ExecQuery("SELECT Name, Version, Vendor, InstallDate FROM Win32_Product")
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Cells.ClearContents
    ws.Range("A1:D1").Value = Array("Yazılım", "Sürüm", "Üretici", "Kurulum Tarihi")
    ws.Range("A1:D1").Font.Bold = True
    Dim r As Long : r = 2
    For Each obj In col
        ws.Cells(r, 1).Value = obj.Name
        ws.Cells(r, 2).Value = obj.Version
        ws.Cells(r, 3).Value = obj.Vendor
        Dim dt As String : dt = CStr(obj.InstallDate)
        If Len(dt) >= 8 Then dt = Left(dt, 4) & "-" & Mid(dt, 5, 2) & "-" & Mid(dt, 7, 2)
        ws.Cells(r, 4).Value = dt
        r = r + 1
    Next obj
    ws.Columns.AutoFit
    MsgBox (r - 2) & " yazılım listelendi.", vbInformation, "GetInstalledSoftwareList"
    Set DynamicFunc = Nothing
End Function` },

// 19
{ name: "GetRunningProcesses",
  desc: "CPU/Bellek kullanımı ile process listesi",
  category: "donanim",
  code: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim objWMI As Object, col As Object, obj As Object
    Set objWMI = GetObject("winmgmts:\\\\.\root\cimv2")
    Set col = objWMI.ExecQuery("SELECT Name, ProcessId, WorkingSetSize, KernelModeTime FROM Win32_Process ORDER BY WorkingSetSize DESC")
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Cells.ClearContents
    ws.Range("A1:D1").Value = Array("Process Adı", "PID", "Bellek (MB)", "CPU Zamanı (s)")
    ws.Range("A1:D1").Font.Bold = True
    Dim r As Long : r = 2
    For Each obj In col
        ws.Cells(r, 1).Value = obj.Name
        ws.Cells(r, 2).Value = obj.ProcessId
        ws.Cells(r, 3).Value = Format(obj.WorkingSetSize / 1048576, "0.00")
        ws.Cells(r, 4).Value = Format(obj.KernelModeTime / 10000000, "0.0")
        r = r + 1
        If r > 102 Then Exit For ' max 100 satır
    Next obj
    ws.Columns.AutoFit
    Set DynamicFunc = Nothing
End Function` },

// 20
{ name: "GetWindowsActivationStatus",
  desc: "Windows aktivasyon durumu (slmgr /xpr çıktısı)",
  category: "donanim",
  code: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim wsh As Object : Set wsh = CreateObject("WScript.Shell")
    Dim exec As Object
    Set exec = wsh.Exec("powershell -NonInteractive -Command ""(Get-CimInstance SoftwareLicensingProduct | Where-Object { $_.PartialProductKey -ne $null -and $_.Name -like '*Windows*' } | Select-Object -First 1).LicenseStatus""")
    Dim out As String
    Do While exec.Status = 0 : Application.Wait Now + TimeValue("00:00:01") : Loop
    out = Trim(exec.StdOut.ReadAll)
    Dim status As String
    Select Case out
        Case "1" : status = "Aktif (Lisanslı)"
        Case "0" : status = "Lisanssız"
        Case "2" : status = "Ek Bilgi Gerekli"
        Case "5" : status = "Bildirim Modu"
        Case Else : status = "Durum: " & out
    End Select
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Range("A1").Value = "Windows Aktivasyon Durumu"
    ws.Range("B1").Value = status
    ws.Range("A1").Font.Bold = True
    ws.Columns.AutoFit
    MsgBox "Aktivasyon: " & status, vbInformation, "GetWindowsActivationStatus"
    Set DynamicFunc = Nothing
End Function` },

// 21
{ name: "GetBatteryInfo",
  desc: "Pil şarj yüzdesi ve kalan süre (laptop)",
  category: "donanim",
  code: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim objWMI As Object, col As Object, obj As Object
    Set objWMI = GetObject("winmgmts:\\\\.\root\cimv2")
    Set col = objWMI.ExecQuery("SELECT Name, BatteryStatus, EstimatedChargeRemaining, EstimatedRunTime FROM Win32_Battery")
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Cells.ClearContents
    ws.Range("A1:D1").Value = Array("Pil Adı", "Durum", "Şarj (%)", "Tahmini Süre (dk)")
    ws.Range("A1:D1").Font.Bold = True
    Dim r As Long : r = 2
    Dim found As Boolean : found = False
    For Each obj In col
        found = True
        Dim bStatus As String
        Select Case obj.BatteryStatus
            Case 1 : bStatus = "Deşarj"
            Case 2 : bStatus = "AC - Şarj Değil"
            Case 3 : bStatus = "Tam Dolu"
            Case 4 : bStatus = "Düşük"
            Case 5 : bStatus = "Kritik"
            Case 6 : bStatus = "Şarj Oluyor"
            Case Else : bStatus = "Bilinmiyor"
        End Select
        ws.Cells(r, 1).Value = obj.Name
        ws.Cells(r, 2).Value = bStatus
        ws.Cells(r, 3).Value = obj.EstimatedChargeRemaining
        ws.Cells(r, 4).Value = IIf(obj.EstimatedRunTime = 71582788, "AC'de", obj.EstimatedRunTime)
        r = r + 1
    Next obj
    If Not found Then
        ws.Range("A2").Value = "Pil bulunamadı (masaüstü bilgisayar)"
    End If
    ws.Columns.AutoFit
    Set DynamicFunc = Nothing
End Function` },

// 22
{ name: "GetPrinterList",
  desc: "Yüklü yazıcılar ve varsayılan yazıcı",
  category: "donanim",
  code: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim objWMI As Object, col As Object, obj As Object
    Set objWMI = GetObject("winmgmts:\\\\.\root\cimv2")
    Set col = objWMI.ExecQuery("SELECT Name, Default, PrinterStatus, DriverName FROM Win32_Printer")
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Cells.ClearContents
    ws.Range("A1:D1").Value = Array("Yazıcı Adı", "Varsayılan", "Durum", "Sürücü")
    ws.Range("A1:D1").Font.Bold = True
    Dim r As Long : r = 2
    For Each obj In col
        ws.Cells(r, 1).Value = obj.Name
        ws.Cells(r, 2).Value = IIf(obj.Default, "Evet", "Hayır")
        ws.Cells(r, 3).Value = obj.PrinterStatus
        ws.Cells(r, 4).Value = obj.DriverName
        r = r + 1
    Next obj
    ws.Columns.AutoFit
    Set DynamicFunc = Nothing
End Function` },

// 23
{ name: "GetUsbDevices",
  desc: "Bağlı USB aygıtları listeler",
  category: "donanim",
  code: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim objWMI As Object, col As Object, obj As Object
    Set objWMI = GetObject("winmgmts:\\\\.\root\cimv2")
    Set col = objWMI.ExecQuery("SELECT Name, Description, Manufacturer, DeviceID FROM Win32_PnPEntity WHERE PNPClass='USB'")
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Cells.ClearContents
    ws.Range("A1:D1").Value = Array("Aygıt Adı", "Açıklama", "Üretici", "Device ID")
    ws.Range("A1:D1").Font.Bold = True
    Dim r As Long : r = 2
    For Each obj In col
        ws.Cells(r, 1).Value = obj.Name
        ws.Cells(r, 2).Value = obj.Description
        ws.Cells(r, 3).Value = IIf(IsNull(obj.Manufacturer), "-", obj.Manufacturer)
        ws.Cells(r, 4).Value = obj.DeviceID
        r = r + 1
    Next obj
    ws.Columns.AutoFit
    Set DynamicFunc = Nothing
End Function` },

// 24
{ name: "GetAudioDevices",
  desc: "Ses aygıtları (varsayılan giriş/çıkış)",
  category: "donanim",
  code: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim objWMI As Object, col As Object, obj As Object
    Set objWMI = GetObject("winmgmts:\\\\.\root\cimv2")
    Set col = objWMI.ExecQuery("SELECT Name, Manufacturer, Status, DeviceID FROM Win32_SoundDevice")
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Cells.ClearContents
    ws.Range("A1:D1").Value = Array("Ses Aygıtı", "Üretici", "Durum", "Device ID")
    ws.Range("A1:D1").Font.Bold = True
    Dim r As Long : r = 2
    For Each obj In col
        ws.Cells(r, 1).Value = obj.Name
        ws.Cells(r, 2).Value = IIf(IsNull(obj.Manufacturer), "-", obj.Manufacturer)
        ws.Cells(r, 3).Value = obj.Status
        ws.Cells(r, 4).Value = obj.DeviceID
        r = r + 1
    Next obj
    ws.Columns.AutoFit
    Set DynamicFunc = Nothing
End Function` },

// 25
{ name: "GetSystemLocale",
  desc: "Sistem dili, klavye düzeni, para birimi biçimi",
  category: "donanim",
  code: `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim wsh As Object : Set wsh = CreateObject("WScript.Shell")
    Dim exec As Object
    Set exec = wsh.Exec("powershell -NonInteractive -Command ""[PSCustomObject]@{Locale=$([System.Globalization.CultureInfo]::CurrentCulture.Name);Language=$([System.Globalization.CultureInfo]::CurrentCulture.DisplayName);Currency=$([System.Globalization.CultureInfo]::CurrentCulture.NumberFormat.CurrencySymbol);DateFormat=$([System.Globalization.CultureInfo]::CurrentCulture.DateTimeFormat.ShortDatePattern);TimeZone=$([System.TimeZoneInfo]::Local.DisplayName)} | ConvertTo-Json""")
    Do While exec.Status = 0 : Application.Wait Now + TimeValue("00:00:01") : Loop
    Dim out As String : out = Trim(exec.StdOut.ReadAll)
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Cells.ClearContents
    ws.Range("A1:B1").Value = Array("Alan", "Değer")
    ws.Range("A1:B1").Font.Bold = True
    ws.Range("A2:B2").Value = Array("Sistem Dili", ExtractJsonValue(out, "Language"))
    ws.Range("A3:B3").Value = Array("Locale Kodu", ExtractJsonValue(out, "Locale"))
    ws.Range("A4:B4").Value = Array("Para Birimi", ExtractJsonValue(out, "Currency"))
    ws.Range("A5:B5").Value = Array("Tarih Formatı", ExtractJsonValue(out, "DateFormat"))
    ws.Range("A6:B6").Value = Array("Saat Dilimi", ExtractJsonValue(out, "TimeZone"))
    ws.Columns.AutoFit
    Set DynamicFunc = Nothing
End Function

Private Function ExtractJsonValue(json As String, key As String) As String
    Dim sk As String : sk = """" & key & """:"
    Dim p1 As Long, p2 As Long
    p1 = InStr(1, json, sk, vbTextCompare)
    If p1 = 0 Then ExtractJsonValue = "" : Exit Function
    p1 = p1 + Len(sk)
    Do While Mid(json, p1, 1) = " " : p1 = p1 + 1 : Loop
    If Mid(json, p1, 1) = """" Then
        p1 = p1 + 1 : p2 = InStr(p1, json, """")
        ExtractJsonValue = Mid(json, p1, p2 - p1)
    Else
        p2 = p1
        Do While p2 <= Len(json)
            If InStr(",}" & Chr(13) & Chr(10), Mid(json, p2, 1)) > 0 Then Exit Do
            p2 = p2 + 1
        Loop
        ExtractJsonValue = Trim(Mid(json, p1, p2 - p1))
    End If
End Function` },

];

/* ── Write .bas files ────────────────────────────────── */
let addedCount   = 0;
let updatedCount = 0;

for (const mod of modules) {
  const filePath = path.join(SRC_DIR, `${mod.name}.bas`);
  fs.writeFileSync(filePath, mod.code, "utf-8");
  console.log(`  ✔  ${mod.name}.bas yazıldı`);
}

/* ── Update modules.json ─────────────────────────────── */
let existing = [];
if (fs.existsSync(JSON_PATH)) {
  existing = JSON.parse(fs.readFileSync(JSON_PATH, "utf-8"));
}

for (const mod of modules) {
  const idx = existing.findIndex(
    (m) => m.methodName.toLowerCase() === mod.name.toLowerCase()
  );
  const entry = {
    methodName: mod.name,
    description: mod.desc,
    category: mod.category,
    active: true,
    code: mod.code,
  };
  if (idx >= 0) {
    existing[idx] = entry;
    updatedCount++;
  } else {
    existing.push(entry);
    addedCount++;
  }
}

fs.writeFileSync(JSON_PATH, JSON.stringify(existing, null, 2), "utf-8");
console.log(`\nToplam ${addedCount} yeni eklendi, ${updatedCount} güncellendi.`);
console.log(`modules.json artık ${existing.length} modül içeriyor.`);
