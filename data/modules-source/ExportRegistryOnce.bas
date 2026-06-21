Attribute VB_Name = "ExportRegistryOnce"
Option Explicit

' TEK SEFERLIK — Keşfedilen tüm GetSetting değerlerini kök klasöre JSON yazar.
' Kök = GetSetting("ilhan","Settings","malzemedizini")  (ör. C:\Belgelerim\Cemex)
'
' Kullanım (Immediate veya makro):
'   ExportRegistryOnceToRoot
'
' Çıktı:
'   {malzemedizini}\registry-settings-full.json
'
' İş bitince bu modülü silebilirsiniz.

Private Const REG_ROOT As String = "Software\VB and VBA Program Settings\"
Private Const HKCU As Long = &H80000001
Private Const OUTPUT_FILE As String = "registry-settings-full.json"

Public Sub ExportRegistryOnceToRoot()
    Dim rootFolder As String
    Dim outputPath As String
    Dim jsonText As String

    rootFolder = ResolveRootFolder()
    outputPath = rootFolder & Application.PathSeparator & OUTPUT_FILE

    jsonText = BuildFullJson(rootFolder)
    WriteUtf8File outputPath, jsonText

    MsgBox "Tüm registry değerleri kaydedildi:" & vbCrLf & vbCrLf & outputPath, vbInformation, "ExportRegistryOnce"
End Sub

Private Function ResolveRootFolder() As String
    Dim p As String

    p = Trim$(GetSetting("ilhan", "Settings", "malzemedizini", ""))
    If Len(p) > 0 Then
        ResolveRootFolder = p
        Exit Function
    End If

    On Error Resume Next
    p = ActiveWorkbook.Path
    If Len(p) = 0 Then p = ThisWorkbook.Path
    On Error GoTo 0

    If Len(p) = 0 Then
        p = CreateObject("WScript.Shell").SpecialFolders("Desktop")
    End If

    ResolveRootFolder = p
End Function

Private Function BuildFullJson(ByVal rootFolder As String) As String
    Dim apps As Variant
    Dim parts As Variant
    Dim appName As String
    Dim sectionName As String
    Dim i As Long
    Dim body As String

    body = "{"
    body = body & """exportedAt"":""" & JsonEscape(Format$(Now, "yyyy-mm-dd hh:nn:ss")) & ""","
    body = body & """rootFolder"":""" & JsonEscape(rootFolder) & ""","
    body = body & """sourceWorkbook"":""" & JsonEscape(GetSourceWorkbookLabel()) & ""","
    body = body & """registryRoot"":""HKCU\\Software\\VB and VBA Program Settings"","
    body = body & """keyCount"":" & CountAllKeys() & ","
    body = body & """applications"":{"

    apps = KnownApplications()
    For i = LBound(apps) To UBound(apps)
        parts = Split(CStr(apps(i)), "|")
        appName = parts(0)
        sectionName = parts(1)
        If i > LBound(apps) Then body = body & ","
        body = body & """" & JsonEscape(appName) & """:"
        body = body & BuildSectionJson(appName, sectionName)
    Next i

    body = body & "}}"
    BuildFullJson = body
End Function

Private Function BuildSectionJson(ByVal appName As String, ByVal sectionName As String) As String
    Dim keys As Object
    Dim keyName As Variant
    Dim first As Boolean
    Dim result As String
    Dim val As String

    Set keys = CollectAllKeys(appName, sectionName)

    result = "{""" & JsonEscape(sectionName) & """:{"
    first = True

    For Each keyName In keys.Keys
        val = ReadRegistryValue(appName, sectionName, CStr(keyName))
        If Not first Then result = result & ","
        first = False
        result = result & """" & JsonEscape(CStr(keyName)) & """:""" & JsonEscape(val) & """"
    Next keyName

    result = result & "}}"
    BuildSectionJson = result
End Function

Private Function CountAllKeys() As Long
    Dim apps As Variant
    Dim parts As Variant
    Dim i As Long
    Dim n As Long
    Dim keys As Object

    apps = KnownApplications()
    For i = LBound(apps) To UBound(apps)
        parts = Split(CStr(apps(i)), "|")
        Set keys = CollectAllKeys(parts(0), parts(1))
        n = n + keys.Count
    Next i

    CountAllKeys = n
End Function

Private Function KnownApplications() As Variant
    KnownApplications = Array( _
        "ilhan|Settings", _
        "sercan|fileOpenWorkBooks", _
        "scngnr|Settings" _
    )
End Function

Private Function KnownKeys(ByVal appName As String, ByVal sectionName As String) As Variant
    Select Case appName & "|" & sectionName
        Case "ilhan|Settings"
            KnownKeys = Array( _
                "malzemedizini", "deposabitdosya", "depodizini", "dtxs", "pfc", _
                "teklifdizini", "tbnoek", "sonteklif", "TBveren", _
                "bara", "baralar", "amb", _
                "panodizini", "panomarka", _
                "misi", "msas", "mbab", "mama", _
                "misia", "msasa", "mbaba", "mamaa", _
                "skdi", "skds", "skdb", "skda", _
                "license", _
                "Cemex", "mdip", "drcp", "drmi", _
                "PTA1", "PTA2", "PTA3", "PTA4", "PTA5", "PTA6", _
                "PTA7", "PTA8", "PTA9", "PTA10", "PTA11", "PTA12", _
                "ddcp", "dhcp", "dxcp", "sdcp", "shcp", "sxcp", "sacp", "kkcp", "tkcp", _
                "ddf2", "ddf3", "ddf4", _
                "cpia", "cpib", "cpic", "cpid", "cpie", "cpif", "cpig", "cpih", _
                "cpsa", "cpsb", "cpsc", "cpsd", "cpse", "cpsf", "cpsg", "cpsh" _
            )
        Case "sercan|fileOpenWorkBooks"
            KnownKeys = Array("nowOpenPropsFile")
        Case "scngnr|Settings"
            KnownKeys = Array("license")
        Case Else
            KnownKeys = Array()
    End Select
End Function

Private Function CollectAllKeys(ByVal appName As String, ByVal sectionName As String) As Object
    Dim dict As Object
    Dim documented As Variant
    Dim scanned As Object
    Dim i As Long
    Dim k As Variant

    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = 1

    documented = KnownKeys(appName, sectionName)
    For i = LBound(documented) To UBound(documented)
        If Not dict.Exists(CStr(documented(i))) Then dict.Add CStr(documented(i)), True
    Next i

    Set scanned = EnumRegistryValueNames(appName, sectionName)
    For Each k In scanned
        If Not dict.Exists(CStr(k)) Then dict.Add CStr(k), True
    Next k

    Set CollectAllKeys = dict
End Function

Private Function EnumRegistryValueNames(ByVal appName As String, ByVal sectionName As String) As Object
    Dim dict As Object
    Dim oReg As Object
    Dim subKey As String
    Dim names As Variant
    Dim types As Variant
    Dim i As Long
    Dim rc As Long

    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = 1
    subKey = REG_ROOT & appName & "\" & sectionName

    On Error Resume Next
    Set oReg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
    If oReg Is Nothing Then
        Set EnumRegistryValueNames = dict
        Exit Function
    End If

    rc = oReg.EnumValues(HKCU, subKey, names, types)
    If rc = 0 And IsArray(names) Then
        For i = LBound(names) To UBound(names)
            If Not dict.Exists(CStr(names(i))) Then dict.Add CStr(names(i)), True
        Next i
    End If
    On Error GoTo 0

    Set EnumRegistryValueNames = dict
End Function

Private Function ReadRegistryValue(ByVal appName As String, ByVal sectionName As String, ByVal keyName As String) As String
    Dim wsh As Object
    Dim regPath As String
    Dim v As Variant
    Dim missing As String

    missing = Chr$(1) & "MISSING" & Chr$(1)

    On Error Resume Next
    v = GetSetting(appName:=appName, Section:=sectionName, key:=keyName, Default:=missing)
    If Err.Number = 0 And CStr(v) <> missing Then
        ReadRegistryValue = CStr(v)
        Exit Function
    End If
    Err.Clear

    Set wsh = CreateObject("WScript.Shell")
    regPath = "HKCU\" & REG_ROOT & appName & "\" & sectionName & "\" & keyName
    v = wsh.RegRead(regPath)
    If Err.Number = 0 Then
        ReadRegistryValue = CStr(v)
    Else
        ReadRegistryValue = ""
    End If
    On Error GoTo 0
End Function

Private Function GetSourceWorkbookLabel() As String
    On Error Resume Next
    If Len(ThisWorkbook.Path) > 0 Then
        GetSourceWorkbookLabel = ThisWorkbook.FullName
    ElseIf Len(ActiveWorkbook.Path) > 0 Then
        GetSourceWorkbookLabel = ActiveWorkbook.FullName
    Else
        GetSourceWorkbookLabel = ThisWorkbook.Name
    End If
    On Error GoTo 0
End Function

Private Sub WriteUtf8File(ByVal filePath As String, ByVal content As String)
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2
    stream.Charset = "utf-8"
    stream.Open
    stream.WriteText content
    stream.SaveToFile filePath, 2
    stream.Close
    Set stream = Nothing
End Sub

Private Function JsonEscape(ByVal s As String) As String
    Dim t As String
    t = CStr(s)
    t = Replace(t, "\", "\\")
    t = Replace(t, """", "\""")
    t = Replace(t, vbCrLf, "\n")
    t = Replace(t, vbCr, "\n")
    t = Replace(t, vbLf, "\n")
    t = Replace(t, vbTab, "\t")
    JsonEscape = t
End Function
