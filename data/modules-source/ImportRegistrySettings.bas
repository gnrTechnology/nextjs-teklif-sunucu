' Uzak modül: zInternet.RunRemoteCode "ImportRegistrySettings"
' teklif.xlam'dan (opsiyonel): ImportAllRegistrySettings
' Not: Option Explicit ExecuteDynamicFunction tarafindan eklenir; burada tekrar yazmayin.

Private Const JSON_FILE As String = "registry-settings-full.json"
Private Const FIRMA_PLACEHOLDER As String = "EPRON"

Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    On Error GoTo ErrHandler
    Debug.Print "[ImportRegistrySettings] DynamicFunc basladi. Kitap: " & targetWb.Name
    EnsureInteractiveUI
    RunImportAllRegistrySettings param
    Set DynamicFunc = Nothing
    Exit Function
ErrHandler:
    MsgBox "ImportRegistrySettings hatasi:" & vbCrLf & Err.Description, vbCritical, "Registry Import"
    Set DynamicFunc = Nothing
End Function

Public Sub ImportAllRegistrySettings()
    RunImportAllRegistrySettings "http://localhost:3000/api/"
End Sub

Private Sub RunImportAllRegistrySettings(Optional apiBaseUrl As Variant)
    Dim baraFiyat As String
    Dim tbVeren As String
    Dim firmaAdi As String
    Dim jsonText As String
    Dim saved As Long
    Dim firmaDefault As String

    EnsureInteractiveUI

    firmaDefault = Trim$(GetSetting("ilhan", "Settings", "mdip", FIRMA_PLACEHOLDER))
    If Len(firmaDefault) = 0 Then firmaDefault = FIRMA_PLACEHOLDER

    baraFiyat = AskInput("Bakir bara birim fiyatini giriniz (TL/kg).", "Registry Import", GetSetting("ilhan", "Settings", "bara", "780"))
    If baraFiyat = "" Then
        MsgBox "Islem iptal edildi.", vbInformation, "Registry Import"
        Exit Sub
    End If

    tbVeren = AskInput("Teklif hazirlayan personel adini giriniz.", "Registry Import", GetSetting("ilhan", "Settings", "TBveren", ""))
    If tbVeren = "" Then
        MsgBox "Islem iptal edildi.", vbInformation, "Registry Import"
        Exit Sub
    End If

    firmaAdi = AskInput("Firma adini giriniz." & vbCrLf & "(JSON icindeki EPRON bu adla degistirilir)", "Registry Import", firmaDefault)
    If firmaAdi = "" Then
        MsgBox "Islem iptal edildi.", vbInformation, "Registry Import"
        Exit Sub
    End If

    jsonText = LoadJsonTemplate(apiBaseUrl)
    If jsonText = "" Then
        MsgBox "registry-settings-full.json bulunamadi ve sablon yuklenemedi.", vbCritical
        Exit Sub
    End If

    saved = ApplyJsonToRegistry(jsonText, baraFiyat, tbVeren, firmaAdi)

    MsgBox saved & " ayar kaydedildi." & vbCrLf & vbCrLf & "bara = " & baraFiyat & vbCrLf & "TBveren = " & tbVeren & vbCrLf & "Firma = " & firmaAdi, vbInformation, "ImportRegistrySettings"
    Debug.Print "[ImportRegistrySettings] Tamamlandi. Kayit: " & saved
End Sub

Private Sub EnsureInteractiveUI()
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Interactive = True
    DoEvents
End Sub

Private Function AskInput(ByVal prompt As String, ByVal title As String, Optional ByVal defaultValue As String = "") As String
    Dim answer As Variant

    EnsureInteractiveUI
    answer = Application.InputBox(prompt, title, defaultValue, Type:=2)

    If VarType(answer) = vbBoolean Then
        AskInput = ""
    Else
        AskInput = Trim$(CStr(answer))
    End If
End Function

Private Function ApplyJsonToRegistry(ByVal jsonText As String, ByVal baraFiyat As String, ByVal tbVeren As String, ByVal firmaAdi As String) As Long
    Dim n As Long

    n = n + SaveAppSection(jsonText, "ilhan", "Settings", baraFiyat, tbVeren, firmaAdi, True)
    n = n + SaveAppSection(jsonText, "sercan", "fileOpenWorkBooks", baraFiyat, tbVeren, firmaAdi, False)
    n = n + SaveAppSection(jsonText, "scngnr", "Settings", baraFiyat, tbVeren, firmaAdi, False)

    ApplyJsonToRegistry = n
End Function

Private Function SaveAppSection(ByVal jsonText As String, ByVal appName As String, ByVal sectionName As String, ByVal baraFiyat As String, ByVal tbVeren As String, ByVal firmaAdi As String, ByVal applyFirma As Boolean) As Long
    Dim sectionJson As String
    Dim pairs As Object
    Dim k As Variant
    Dim v As String
    Dim n As Long

    sectionJson = ExtractSectionJson(jsonText, appName, sectionName)
    If sectionJson = "" Then Exit Function

    Set pairs = ParseFlatStringMap(sectionJson)
    For Each k In pairs.Keys
        v = CStr(pairs(k))

        If applyFirma Then v = ReplaceFirma(v, firmaAdi)

        Select Case LCase$(CStr(k))
            Case "bara"
                v = baraFiyat
            Case "tbveren"
                v = tbVeren
        End Select

        If appName = "sercan" And LCase$(CStr(k)) = "nowopenpropsfile" Then
            If Len(Trim$(v)) = 0 Then GoTo NextKey
        End If

        SaveSetting appName, sectionName, CStr(k), v
        n = n + 1
NextKey:
    Next k

    SaveAppSection = n
End Function

Private Function ReplaceFirma(ByVal value As String, ByVal firmaAdi As String) As String
    If InStr(1, value, FIRMA_PLACEHOLDER, vbTextCompare) > 0 Then
        ReplaceFirma = Replace(value, FIRMA_PLACEHOLDER, firmaAdi, , , vbTextCompare)
    Else
        ReplaceFirma = value
    End If
End Function

Private Function LoadJsonTemplate(Optional apiBaseUrl As Variant) As String
    Dim paths As Variant
    Dim i As Long
    Dim p As String
    Dim root As String
    Dim jsonText As String

    jsonText = DownloadRegistryJson(apiBaseUrl)
    If Len(jsonText) > 0 Then
        LoadJsonTemplate = jsonText
        Exit Function
    End If

    root = Trim$(GetSetting("ilhan", "Settings", "malzemedizini", ""))
    paths = Array(IIf(Len(root) > 0, root & Application.PathSeparator & JSON_FILE, ""), Application.UserLibraryPath & JSON_FILE)

    For i = LBound(paths) To UBound(paths)
        p = CStr(paths(i))
        If Len(p) > 0 Then
            If Len(Dir$(p)) > 0 Then
                LoadJsonTemplate = ReadUtf8File(p)
                If Len(LoadJsonTemplate) > 0 Then Exit Function
            End If
        End If
    Next i

    LoadJsonTemplate = GetEmbeddedJsonTemplate()
End Function

Private Function DownloadRegistryJson(Optional apiBaseUrl As Variant) As String
    Dim http As Object
    Dim url As String

    On Error GoTo Fail
    url = ResolveApiBaseUrl(apiBaseUrl) & "registry-settings"
    Debug.Print "[ImportRegistrySettings] JSON URL: " & url

    Set http = CreateObject("MSXML2.XMLHTTP.6.0")
    http.Open "GET", url, False
    http.send

    If http.Status = 200 And Len(http.responseText) > 0 Then
        DownloadRegistryJson = http.responseText
        Debug.Print "[ImportRegistrySettings] JSON sunucudan indirildi."
    End If

    Set http = Nothing
    Exit Function

Fail:
    Debug.Print "[ImportRegistrySettings] JSON indirme hatasi: " & Err.Description
    DownloadRegistryJson = ""
End Function

Private Function ResolveApiBaseUrl(apiBaseUrl As Variant) As String
    Dim url As String

    If Not IsMissing(apiBaseUrl) Then url = Trim(CStr(apiBaseUrl))
    If Len(url) = 0 Then url = "http://localhost:3000/api/"
    If Right(url, 1) <> "/" Then url = url & "/"
    ResolveApiBaseUrl = url
End Function

Private Function ExtractSectionJson(ByVal jsonText As String, ByVal appName As String, ByVal sectionName As String) As String
    Dim appBlock As String
    Dim p As Long
    Dim q As Long
    Dim token As String

    token = """" & appName & """:"
    p = InStr(1, jsonText, token, vbTextCompare)
    If p = 0 Then Exit Function

    appBlock = Mid$(jsonText, p)
    token = """" & sectionName & """:"
    p = InStr(1, appBlock, token, vbTextCompare)
    If p = 0 Then Exit Function

    p = InStr(p, appBlock, "{")
    If p = 0 Then Exit Function

    q = FindMatchingBrace(appBlock, p)
    If q = 0 Then Exit Function

    ExtractSectionJson = Mid$(appBlock, p + 1, q - p - 1)
End Function

Private Function FindMatchingBrace(ByVal s As String, ByVal openPos As Long) As Long
    Dim i As Long
    Dim depth As Long
    Dim ch As String
    Dim inString As Boolean
    Dim escaped As Boolean

    depth = 0
    inString = False
    escaped = False

    For i = openPos To Len(s)
        ch = Mid$(s, i, 1)

        If inString Then
            If escaped Then
                escaped = False
            ElseIf ch = "\" Then
                escaped = True
            ElseIf ch = """" Then
                inString = False
            End If
        Else
            Select Case ch
                Case """"
                    inString = True
                Case "{"
                    depth = depth + 1
                Case "}"
                    depth = depth - 1
                    If depth = 0 Then
                        FindMatchingBrace = i
                        Exit Function
                    End If
            End Select
        End If
    Next i
End Function

Private Function ParseFlatStringMap(ByVal sectionBody As String) As Object
    Dim dict As Object
    Dim i As Long
    Dim ch As String
    Dim inString As Boolean
    Dim escaped As Boolean
    Dim mode As String
    Dim keyBuf As String
    Dim valBuf As String

    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = 1

    mode = "seekKey"
    inString = False
    escaped = False

    For i = 1 To Len(sectionBody)
        ch = Mid$(sectionBody, i, 1)

        If inString Then
            If escaped Then
                valBuf = valBuf & ch
                escaped = False
            ElseIf ch = "\" Then
                escaped = True
            ElseIf ch = """" Then
                inString = False
                If mode = "inValue" Then
                    dict(keyBuf) = UnescapeJsonString(valBuf)
                    keyBuf = ""
                    valBuf = ""
                    mode = "seekKey"
                End If
            Else
                If mode = "inKey" Then keyBuf = keyBuf & ch
                If mode = "inValue" Then valBuf = valBuf & ch
            End If
        Else
            If ch = """" Then
                inString = True
                If mode = "seekKey" Or mode = "afterComma" Then
                    mode = "inKey"
                    keyBuf = ""
                ElseIf mode = "waitValue" Then
                    mode = "inValue"
                    valBuf = ""
                End If
            ElseIf ch = ":" Then
                If mode = "inKey" Then mode = "waitValue"
            ElseIf ch = "," Then
                mode = "afterComma"
            End If
        End If
    Next i

    Set ParseFlatStringMap = dict
End Function

Private Function UnescapeJsonString(ByVal s As String) As String
    Dim t As String
    t = s
    t = Replace(t, "\n", vbLf)
    t = Replace(t, "\r", vbCr)
    t = Replace(t, "\t", vbTab)
    t = Replace(t, "\\", "\")
    t = Replace(t, "\""", """")
    UnescapeJsonString = t
End Function

Private Function ReadUtf8File(ByVal filePath As String) As String
    Dim stream As Object
    On Error GoTo Fail
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2
    stream.Charset = "utf-8"
    stream.Open
    stream.LoadFromFile filePath
    ReadUtf8File = stream.ReadText
    stream.Close
    Set stream = Nothing
    Exit Function
Fail:
    ReadUtf8File = ""
End Function

Private Function GetEmbeddedJsonTemplate() As String
    GetEmbeddedJsonTemplate = "{""applications"":{""ilhan"":{""Settings"":{""malzemedizini"":""C:\\Belgelerim\\Cemex"",""deposabitdosya"":"""",""depodizini"":"""",""dtxs"":"""",""pfc"":""1"",""teklifdizini"":""C:\\Users\\onurm\\Desktop\\Sunucuya yüklenecekler"",""tbnoek"":""EPR-020524-1043\nEPR-020524-1043\nEPR-\n"",""sonteklif"":"""",""TBveren"":""Onur MEMİŞ"",""bara"":""780"",""baralar"":""3_3_3_3_3_3_3_3_3_3_3_3_3_"",""amb"":""12"",""panodizini"":""LOKAL Pano (Montajlı Panolar) Fiyat Listesi - 2021.xlsb"",""panomarka"":""1"",""misi"":""EPRON İŞÇ."",""msas"":""AKSESUAR"",""mbab"":""SARKUYSAN"",""mama"":""PANO AMB."",""misia"":""Pano Montaj İşçilik Bedeli"",""msasa"":""Pano içi kablolama, klemens, kanal ve aksesuarları"",""mbaba"":""Elektrolitik Bakır Bara"",""mamaa"":""Standart Karayolu Taşımasına uygun Ambalaj Bedeli"",""skdi"":""PM-MP"",""skds"":""PM-MS"",""skdb"":""LAMA"",""skda"":""PM-MA"",""license"":""False"",""Cemex"":""12642"",""mdip"":""EPRON"",""drcp"":""12"",""drmi"":""600"",""PTA1"":""DD"",""PTA2"":""DH"",""PTA3"":""DX"",""PTA4"":""SD"",""PTA5"":""SH"",""PTA6"":""SX"",""PTA7"":""SA"",""PTA8"":""KK"",""PTA9"":""TK"",""PTA10"":""DF"",""PTA11"":""DF"",""PTA12"":""DF"",""ddcp"":""575"",""dhcp"":""750"",""dxcp"":""1375"",""sdcp"":""575"",""shcp"":""750"",""sxcp"":""1375"",""sacp"":""375"",""kkcp"":""625"",""tkcp"":""375"",""ddf2"":""875"",""ddf3"":""1125"",""ddf4"":""1500"",""cpia"":""0"",""cpib"":""120"",""cpic"":""140"",""cpid"":""160"",""cpie"":""180"",""cpif"":""200"",""cpig"":""220"",""cpih"":""240"",""cpsa"":""0"",""cpsb"":""120"",""cpsc"":""140"",""cpsd"":""140"",""cpse"":""160"",""cpsf"":""180"",""cpsg"":""180"",""cpsh"":""200""}},""sercan"":{""fileOpenWorkBooks"":{""nowOpenPropsFile"":""""}},""scngnr"":{""Settings"":{""license"":""false""}}}}"
End Function
