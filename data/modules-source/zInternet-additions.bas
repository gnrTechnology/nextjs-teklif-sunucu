' teklif.xlam içindeki zInternet modülüne ekleyin (veya mevcut kodu bununla değiştirin).
' Excel açılışı: ThisWorkbook → Auto_Open → RunRemoteCode "AutoStartOnExcelOpen"

Public Const GET_LICENSE_URL As String = "http://localhost:3000/api/"

Public Sub RunRemoteCode(methodName As String)
    Dim http As Object
    Dim rawResponse As String
    Dim cleanVbaCode As String
    Dim jsonBody As String
    Dim hostWb As Workbook
    Dim apiUrl As String

    Debug.Print "[zInternet] RunRemoteCode basladi. methodName: " & methodName

    apiUrl = GET_LICENSE_URL
    If Right(apiUrl, 1) <> "/" Then apiUrl = apiUrl & "/"
    apiUrl = apiUrl & "module/"

    jsonBody = "{""methodName"":""" & methodName & """}"
    Debug.Print "[zInternet] API URL: " & apiUrl

    Set hostWb = GetHostWorkbook(ActiveWorkbook)
    If hostWb Is Nothing Then
        Debug.Print "[zInternet] Ana dosya bulunamadi."
        MsgBox "Ana teklif dosyası bulunamadı.", vbCritical
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Set http = CreateObject("MSXML2.XMLHTTP.6.0")

    On Error GoTo ErrHandler
    With http
        .Open "POST", apiUrl, False
        .setRequestHeader "Content-Type", "application/json;charset=UTF-8"
        .send jsonBody

        Debug.Print "[zInternet] HTTP Status: " & .Status

        If .Status = 200 Then
            rawResponse = .responseText
            cleanVbaCode = ExtractCodeFromJSON(rawResponse)
            Debug.Print "[zInternet] Kod uzunlugu: " & Len(cleanVbaCode)

            If Len(cleanVbaCode) > 0 Then
                Call ExecuteDynamicFunction(cleanVbaCode, hostWb, GET_LICENSE_URL)
            Else
                MsgBox "Sunucudan kod içeriği boş döndü.", vbExclamation
            End If
        Else
            MsgBox "Sunucu Hatası (" & .Status & "): " & .responseText, vbCritical
        End If
    End With

    Set http = Nothing
    Application.ScreenUpdating = True
    Debug.Print "[zInternet] RunRemoteCode tamamlandi."
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    Debug.Print "[zInternet] Baglanti hatasi: " & Err.Description
    MsgBox "Bağlantı Hatası: " & Err.Description, vbCritical
    Set http = Nothing
End Sub

Public Function ExecuteDynamicFunction(codeContent As String, targetWb As Workbook, Optional param As Variant) As Object
    Dim tempWb As Workbook
    Dim vbComp As Object
    Dim modName As String
    Dim result As Object
    Dim fullCode As String

    Debug.Print "[zInternet] ExecuteDynamicFunction basladi. targetWb: " & targetWb.Name

    If IsMissing(param) Then param = ""

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    Set tempWb = Workbooks.Add

    On Error Resume Next
    tempWb.Windows(1).Visible = False
    On Error GoTo Cleanup

    Set vbComp = tempWb.VBProject.VBComponents.Add(1)
    modName = "TempMod"
    vbComp.Name = modName

    fullCode = PrepareModuleCode(codeContent)
    vbComp.CodeModule.AddFromString fullCode

    Debug.Print "[zInternet] DynamicFunc cagriliyor..."
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Interactive = True
    On Error GoTo Cleanup
    Set result = Application.Run("'" & tempWb.Name & "'!" & modName & ".DynamicFunc", targetWb, param)
    Set ExecuteDynamicFunction = result

Cleanup:
    If Not tempWb Is Nothing Then
        tempWb.Close SaveChanges:=False
        Set tempWb = Nothing
    End If

    Application.EnableEvents = True
    Application.ScreenUpdating = True

    If Err.Number <> 0 Then
        Debug.Print "[zInternet] ExecuteDynamicFunction hata: " & Err.Description
        MsgBox "Uzak modul hatasi:" & vbCrLf & Err.Description, vbCritical, "RunRemoteCode"
        Err.Clear
    End If
End Function

Private Function PrepareModuleCode(codeContent As String) As String
    Dim s As String

    s = codeContent
    Do While Len(s) > 0
        If Left$(s, 2) = vbCrLf Then
            s = Mid$(s, 3)
        ElseIf Left$(s, 1) = vbCr Or Left$(s, 1) = vbLf Then
            s = Mid$(s, 2)
        Else
            Exit Do
        End If
    Loop

    If StrComp(Left$(s, 14), "Option Explicit", vbTextCompare) = 0 Then
        s = Trim$(Mid$(s, 15))
    End If

    PrepareModuleCode = "Option Explicit" & vbCrLf & vbCrLf & s
End Function

Public Function ExtractCodeFromJSON(jsonText As String) As String
    Dim p1 As Long, p2 As Long
    Dim tempStr As String

    p1 = InStr(1, jsonText, """code""", vbTextCompare)
    If p1 = 0 Then
        ExtractCodeFromJSON = jsonText
        Exit Function
    End If

    p1 = InStr(p1, jsonText, ":")
    p1 = InStr(p1, jsonText, """") + 1
    p2 = InStrRev(jsonText, """")

    If p2 > p1 Then
        tempStr = Mid(jsonText, p1, p2 - p1)
        tempStr = Replace(tempStr, "\""", """")
        tempStr = Replace(tempStr, "\r\n", vbCrLf)
        tempStr = Replace(tempStr, "\n", vbCrLf)
        tempStr = Replace(tempStr, "\t", vbTab)
        tempStr = Replace(tempStr, "\\", "\")
        ExtractCodeFromJSON = tempStr
    Else
        ExtractCodeFromJSON = ""
    End If
End Function

Private Function GetHostWorkbook(Optional preferred As Workbook) As Workbook
    Dim wb As Workbook

    If Not preferred Is Nothing Then
        If Not preferred.IsAddin Then
            Set GetHostWorkbook = preferred
            Exit Function
        End If
    End If

    For Each wb In Application.Workbooks
        If Not wb.IsAddin Then
            Set GetHostWorkbook = wb
            Exit Function
        End If
    Next wb
End Function

' teklif.xlam → ThisWorkbook modülüne ekleyin:
' Public Sub Auto_Open()
'     zInternet.RunRemoteCode "AutoStartOnExcelOpen"
' End Sub
