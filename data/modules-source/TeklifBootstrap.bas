Option Explicit

' Ana teklif dosyasına (xlsm) ekleyin. Sadece eklenti kurulum işlemleri.

Public Sub ApplyPendingAddinUpdate(sourceFile As String, destFile As String, addinName As String, fileName As String)
    Debug.Print "[TeklifBootstrap] ApplyPendingAddinUpdate basladi."
    Debug.Print "[TeklifBootstrap] source: " & sourceFile
    Debug.Print "[TeklifBootstrap] dest: " & destFile
    Call SafeUpdateAddin(sourceFile, destFile, addinName, fileName)
    Debug.Print "[TeklifBootstrap] ApplyPendingAddinUpdate tamamlandi."
End Sub

Public Sub SafeUpdateAddin(sourceFile As String, destFile As String, addinName As String, fileName As String)
    Dim fso As Object
    Dim wb As Workbook
    Dim ai As AddIn
    Dim eklentiListedeVar As Boolean

    Debug.Print "[TeklifBootstrap] SafeUpdateAddin basladi."
    Debug.Print "[TeklifBootstrap] addinName: " & addinName & " | fileName: " & fileName

    Set fso = CreateObject("Scripting.FileSystemObject")

    On Error Resume Next
    Set ai = AddIns(addinName)
    If Err.Number = 0 Then
        eklentiListedeVar = True
        Debug.Print "[TeklifBootstrap] Eklenti listede mevcut."
    Else
        eklentiListedeVar = False
        Err.Clear
        Debug.Print "[TeklifBootstrap] Eklenti listede yok, ilk kurulum."
    End If
    On Error GoTo 0

    If eklentiListedeVar Then
        On Error Resume Next
        If ai.Installed Then
            Debug.Print "[TeklifBootstrap] Eklenti pasif yapiliyor..."
            ai.Installed = False
        End If
        On Error GoTo 0
    End If

    DoEvents

    On Error Resume Next
    Set wb = Workbooks(fileName)
    If Not wb Is Nothing Then
        Debug.Print "[TeklifBootstrap] Acik eklenti kitabi kapatiliyor: " & fileName
        wb.Close SaveChanges:=False
    End If
    On Error GoTo 0

    On Error GoTo KopyalamaHatasi
    Debug.Print "[TeklifBootstrap] Dosya kopyalaniyor..."
    fso.CopyFile sourceFile, destFile, True
    Debug.Print "[TeklifBootstrap] Kopyalama tamam."

    On Error Resume Next
    fso.DeleteFile sourceFile
    Debug.Print "[TeklifBootstrap] Temp dosya silindi."
    On Error GoTo 0

    On Error GoTo AktivasyonHatasi
    Debug.Print "[TeklifBootstrap] Eklenti aktif ediliyor..."

    If eklentiListedeVar Then
        ai.Installed = True
    Else
        Set ai = AddIns.Add(destFile, False)
        ai.Installed = True
    End If

    Debug.Print "[TeklifBootstrap] Eklenti basariyla guncellendi."
    MsgBox "Eklenti başarıyla güncellendi.", vbInformation
    Exit Sub

KopyalamaHatasi:
    Debug.Print "[TeklifBootstrap] Kopyalama hatasi: " & Err.Description
    MsgBox "Dosya kopyalanırken hata oluştu. Dosya kullanımda olabilir.", vbCritical
    Exit Sub

AktivasyonHatasi:
    Debug.Print "[TeklifBootstrap] Aktivasyon hatasi: " & Err.Description
    MsgBox "Aktif etme hatası!" & vbCrLf & _
           "İpucu: İndirilen dosya bozuk olabilir (0 KB veya metin dosyası)." & vbCrLf & _
           "Hata: " & Err.Description, vbExclamation
End Sub

Public Function GetHostWorkbook(Optional preferred As Workbook) As Workbook
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

' RunRemoteCode ve ExecuteDynamicFunction ana dosyada olmalı (eklentide değil).
Public Sub RunRemoteCode(methodName As String)
    Dim http As Object
    Dim rawResponse As String
    Dim cleanVbaCode As String
    Dim jsonBody As String
    Dim hostWb As Workbook
    Dim apiUrl As String

    apiUrl = GetApiBaseUrl() & "module/"
    jsonBody = "{""methodName"":""" & methodName & """}"

    Set hostWb = GetHostWorkbook(ActiveWorkbook)
    If hostWb Is Nothing Then
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

        If .Status = 200 Then
            rawResponse = .responseText
            cleanVbaCode = ExtractCodeFromJSON(rawResponse)

            If Len(cleanVbaCode) > 0 Then
                Call ExecuteDynamicFunction(cleanVbaCode, hostWb, GetApiBaseUrl())
            Else
                MsgBox "Sunucudan kod içeriği boş döndü.", vbExclamation
            End If
        Else
            MsgBox "Sunucu Hatası (" & .Status & "): " & .responseText, vbCritical
        End If
    End With

    Set http = Nothing
    Application.ScreenUpdating = True
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "Bağlantı Hatası: " & Err.Description, vbCritical
    Set http = Nothing
End Sub

Public Function ExecuteDynamicFunction(codeContent As String, targetWb As Workbook, Optional param As Variant) As Object
    Dim tempWb As Workbook
    Dim vbComp As Object
    Dim modName As String
    Dim result As Object
    Dim fullCode As String

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

    fullCode = "Option Explicit" & vbCrLf & vbCrLf & codeContent
    vbComp.CodeModule.AddFromString fullCode

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
        Err.Clear
    End If
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

Private Function GetApiBaseUrl() As String
    On Error Resume Next
    GetApiBaseUrl = zInternet.GET_LICENSE_URL
    If Len(GetApiBaseUrl) = 0 Then GetApiBaseUrl = "http://localhost:3000/api/"
    If Right(GetApiBaseUrl, 1) <> "/" Then GetApiBaseUrl = GetApiBaseUrl & "/"
End Function

