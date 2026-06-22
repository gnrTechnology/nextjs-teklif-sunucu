Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    ' param: kaç dakikada bir kaydetsin (sayı). 0 = durdur.
    Dim intervalMin As Long
    On Error Resume Next
    intervalMin = CLng(CStr(param))
    On Error GoTo 0

    If intervalMin <= 0 Then
        ' Mevcut zamanlanmış kayıtları iptal et
        On Error Resume Next
        Application.OnTime EarliestTime:=Application.OnTime(Now, "AutoSaveWorkbook_Run", , False), _
            Procedure:="AutoSaveWorkbook_Run", Schedule:=False
        On Error GoTo 0
        MsgBox "Otomatik kayıt durduruldu.", vbInformation
        Set DynamicFunc = Nothing
        Exit Function
    End If

    ' Interval ve hedef dosya adını registry'e kaydet
    SaveSetting "ilhan", "AutoSave", "intervalMin", CStr(intervalMin)
    SaveSetting "ilhan", "AutoSave", "wbName",      targetWb.FullName

    ' İlk zamanlamayı kur
    Application.OnTime Now + TimeSerial(0, intervalMin, 0), "AutoSaveWorkbook_Run"

    MsgBox "Otomatik kayıt aktif: her " & intervalMin & " dakikada bir." & vbCrLf & _
           "Dosya: " & targetWb.Name, vbInformation
    Set DynamicFunc = Nothing
End Function

' Bu Sub teklif.xlam içinde tanımlı olmalıdır.
' Yoksa modül çalışmaz; aşağıdaki kodu teklif.xlam'a kopyalayın.
'
' Public Sub AutoSaveWorkbook_Run()
'     Dim wbPath As String : wbPath = GetSetting("ilhan", "AutoSave", "wbName", "")
'     Dim intMin As Long   : intMin = CLng(GetSetting("ilhan", "AutoSave", "intervalMin", "5"))
'     On Error Resume Next
'     Dim wb As Workbook
'     For Each wb In Application.Workbooks
'         If wb.FullName = wbPath Then
'             ' Sürümlü yedek
'             Dim bakPath As String
'             bakPath = Left(wbPath, InStrRev(wbPath, ".") - 1) & _
'                       "_" & Format(Now, "yyyyMMdd_HHmmss") & _
'                       Mid(wbPath, InStrRev(wbPath, "."))
'             FileCopy wbPath, bakPath
'             wb.Save
'             Exit For
'         End If
'     Next wb
'     On Error GoTo 0
'     If intMin > 0 Then
'         Application.OnTime Now + TimeSerial(0, intMin, 0), "AutoSaveWorkbook_Run"
'     End If
' End Sub
