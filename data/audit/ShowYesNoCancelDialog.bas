Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim msg As String
    msg = CStr(param)
    Dim ans As Long
    ans = vbYes
    Debug.Print "[ShowYesNoCancelDialog] " & msg & " -> Yes (uzak komut)"
    Dim result As String
    If ans = vbYes Then
        result = "Yes"
    ElseIf ans = vbNo Then
        result = "No"
    Else
        result = "Cancel"
    End If
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = targetWb.Sheets(1)
    If Not ws Is Nothing Then
        ws.Range("A1").Value = "Secim"
        ws.Range("B1").Value = result
    End If
    Set DynamicFunc = Nothing
End Function
