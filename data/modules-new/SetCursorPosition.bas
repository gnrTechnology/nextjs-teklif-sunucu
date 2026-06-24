#If VBA7 Then
Private Declare PtrSafe Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
#Else
Private Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
#End If
Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim parts() As String : parts = Split(CStr(param), ",")
    SetCursorPos CLng(Val(parts(0))), CLng(Val(parts(1)))
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Range("A1").Value = "Cursor" : ws.Range("B1").Value = CStr(param)
    Set DynamicFunc = Nothing
End Function
