#If VBA7 Then
Private Declare PtrSafe Function MoveWindow Lib "user32" (ByVal hwnd As LongPtr, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
#Else
Private Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
#End If
Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim parts() As String : parts = Split(CStr(param), ",")
    Dim x As Long : x = CLng(Val(parts(0)))
    Dim y As Long : y = IIf(UBound(parts) >= 1, CLng(Val(parts(1))), 0)
    Dim w As Long : w = IIf(UBound(parts) >= 2, CLng(Val(parts(2))), 800)
    Dim h As Long : h = IIf(UBound(parts) >= 3, CLng(Val(parts(3))), 600)
    MoveWindow Application.hwnd, x, y, w, h, 1
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Range("A1").Value = "Moved" : ws.Range("B1").Value = x & "," & y
    Set DynamicFunc = Nothing
End Function
