#If VBA7 Then
Private Declare PtrSafe Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As Long
#Else
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
#End If
Private Const WM_CLOSE As Long = &H10
Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim hwnd As LongPtr : hwnd = CLng(Val(CStr(param)))
    If hwnd = 0 Then hwnd = Application.hwnd
    PostMessage hwnd, WM_CLOSE, 0, 0
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Range("A1").Value = "WM_CLOSE" : ws.Range("B1").Value = hwnd
    Set DynamicFunc = Nothing
End Function
