#If VBA7 Then
Private Declare PtrSafe Function GetForegroundWindow Lib "user32" () As LongPtr
Private Declare PtrSafe Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As LongPtr, ByVal lpString As String, ByVal cch As Long) As Long
#Else
Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
#End If
Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim buf As String : buf = String$(256, 0)
    GetWindowText GetForegroundWindow(), buf, 255
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Range("A1").Value = "Baslik" : ws.Range("B1").Value = Trim$(buf)
    Set DynamicFunc = Nothing
End Function
