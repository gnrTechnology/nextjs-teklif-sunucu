#If VBA7 Then
Private Declare PtrSafe Function GetWindowRect Lib "user32" (ByVal hwnd As LongPtr, rc As RECT) As Long
Private Type RECT
    Left As Long: Top As Long: Right As Long: Bottom As Long
End Type
#Else
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, rc As RECT) As Long
Private Type RECT
    Left As Long: Top As Long: Right As Long: Bottom As Long
End Type
#End If
Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim rc As RECT
    GetWindowRect Application.hwnd, rc
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Range("A1").Value = "Left" : ws.Range("B1").Value = rc.Left
    ws.Range("A2").Value = "Top" : ws.Range("B2").Value = rc.Top
    ws.Range("A3").Value = "Width" : ws.Range("B3").Value = rc.Right - rc.Left
    ws.Range("A4").Value = "Height" : ws.Range("B4").Value = rc.Bottom - rc.Top
    Set DynamicFunc = Nothing
End Function
