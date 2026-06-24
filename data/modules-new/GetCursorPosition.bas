#If VBA7 Then
Private Declare PtrSafe Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Type POINTAPI
    x As Long: y As Long
End Type
#Else
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Type POINTAPI
    x As Long: y As Long
End Type
#End If
Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim pt As POINTAPI : GetCursorPos pt
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Range("A1").Value = "X" : ws.Range("B1").Value = pt.x
    ws.Range("A2").Value = "Y" : ws.Range("B2").Value = pt.y
    Set DynamicFunc = Nothing
End Function
