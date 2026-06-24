#If VBA7 Then
Private Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long
Private Declare PtrSafe Function GetLastInputInfo Lib "user32" (plii As LASTINPUTINFO) As Long
Private Type LASTINPUTINFO
    cbSize As Long: dwTime As Long
End Type
#Else
Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Function GetLastInputInfo Lib "user32" (plii As LASTINPUTINFO) As Long
Private Type LASTINPUTINFO
    cbSize As Long: dwTime As Long
End Type
#End If
Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim lii As LASTINPUTINFO : lii.cbSize = Len(lii)
    GetLastInputInfo lii
    Dim idleMs As Long : idleMs = GetTickCount() - lii.dwTime
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Range("A1").Value = "IdleMs" : ws.Range("B1").Value = idleMs
    Set DynamicFunc = Nothing
End Function
