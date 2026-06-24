#If VBA7 Then
Private Declare PtrSafe Function GetTickCount64 Lib "kernel32" () As LongLong
#Else
Private Declare Function GetTickCount Lib "kernel32" () As Long
#End If
Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    #If VBA7 Then
    ws.Range("A1").Value = "Tick64" : ws.Range("B1").Value = GetTickCount64()
    #Else
    ws.Range("A1").Value = "Tick" : ws.Range("B1").Value = GetTickCount()
    #End If
    Set DynamicFunc = Nothing
End Function
