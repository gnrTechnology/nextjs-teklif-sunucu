Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Application.Run "CheckDebuggerAttached", targetWb, param
    Set DynamicFunc = Nothing
End Function
