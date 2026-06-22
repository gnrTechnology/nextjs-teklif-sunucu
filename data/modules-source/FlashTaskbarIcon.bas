Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    ' Windows API FlashWindow ile
    Declare PtrSafe Function FlashWindow Lib "user32" (ByVal hwnd As LongPtr, ByVal bInvert As Long) As Long
    Dim i As Long
    For i = 1 To 6
        FlashWindow Application.hwnd, 1
        Application.Wait Now + TimeValue("00:00:01")
    Next i
    FlashWindow Application.hwnd, 0
    Set DynamicFunc = Nothing
End Function