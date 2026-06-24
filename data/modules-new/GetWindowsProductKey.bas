Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim path As String
    path = "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\DigitalProductId"
    If Not IsMissing(param) And Not IsEmpty(param) Then
        If Len(Trim$(CStr(param))) > 0 Then path = Trim$(CStr(param))
    End If

    Dim wsh As Object
    Set wsh = CreateObject("WScript.Shell")
    Dim raw As Variant
    On Error Resume Next
    raw = wsh.RegRead(path)
  Dim errTxt As String
    If Err.Number <> 0 Then errTxt = "HATA: " & Err.Description
    On Error GoTo 0

    Dim key As String
    key = ""
    If IsArray(raw) Then key = DecodeProductKey(raw)

    Dim ws As Worksheet
    Set ws = targetWb.Sheets(1)
    ws.Cells.ClearContents
    ws.Range("A1").Value = "Registry yolu"
    ws.Range("B1").Value = path
    ws.Range("A2").Value = "Urun anahtari"
    If Len(errTxt) > 0 Then
        ws.Range("B2").Value = errTxt
    ElseIf Len(key) > 0 Then
        ws.Range("B2").Value = key
    Else
        ws.Range("B2").Value = "Anahtar cozulemedi (byte dizisi okundu)"
    End If
    ws.Columns.AutoFit
    Set DynamicFunc = Nothing
End Function

Private Function DecodeProductKey(digits() As Byte) As String
    On Error GoTo Fail
    Dim keyChars As Variant
    keyChars = Array("B", "C", "D", "F", "G", "H", "J", "K", "M", "P", "Q", "R", "T", "V", "W", "X", "Y", "2", "3", "4", "6", "7", "8", "9")
    Dim isWin8 As Boolean
    isWin8 = (digits(66) \ 6) And 1
    Dim start As Long
    start = 52
    Dim key As String
    Dim i As Long, j As Long, cur As Long
    key = ""
    For i = 24 To 0 Step -1
        cur = 0
        For j = 14 To 0 Step -1
            cur = cur * 256 + digits(start + j)
            digits(start + j) = cur \ 24
            cur = cur Mod 24
        Next j
        key = key & keyChars(cur)
    Next i
    If isWin8 Then
        Dim insertPos As Long
        insertPos = 1
        For i = 1 To Len(key)
            If i Mod 6 = 0 And i < Len(key) Then
                DecodeProductKey = DecodeProductKey & Mid$(key, insertPos, 5) & "-"
                insertPos = insertPos + 5
            End If
        Next i
        DecodeProductKey = DecodeProductKey & Mid$(key, insertPos)
    Else
        DecodeProductKey = key
    End If
    Exit Function
Fail:
    DecodeProductKey = ""
End Function
