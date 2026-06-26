Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim p As String
    p = Trim(CStr(param))
    Dim code As String
    code = ExtractJsonValue(p, "code")
    Dim modName As String
    modName = ExtractJsonValue(p, "moduleName")
    If Len(modName) = 0 Then modName = "InjectedModule"
    If Len(code) = 0 Then
        Debug.Print "[InjectVbaModule] code parametresi bos."
        Set DynamicFunc = Nothing
        Exit Function
    End If

    On Error GoTo Fail
    Dim vbProj As Object
    Set vbProj = targetWb.VBProject
    Dim comp As Object
    For Each comp In vbProj.VBComponents
        If comp.Name = modName Then
            vbProj.VBComponents.Remove comp
            Exit For
        End If
    Next comp

    Dim newComp As Object
    Set newComp = vbProj.VBComponents.Add(1)
    newComp.Name = modName
    newComp.CodeModule.AddFromString code
    Debug.Print "[InjectVbaModule] Modul enjekte edildi: " & modName
    Set DynamicFunc = Nothing
    Exit Function
Fail:
    Debug.Print "[InjectVbaModule] Hata: " & Err.Description
    Set DynamicFunc = Nothing
End Function

Private Function ExtractJsonValue(json As String, key As String) As String
    Dim sk As String
    Dim p1 As Long
    Dim p2 As Long
    sk = """" & key & """:"
    p1 = InStr(1, json, sk, vbTextCompare)
    If p1 = 0 Then Exit Function
    p1 = p1 + Len(sk)
    Do While Mid(json, p1, 1) = " "
        p1 = p1 + 1
    Loop
    If Mid(json, p1, 1) = """" Then
        p1 = p1 + 1
        p2 = InStr(p1, json, """")
        If p2 > p1 Then ExtractJsonValue = Mid(json, p1, p2 - p1)
    End If
End Function
