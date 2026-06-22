Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    ' param: JSON  {"moduleName":"GetCpuInfo","output":{"key":"value",...}}
    ' Modül çıktılarını /api/module-output endpoint'ine gönderir.
    ' Herhangi bir modülün sonunda çağrılabilir.
    Dim p As String : p = Trim(CStr(param))
    Dim baseUrl As String
    baseUrl = GetSetting("ilhan", "Settings", "apiBaseUrl", "https://nextjs-teklif-sunucu.vercel.app")
    If Right(baseUrl, 1) = "/" Then baseUrl = Left(baseUrl, Len(baseUrl) - 1)

    Dim mac      As String : mac      = GetSetting("ilhan", "Settings", "mac", "")
    Dim hostname As String : hostname = Environ("COMPUTERNAME")
    Dim firmaAdi As String : firmaAdi = GetSetting("ilhan", "Settings", "mdip", "")

    If mac = "" Then
        On Error Resume Next
        Dim wmiObj As Object, col As Object, obj As Object
        Set wmiObj = GetObject("winmgmts:\\.\root\cimv2")
        Set col = wmiObj.ExecQuery("SELECT MACAddress FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled=True")
        For Each obj In col
            If obj.MACAddress <> "" Then mac = obj.MACAddress : Exit For
        Next
        On Error GoTo 0
    End If

    If mac = "" Or p = "" Then
        Set DynamicFunc = Nothing
        Exit Function
    End If

    Dim moduleName As String
    Dim outputJson As String

    If Left(p, 1) = "{" Then
        moduleName = EscapeJson(ExtractJsonVal(p, "moduleName"))
        outputJson = ExtractOutputJson(p)
    Else
        moduleName = "Unknown"
        outputJson = "{""text"":""" & EscapeJson(p) & """}"
    End If

    Dim postBody As String
    postBody = "{" & _
        """mac"":""" & EscapeJson(mac) & """," & _
        """moduleName"":""" & moduleName & """," & _
        """hostname"":""" & EscapeJson(hostname) & """," & _
        """firmaAdi"":""" & EscapeJson(firmaAdi) & """," & _
        """output"":" & outputJson & _
        "}"

    On Error Resume Next
    Dim http As Object
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.Open "POST", baseUrl & "/api/module-output", False
    http.setRequestHeader "Content-Type", "application/json"
    http.setTimeouts 5000, 5000, 15000, 15000
    http.send postBody
    On Error GoTo 0

    Set DynamicFunc = Nothing
End Function

' Excel sayfasından verilen alanları JSON objesine dönüştürüp POST eder.
' Kullanım: Call zInternet.RunRemoteCode("PostToServer", "{""moduleName"":""GetCpuInfo"",""output"":{""cpu"":""..."""}}")
' Veya daha basit: aşağıdaki yardımcı Sub kullanılabilir:
Public Sub PostSheetDataToServer(moduleName As String, ws As Worksheet)
    Dim baseUrl As String
    baseUrl = GetSetting("ilhan", "Settings", "apiBaseUrl", "https://nextjs-teklif-sunucu.vercel.app")
    If Right(baseUrl, 1) = "/" Then baseUrl = Left(baseUrl, Len(baseUrl) - 1)

    Dim mac As String : mac = GetSetting("ilhan", "Settings", "mac", "")
    Dim hostname As String : hostname = Environ("COMPUTERNAME")

    ' Sayfanın A:B sütunlarını JSON'a çevir
    Dim json As String : json = "{"
    Dim r As Long : r = 1
    Dim lastRow As Long : lastRow = ws.Cells(ws.Rows.Count, 1).End(-4162).Row ' xlUp
    Dim sep As String : sep = ""
    Do While r <= lastRow
        Dim k As String : k = Trim(CStr(ws.Cells(r, 1).Value))
        Dim v As String : v = Trim(CStr(ws.Cells(r, 2).Value))
        If k <> "" Then
            json = json & sep & """" & EscapeJson(k) & """:""" & EscapeJson(v) & """"
            sep = ","
        End If
        r = r + 1
    Loop
    json = json & "}"

    Dim postBody As String
    postBody = "{""mac"":""" & EscapeJson(mac) & """," & _
               """moduleName"":""" & EscapeJson(moduleName) & """," & _
               """hostname"":""" & EscapeJson(hostname) & """," & _
               """output"":" & json & "}"

    On Error Resume Next
    Dim http As Object
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.Open "POST", baseUrl & "/api/module-output", False
    http.setRequestHeader "Content-Type", "application/json"
    http.setTimeouts 5000, 5000, 15000, 15000
    http.send postBody
    On Error GoTo 0
End Sub

Private Function EscapeJson(s As String) As String
    s = Replace(s, "\", "\\") : s = Replace(s, """", "\""")
    s = Replace(s, Chr(10), "\n") : s = Replace(s, Chr(13), "")
    EscapeJson = s
End Function

Private Function ExtractJsonVal(json As String, key As String) As String
    Dim sk As String : sk = """" & key & """:"
    Dim p1 As Long : p1 = InStr(1, json, sk, vbTextCompare)
    If p1 = 0 Then Exit Function
    p1 = p1 + Len(sk)
    Do While Mid(json, p1, 1) = " " : p1 = p1 + 1 : Loop
    If Mid(json, p1, 1) = """" Then
        p1 = p1 + 1 : Dim p2 As Long : p2 = InStr(p1, json, """")
        If p2 > p1 Then ExtractJsonVal = Mid(json, p1, p2 - p1)
    End If
End Function

Private Function ExtractOutputJson(json As String) As String
    Dim sk As String : sk = """output"":"
    Dim p1 As Long : p1 = InStr(1, json, sk, vbTextCompare)
    If p1 = 0 Then ExtractOutputJson = "{}": Exit Function
    p1 = p1 + Len(sk)
    Do While Mid(json, p1, 1) = " " : p1 = p1 + 1 : Loop
    If Mid(json, p1, 1) <> "{" Then ExtractOutputJson = "{}": Exit Function
    Dim depth As Long : depth = 0 : Dim p2 As Long : p2 = p1
    Do While p2 <= Len(json)
        If Mid(json, p2, 1) = "{" Then depth = depth + 1
        If Mid(json, p2, 1) = "}" Then
            depth = depth - 1
            If depth = 0 Then ExtractOutputJson = Mid(json, p1, p2 - p1 + 1) : Exit Function
        End If
        p2 = p2 + 1
    Loop
    ExtractOutputJson = "{}"
End Function
