Public r() As New EVN
Public say As Integer
Sub isimler1()
Dim im As control
    say = 0
    For Each im In UFOPAN11.MultiPage1.Pages(0).Controls
        If TypeName(im) = "Image" Then
            say = say + 1
            ReDim Preserve r(1 To say)
            Set r(say).Resimler = im
        End If
    Next im
End Sub
Sub renkeo()
On Error Resume Next
For i = 1 To say
UFOPAN11.MultiPage1.Pages(0).Controls("EGOM" & i).BackColor = &H8000000F
UFOPAN11.MultiPage1.Pages(0).Controls("EGOD" & i).BackColor = &H8000000F
Next i
End Sub
Sub Alt_Klasör_İsimlerixxx()
 Dim ds, f, f1, fc, s
 Set ds = CreateObject("Scripting.FileSystemObject")
 Set f = ds.GetFolder("C:\Belgelerim\Cemex\Şablonlar")
 Set fc = f.SubFolders
 For Each f1 In fc
 s = s & f1.Name
 s = s & vbCrLf
 Next
 MsgBox s
 End Sub