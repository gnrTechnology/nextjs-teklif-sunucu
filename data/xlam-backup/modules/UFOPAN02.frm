Private Sub CBPML1_Click()
Call panolistegir1
End Sub
Private Sub CBPML2_Click()
Unload Me
End Sub
Private Sub LBPML1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Call panolistegir1
End Sub
Sub panolistegir1()
If LBPML1.listIndex >= 0 Then
SaveSetting "ilhan", "Settings", "panodizini", LBPML1.List(LBPML1.listIndex)
Call panogir0
Unload Me
End If
End Sub
Private Sub UserForm_Initialize()
ControlTipText = "Hazırlayan: İlhan Şirin"
LBPML1.Clear
Dim f, fm1
fm1 = GetSetting("ilhan", "Settings", "malzemedizini")
f = fm1 & "\Malzeme Listeleri\4\"
dosya = dir(f & "*.xlsb")
Do While dosya <> ""
If dosya Like "*" & "Montajlı Pano" & "*" Then LBPML1.AddItem dosya
   dosya = dir
Loop
End Sub