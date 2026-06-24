Private Sub CBPC2_Click()
On Error Resume Next
If LBPSA3.ListCount < 1 Then Exit Sub
If CBPC2.Value = True Then
LBPSA3.ListStyle = 1: LBPSA3.MultiSelect = 1
Else
LBPSA3.ListStyle = 0: LBPSA3.MultiSelect = 0
End If
End Sub
Private Sub LBPSA5_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
CommandButton14_Click
End Sub
Private Sub LBPSA5_Click()
On Error Resume Next
If LBPSA5.listIndex < 0 Then Exit Sub
TBMH11.Text = "": TBMH11.Text = LBPSA5.List(LBPSA5.listIndex)
'--
LBLPSA3 = "": LBLPSA3 = "Yoğunluk " & LBPSA5.List(LBPSA5.listIndex, 1)
LBLPSA212 = LBPSA5.List(LBPSA5.listIndex, 3) & " - " & LBPSA5.List(LBPSA5.listIndex, 4)
'--montaj seviyesi
If Len(TBMH11) > 1 Then
LBLPSA2.Height = 92
LBLPSA2.BackColor = &HBB9677
LBLPSA2.Top = 92 - LBLPSA2.Height
Else
ayır = Split(LBPSA5.List(LBPSA5.listIndex, 1), "/")
LBLPSA2.Height = 0: LBLPSA2.Height = 92 * CDbl(ayır(0)) / CDbl(ayır(1))
LBLPSA2.BackColor = &HD5D5FF
LBLPSA2.Top = 92 - LBLPSA2.Height
If LBLPSA2.Height > 99 Then LBLPSA2.BackColor = &HAAAAFF: GoTo git1
If LBLPSA2.Height > 66 Then LBLPSA2.BackColor = &HBFBFFF: GoTo git1
If LBLPSA2.Height > 33 Then LBLPSA2.BackColor = &HD5D5FF: GoTo git1
End If
git1:
End Sub
Private Sub UserForm_Initialize()
On Error Resume Next
Call ara1
If UFOPAN00.LBPSA1.ListCount > 0 Then
TBMH1.Text = UFOPAN00.TBMH11.Text
LBLPSA311 = TBMH1.Text
LBLPSA31 = UFOPAN00.LBLPSA3
LBLPSA21.Height = (LBLPSA21.Height / 132) * UFOPAN00.LBLPSA2.Height
LBLPSA21.Top = Frame23.Height - LBLPSA21.Height
LBLPSA21.BackColor = &HD5D5FF
LBLPSA211 = UFOPAN00.LBMSC2.List(UFOPAN00.LBMSC2.listIndex, 1) & " - " & UFOPAN00.LBMSC2.List(UFOPAN00.LBMSC2.listIndex, 2)
For n = 0 To UFOPAN00.LBMSC1.ListCount - 1
  LBPSA5.AddItem UFOPAN00.LBMSC1.List(n, 0)

  UFOPAN00.LBPSA1.Object = "": UFOPAN00.LBPSA1.Object = LBPSA5.List(n, 0)
  If UFOPAN00.LBPSA1.listIndex < 0 Then LBPSA5.List(LBPSA5.ListCount - 1, 2) = "TANIMSIZ": GoTo atla1
  LBPSA5.List(LBPSA5.ListCount - 1, 2) = UFOPAN00.LBPSA1.List(UFOPAN00.LBPSA1.listIndex, 2)
  LBPSA5.List(LBPSA5.ListCount - 1, 1) = UFOPAN00.LBPSA1.List(UFOPAN00.LBPSA1.listIndex, 1)
atla1:
  LBPSA5.List(LBPSA5.ListCount - 1, 3) = UFOPAN00.LBMSC1.List(n, 1)
  LBPSA5.List(LBPSA5.ListCount - 1, 4) = UFOPAN00.LBMSC1.List(n, 2)
Next
End If
UFOPAN00.CBPSC1.BackColor = &HD7BBA2
End Sub
Private Sub CommandButton13_Click()
Unload Me
End Sub
Private Sub CommandButton14_Click()
If CBPC2.Value = True Then
listedeara
Else
listetüm
End If
If LBPSA3.ListCount = 0 Then
UFOPAN00.LBPST1 = ""
On Error GoTo hata
UFOPAN00.LBMSC2.Object = TBMH1
Unload Me
Exit Sub
hata:
UFOPAN00.LBMSC2.Object = TBMH11
Unload Me
End If
End Sub
Sub listetüm()
If LBPSA4.ListCount > 0 And LBPSA3.ListCount > 0 Then
If LBPSA5.listIndex < 0 Then Exit Sub
If TBMH1 = LBPSA5.List(LBPSA5.listIndex) Then Exit Sub
harf1 = TBMH1: harf2 = LBPSA5.List(LBPSA5.listIndex, 0)
For n = 0 To LBPSA4.ListCount - 1
Y1 = CDbl(LBPSA4.List(n, 0))
 Range("I" & Y1) = Replace(Range("I" & Y1), harf1, harf2)
Next
LBPSA3.Clear
Call UFOPAN00.MontajCarpankontrol
End If
End Sub
Sub listedeara()
If LBPSA4.ListCount > 0 And LBPSA3.ListCount > 0 Then
If LBPSA5.listIndex < 0 Then Exit Sub
If TBMH1 = LBPSA5.List(LBPSA5.listIndex) Then Exit Sub
harf1 = TBMH1: harf2 = LBPSA5.List(LBPSA5.listIndex, 0)
For i = LBPSA3.ListCount - 1 To 0 Step -1
If LBPSA3.Selected(i) = True Then
  tip = LBPSA3.List(i): marka = LBPSA3.List(i, 2)
  For n = 0 To LBPSA4.ListCount - 1
   Y1 = CDbl(LBPSA4.List(n, 0))
   If Replace(Range("A" & Y1), "-auto", "") = tip And Range("D" & Y1) = marka Then
   Range("I" & Y1) = Replace(Range("I" & Y1), harf1, harf2)
   End If
  Next n
  LBPSA3.RemoveItem (i)
End If
Next i
Call UFOPAN00.MontajCarpankontrol
End If
End Sub
Sub ara1()
TBMH1 = UFOPAN00.TBMH1
LBLPSA1 = UFOPAN00.LBLPSA1
If UFOPAN00.LBMSC2.listIndex < 0 Then Exit Sub
If UFOPAN00.CBPSC1.BackColor = &HCAE3BF Then UFOPAN01.Caption = UFOPAN00.LBPST1 & ". Satır Pano Harf Değiştirme"
Call panotipleri1
End Sub
Sub panotipleri1()
On Error Resume Next
Dim ls As Integer
If UFOPAN00.CBPSC1.BackColor = &HCAE3BF Then dizi1 = UFOPAN00.LBPST1 & "." Else dizi1 = UFOPAN00.TBPST1
ayır1 = Split(dizi1, ".")
tsay = Len(dizi1) - Len(Replace(dizi1, ".", ""))
For n = 1 To tsay '
s = ayır1(n - 1)
       LBPSA3.AddItem Replace(Range("A" & s), "-auto", "")
       LBPSA3.List(LBPSA3.ListCount - 1, 2) = Range("D" & s)
       LBPSA3.List(LBPSA3.ListCount - 1, 3) = Range("I" & s)
       LBPSA4.AddItem ayır1(n - 1)
Next n
For n = 0 To LBPSA3.ListCount - 1
  tip = LBPSA3.List(n): marka = LBPSA3.List(n, 2)
  For i = LBPSA3.ListCount - 1 To n + 1 Step -1
   If LBPSA3.List(i, 0) = tip And LBPSA3.List(i, 2) = marka Then
   LBPSA3.RemoveItem (i)
   End If
  Next i
Next n
For n = 0 To LBPSA3.ListCount - 1
  For i = 0 To UFOPAN00.ListBoxKT1.ListCount - 1
   LBPSA3.List(n, 1) = LBPSA3.List(n, 2)
   If UFOPAN00.ListBoxKT1.List(i, 1) = LBPSA3.List(n) And UFOPAN00.ListBoxKT1.List(i) = LBPSA3.List(n, 2) Then
   LBPSA3.List(n, 1) = UFOPAN00.ListBoxKT1.List(i, 3)
   Exit For
   End If
  Next i
Next n
dizi1 = ""
End Sub