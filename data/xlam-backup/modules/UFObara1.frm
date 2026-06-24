Dim s
Dim sd
Public bfyt As String
Private Sub CKNT_Click()
If CKNT.Value = True Then nötrsec: Exit Sub
For X = 1 To 13
Controls("TBNA" & X) = ""
Next
End Sub
Private Sub CKPE_Click()
If CKPE.Value = True Then topraksec: Exit Sub
For X = 1 To 13
Controls("TBTA" & X) = ""
Next
End Sub
Sub CommandButton4_Click()
If MultiPage1.Value = 1 Then CommandButton6_Click: Exit Sub
If TBGT = 0 Or TBGT = "" Then Exit Sub
'If Not ActiveSheet.Name = "Sayfa1" Then MsgBox (" Teklif sayfasına geçiniz "), vbInformation, "scngnr@hotmail.com": Exit Sub
If Selection.row < 2 Then Exit Sub
If CKM1.Value = True Then Call baraliste1: Exit Sub
Call bakir_gir
a = Selection.row
If Range("E" & a).FormulaR1C1 = "1" Then Range("E" & a).FormulaR1C1 = WorksheetFunction.RoundUp(TBGT.Value, -1)
End Sub
Private Sub Image2_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
On Error Resume Next
CreateObject("Shell.Application").Open "C:\Belgelerim\Cemex\Ayarlar\Bakır Hesabı"
End Sub
Private Sub ListBoxBT_Click()
On Error Resume Next
LBB1.Selected(LBB1.listIndex) = False: LBB2.Selected(LBB2.listIndex) = False
CommandButtonB13.Enabled = False: CommandButtonB14.Enabled = True
TextBoxB2.Visible = False: TextBoxB3.Visible = False
TextBoxBL2.Visible = True: TextBoxBL3.Visible = True

TextBoxBy.Value = ListBoxBT.List(ListBoxBT.listIndex, 5)
TextBoxBL3.Value = ListBoxBT.List(ListBoxBT.listIndex, 3)
TextBoxBL2.Value = ListBoxBT.List(ListBoxBT.listIndex, 2)
End Sub
Sub BiP()
On Error Resume Next
Dim stp
stp = 0
For X = 1 To 13
If Controls("TBSK" & X) <> "" Then
Controls("TBSK" & X) = Format(Controls("LB" & X) * Controls("TBSM" & X) * Controls("TBSP" & X) * Controls("TBSA" & X), "#,##0.00")
Controls("TBSM" & s).ControlTipText = Controls("TBSK" & s) & " kg"
stp = Controls("TBSK" & X) + stp
End If
Next
If stp = 0 Then TBSK14 = "" Else TBSK14 = Format(stp, "#,##0.00")
stp = 0
For X = 1 To 13
If Controls("TBAK" & X) <> "" Then
Controls("TBAK" & X) = Format(Controls("LA" & X) * Controls("TBAP" & X) * Controls("TBAA" & X), "#,##0.00")
stp = Controls("TBAK" & X) + stp
End If
Next
If stp = 0 Then TBAK14 = "" Else TBAK14 = Format(stp, "#,##0.00")
stp = 0
For X = 1 To 10
If Controls("TBNK" & X) <> "" Then
Controls("TBNK" & X) = Format(Controls("LA" & X) * Controls("TBNA" & X), "#,##0.00")
stp = Controls("TBNK" & X) + stp
End If
Next
If stp = 0 Then TBNK14 = "" Else TBNK14 = Format(stp, "#,##0.00")
stp = 0
For X = 1 To 8
If Controls("TBTK" & X) <> "" Then
Controls("TBTK" & X) = Format(Controls("LA" & X) * Controls("TBTA" & X), "#,##0.00")
stp = Controls("TBTK" & X) + stp
End If
Next
If stp = 0 Then TBTK14 = "" Else TBTK14 = Format(stp, "#,##0.00")
End Sub
Private Sub SBSBM_SpinUp()
On Error GoTo hata
For X = 1 To 13
Controls("TBSM" & X).Value = Format(Controls("TBSM" & X) + 0.1, "#,##0.0")
Next
hata:
End Sub
Private Sub SBSBM_SpinDown()
On Error GoTo hata
For X = 1 To 13
If Controls("TBSM" & X).Value <= 0 Then Exit Sub
Controls("TBSM" & X).Value = Format(Controls("TBSM" & X) - 0.1, "#,##0.0")
Next
hata:
End Sub
Private Sub SBtolr_Change() 'akım tolerans
On Error Resume Next
TBtolr.Text = SBtolr.Value
Call topkg
End Sub
Sub salter()
On Error Resume Next
If IsNumeric(Controls("TBSA" & s)) Then
pf = Controls("TBSP" & s) * Controls("LB" & s)
If Controls("TBSM" & s).BackColor <> &HC0C0FF Then Controls("TBSM" & s).BackColor = &HC0C0FF
If Controls("TBSA" & s).BackColor <> &HC0C0FF Then Controls("TBSA" & s).BackColor = &HC0C0FF
If Controls("TBSP" & s).BackColor <> &HC0C0FF Then Controls("TBSP" & s).BackColor = &HC0C0FF
If Controls("LBA" & s).BackColor <> &H926A36 Then Controls("LBA" & s).BackColor = &H926A36
If Controls("TBSK" & s).BackColor <> &H926A36 Then Controls("TBSK" & s).BackColor = &HC0C0FF
If s > 2 And Controls("TBSP" & s) > 3 Then
If Left(Controls("TBK" & s), 2) = "3x" Then pf = (Controls("LB" & s) * 2 / 3) + (Controls("LB" & s) * 3) Else pf = Controls("LB" & s) * 3.5
Else
End If
Controls("TBSK" & s) = Format(Controls("TBSM" & s) * pf * Controls("TBSA" & s), "#,##0.00")
Controls("TBSM" & s).ControlTipText = Controls("TBSK" & s) & " kg"
Else
Controls("TBSA" & s) = "": Controls("TBSK" & s) = "": Controls("TBSM" & s).ControlTipText = ""
Controls("TBSA" & s).BackColor = &HFFFFFF: Controls("TBSP" & s).BackColor = &HFFFFFF
Controls("TBSM" & s).BackColor = &HE0E0E0: Controls("LBA" & s).BackColor = &H86B57D
Controls("TBSK" & s).BackColor = &HB1B4B8
End If
End Sub
Sub saltertpkg()
On Error Resume Next
Dim stp
stp = 0
For X = 1 To 13
If Controls("TBSK" & X) <> "" Then stp = Controls("TBSK" & X) + stp
Next
If stp = 0 Then TBSK14 = "" Else TBSK14 = Format(stp, "#,##0.00")
End Sub
Private Sub TBGT_Change()
If TBGT <> "" And TextBoxbara <> "" Then TextBoxbaratp = Format(TBGT * TextBoxbara, "#,##0.00")
End Sub
Private Sub TextBoxbara_Change()
On Error Resume Next
If TextBoxBtkg <> "" And TextBoxbara <> "" Then TextBoxbaratp = Format(TBGT * TextBoxbara, "#,##0.00")
End Sub
Private Sub TBSM1_Change()
If TBSA1 = "" Then Exit Sub Else s = 1: Call salter: Call saltertpkg
End Sub
Private Sub TBSM2_Change()
s = 2: If TBSA2 = "" Then Exit Sub: Call salter: Call saltertpkg
End Sub
Private Sub TBSM3_Change()
s = 3: If TBSA3 = "" Then Exit Sub Else Call salter: Call saltertpkg
End Sub
Private Sub TBSM4_Change()
s = 4: If TBSA4 = "" Then Exit Sub Else Call salter: Call saltertpkg
End Sub
Private Sub TBSM5_Change()
s = 5: If TBSA5 = "" Then Exit Sub Else Call salter: Call saltertpkg
End Sub
Private Sub TBSM6_Change()
s = 6: If TBSA6 = "" Then Exit Sub Else Call salter: Call saltertpkg
End Sub
Private Sub TBSM7_Change()
s = 7: If TBSA7 = "" Then Exit Sub Else Call salter: Call saltertpkg
End Sub
Private Sub TBSM8_Change()
s = 8: If TBSA8 = "" Then Exit Sub Else Call salter: Call saltertpkg
End Sub
Private Sub TBSM9_Change()
s = 9: If TBSA9 = "" Then Exit Sub Else Call salter: Call saltertpkg
End Sub
Private Sub TBSM10_Change()
s = 10: If TBSA10 = "" Then Exit Sub Else Call salter: Call saltertpkg
End Sub
Private Sub TBSM11_Change()
s = 11: If TBSA11 = "" Then Exit Sub Else Call salter: Call saltertpkg
End Sub
Private Sub TBSM12_Change()
s = 12: If TBSA12 = "" Then Exit Sub Else Call salter: Call saltertpkg
End Sub
Private Sub TBSM13_Change()
s = 13: If TBSA13 = "" Then Exit Sub Else Call salter: Call saltertpkg
End Sub
Private Sub TBSP1_Change()
s = 1: Call salter: Call saltertpkg
End Sub
Private Sub TBSP2_Change()
s = 2: Call salter: Call saltertpkg
End Sub
Private Sub TBSP3_Change()
s = 3: Call salter: Call saltertpkg
End Sub
Private Sub TBSP4_Change()
s = 4: Call salter: Call saltertpkg
End Sub
Private Sub TBSP5_Change()
s = 5: Call salter: Call saltertpkg
End Sub
Private Sub TBSP6_Change()
s = 6: Call salter: Call saltertpkg
End Sub
Private Sub TBSP7_Change()
s = 7: Call salter: Call saltertpkg
End Sub
Private Sub TBSP8_Change()
s = 8: Call salter: Call saltertpkg
End Sub
Private Sub TBSP9_Change()
s = 9: Call salter: Call saltertpkg
End Sub
Private Sub TBSP10_Change()
s = 10: Call salter: Call saltertpkg
End Sub
Private Sub TBSP11_Change()
s = 11: Call salter: Call saltertpkg
End Sub
Private Sub TBSP12_Change()
s = 12: Call salter: Call saltertpkg
End Sub
Private Sub TBSP13_Change()
s = 13: Call salter: Call saltertpkg
End Sub
Private Sub TBSA1_Change()
s = 1: Call salter: Call saltertpkg
End Sub
Private Sub TBSA2_Change()
s = 2: Call salter: Call saltertpkg
End Sub
Private Sub TBSA3_Change()
s = 3: Call salter: Call saltertpkg
End Sub
Private Sub TBSA4_Change()
s = 4: Call salter: Call saltertpkg
End Sub
Private Sub TBSA5_Change()
s = 5: Call salter: Call saltertpkg
End Sub
Private Sub TBSA6_Change()
s = 6: Call salter: Call saltertpkg
End Sub
Private Sub TBSA7_Change()
s = 7: Call salter: Call saltertpkg
End Sub
Private Sub TBSA8_Change()
s = 8: Call salter: Call saltertpkg
End Sub
Private Sub TBSA9_Change()
s = 9: Call salter: Call saltertpkg
End Sub
Private Sub TBSA10_Change()
s = 10: Call salter: Call saltertpkg
End Sub
Private Sub TBSA11_Change()
s = 11: Call salter: Call saltertpkg
End Sub
Private Sub TBSA12_Change()
s = 12: Call salter: Call saltertpkg
End Sub
Private Sub TBSA13_Change()
s = 13: Call salter: Call saltertpkg
End Sub
Sub anabara()
On Error Resume Next
If IsNumeric(Controls("TBAA" & s)) Then
Akgb = Controls("TAK" & s).ControlTipText
Controls("TBAK" & s) = Format(Akgb * Controls("TBAP" & s) * Controls("TBAA" & s), "#,##0.00")
Controls("TAK" & s).BackColor = &HC0C0FF: Controls("TBAK" & s).BackColor = &HC0C0FF
Controls("TBAA" & s).BackColor = &HC0C0FF: Controls("TBAP" & s).BackColor = &HC0C0FF
Else
Controls("TBAA" & s) = "": Controls("TBAK" & s) = ""
Controls("TAK" & s).BackColor = &HAEDFF7: Controls("TBAK" & s).BackColor = &HAEDFF7
Controls("TBAA" & s).BackColor = &HFFFFFF: Controls("TBAP" & s).BackColor = &HFFFFFF
End If
End Sub
Sub anabaratpkg()
On Error Resume Next
Dim stp
stp = 0
For X = 1 To 13
If Controls("TBAK" & X) <> "" Then stp = Controls("TBAK" & X) + stp
Next
If stp = 0 Then TBAK14 = "" Else TBAK14 = Format(stp, "#,##0.00")
If CKNT.Value = True Then Controls("TBNA" & s) = Controls("TBAA" & s)
If CKPE.Value = True Then Controls("TBTA" & s) = Controls("TBAA" & s)
'If CKPE.Value = True Then Controls("TBTA" & s) = Controls("TBAA" & s)
End Sub
Private Sub TBAP1_Change()
s = 1: Call anabara: Call anabaratpkg
End Sub
Private Sub TBAP2_Change()
s = 2: Call anabara: Call anabaratpkg
End Sub
Private Sub TBAP3_Change()
s = 3: Call anabara: Call anabaratpkg
End Sub
Private Sub TBAP4_Change()
s = 4: Call anabara: Call anabaratpkg
End Sub
Private Sub TBAP5_Change()
s = 5: Call anabara: Call anabaratpkg
End Sub
Private Sub TBAP6_Change()
s = 6: Call anabara: Call anabaratpkg
End Sub
Private Sub TBAP7_Change()
s = 7: Call anabara: Call anabaratpkg
End Sub
Private Sub TBAP8_Change()
s = 8: Call anabara: Call anabaratpkg
End Sub
Private Sub TBAP9_Change()
s = 9: Call anabara: Call anabaratpkg
End Sub
Private Sub TBAP10_Change()
s = 10: Call anabara: Call anabaratpkg
End Sub
Private Sub TBAP11_Change()
s = 11: Call anabara: Call anabaratpkg
End Sub
Private Sub TBAP12_Change()
s = 12: Call anabara: Call anabaratpkg
End Sub
Private Sub TBAP13_Change()
s = 13: Call anabara: Call anabaratpkg
End Sub
Private Sub TBAA1_Change()
s = 1: Call anabara: Call anabaratpkg
End Sub
Private Sub TBAA2_Change()
s = 2: Call anabara: Call anabaratpkg
End Sub
Private Sub TBAA3_Change()
s = 3: Call anabara: Call anabaratpkg
End Sub
Private Sub TBAA4_Change()
s = 4: Call anabara: Call anabaratpkg
End Sub
Private Sub TBAA5_Change()
s = 5: Call anabara: Call anabaratpkg
End Sub
Private Sub TBAA6_Change()
s = 6: Call anabara: Call anabaratpkg
End Sub
Private Sub TBAA7_Change()
s = 7: Call anabara: Call anabaratpkg
End Sub
Private Sub TBAA8_Change()
s = 8: Call anabara: Call anabaratpkg
End Sub
Private Sub TBAA9_Change()
s = 9: Call anabara: Call anabaratpkg
End Sub
Private Sub TBAA10_Change()
s = 10: Call anabara: Call anabaratpkg
End Sub
Private Sub TBAA11_Change()
s = 11: Call anabara: Call anabaratpkg
End Sub
Private Sub TBAA12_Change()
s = 12: Call anabara: Call anabaratpkg
End Sub
Private Sub TBAA13_Change()
s = 13: Call anabara: Call anabaratpkg
End Sub
Sub Nötr()
On Error Resume Next
If IsNumeric(Controls("TBNA" & s)) Then
Nkgb = Controls("TNK" & s).ControlTipText
Dim no As Double
no = SBNORAN / 100
Controls("TBNK" & s) = Format(no * Nkgb * Controls("TBNA" & s), "#,##0.00")
Controls("TNK" & s).BackColor = &HC0C0FF: Controls("TBNK" & s).BackColor = &HC0C0FF
Controls("TBNA" & s).BackColor = &HC0C0FF
Else
Controls("TBNA" & s) = "": Controls("TBNK" & s) = ""
Controls("TNK" & s).BackColor = &HF0D89F: Controls("TBNK" & s).BackColor = &HF0D89F
Controls("TBNA" & s).BackColor = &HFFFFFF
End If
End Sub
Sub nötrtpkg()
On Error Resume Next
Dim no As Double
Dim stp
stp = 0
For X = 1 To 13
If Controls("TBNK" & X) <> "" Then
Nkgb = Controls("TNK" & X).ControlTipText
no = SBNORAN / 100
Controls("TBNK" & X) = Format(no * Nkgb * Controls("TBNA" & X), "#,##0.00")
stp = (Controls("TBNK" & X)) + stp
End If
Next
If stp = 0 Then TBNK14 = "" Else TBNK14 = Format(stp, "#,##0.00")
End Sub
Private Sub TBNA1_Change()
s = 1: Call Nötr: Call nötrtpkg
End Sub
Private Sub TBNA2_Change()
s = 2: Call Nötr: Call nötrtpkg
End Sub
Private Sub TBNA3_Change()
s = 3: Call Nötr: Call nötrtpkg
End Sub
Private Sub TBNA4_Change()
s = 4: Call Nötr: Call nötrtpkg
End Sub
Private Sub TBNA5_Change()
s = 5: Call Nötr: Call nötrtpkg
End Sub
Private Sub TBNA6_Change()
s = 6: Call Nötr: Call nötrtpkg
End Sub
Private Sub TBNA7_Change()
s = 7: Call Nötr: Call nötrtpkg
End Sub
Private Sub TBNA8_Change()
s = 8: Call Nötr: Call nötrtpkg
End Sub
Private Sub TBNA9_Change()
s = 9: Call Nötr: Call nötrtpkg
End Sub
Private Sub TBNA10_Change()
s = 10: Call Nötr: Call nötrtpkg
End Sub
Private Sub TBNA11_Change()
s = 11: Call Nötr: Call nötrtpkg
End Sub
Private Sub TBNA12_Change()
s = 12: Call Nötr: Call nötrtpkg
End Sub
Private Sub TBNA13_Change()
s = 13: Call Nötr: Call nötrtpkg
End Sub
Sub PE()
On Error Resume Next
If IsNumeric(Controls("TBTA" & s)) Then
Pkgb = Controls("TTK" & s).ControlTipText
Dim no As Double
no = SBTORAN / 100
Controls("TBTK" & s) = Format(no * Pkgb * Controls("TBTA" & s), "#,##0.00")
Controls("TTK" & s).BackColor = &HC0C0FF: Controls("TBTK" & s).BackColor = &HC0C0FF
Controls("TBTA" & s).BackColor = &HC0C0FF
Else
Controls("TBTA" & s) = "": Controls("TBTK" & s) = ""
Controls("TTK" & s).BackColor = &HCAE3BF: Controls("TBTK" & s).BackColor = &HCAE3BF
Controls("TBTA" & s).BackColor = &HFFFFFF
End If
End Sub
Sub petpkg()
On Error Resume Next
Dim no As Double
Dim stp
stp = 0
For X = 1 To 13
If Controls("TBTK" & X) <> "" Then
Pkgb = Controls("TTK" & X).ControlTipText
no = SBTORAN / 100
Controls("TBTK" & X) = Format(no * Pkgb * Controls("TBTA" & X), "#,##0.00")
stp = (Controls("TBTK" & X)) + stp
End If
Next
If stp = 0 Then TBTK14 = "" Else TBTK14 = Format(stp, "#,##0.00")
End Sub
Private Sub TBTA1_Change()
s = 1: Call PE: Call petpkg
End Sub
Private Sub TBTA2_Change()
s = 2: Call PE: Call petpkg
End Sub
Private Sub TBTA3_Change()
s = 3: Call PE: Call petpkg
End Sub
Private Sub TBTA4_Change()
s = 4: Call PE: Call petpkg
End Sub
Private Sub TBTA5_Change()
s = 5: Call PE: Call petpkg
End Sub
Private Sub TBTA6_Change()
s = 6: Call PE: Call petpkg
End Sub
Private Sub TBTA7_Change()
s = 7: Call PE: Call petpkg
End Sub
Private Sub TBTA8_Change()
s = 8: Call PE: Call petpkg
End Sub
Private Sub TBTA9_Change()
s = 9: Call PE: Call petpkg
End Sub
Private Sub TBTA10_Change()
s = 10: Call PE: Call petpkg
End Sub
Private Sub TBTA11_Change()
s = 11: Call PE: Call petpkg
End Sub
Private Sub TBTA12_Change()
s = 12: Call PE: Call petpkg
End Sub
Private Sub TBTA13_Change()
s = 13: Call PE: Call petpkg
End Sub
Sub topkg()
On Error Resume Next
Dim SK14, TK14, AK14, NK14 As Integer
If TBSK14 = "" Then SK14 = 0 Else SK14 = TBSK14
If TBTK14 = "" Then TK14 = 0 Else TK14 = TBTK14
If TBAK14 = "" Then AK14 = 0 Else AK14 = TBAK14
If TBNK14 = "" Then NK14 = 0 Else NK14 = TBNK14
Dim tkg
tkg = CDbl(SK14) + CDbl(TK14) + CDbl(AK14) + CDbl(NK14)
If tkg = 0 Then TBGT.Value = "" Else TBGT.Value = Format(tkg + (tkg * TBtolr / 100), "#,##0.00")
End Sub
Private Sub TBSK14_Change()
Call topkg
End Sub
Private Sub TBTK14_Change()
Call topkg
End Sub
Private Sub TBAK14_Change()
Call topkg
End Sub
Private Sub TBNK14_Change()
Call topkg
End Sub
Private Sub SBNORAN_Change()
On Error Resume Next
TBNORAN = SBNORAN.Value
If CKNT.Value = True Then: Call nötrtpkg
End Sub
Private Sub SBTORAN_Change()
On Error Resume Next
TBTORAN = SBTORAN.Value
If CKPE.Value = True Then: Call petpkg
End Sub
Sub nötrsec()
On Error Resume Next
For Y = 1 To 13
If Controls("TBAA" & Y) <> "" And IsNumeric(Controls("TBAA" & Y)) Then
Controls("TBNA" & Y) = Controls("TBAA" & Y)
End If
Next
End Sub
Sub topraksec()
On Error Resume Next
For Y = 1 To 13
If Controls("TBAA" & Y) <> "" And IsNumeric(Controls("TBAA" & Y)) Then
Controls("TBTA" & Y) = Controls("TBAA" & Y)
End If
Next
End Sub
'BÖLÜM-2
Private Sub CBBS2_Click()
MultiPage1.Value = 2
CommandButton4.Visible = False
Call Blist
End Sub
Private Sub Blist() 'form yükleme
On Error Resume Next
If ListBoxB1.ListCount > 1 Then Exit Sub
ListBoxB1.Clear
Dim Rky As String, Dosyam As String
    Dim Ert As Long, satır As Long, i As Long
    Dim ayır As Variant
    Dosyam = "C:\Belgelerim\Cemex\Ayarlar\Bakır Hesabı\Bara Akım Taşıma Kapasiteleri.txt"
    Ert = FreeFile
    Open Dosyam For Input As #Ert
    If Err.Number <> 0 Then
        MsgBox "Bara Akım Taşıma Kapasiteleri.txt" & " Dosyası Bulunamadı !", vbCritical, "Hata !"
        Exit Sub
    End If
    On Error GoTo 0
    satır = 1
    ListBoxB1.Clear
Do While Not EOF(Ert)
        Line Input #Ert, Rky
        ayır = Split(Rky, ";")
        ListBoxB1.AddItem ayır(i)
tsay = Len(Rky) - Len(Replace(Rky, ";", ""))
For n = 1 To tsay '
If UBound(ayır) <> 0 Then ListBoxB1.List(satır - 1, n) = ayır(n)
Next n
     satır = satır + 1
Loop
    kt = 1
Close #Ert
ListBoxB1.Width = 826
'TextBoxBtkg.Value = "0,00"
End Sub
Private Sub CBBS0_Click()
On Error Resume Next
If MultiPage1.Value = 0 Then Exit Sub
MultiPage1.Value = 0
CommandButton4.Visible = True
TextBoxbaratp = Format(TBGT * TextBoxbara, "#,##0.00")
End Sub
Private Sub CBBS1_Click()
On Error Resume Next
MultiPage1.Value = 1
CommandButton4.Visible = True
TextBoxbaratp = Format(TextBoxBtkg * TextBoxbara, "#,##0.00")
If LBamper.ListCount > 1 Then Exit Sub
Call Blist
Call bakır1
End Sub
Private Sub TextBoxBtkg_Change()
CBBS1_Click
End Sub
Sub bakır1() '
LBamper.Clear
With LBamper
.AddItem "100": .AddItem "160": .AddItem "200": .AddItem "250": .AddItem "320": .AddItem "400": .AddItem "500"
.AddItem "600": .AddItem "800": .AddItem "1000": .AddItem "1250": .AddItem "1600": .AddItem "2000": .AddItem "2500"
.AddItem "3000": .AddItem "4000": .AddItem "5000": .AddItem "6300"
End With
End Sub
Private Sub LBamper_Click()
On Error Resume Next
TextBoxB3.Value = Format(0, "#,##0.00"): TextBoxBy.Value = "Boyut"
s = CDbl(LBamper.Text)
n = ListBoxB1.ListCount - 1
LBB1.Clear
For m = 2 To 5
f = s + (s * CDbl(TBtolr1.Value)) / 100 'tolerans
i = 0
Do Until i > n
        If ListBoxB1.List(i, m) = "-" Then GoTo yok1
        a = ListBoxB1.List(i, 0): k = ListBoxB1.List(i, 1): b = CDbl(ListBoxB1.List(i, m))
        If s > b Then GoTo yok1
        If b <= f Then LBB1.AddItem (m - 1) & "x(" & a & ")": LBB1.List(LBB1.ListCount - 1, 1) = b & " A": _
        LBB1.List(LBB1.ListCount - 1, 2) = k: LBB1.List(LBB1.ListCount - 1, 3) = a
yok1:
i = i + 1
Loop
git1:
Next m
'--
n = ListBoxB1.ListCount - 1
LBB2.Clear
For m = 6 To 9
i = 0
Do Until i > n
        If ListBoxB1.List(i, m) = "-" Then GoTo yok2
        a = ListBoxB1.List(i, 0): k = ListBoxB1.List(i, 1): b = CDbl(ListBoxB1.List(i, m))
        If s > b Then GoTo yok2
        If b <= f Then LBB2.AddItem (m - 5) & "x(" & a & ")": LBB2.List(LBB2.ListCount - 1, 1) = b & " A": _
        LBB2.List(LBB2.ListCount - 1, 2) = k: LBB2.List(LBB2.ListCount - 1, 3) = a
yok2:
i = i + 1
Loop
git2:
Next m
Call sırala1
Call sırala2
CommandButtonB13.Enabled = False: CommandButtonB14.Enabled = False
End Sub
Private Sub SBtolr1_Change() 'akım tolerans
On Error Resume Next
TBtolr1.Text = SBtolr1.Value
Call LBamper_Click
End Sub
Private Sub LBB1_Click()
On Error Resume Next
s = CDbl(LBB1.listIndex)
sd = LBB1.List(s, 2)
TextBoxBy.Value = LBB1.List(s, 3)
TextBoxBkg.Value = sd
TextBoxB3.Value = Format(sd * TextBoxB2.Value, "#,##0.00")
LBB2.Selected(LBB2.listIndex) = False: ListBoxBT.Selected(ListBoxBT.listIndex) = False
CommandButtonB13.Enabled = True: CommandButtonB14.Enabled = False
TextBoxBL2.Visible = False: TextBoxBL3.Visible = False
TextBoxB2.Visible = True: TextBoxB3.Visible = True
End Sub
Private Sub LBB2_Click()
On Error Resume Next
s = CDbl(LBB2.listIndex)
sd = LBB2.List(s, 2)
TextBoxBy.Value = LBB2.List(s, 3)
TextBoxBkg.Value = sd
TextBoxB3.Value = Format(sd * TextBoxB2.Value, "#,##0.00")
LBB1.Selected(LBB1.listIndex) = False: ListBoxBT.Selected(ListBoxBT.listIndex) = False
CommandButtonB13.Enabled = True: CommandButtonB14.Enabled = False
TextBoxBL2.Visible = False: TextBoxBL3.Visible = False
TextBoxB2.Visible = True: TextBoxB3.Visible = True
End Sub
Private Sub TextBoxB2_Change()
On Error Resume Next
If Not IsNumeric(TextBoxB2.Value) Then MsgBox " Harf girilmeyecek,Sadece Rakam Giriniz.": TextBoxB2.Value = 1
TextBoxB3.Value = Format(sd * TextBoxB2.Value, "#,##0.00")
End Sub
Private Sub TextBoxBL2_Change()
ListBoxBT.List(ListBoxBT.listIndex, 2) = TextBoxBL2.Value
ListBoxBT.List(ListBoxBT.listIndex, 3) = Format(TextBoxBL2.Value * ListBoxBT.List(ListBoxBT.listIndex, 1), "#,##0.00")
TextBoxBL3.Value = ListBoxBT.List(ListBoxBT.listIndex, 3)
For i = 1 To ListBoxBT.ListCount
tkg = CDbl(ListBoxBT.List(i - 1, 3)) + tkg
Next i
TextBoxBtkg.Value = Format(tkg, "#,##0.00")
End Sub
Private Sub CommandButtonB13_Click()
On Error Resume Next
If TextBoxB3 = 0 Then Exit Sub
a = ListBoxBT.ListCount
With ListBoxBT
.AddItem TextBoxBy & "mm. Elektrolitik Bakır Bara": .List(a, 1) = TextBoxBkg: .List(a, 2) = TextBoxB2: .List(a, 3) = TextBoxB3: .List(a, 4) = "kg."
.List(a, 5) = TextBoxBy
End With
For i = 1 To ListBoxBT.ListCount
tkg = CDbl(ListBoxBT.List(i - 1, 3)) + tkg
Next i
TextBoxBtkg.Value = Format(tkg, "#,##0.00")
End Sub
Private Sub CommandButtonB14_Click()
On Error Resume Next
If ListBoxBT.listIndex < 0 Then Exit Sub
ListBoxBT.RemoveItem (ListBoxBT.listIndex)
tkg = 0
For i = 1 To ListBoxBT.ListCount
tkg = CDbl(ListBoxBT.List(i - 1, 3)) + tkg
Next i
TextBoxBtkg.Value = Format(tkg, "#,##0.00")
End Sub
Private Sub CBbakır_Click() '++ yeni bakır fiyat teklife aktarma
On Error Resume Next
Application.Calculation = xlCalculationManual
Application.ScreenUpdating = False
Windows(dt).Activate: Sheets("Sayfa1").Select '
son = Range("B65536").End(xlUp).row + 1
Dim a
a = 0
'--
For i = 2 To son
If Left(Range("A" & i), 5) = "PM-MB" Then
Range("F" & i) = CDbl(TextBoxbara.Value)
a = 1
End If
Next i
Application.Calculation = xlCalculationAutomatic
Application.ScreenUpdating = True
If a = 0 Then MsgBox " Teklife Bakır girilmemiş. ", vbInformation, "scngnr@hotmail.com": Exit Sub
MsgBox " Güncel fiyat teklife aktarıldı. ", vbOKOnly, "scngnr@hotmail.com"
End Sub
Private Sub CommandButton22_Click()
On Error Resume Next
bara = TextBoxbara: SaveSetting "ilhan", "Settings", "bara", bara
dizi1 = ""
For i = 1 To 13
 dizi1 = dizi1 & Controls("TBSM" & i).Text & "_"
Next i
SaveSetting "ilhan", "Settings", "baralar", dizi1
End Sub
Private Sub CommandButtonB15_Click()
On Error Resume Next
ListBoxBT.Clear
TextBoxBtkg.Value = "0,00"
End Sub
Private Sub CommandButton6_Click()
If ListBoxBT.ListCount <= 0 Then Exit Sub
If Not ActiveSheet.Name = "Sayfa1" Then MsgBox (" Teklif sayfasına geçiniz "), vbInformation, "scngnr@hotmail.com": Exit Sub
Dim a
a = Selection.row
If a < 2 Then Exit Sub
Call bakir_gir
If Range("E" & a).FormulaR1C1 = "1" Then Range("E" & a).FormulaR1C1 = WorksheetFunction.RoundUp(TextBoxBtkg.Value, 0)
End Sub
Sub sırala1()
If Me.LBB1.ColumnCount <> 0 Then
        LBB1.List = Diz(LBB1.List, 2)
End If
End Sub
Sub sırala2()
If Me.LBB2.ColumnCount <> 0 Then
        LBB2.List = Diz(LBB2.List, 2)
End If
End Sub
Private Function Diz(ByVal Dizim As Variant, Stn As Integer) As Variant
    Dim i, j, k As Long
    Dim Tmp As Variant
    Stn = Stn - 1
    For i = LBound(Dizim, 1) To UBound(Dizim, 1)
        For j = i + 1 To UBound(Dizim, 1)
            If Dizim(i, Stn) > Dizim(j, Stn) Then
                For k = LBound(Dizim, 2) To UBound(Dizim, 2)
                    Tmp = Dizim(j, k)
                    Dizim(j, k) = Dizim(i, k)
                    Dizim(i, k) = Tmp
                Next
            End If
        Next
    Next
    Diz = Dizim
End Function
Private Sub UserForm_Initialize()
On Error Resume Next
TextBoxbara = GetSetting("ilhan", "Settings", "bara") 'bara
GoTo atla1
'--
dizi1 = GetSetting("ilhan", "Settings", "baralar") 'baralar
For i = 0 To 12
 Controls("TBSM" & i + 1).Value = Split(dizi1, "_")(i)
Next i
'--
atla1:
Call salterayakuzunluk1
For n = 2 To ListBox0.ListCount
ListBoxPD1.AddItem ListBox0.List(n - 1, 0)
Next n
'--
Call salterbakırkesit1
bamper = ListBox1.List(0, 1)
For X = 1 To 13
Controls("LBA" & X).Caption = Split(bamper, ";")(X - 1)
Next
For n = 2 To ListBox1.ListCount
ListBoxmarka.AddItem ListBox1.List(n - 1, 0)
Next n
'--
Call anabakırkesit1
For n = 1 To ListBox2.ListCount
ListBoxBF1.AddItem ListBox2.List(n - 1, 0)
Next n
Call Nbakırkesit1
Call PEbakırkesit1
ListBox3.Selected(0) = True
ListBox4.Selected(0) = True
ListBoxmarka.Selected(0) = True
ListBoxPD1.Selected(2) = True
ListBoxBF1.Selected(0) = True
End Sub
Sub barakg() 'barakg
On Error GoTo hata
For X = 1 To 13
Controls("TBK" & X).Caption = Split(bkesit, ";")(X - 1)
Controls("TAK" & X).Caption = Split(bakesit, ";")(X - 1)
Next
hata:
End Sub
Private Sub ListBoxMARKA_Click()
On Error GoTo hata
ListBox1.Object = "": ListBox1.Object = ListBoxmarka.List(ListBoxmarka.listIndex)
bkesit = ListBox1.List(ListBox1.listIndex, 1)
For X = 1 To 13
Controls("TBK" & X).Caption = Split(bkesit, ";")(X - 1)
a = Replace(Controls("TBK" & X).Caption, "x", " * ")
Controls("LB" & X).Caption = Evaluate(a) * 0.0089
Next
Call BiP
hata:
End Sub
Sub salterbakırkesit1() 'txt dosyası
Dim Rky As String, Dosyam As String
    Dim Ert As Long, satır As Long
    Dim ayır As Variant
    Dosyam = "C:\Belgelerim\Cemex\Ayarlar\Bakır Hesabı\Şalter Bara Kesitleri.txt"
    Ert = FreeFile
    On Error Resume Next
    Open Dosyam For Input As #Ert
    If Err.Number <> 0 Then MsgBox "Şalter Bara Kesitleri.txt / Dosyası Bulunamadı !", vbCritical, "Hata !": Exit Sub
    On Error GoTo 0
    satır = 1
    ListBox1.Clear
Do While Not EOF(Ert)
        Line Input #Ert, Rky
        ayır = Split(Rky, ";")
        ListBox1.AddItem ayır(0): ListBox1.List(ListBox1.ListCount - 1, 1) = Replace(Rky, ayır(0) & ";", "")
     satır = satır + 1
Loop
Close #Ert
End Sub
Private Sub ListBoxPD1_Click()
On Error GoTo hata
ListBox0.Object = "": ListBox0.Object = ListBoxPD1.List(ListBoxPD1.listIndex)
bsboy = ListBox0.List(ListBox0.listIndex, 1)
For X = 1 To 13
Controls("TBSM" & X) = Split(bsboy, ";")(X - 1)
Next
hata:
End Sub
Sub salterayakuzunluk1()
Dim Rky As String, Dosyam As String
    Dim Ert As Long, satır As Long
    Dim ayır As Variant
    Dosyam = "C:\Belgelerim\Cemex\Ayarlar\Bakır Hesabı\Şalter Bara Ayak Uzunlukları.txt"
    Ert = FreeFile
    On Error Resume Next
    Open Dosyam For Input As #Ert
    If Err.Number <> 0 Then MsgBox "Şalter Bara Ayak Uzunlukları.txt / Dosyası Bulunamadı !", vbCritical, "Hata !": Exit Sub
    On Error GoTo 0
    satır = 1
    ListBox0.Clear
Do While Not EOF(Ert)
        Line Input #Ert, Rky
        ayır = Split(Rky, ";")
        ListBox0.AddItem ayır(0): ListBox0.List(ListBox0.ListCount - 1, 1) = Replace(Rky, ayır(0) & ";", "")
     satır = satır + 1
Loop
Close #Ert
End Sub
Private Sub ListBoxBF1_Click()
On Error GoTo hata
ListBox2.Selected(ListBoxBF1.listIndex) = True: Lbara = ListBox2.List(ListBox2.listIndex, 1)
If ListBoxBF1.listIndex < ListBox3.ListCount Then
ListBox3.Selected(ListBoxBF1.listIndex) = True: Nbara = ListBox3.List(ListBox3.listIndex, 1)
Else
Nbara = ListBox3.List(ListBox3.listIndex, 1)
End If
If ListBoxBF1.listIndex < ListBox4.ListCount Then
ListBox4.Selected(ListBoxBF1.listIndex) = True: Pbara = ListBox4.List(ListBox4.listIndex, 1)
Else
Pbara = ListBox4.List(ListBox4.listIndex, 1)
End If
For X = 1 To 13
Controls("TAK" & X).Caption = Split(Lbara, ";")(X - 1) 'anabara
ba = Replace(Controls("TAK" & X).Caption, "x", " * ")
Controls("TAK" & X).ControlTipText = Evaluate(ba) * 0.0089
'---
Controls("TNK" & X) = Split(Nbara, ";")(X - 1) 'nötr
bn = Replace(Controls("TNK" & X).Caption, "x", " * ")
Controls("TNK" & X).ControlTipText = Evaluate(bn) * 0.0089
'---
Controls("TTK" & X) = Split(Pbara, ";")(X - 1) 'pe
bt = Replace(Controls("TTK" & X).Caption, "x", " * ")
Controls("TTK" & X).ControlTipText = Evaluate(bt) * 0.0089
'---
hata:
Next
If CKNT.Value = True Then Call Nötr
If CKPE.Value = True Then Call PE
Call BiP
End Sub
Sub anabakırkesit1() 'txt dosyası
Dim Rky As String, Dosyam As String
    Dim Ert As Long, satır As Long
    Dim ayır As Variant
    Dosyam = "C:\Belgelerim\Cemex\Ayarlar\Bakır Hesabı\Ana Bara Kesitleri.txt"
    Ert = FreeFile
    On Error Resume Next
    Open Dosyam For Input As #Ert
    If Err.Number <> 0 Then MsgBox "Ana Bara Kesitleri.txt / Dosyası Bulunamadı !", vbCritical, "Hata !": Exit Sub
    On Error GoTo 0
    satır = 1
    ListBox2.Clear
Do While Not EOF(Ert)
        Line Input #Ert, Rky
        ayır = Split(Rky, ";")
        ListBox2.AddItem ayır(0): ListBox2.List(ListBox2.ListCount - 1, 1) = Replace(Rky, ayır(0) & ";", "")
     satır = satır + 1
Loop
Close #Ert
End Sub
Sub Nbakırkesit1() 'txt dosyası
Dim Rky As String, Dosyam As String
    Dim Ert As Long, satır As Long
    Dim ayır As Variant
    Dosyam = "C:\Belgelerim\Cemex\Ayarlar\Bakır Hesabı\Nötr Bara Kesitleri.txt"
    Ert = FreeFile
    On Error Resume Next
    Open Dosyam For Input As #Ert
    If Err.Number <> 0 Then MsgBox "Nötr Bara Kesitleri.txt / Dosyası Bulunamadı !", vbCritical, "Hata !": Exit Sub
    On Error GoTo 0
    satır = 1
    ListBox3.Clear
Do While Not EOF(Ert)
        Line Input #Ert, Rky
        ayır = Split(Rky, ";")
        ListBox3.AddItem ayır(0): ListBox3.List(ListBox3.ListCount - 1, 1) = Replace(Rky, ayır(0) & ";", "")
     satır = satır + 1
Loop
Close #Ert
End Sub
Sub PEbakırkesit1() 'txt dosyası
Dim Rky As String, Dosyam As String
    Dim Ert As Long, satır As Long
    Dim ayır As Variant
    Dosyam = "C:\Belgelerim\Cemex\Ayarlar\Bakır Hesabı\Toprak Bara Kesitleri.txt"
    Ert = FreeFile
    On Error Resume Next
    Open Dosyam For Input As #Ert
    If Err.Number <> 0 Then MsgBox "Toprak Bara Kesitleri.txt / Dosyası Bulunamadı !", vbCritical, "Hata !": Exit Sub
    On Error GoTo 0
    satır = 1
    ListBox4.Clear
Do While Not EOF(Ert)
        Line Input #Ert, Rky
        ayır = Split(Rky, ";")
        ListBox4.AddItem ayır(0): ListBox4.List(ListBox4.ListCount - 1, 1) = Replace(Rky, ayır(0) & ";", "")
     satır = satır + 1
Loop
Close #Ert
End Sub
Sub baraliste1()
On Error Resume Next
a = Selection.row
If Range("B" & a).FormulaR1C1 = "BÖLÜM ADI/NO:" Then a = a + 1: Range("B" & a).Select
If TBSK14 > 0 Then
bkod = GetSetting("ilhan", "Settings", "skdb")
For X = 1 To 13
  If Not Controls("TBSK" & X).Value = "" Then
  bkesit = Controls("TBK" & X).Caption: bkg = Controls("TBSK" & X).Value
  bkesit1 = Replace(bkesit, "x", "."): bkesit1 = Replace(bkesit1, ")", ""): bkesit1 = Split(bkesit1, "(")(1)
  Call bakir_gir
  Range("B" & a).FormulaR1C1 = bkod & "." & bkesit1
  Range("C" & a).FormulaR1C1 = "Elektrolitik Bakır Bara" & " (" & bkesit & " - " & Controls("LBA" & X) & " - " & ListBoxmarka.List(ListBoxmarka.listIndex) & ")"
  Range("E" & a).FormulaR1C1 = WorksheetFunction.RoundUp(bkg + (bkg * TBtolr / 100), 2)
  a = a + 1
  Range("A" & a).Select
  End If
Next
End If
'--
If TBAK14 > 0 Then
For X = 1 To 13
  If Not Controls("TBAK" & X).Value = "" Then
  bkesit = Controls("TAK" & X).Caption: bkg = Controls("TBAK" & X).Value
  bkesit1 = Replace(bkesit, "x", "."): bkesit1 = Replace(bkesit1, ")", ""): bkesit1 = Split(bkesit1, "(")(1)
  Call bakir_gir
  Range("B" & a).FormulaR1C1 = bkod & "." & bkesit1
  Range("C" & a).FormulaR1C1 = "Elektrolitik Bakır Bara" & " (" & bkesit & " - " & Controls("LBA" & X) & " için " & "Ana Bara)"
  Range("E" & a).FormulaR1C1 = WorksheetFunction.RoundUp(bkg + (bkg * TBtolr / 100), 2)
  a = a + 1
  Range("A" & a).Select
  End If
Next
End If
'--
If TBNK14 > 0 Then
For X = 1 To 13
  If Not Controls("TBNK" & X).Value = "" Then
  bkesit = Controls("TNK" & X).Caption: bkg = Controls("TBNK" & X).Value
  bkesit1 = Replace(bkesit, "x", "."): bkesit1 = Replace(bkesit1, ")", ""): bkesit1 = Split(bkesit1, "(")(1)
  Call bakir_gir
  Range("B" & a).FormulaR1C1 = bkod & "." & bkesit1
  Range("C" & a).FormulaR1C1 = "Elektrolitik Bakır Bara" & " (" & bkesit & " - " & Controls("LBA" & X) & " için " & "Nötr)"
  Range("E" & a).FormulaR1C1 = WorksheetFunction.RoundUp(bkg + (bkg * TBtolr / 100), 2)
  a = a + 1
  Range("A" & a).Select
  End If
Next
End If
'--
If TBTK14 > 0 Then
For X = 1 To 13
  If Not Controls("TBTK" & X).Value = "" Then
  bkesit = Controls("TTK" & X).Caption: bkg = Controls("TBTK" & X).Value
  bkesit1 = Replace(bkesit, "x", "."): bkesit1 = Replace(bkesit1, ")", ""): bkesit1 = Split(bkesit1, "(")(1)
  Call bakir_gir
  Range("B" & a).FormulaR1C1 = bkod & "." & bkesit1
  Range("C" & a).FormulaR1C1 = "Elektrolitik Bakır Bara" & " (" & bkesit & " - " & Controls("LBA" & X) & " için " & "için Toprak)"
  Range("E" & a).FormulaR1C1 = WorksheetFunction.RoundUp(bkg + (bkg * TBtolr / 100), 2)
  a = a + 1
  Range("A" & a).Select
  End If
Next
End If
End Sub