Public bfyt As String
Dim ir As String, ir1 As Byte
Public prs As Byte
Dim r 'renk
Dim ptdf As Byte
Dim lr1 As Integer, lr2 As Integer
Dim marka, dts, se1, pmlz
Dim panotip, prgpanokod
Dim secim
Dim prgkod, ptipkod, ptip, pacıklama, pacıklama2, pformat, pcarpan, pdermesafe, pdercarpan, kodform, pmsharf
Dim prcarp
Dim Cphtv, Cpptv ', Cpatv
Private Sub CBBPM1_Change()
ListBoxmarka.listIndex = CBBPM1.listIndex
SaveSetting "ilhan", "Settings", "panomarka", CBBPM1.listIndex
End Sub
Private Sub CBPC2_Click() '2021 +21 PANO CARPAN
On Error Resume Next
If LBPC1.ListCount < 1 Then Exit Sub
If CBPC2.Value = True Then
TBPLIST = 1
TBCAK.Value = "Teklif Varsayılan Pano Çarpanları"
LBPC2.Object = "": LBPC2.Visible = False: LBPC1.Visible = True
If LBPC1.listIndex >= 0 Then Frame24.Enabled = True Else Frame24.Enabled = False
LBPC1.IntegralHeight = False: LBPC1.Height = 145: LBPC1.IntegralHeight = True
Else
TBPLIST = 2
LBPC1.Object = "": LBPC1.Visible = False: LBPC2.Visible = True
LBPC2.IntegralHeight = False: LBPC2.Height = 145: LBPC2.IntegralHeight = True
If LBPC2.BackColor = &HE6D3C4 Then TBCAK.Value = "Teklife Girilmiş Pano Çarpanları" 'mavi
If LBPC2.listIndex >= 0 Then Frame24.Enabled = True Else Frame24.Enabled = False
End If
End Sub
Sub Panoteklifsüz() ' mevcutlar pano süz 2021++
On Error Resume Next
LBMSC2.Visible = True
LBPC2.Clear
For n = 0 To LBPC1.ListCount - 1
  If LBPC1.List(n, 3) <> "" Then
     LBPC2.AddItem LBPC1.List(n)
     LBPC2.List(LBPC2.ListCount - 1, 1) = LBPC1.List(n, 1)
     LBPC2.List(LBPC2.ListCount - 1, 2) = LBPC1.List(n, 2)
     LBPC2.List(LBPC2.ListCount - 1, 3) = LBPC1.List(n, 3)
     LBPC2.List(LBPC2.ListCount - 1, 5) = LBPC1.List(n, 4)
     LBPC2.List(n, 5) = LBPC1.ListCount - 1
  End If
Next n
LBMSC2.IntegralHeight = False: LBMSC2.Height = 145: LBMSC2.IntegralHeight = True
End Sub
Private Sub CommandButton25_Click() '2021 +21 PANO CARPAN
If LBPC1.ListCount < 0 Then Exit Sub
If LBPC1.BackColor = &H80000004 Then 'PC çarpanları
'ListBoxKT1.List(LBPC1.ListIndex, 6) = LBPC1.List(LBPC1.ListIndex, 2) 'tek tek
n1 = CDbl(LBM2A1)
For i = 0 To LBPC1.ListCount - 1
ListBoxKT1.List(n1 + i, 6) = LBPC1.List(i, 2)  'tüm liste
Next
Call PcpanoCarpanver
Else
Call CarpanlarPCTV 'teklif çarpanları
End If
KDtipCP = TBPC2: Call hesapla
'Call panotipcarpan1
End Sub
Sub PcpanoCarpanver() 'pano çarpanlar listeden -> txt ye '2021 +21 PANO CARPAN
On Error Resume Next
'Const MyFile As String = "C:\Belgelerim\Cemex\Ayarlar\Panolar\Pano Tip Tanımlamalar.txt"
Dim Myfile As String
Myfile = "C:\Belgelerim\Cemex\Ayarlar\Panolar\" & marka & "\Pano Tip Tanımlamalar.txt"
If Len(dir(Myfile)) > 0 Then
     Kill Myfile
 End If
    Open Myfile For Append As #1
    For i = 0 To ListBoxKT1.ListCount - 1
    If ListBoxKT1.List(i, 0) = marka Then
        Print #1, ListBoxKT1.List(i, 0) & ";" & ListBoxKT1.List(i, 1) & ";" & ListBoxKT1.List(i, 2) & ";" & ListBoxKT1.List(i, 3) & ";" & _
        ListBoxKT1.List(i, 4) & ";" & ListBoxKT1.List(i, 5) & ";" & ListBoxKT1.List(i, 6) & ";" & ListBoxKT1.List(i, 7) & ";" & _
        ListBoxKT1.List(i, 8) & ";" & ListBoxKT1.List(i, 9)
    End If
    Next
    Close #1
End Sub
Private Sub CommandButton26_Click() '2021 +21 PANO CARPAN
C1 = LBPC1.listIndex: C2 = LBPC2.listIndex
If Controls("LBPC" & TBPLIST.Value).ListCount < 1 Then Exit Sub
'Call PanoCarpanlar
Call pano_gir1
Call CarpanlarPCTV
If C1 >= 0 Then LBPC1.Selected(C1) = False: LBPC1.Selected(C1) = True
If C2 >= 0 Then LBPC2.Selected(C2) = False: LBPC2.Selected(C2) = True
If LBPC1.BackColor = &HDBE8DB Then TBCAK.Value = "Teklif Pano Çarpanlar" 'yeşil
Label465.Caption = Format(WorksheetFunction.SumIfs(dts.Range("P1:P65536"), dts.Range("A1:A65536"), "PP-" & "*", dts.Range("D1:D65536"), marka), "#,##0.00") & " " & "TL"
MsgBox "  Teklife aktarıldı. ", vbOKOnly, "scngnr@hotmail.com"
End Sub
Private Sub CommandButton27_Click()
Call panoyapıparametre
End Sub
Private Sub Image76_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single) '2021 +21 PANO CARPAN
Dim rds
rds = CreateObject("Scripting.FileSystemObject").FolderExists("C:\Belgelerim\CEMEX\Resimler\" & marka)
If rds = True Then
CreateObject("Shell.Application").Open "C:\Belgelerim\CEMEX\Resimler\" & marka
Else
MsgBox "C:\Belgelerim\CEMEX\Resimler\" & marka & " klasörü yok", vbInformation, "scngnr@hotmail.com"
End If
End Sub
Private Sub Image449_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single) '2021 +21 PANO CARPAN
Dim rds
rds = CreateObject("Scripting.FileSystemObject").FolderExists("C:\Belgelerim\CEMEX\Resimler\" & ir)
If rds = True Then
CreateObject("Shell.Application").Open "C:\Belgelerim\CEMEX\Resimler\" & ir
Else
MsgBox "C:\Belgelerim\CEMEX\Resimler\" & ir & " klasörü yok", vbInformation, "scngnr@hotmail.com"
End If
End Sub
Private Sub Image448_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single) '2021 +21 PANO CARPAN
Dim rds
rds = CreateObject("Scripting.FileSystemObject").FolderExists("C:\Belgelerim\CEMEX\Resimler\" & Left(marka, 3))
If rds = True Then
CreateObject("Shell.Application").Open "C:\Belgelerim\CEMEX\Resimler\" & Left(marka, 3)
Else
MsgBox "C:\Belgelerim\CEMEX\Resimler\" & Left(marka, 3) & " klasörü yok", vbInformation, "scngnr@hotmail.com"
End If
End Sub
Private Sub Label363_Click() '2021 +21 PANO CARPAN
MsgBox "FORMATLAR:" & vbCr & "YGD1 : Y200 G06 D06" & vbCr & "YGD2 : Y080 G60 D20" & vbCr & "YG1 : Y080 G06" _
& vbCr & vbCr & "YGD1 için örnek:" & vbCr & "DD2000604D041+OS", vbOKOnly, "scngnr@hotmail.com"
End Sub
Private Sub ListBoxKT1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
On Error Resume Next
CreateObject("Shell.Application").Open "C:\Belgelerim\Cemex\Ayarlar\Panolar\" & _
ListBoxKT1.List(ListBoxKT1.listIndex, 0) & "\Pano Tip Tanımlamalar.txt"
End Sub
Private Sub ListBoxKT2_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
On Error Resume Next
CreateObject("Shell.Application").Open "C:\Belgelerim\Cemex\Ayarlar\Panolar\" & marka & "\Pano Yapısal Parametreler.txt"
End Sub
Private Sub ListBoxMARKA2_Click() '2021 +21 PANO CARPAN
On Error Resume Next
marka = ListBoxMARKA2.List(ListBoxMARKA2.listIndex)
'panocarpangetir
If ListBoxKT1.ListCount < 1 Then Exit Sub
LBPC1.BackColor = &H80000004 'gri renk PC
LBM2A1 = "": LBM2A2 = "" 'listede başlangıc-bitiş sırası
LBPC1.Clear
For n = 0 To ListBoxKT1.ListCount - 1
  If ListBoxKT1.List(n, 0) = ListBoxMARKA2.List(ListBoxMARKA2.listIndex) Then
If LBM2A1 = "" Then LBM2A1 = n
     LBPC1.AddItem ListBoxKT1.List(n, 1)
     LBPC1.List(LBPC1.ListCount - 1, 1) = ListBoxKT1.List(n, 3)
     LBPC1.List(LBPC1.ListCount - 1, 2) = ListBoxKT1.List(n, 6)
     LBPC1.List(LBPC1.ListCount - 1, 6) = ListBoxKT1.List(n, 5)
     LBPC1.List(LBPC1.ListCount - 1, 7) = ListBoxKT1.List(n, 7)
     LBPC1.List(LBPC1.ListCount - 1, 8) = ListBoxKT1.List(n, 8)
     LBM2A2 = n
  End If
Next n
TBCAK.Value = "Firma Varsayılan Pano Çarpanlar"
TBPLIST = 1
End Sub
Private Sub ListBoxMARKA21_Click() '2021 +21 PANO CARPAN
On Error Resume Next
TBPLIST = 2
marka = ListBoxMARKA21.List(ListBoxMARKA21.listIndex)
Call panoyapıparametre
LBPC2.IntegralHeight = False: LBPC2.Height = 145: LBPC2.IntegralHeight = True
LBPC2.Clear
For n = 0 To LBPC3.ListCount - 1
  If LBPC3.List(n, 1) = marka Then
     LBPC2.AddItem Replace(LBPC3.List(n), "-auto", "")
     LBPC2.List(LBPC2.ListCount - 1, 1) = LBPC3.List(n, 1)
     LBPC2.List(LBPC2.ListCount - 1, 3) = _
     WorksheetFunction.SumIfs(dts.Range("E1:E65536"), dts.Range("A1:A65536"), LBPC3.List(n), dts.Range("D1:D65536"), marka) & " " & "Ad."
   End If
Next n
'--
GoTo hepsidahil
tkmarka = marka
ListBoxKT11.Object = tkmarka
i = ListBoxKT11.List(ListBoxKT11.listIndex, 1)
For n = 0 To ListBoxKT11.ListCount - 1
  If ListBoxKT11.List(n, 0) = tkmarka Then
  i = ListBoxKT11.List(n, 1): m = ListBoxKT11.List(n, 2)
  GoTo atla1
  End If
Next n
'--
tkmarka = "LOKAL"
ListBoxKT11.Object = tkmarka
i = ListBoxKT11.List(ListBoxKT11.listIndex, 1)
For n = 0 To ListBoxKT11.ListCount - 1
  If ListBoxKT11.List(n, 0) = tkmarka Then
  i = ListBoxKT11.List(n, 1): m = ListBoxKT11.List(n, 2)
  GoTo atla1
  End If
Next n
hepsidahil:
'i = 0
'--
m = ListBoxKT1.ListCount - 1
atla1:
ListBoxKT11.Object = ""
ListBoxKT11.Object = marka
If ListBoxKT11.Object = "" Then tkmarka = "LOKAL"
i = ListBoxKT11.List(ListBoxKT11.listIndex, 1)
For n = 0 To LBPC2.ListCount - 1
  For X = i To m
  If LBPC2.List(n, 0) = ListBoxKT1.List(X, 1) Then
     LBPC2.List(n, 1) = ListBoxKT1.List(X, 3)
     LBPC2.List(n, 2) = ListBoxKT1.List(X, 6)
     LBPC2.List(n, 6) = ListBoxKT1.List(X, 5)
     LBPC2.List(n, 7) = ListBoxKT1.List(X, 7)
     LBPC2.List(n, 8) = ListBoxKT1.List(X, 8)
     Exit For
    End If
  Next X
Next n
LBPC2.IntegralHeight = False: LBPC2.Height = 145: LBPC2.IntegralHeight = True
'--
Label441.Caption = "      " & marka
Call PanoCarpanlar
Label455.Caption = ""
Label465.Caption = ""
Label465.Caption = Format(WorksheetFunction.SumIfs(dts.Range("P1:P65536"), dts.Range("A1:A65536"), "PP-" & "*", dts.Range("D1:D65536"), marka), "#,##0.00") & " " & "TL"
'TBPT1 = marka
Frame24.Enabled = False
LBPC1.BackColor = &HDBE8DB
End Sub
Sub PanoCarpanlar() '2021 CARPAN KONTROL '2021 +21 PANO CARPAN
On Error Resume Next
Dim Cphtv, Cpptv ', Cpatv
mr1 = Replace(marka, " ", "_")
mr1 = Replace(mr1, "-", "_")
Cphtv = "": Cpptv = ""
Cphtv = ActiveWorkbook.names("Cphtv_" & mr1).RefersToR1C1 'tipler
Cpptv = ActiveWorkbook.names("Cpptv_" & mr1).RefersToR1C1: 'Cpatv = ActiveWorkbook.Names("Cpatv_" & marka).RefersToR1C1
If Cphtv = Empty Then
If Workbooks(dt).ActiveSheet.CodeName = "T1" Then GoTo atla0
'--
Cpptv = ActiveWorkbook.names("Cptv").RefersToR1C1 'tipler
If Not Cpptv = Empty Then  'eskitipler
diziP0 = "PP-DD;PP-DH;PP-DX;PP-F2;PP-F3;PP-F4;PP-SD;PP-SH;PP-SX;PP-SA;PP-KK;PP-TK"
dizi1 = ActiveWorkbook.names("Cptv").Comment
ayır1 = Split(dizi1, "-")
diziP1 = ayır1(0) & "-" & ayır1(4) & "-" & ayır1(5) & "-" & ayır1(1) & "-" & ayır1(2) & "-" & ayır1(3) _
& "-" & ayır1(6) & "-" & ayır1(7) & "-" & ayır1(8) & "-" & ayır1(9) & "-" & ayır1(10) & "-" & ayır1(11)
ActiveWorkbook.names.Add Name:="Cphtv_" & mr1, RefersToR1C1:=mr1 & " Pano Tipler"
ActiveWorkbook.names("Cphtv_" & mr1).Comment = diziP0
ActiveWorkbook.names.Add Name:="Cpptv_" & mr1, RefersToR1C1:=mr1 & " Pano Çarpanlar"
ActiveWorkbook.names("Cpptv_" & mr1).Comment = diziP1
GoTo atla3
End If
atla0:
'--
If ListBoxKT1.ListCount < 1 Then Exit Sub
LBPC1.Clear
'--
For n = 0 To ListBoxKT11.ListCount - 1
  If ListBoxKT11.List(n, 0) = marka Then
  i = ListBoxKT11.List(n, 1): m = ListBoxKT11.List(n, 2)
  GoTo atla1
  End If
Next n
'--
For n = 0 To LBPC2.ListCount - 1
  LBPC1.AddItem LBPC2.List(n, 0)
  LBPC1.List(LBPC1.ListCount - 1, 1) = LBPC2.List(n, 1)
  LBPC1.List(LBPC1.ListCount - 1, 2) = LBPC2.List(n, 2)
  LBPC1.List(LBPC1.ListCount - 1, 3) = LBPC2.List(n, 3)
  LBPC1.List(LBPC1.ListCount - 1, 5) = LBPC1.ListCount - 1
  LBPC1.List(LBPC1.ListCount - 1, 6) = LBPC2.List(n, 6)
  LBPC1.List(LBPC1.ListCount - 1, 7) = LBPC2.List(n, 7)
  LBPC1.List(LBPC1.ListCount - 1, 8) = LBPC2.List(n, 8)
Next n
GoTo atla2
atla1:
For n = i To m
     LBPC1.AddItem ListBoxKT1.List(n, 1)
     LBPC1.List(LBPC1.ListCount - 1, 1) = ListBoxKT1.List(n, 3)
     LBPC1.List(LBPC1.ListCount - 1, 2) = ListBoxKT1.List(n, 6)
     LBPC1.List(LBPC1.ListCount - 1, 6) = ListBoxKT1.List(n, 5)
     LBPC1.List(LBPC1.ListCount - 1, 7) = ListBoxKT1.List(n, 7)
     LBPC1.List(LBPC1.ListCount - 1, 8) = ListBoxKT1.List(n, 8)
Next n
For n = 0 To LBPC2.ListCount - 1
  For i = 0 To LBPC1.ListCount - 1
  If LBPC1.List(i, 0) = LBPC2.List(n, 0) Then
  LBPC1.List(i, 3) = LBPC2.List(n, 3)
  LBPC1.List(i, 5) = n
  LBPC2.List(n, 5) = i
  GoTo atla2
  End If
  Next i
  LBPC1.AddItem LBPC2.List(n, 0)
  LBPC1.List(LBPC1.ListCount - 1, 1) = LBPC2.List(n, 1)
  LBPC1.List(LBPC1.ListCount - 1, 2) = LBPC2.List(n, 2)
  LBPC1.List(LBPC1.ListCount - 1, 3) = LBPC2.List(n, 3)
  LBPC1.List(LBPC1.ListCount - 1, 5) = n
  LBPC2.List(n, 5) = LBPC1.ListCount - 1
atla2:
Next n
Exit Sub
Call CarpanlarPCTV
Else
atla3:
TBCAK.Value = "Teklif Pano Çarpanlar"
Call CarpanlarCPTA
LBPC1.BackColor = &HDBE8DB: LBPC2.BackColor = &HDBE8DB   'açık yeşil renk teklif
End If
LBPC1.IntegralHeight = False: LBPC1.Height = 145: LBPC1.IntegralHeight = True
End Sub
Sub panocarpangetir() 'xxxxx
Call panoyapıparametre: Call panotipparametre
Label441.Caption = "      " & marka
Call PanoCarpanlar
Call teklifpanocarpan_mevcutlar
Label455.Caption = ""
'TBPT1 = marka
Frame24.Enabled = False
If LBPC2.Visible = True Then
If LBPC2.BackColor = &HE6D3C4 Then TBCAK.Value = "Teklif Pano Çarpanlar" 'mavi
Else
If LBPC1.BackColor = &H80000004 Then TBCAK.Value = "Firma Varsayılan Pano Çarpanlar" 'gri
If LBPC1.BackColor = &HDBE8DB Then TBCAK.Value = "Teklif Pano Çarpanlar" 'yeşil
End If
End Sub
Sub teklifpanocarpan_mevcutlar() ' pano çarpan mevcutlar 2021 PP-ED şeklinde
On Error Resume Next
son = dts.Range("B65536").End(xlUp).row
T1 = "A" & 2 & ":" & "A" & son
LBPC2.Clear
 For n = 0 To LBPC1.ListCount - 1
 b = LBPC1.List(n, 1) & "*"
 Set a = dts.Range(T1).Find(b, LookIn:=xlValues, LookAt:=xlPart)
 If a Is Nothing Then GoTo git:
 c = a.Address
   If marka = dts.Range("D" & a.row) Then
GoTo var:
   Else
   Do
   Set a = dts.Range(T1).FindNext(a)
   If marka = dts.Range("D" & a.row) Then GoTo var:
   Loop While c <> a.Address
   End If
GoTo git:
var:
     LBPC1.List(n, 3) = WorksheetFunction.SumIfs(dts.Range("E1:E65536"), dts.Range("A1:A65536"), b, dts.Range("D1:D65536"), marka) & " " & "Ad."
     LBPC2.AddItem LBPC1.List(n)
     LBPC2.List(LBPC2.ListCount - 1, 1) = LBPC1.List(n, 1)
     LBPC2.List(LBPC2.ListCount - 1, 2) = LBPC1.List(n, 2)
     LBPC2.List(LBPC2.ListCount - 1, 3) = LBPC1.List(n, 3)
     LBPC2.List(LBPC2.ListCount - 1, 5) = n
     
     LBPC2.List(LBPC2.ListCount - 1, 6) = LBPC1.List(n, 6)
     LBPC2.List(LBPC2.ListCount - 1, 7) = LBPC1.List(n, 7)
     LBPC2.List(LBPC2.ListCount - 1, 8) = LBPC1.List(n, 8)
     
     LBPC1.List(n, 4) = n
     LBPC1.List(n, 5) = LBPC2.ListCount - 1
git:
 Next n
LBPC2.IntegralHeight = False: LBPC2.Height = 145: LBPC2.IntegralHeight = True
End Sub
Sub CarpanlarCPTA() 'pano çarpanlar ad yöneticisinden -> listeye 2021 '2021 +21 PANO CARPAN
On Error Resume Next
Dim Cphtv, Cpptv ', Cpatv
mr1 = Replace(marka, " ", "_")
mr1 = Replace(mr1, "-", "_")
Cphtv = "": Cpptv = ""
Cphtv = ActiveWorkbook.names("Cphtv_" & mr1).Comment 'tipler
Cpptv = ActiveWorkbook.names("Cpptv_" & mr1).Comment
diziP0 = Cphtv: diziP1 = Cpptv
ayır0 = Split(diziP0, ";"): ayır1 = Split(diziP1, "-") ': ayır2 = Split(dizi2, "-")
tsay = Len(diziP0) - Len(Replace(diziP0, ";", "")) + 1
LBPC1.Clear
For n = 1 To tsay '
  LBPC1.AddItem ayır0(n - 1)
  LBPC1.List(n - 1, 1) = marka
  LBPC1.List(n - 1, 2) = ayır1(n - 1)
Next n
'--
i = 0: m = ListBoxKT1.ListCount - 1
GoTo atla1
For n = 0 To ListBoxKT11.ListCount - 1
  If ListBoxKT11.List(n, 0) = tkmarka Then
  i = ListBoxKT11.List(n, 1): m = ListBoxKT11.List(n, 2)
  GoTo atla1
  End If
Next n
'--
tkmarka = "LOKAL"
For n = 0 To ListBoxKT11.ListCount - 1
  If ListBoxKT11.List(n, 0) = tkmarka Then
  i = ListBoxKT11.List(n, 1): m = ListBoxKT11.List(n, 2)
  GoTo atla1
  End If
Next n

For n = 0 To LBPC1.ListCount - 1
  For i = 0 To LBPC2.ListCount - 1
  If LBPC1.List(n, 0) = LBPC2.List(i, 0) Then
  LBPC1.List(n, 1) = LBPC2.List(i, 1) 'pano tip
  LBPC1.List(n, 3) = LBPC2.List(i, 3)
  LBPC1.List(n, 6) = LBPC2.List(i, 6)
  LBPC1.List(n, 7) = LBPC2.List(i, 7)
  LBPC1.List(n, 8) = LBPC2.List(i, 8)
  LBPC1.List(n, 5) = i
  LBPC2.List(i, 5) = n
  LBPC2.List(i, 2) = LBPC1.List(n, 2)
  Exit For
  End If
  Next i
Next n
GoTo atla2

atla1:
ListBoxKT11.Object = ""
ListBoxKT11.Object = marka
If ListBoxKT11.Object = "" Then tkmarka = "LOKAL"
i = ListBoxKT11.List(ListBoxKT11.listIndex, 1)
For n = 0 To LBPC1.ListCount - 1
  For X = i To m
  If LBPC1.List(n, 0) = ListBoxKT1.List(X, 1) Then
  LBPC1.List(n, 1) = ListBoxKT1.List(X, 3) 'pano tip
  LBPC1.List(n, 6) = ListBoxKT1.List(X, 5)
  LBPC1.List(n, 7) = ListBoxKT1.List(X, 7)
  LBPC1.List(n, 8) = ListBoxKT1.List(X, 8)
  Exit For
  End If
  Next X
Next n

For n = 0 To LBPC2.ListCount - 1
  For i = 0 To LBPC1.ListCount - 1
  If LBPC2.List(n, 0) = LBPC1.List(i, 0) Then
  'LBPC2.List(n, 1) = LBPC1.List(i, 1)
  LBPC1.List(i, 3) = LBPC2.List(n, 3)
  LBPC1.List(i, 5) = n
  LBPC2.List(n, 5) = i
  LBPC2.List(n, 2) = LBPC1.List(i, 2)
  GoTo atla2
  End If
  Next i
  LBPC1.AddItem LBPC2.List(n, 0)
  LBPC1.List(LBPC1.ListCount - 1, 1) = LBPC2.List(n, 1)
  LBPC1.List(LBPC1.ListCount - 1, 2) = LBPC2.List(n, 2)
  LBPC1.List(LBPC1.ListCount - 1, 3) = LBPC2.List(n, 3)
  LBPC1.List(LBPC1.ListCount - 1, 5) = n
  LBPC2.List(n, 5) = LBPC1.ListCount - 1
atla2:
Next n
End Sub
Sub CarpanlarPCTV() '2021 +21 PANO CARPAN
On Error Resume Next
Dim Cphtv, Cpptv ', Cpatv
Cphtv = "": Cpptv = ""
mr1 = Replace(marka, " ", "_")
mr1 = Replace(mr1, "-", "_")
Cphtv = ActiveWorkbook.names("Cphtv_" & mr1).RefersToR1C1 'tipler
Cpptv = ActiveWorkbook.names("Cpptv_" & mr1).RefersToR1C1 'carpanlar
If LBPC1.ListCount < 1 Then Exit Sub
For n = 0 To LBPC1.ListCount - 1
     diziP0 = diziP0 & LBPC1.List(n, 0) & ";": diziP1 = diziP1 & LBPC1.List(n, 2) & "-"
Next n
diziP0 = Left(diziP0, Len(diziP0) - 1): diziP1 = Left(diziP1, Len(diziP1) - 1)
ActiveWorkbook.names.Add Name:="Cphtv_" & mr1, RefersToR1C1:=mr1 & " Pano Tipler": ActiveWorkbook.names("Cphtv_" & mr1).Comment = diziP0
ActiveWorkbook.names.Add Name:="Cpptv_" & mr1, RefersToR1C1:=mr1 & " Pano Çarpanlar": ActiveWorkbook.names("Cpptv_" & mr1).Comment = diziP1
TBCAK.Value = "Teklif Varsayılan Çarpanları" 'yeşil
End Sub
Private Sub LBPC1_Click() '2021 +21 PANO CARPAN
TBPLIST = 1
teklifpanocarpanliste1
TBPT1 = LBPC1.List(LBPC1.listIndex, 1)
If LBPC1.listIndex >= 0 Then Frame24.Enabled = True Else Frame24.Enabled = False
Label455.Caption = Format(WorksheetFunction.SumIfs(dts.Range("P1:P65536"), dts.Range("A1:A65536"), TBPC1 & "*", dts.Range("D1:D65536"), marka), "#,##0.00") & " " & "TL"
End Sub
Private Sub LBPC2_Click() '2021 +21 PANO CARPAN
TBPLIST = 2
teklifpanocarpanliste2
TBPT1 = LBPC2.List(LBPC2.listIndex, 1)
If LBPC2.listIndex >= 0 Then Frame24.Enabled = True Else Frame24.Enabled = False
Label455.Caption = Format(WorksheetFunction.SumIfs(dts.Range("P1:P65536"), dts.Range("A1:A65536"), TBPC1 & "*", dts.Range("D1:D65536"), marka), "#,##0.00") & " " & "TL"
End Sub
Sub teklifpanocarpanliste1() '2021 +21 PANO CARPAN
On Error Resume Next
If LBPC1.ListCount < 1 Then Exit Sub
SPPC1.Max = LBPC1.ListCount - 1
TBPC1.Text = LBPC1.List(LBPC1.listIndex, 0)
If LBPC1.List(LBPC1.listIndex, 2) >= 0 Then TBPC2.Text = LBPC1.List(LBPC1.listIndex, 2) Else TBPC2.Text = ""
SPPC1.Value = LBPC1.listIndex
End Sub
Sub teklifpanocarpanliste2() '2021 +21 PANO CARPAN
On Error Resume Next
If LBPC2.ListCount < 1 Then Exit Sub
SPPC1.Max = LBPC2.ListCount - 1
TBPC1.Text = LBPC2.List(LBPC2.listIndex, 0)
If LBPC2.List(LBPC2.listIndex, 2) >= 0 Then TBPC2.Text = LBPC2.List(LBPC2.listIndex, 2) Else TBPC2.Text = ""
SPPC1.Value = LBPC2.listIndex
End Sub
Private Sub ListBoxPTD10_Click()
On Error Resume Next
ListBoxPTD12.Clear
For n = 0 To LBPSA1.ListCount - 1
  If ListBoxPTD10.List(ListBoxPTD10.listIndex) = LBPSA1.List(n, 2) Then
     ListBoxPTD12.AddItem LBPSA1.List(n, 0)
     ListBoxPTD12.List(ListBoxPTD12.ListCount - 1, 1) = "Yoğunluk " & LBPSA1.List(n, 1)
     If Not LBPSA1.List(n, 3) = "" Then s = ListBoxPTD12.ListCount - 1
  End If
Next n
If ListBoxPTD12.ListCount >= 0 Then
 If s = "" Then
 ListBoxPTD12.Selected(ListBoxPTD12.ListCount - 1) = True
 Else
 ListBoxPTD12.Selected(s) = True
 End If
End If
End Sub
Private Sub SPPC1_Change() '2021 +21 PANO CARPAN
If TBPLIST = 2 Then LBPC2.Selected(SPPC1.Value) = True
If TBPLIST = 1 Then LBPC1.Selected(SPPC1.Value) = True
End Sub
Private Sub SPPC2_Change() '2021 +21 PANO CARPAN
On Error Resume Next: TBPC2.Text = SPPC2.Value
End Sub
Private Sub TBPC1_Change() '2021 +21 PANO CARPAN
If TBPLIST = 2 Then SPPC1.Value = LBPC2.listIndex
If TBPLIST = 1 Then SPPC1.Value = LBPC1.listIndex
End Sub
Private Sub TBPC2_Change() '2021 +21 PANO CARPAN
On Error Resume Next
SPPC2.Value = TBPC2.Text
If TBPLIST = 2 Then
LBPC2.List(LBPC2.listIndex, 2) = TBPC2.Text
LBPC1.List(LBPC2.List(LBPC2.listIndex, 5), 2) = TBPC2.Text
Else
LBPC1.List(LBPC1.listIndex, 2) = TBPC2.Text
LBPC2.List(LBPC1.List(LBPC1.listIndex, 5), 2) = TBPC2.Text
End If
End Sub
Private Sub TBPT1_Change() '2021 +21 PANO CARPAN
On Error GoTo hata
Dim rd, a
Set rd = CreateObject("Scripting.FileSystemObject")
a = rd.FileExists("C:\Belgelerim\CEMEX\Resimler\" & marka & "\" & TBPT1 & ".gif")
If a = True Then
Image1.Picture = LoadPicture("C:\Belgelerim\CEMEX\Resimler\" & marka & "\" & TBPT1 & ".gif")
Else
 a = rd.FileExists("C:\Belgelerim\CEMEX\Resimler\" & marka & "\" & marka & ".gif")
 If a = True Then
 Image1.Picture = LoadPicture("C:\Belgelerim\CEMEX\Resimler\" & marka & "\" & marka & ".gif")
 Else
 Image1.Picture = LoadPicture("C:\Belgelerim\CEMEX\Resimler\noimagepano.gif")
 End If
End If
hata:
End Sub
Private Sub TBSC1_Change() '2021+
On Error Resume Next
SPSC1.Value = TBSC1.Text
If TBMSLIST = 2 Then
LBMSC2.List(LBMSC2.listIndex, 2) = TBSC1.Text
LBMSC1.List(LBMSC2.List(LBMSC2.listIndex, 5), 2) = TBSC1.Text
Else
LBMSC1.List(LBMSC1.listIndex, 2) = TBSC1.Text
LBMSC2.List(LBMSC1.List(LBMSC1.listIndex, 5), 2) = TBSC1.Text
End If
End Sub
Private Sub TBskd_Change()
On Error GoTo hata
If Toolbar1.Buttons.Item(6).Image = ImageList1.ListImages.Item(9).Index Then UFOPAN00P2.TBRS01 = TBskd
hata:
End Sub
Private Sub TBtip_Change() '2021 +21 PANO
On Error GoTo hata
Dim rd, a
Set rd = CreateObject("Scripting.FileSystemObject")
a = rd.FileExists("C:\Belgelerim\CEMEX\Resimler\" & Left(marka, 3) & "\" & TBtip & ".gif")
If a = True Then
Image01.Picture = LoadPicture("C:\Belgelerim\CEMEX\Resimler\" & Left(marka, 3) & "\" & TBtip & ".gif")
Else
 a = rd.FileExists("C:\Belgelerim\CEMEX\Resimler\" & Left(marka, 3) & "\" & marka & ".gif")
 If a = True Then
 Image01.Picture = LoadPicture("C:\Belgelerim\CEMEX\Resimler\" & Left(marka, 3) & "\" & marka & ".gif")
 Else
 Image01.Picture = LoadPicture("C:\Belgelerim\CEMEX\Resimler\noimagepano.gif")
 End If
End If
hata:
End Sub
Private Sub UserForm_Initialize() 'form yükleme '2021 +21
On Error Resume Next
Windows(dt).Activate: Sheets("Sayfa1").Select '
Set dts = Workbooks(dt).Sheets("Sayfa1")
pmlz = mlz
Dim Ckar
Ckar = ActiveWorkbook.names("CkarO").RefersToR1C1
If Ckar = Empty Then ActiveWorkbook.names.Add Name:="CkarO", RefersToR1C1:="=""Liste Fiyatı""": bfyt = "=RC[-6]"
If Ckar = "=""Net Fiyatı""" Then bfyt = "=RC[-1]" Else bfyt = "=RC[-6]"
'--
PBar1.Value = 2
Toolbar1.ImageList = ImageList1
Toolbar1.Buttons.Item(1).Image = ImageList1.ListImages.Item(1).Index
Toolbar1.Buttons.Item(2).Image = ImageList1.ListImages.Item(2).Index
Toolbar1.Buttons.Item(3).Image = ImageList1.ListImages.Item(3).Index
Toolbar1.Buttons.Item(4).Image = ImageList1.ListImages.Item(4).Index
Toolbar1.Buttons.Item(5).Image = ImageList1.ListImages.Item(5).Index
Toolbar1.Buttons.Item(6).Image = ImageList1.ListImages.Item(6).Index
Toolbar1.Buttons.Item(7).Image = ImageList1.ListImages.Item(7).Index
Toolbar1.Buttons.Item(8).Image = ImageList1.ListImages.Item(8).Index
'--
'GoTo atla1:
If pfc = 0 Then '+++++++++++++++
Toolbar1.Buttons(1).Visible = False
Toolbar1.Buttons(2).Visible = False
Toolbar1.Buttons(5).ButtonMenus(1).Visible = False
Toolbar1.Buttons(5).ButtonMenus(2).Visible = False
 Set msayfa = Workbooks(pmlz).Worksheets("Sayfa1")
 Dim rxs
 rx = ""
 rr = 1
 rs = msayfa.Range("B65536").End(xlUp).row + 1
 TBPM01 = msayfa.Cells(2, 4).End(4)
 ir = Trim(Left(pmlz, 3))
 'ir = TBPM01
tekrar:
 ra = "B" & rr & ":" & "B" & rs
 Set rxs = msayfa.Range(ra).Find(rx, LookIn:=xlValues, LookAt:=xlPart)
 If rxs.row >= rs Then GoTo son
 If Not rxs Is Nothing Then
 ListBoxP1.AddItem msayfa.Cells(rxs.row, 1) 'ListBox
 ListBoxP1.List(ListBoxP1.ListCount - 1, 1) = msayfa.Cells(rxs.row, 1).row 'ListBox
 If ListBoxP1.ListCount > 1 Then ListBoxP1.List(ListBoxP1.ListCount - 2, 2) = msayfa.Cells(rxs.row, 1).row - 1 'ListBox
 If rxs.row > 10000 Then MsgBox ("Kodları kontrol et!  "), vbCritical, "Uyarı": Exit Sub
 rr = rxs.row
 GoTo tekrar
 End If
son:
 Call tipmontajsarfacıklamalar: Call montajsarfacıklamalar: Call MontajCarpanlar01
 Call montajsarfacıklamalar
 Call MontajCarpanlar01
 ListBoxP1.List(ListBoxP1.ListCount - 1, 2) = rs + 1 'ListBox
 Image011.Picture = LoadPicture("C:\Belgelerim\CEMEX\Resimler\" & ir & "\" & ir & ".gif")
 MultiPage1.Value = 6
 ListBoxP1.Height = 144.35: ListBoxP1.IntegralHeight = True
 LBBPM11 = "  " & Workbooks(pmlz).Name
Else
 Toolbar1.Buttons(3).Visible = False
 Call montajsarfacıklamalar: Call MontajCarpanlar01
 MultiPage1.Value = 0
 Dim ds, fs, f1, fc
 Set ds = CreateObject("Scripting.FileSystemObject")
 Set fs = ds.GetFolder("C:\Belgelerim\Cemex\Ayarlar\Panolar")
 Set fc = fs.SubFolders
 For Each f1 In fc
 ListBoxmarka.AddItem f1.Name
 ListBoxMARKA2.AddItem f1.Name
 CBBPM1.AddItem f1.Name
 Next
 'Call mevcutmarkalar 'teklifte bulunan pano markları
 Call panotipparametreler
 Call panoeskicarpanceviri
 Call panoeskicarpanlar
'marka = ListBoxMARKA.List(0)
'atla1:
'--
 PM1 = GetSetting("ilhan", "Settings", "panomarka")
 If Not PM1 = "" Then CBBPM1.listIndex = PM1 Else CBBPM1.listIndex = 0
'ListBoxMARKA.Selected(0) = True
 If ListBoxmarka.ListCount > 0 Then
 FYTOK.Value = True
 End If
End If
TBimarka = GetSetting("ilhan", "Settings", "misi") 'işçilik-montaj markası
TBsmarka = GetSetting("ilhan", "Settings", "msas") 'sarf markası
TBbmarka = GetSetting("ilhan", "Settings", "mbab") 'bara markası
TBamarka = GetSetting("ilhan", "Settings", "mama") 'ambalaj markası
TBimarkaA = GetSetting("ilhan", "Settings", "misia") 'işçilik-montaj açıklama
TBsmarkaA = GetSetting("ilhan", "Settings", "msasa") 'sarf açıklama
TBbmarkaA = GetSetting("ilhan", "Settings", "mbaba") 'bara açıklama
TBamarkaA = GetSetting("ilhan", "Settings", "mamaa") 'ambalaj açıklama
TBskdi = GetSetting("ilhan", "Settings", "skdi") 'işçilik-montaj markası kod.
TBskds = GetSetting("ilhan", "Settings", "skds") 'sarf markası kod.
TBskdb = GetSetting("ilhan", "Settings", "skdb") 'bara markası kod.
TBskda = GetSetting("ilhan", "Settings", "skda") 'ambalaj markası kod.
TextBoxamb = ActiveWorkbook.names("Catv").Comment
End Sub
Sub panoeskicarpanceviri() ''2021 +21 PANO
On Error Resume Next
Dim Cptv, Cpptv
Cptv = ActiveWorkbook.names("Cptv").RefersToR1C1 'eski carpanlar
If Cptv = Empty Then Exit Sub
Cpptv = ActiveWorkbook.names("Cptv").Comment
dizi1 = Cpptv
dizi2 = "DD-F2-F3-F4-DH-DX-SD-SH-SX-SA-KK-TK"
ayır1 = Split(dizi1, "-")
ayır2 = Split(dizi2, "-")
tsay = Len(dizi1) - Len(Replace(dizi1, "-", "")) + 1

LBPC4.Clear
For n = 1 To tsay '
    LBPC4.AddItem "PP-" & ayır2(n - 1)
    LBPC4.List(n - 1, 1) = ""
    LBPC4.List(n - 1, 2) = ayır1(n - 1)
    LBPC4.List(n - 1, 3) = "YGD1"
    LBPC4.List(n - 1, 4) = ayır1(12)
    LBPC4.List(n - 1, 5) = ayır1(13)
Next n
End Sub
Sub panoeskicarpanlar() ''2021 +21 PANO
On Error Resume Next
Dim Cptv, Cpptv
Cptv = ActiveWorkbook.names("Cptv").RefersToR1C1 'eski carpanlar
If Cptv = Empty Then Exit Sub
Cpptv = ActiveWorkbook.names("Cptv").Comment
dizi1 = Cpptv
dizi2 = "DD-F2-F3-F4-DH-DX-SD-SH-SX-SA-KK-TK"
ayır1 = Split(dizi1, "-")
ayır2 = Split(dizi2, "-")
tsay = Len(dizi1) - Len(Replace(dizi1, "-", "")) + 1

LBPC4.Clear
For n = 1 To tsay '
    LBPC4.AddItem "PP-" & ayır2(n - 1)
    LBPC4.List(n - 1, 1) = ""
    LBPC4.List(n - 1, 2) = ayır1(n - 1)
    LBPC4.List(n - 1, 3) = "YGD1"
    LBPC4.List(n - 1, 4) = ayır1(12)
    LBPC4.List(n - 1, 5) = ayır1(13)
Next n
End Sub
Sub panotipparametreler() '2021 +21 PANO
On Error Resume Next
Dim Rky As String, Dosyam As String
    Dim Ert As Long, satır As Long, i As Long
    Dim ayır As Variant
    ListBoxKT1.Clear
    satır = 1
For m = 0 To ListBoxmarka.ListCount - 1
 marka21 = ListBoxmarka.List(m, 0)
         ListBoxKT11.AddItem marka21: ListBoxKT11.List(ListBoxKT11.ListCount - 1, 1) = satır - 1
 Dosyam = "C:\Belgelerim\Cemex\Ayarlar\Panolar\" & marka21 & "\Pano Tip Tanımlamalar.txt"
 If Len(dir(Dosyam)) > 0 Then
    Ert = FreeFile
    Open Dosyam For Input As #Ert
Do While Not EOF(Ert)
        Line Input #Ert, Rky
        ayır = Split(Rky, ";")
        ListBoxKT1.AddItem marka21
tsay = Len(Rky) - Len(Replace(Rky, ";", ""))
For n = 1 To tsay '
If UBound(ayır) <> 0 Then ListBoxKT1.List(satır - 1, n) = ayır(n)
Next n
     satır = satır + 1
Loop
Close #Ert
 End If
           ListBoxKT11.List(ListBoxKT11.ListCount - 1, 2) = satır - 2
Next m
End Sub
Sub mevcutmarkalar21() 'teklifte bulunan pano markları PP-*-auto '2021 +21 PANO CARPAN
On Error Resume Next
ListBoxMARKA21.Clear: LBPC3.Clear
Dim son As Integer
son = dts.Range("D65536").End(xlUp).row
For n = 2 To son
If Left(dts.Range("A" & n), 3) = "PP-" And Right(dts.Range("A" & n), 5) = "-auto" Then
       mss = WorksheetFunction.CountIf(dts.Range("D2:D" & n), dts.Cells(n, 4).Value)
       If mss = 1 Then ListBoxMARKA21.AddItem dts.Range("D" & n)
       tip = dts.Range("A" & n): marka = dts.Range("D" & n)
       tss = WorksheetFunction.CountIfs(dts.Range("A2:A" & n), tip, dts.Range("D2:D" & n), marka)
       If tss = 1 Then
       LBPC3.AddItem dts.Range("A" & n)
       LBPC3.List(LBPC3.ListCount - 1, 1) = dts.Range("D" & n)
       Call eskicarpanvar
       End If
End If
Next n
ActiveWorkbook.names("Cptv").Delete
End Sub
Sub eskicarpanvar() '2021 eski CARPAN KONTROL '2021 +21 eski PANO CARPAN
On Error Resume Next
Dim Cphtv, Cpptv ', Cpatv
mr1 = Replace(marka, " ", "_"): mr1 = Replace(mr1, "-", "_")
Cphtv = "": Cpptv = ""
Cphtv = ActiveWorkbook.names("Cphtv_" & mr1).RefersToR1C1 'tipler
Cpptv = ActiveWorkbook.names("Cpptv_" & mr1).RefersToR1C1: 'Cpatv = ActiveWorkbook.Names("Cpatv_" & marka).RefersToR1C1
If Cphtv = Empty Then
'--
Cpptv = ActiveWorkbook.names("Cptv").RefersToR1C1 'tipler
If Not Cpptv = Empty Then  'eskitipler
diziP0 = "PP-DD;PP-DH;PP-DX;PP-F2;PP-F3;PP-F4;PP-SD;PP-SH;PP-SX;PP-SA;PP-KK;PP-TK"
dizi1 = ActiveWorkbook.names("Cptv").Comment
ayır1 = Split(dizi1, "-")
diziP1 = ayır1(0) & "-" & ayır1(4) & "-" & ayır1(5) & "-" & ayır1(1) & "-" & ayır1(2) & "-" & ayır1(3) _
& "-" & ayır1(6) & "-" & ayır1(7) & "-" & ayır1(8) & "-" & ayır1(9) & "-" & ayır1(10) & "-" & ayır1(11)
ActiveWorkbook.names.Add Name:="Cphtv_" & mr1, RefersToR1C1:=mr1 & " Pano Tipler"
ActiveWorkbook.names("Cphtv_" & mr1).Comment = diziP0
ActiveWorkbook.names.Add Name:="Cpptv_" & mr1, RefersToR1C1:=mr1 & " Pano Çarpanlar"
ActiveWorkbook.names("Cpptv_" & mr1).Comment = diziP1
End If
End If
End Sub
Sub mevcutmarkalarPC21() 'PC de bulunan pano markları'2021 +21 PANO CARPAN
On Error Resume Next
m = ListBoxMARKA21.ListCount - 1
For n = 0 To ListBoxMARKA2.ListCount - 1
  For i = 0 To m
  If ListBoxMARKA21.List(i, 0) = ListBoxMARKA2.List(n, 0) Then
  GoTo atla1
  End If
  Next i
  ListBoxMARKA21.AddItem ListBoxMARKA2.List(n, 0)
atla1:
Next n
End Sub
Sub teklifpanocarpan_mevcutlar21() ' pano çarpan mevcutlar
On Error Resume Next
son = dts.Range("B65536").End(xlUp).row
T1 = "A" & 2 & ":" & "A" & son
LBPC2.Clear
 For n = 0 To LBPC1.ListCount - 1
 b = LBPC1.List(n, 1) & "*"
 Set a = dts.Range(T1).Find(b, LookIn:=xlValues, LookAt:=xlPart)
 If a Is Nothing Then GoTo git:
 c = a.Address
   If marka = dts.Range("D" & a.row) Then
GoTo var:
   Else
   Do
   Set a = dts.Range(T1).FindNext(a)
   If marka = dts.Range("D" & a.row) Then GoTo var:
   Loop While c <> a.Address
   End If
GoTo git:
var:
     LBPC1.List(n, 3) = WorksheetFunction.SumIfs(dts.Range("E1:E65536"), dts.Range("A1:A65536"), b, dts.Range("D1:D65536"), marka) & " " & "Ad."
     LBPC2.AddItem LBPC1.List(n)
     LBPC2.List(LBPC2.ListCount - 1, 1) = LBPC1.List(n, 1)
     LBPC2.List(LBPC2.ListCount - 1, 2) = LBPC1.List(n, 2)
     LBPC2.List(LBPC2.ListCount - 1, 3) = LBPC1.List(n, 3)
     LBPC2.List(LBPC2.ListCount - 1, 5) = n
     
     LBPC2.List(LBPC2.ListCount - 1, 6) = LBPC1.List(n, 6)
     LBPC2.List(LBPC2.ListCount - 1, 7) = LBPC1.List(n, 7)
     LBPC2.List(LBPC2.ListCount - 1, 8) = LBPC1.List(n, 8)
     
     LBPC1.List(n, 4) = n
     LBPC1.List(n, 5) = LBPC2.ListCount - 1
git:
 Next n
 LBPC2.IntegralHeight = False: LBPC2.Height = 145: LBPC2.IntegralHeight = True
End Sub
Sub montajsarfacıklamalar() '2021 +21 PANO
On Error Resume Next
Dim Rky As String, Dosyam As String
    Dim Ert As Long, satır As Long, i As Long
    Dim ayır As Variant
    Dosyam = "C:\Belgelerim\Cemex\Ayarlar\Montaj ve Sarf\Montaj ve Sarf Açıklamalar.txt"
    Ert = FreeFile
    Open Dosyam For Input As #Ert
    If Err.Number <> 0 Then
        MsgBox "Montaj ve Sarf Açıklamalar.txt" & " Dosyası Bulunamadı !", vbCritical, "Hata !"
        Exit Sub
    End If
    On Error GoTo 0
    satır = 1
    LBPSA1.Clear
Do While Not EOF(Ert)
        Line Input #Ert, Rky
        ayır = Split(Rky, ";")
        LBPSA1.AddItem ayır(i)
tsay = Len(Rky) - Len(Replace(Rky, ";", ""))
For n = 1 To tsay '
If UBound(ayır) <> 0 Then LBPSA1.List(satır - 1, n) = ayır(n)
Next n
     satır = satır + 1
Loop
    kt = 1
Close #Ert
End Sub
Private Sub ListBoxMARKA_Click() '2021 +21 PANO
On Error Resume Next
marka = ListBoxmarka.List(ListBoxmarka.listIndex)
Call panoyapıparametre
LBKTM2 = marka & " Marka Yapısal Özellikleri"
Call panotipler1
If ListBoxPT.ListCount > 0 Then ListBoxPT.Selected(0) = True Else panodetaysil
End Sub
Private Sub ListBoxPT_Click() '2021 +21 PANO
On Error Resume Next
panotip = ListBoxPT.List(ListBoxPT.listIndex)
prgpanokod = ListBoxPT.List(ListBoxPT.listIndex, 1)
Call panodetay1
se1 = "": ptdf = 0
If ListBoxPTD1.ListCount > 0 Then se1 = ListBoxPTD11.List(0)
If Not se1 = "" Then
ListBoxPTD1.Selected(se1) = True
Else
Call panodetay1getir
FYTOK.Value = False: Call varsayılan1
Call isimyarat
FYTOK.Value = True: secim = ""
End If
End Sub
Private Sub ListBoxPTD1_Click() '2021 +21 PANO
On Error Resume Next
se1 = ListBoxPTD1.listIndex: ptdf = 0: KDform1 = "": YPOEK1 = "": KDformCP1 = 0
Call panodetay1getir
secim = ListBoxPTD1.List(ListBoxPTD1.listIndex)
Call panoparametre1
YPOEK1 = pacıklama: YPSEK1 = pacıklama2: KDform1 = ptipkod
If ptdf = 0 Then KDformCP1 = pformat
FYTOK.Value = False: Call varsayılan1
Call isimyarat
FYTOK.Value = True: secim = "": ptdf = 0
End Sub
Sub panodetay1getir() ' pano tip kodları txt '2021 +21 PANO
On Error Resume Next
If Not se1 = "" Or ListBoxPTD1.ListCount > 0 Then tipdetay = ListBoxPTD1.List(se1)
ListBoxKT2.Object = "": ListBoxKT2.Object = tipdetay
tipkod1 = ListBoxKT2.List(ListBoxKT2.listIndex, 6)
If Not tipkod1 = "" Then secim = tipkod1: ptdf = 1 Else secim = prgpanokod
Call panotipkodlar1
TBtip = ptip
'--
Call panotipcarpan1
KDprgkod = prgkod: KDtip = ptipkod: KDtipkod = ptip: KDacıklama = pacıklama: KDformat = pformat
KDtipCP = pcarpan: KDO5 = pdermesafe: KDO4CP = pdercarpan: KDMSCP = pmsharf
'FYTOK.Value = False: Call varsayılan1
End Sub
Sub panotipkodlar1() ' pano tip kodları txt '2021 +21 PANO
On Error Resume Next
Dim X As Integer
ptipkod = "": ptip = ""
'--
Dim i As Integer, m As Integer
i = 0: m = ListBoxKT1.ListCount
ListBoxKT11.Object = marka
i = ListBoxKT11.List(ListBoxKT11.listIndex, 1): m = ListBoxKT11.List(ListBoxKT11.listIndex, 2)
'--
For X = i To m
  If ListBoxKT1.List(X, 2) = secim Then
  prgkod = ListBoxKT1.List(X, 1) & "-auto": ptipkod = ListBoxKT1.List(X, 2): ptip = ListBoxKT1.List(X, 3): pacıklama = ListBoxKT1.List(X, 4)
  pformat = ListBoxKT1.List(X, 5): pcarpan = ListBoxKT1.List(X, 6): pdermesafe = ListBoxKT1.List(X, 7)
  pdercarpan = ListBoxKT1.List(X, 8): pmsharf = ListBoxKT1.List(X, 9)
  Exit For
  End If
Next X
End Sub
Sub panotipcarpan1() ' pano tip kodları ad yöneticisi carpanlar '2021 +21 PANO
On Error Resume Next
mr1 = Replace(marka, " ", "_")
mr1 = Replace(mr1, "-", "_")
Cphtv = "": Cpptv = ""
Cphtv = ActiveWorkbook.names("Cphtv_" & mr1).Comment 'tipler
Cpptv = ActiveWorkbook.names("Cpptv_" & mr1).Comment
If Cphtv = Empty Then Exit Sub
tip = Replace(prgkod, "-auto", "")
  arr1 = Split(Cphtv, ";")
  aranan = Application.Match(tip, arr1, False)
  If Not IsError(aranan) Then
  If Split(Cpptv, "-")(aranan - 1) = "" Then Exit Sub
  pcarpan = Split(Cpptv, "-")(aranan - 1)
  KDtipCP = pcarpan
  End If
End Sub
Sub panoparametre1() ' parametreler '2021 +21 PANO
On Error Resume Next
Dim X, prson As Integer
ptipkod = "": ptip = "": pacıklama = "": pacıklama2 = "": pformat = ""
For X = 0 To ListBoxKT2.ListCount - 1
  prkod = ListBoxKT2.List(X, 0)
  If prkod = secim Then
  ptipkod = ListBoxKT2.List(X, 1):  ptip = ListBoxKT2.List(X, 2): pacıklama = ListBoxKT2.List(X, 3): pacıklama2 = ListBoxKT2.List(X, 4)
  pformat = ListBoxKT2.List(X, 5)
  If Not ListBoxKT2.List(X, 7) = "" Then pmsharf = ListBoxKT2.List(X, 7)
  Exit For
  End If
Next X
End Sub
Sub panotipler1() '2021 +21 PANO
On Error Resume Next
ListBoxPT.Clear
If ListBoxKT1.ListCount < 1 Then Exit Sub
'--
Dim i As Integer, m As Integer
i = 0: m = ListBoxKT1.ListCount
ListBoxKT11.Object = ListBoxmarka.List(ListBoxmarka.listIndex)
i = ListBoxKT11.List(ListBoxKT11.listIndex, 1): m = ListBoxKT11.List(ListBoxKT11.listIndex, 2)
'--
Dim ds, pds
Set ds = CreateObject("Scripting.FileSystemObject")
For n = i To m
  If ListBoxKT1.List(n, 0) = ListBoxmarka.List(ListBoxmarka.listIndex) Then
  pds = ds.FileExists("C:\Belgelerim\Cemex\Ayarlar\Panolar\" & ListBoxKT1.List(n, 0) & "\" & ListBoxKT1.List(n, 3) & ".txt")
  If pds = True Then ListBoxPT.AddItem ListBoxKT1.List(n, 3): ListBoxPT.List(ListBoxPT.ListCount - 1, 1) = ListBoxKT1.List(n, 2)
  End If
Next n
End Sub
Sub panotipler1xxx() '2021 txt dosya 1.2. sıra no vererek listeye alma
On Error Resume Next
ListBoxPT.Clear
Dim dosya
Dim n As Integer
dosya = dir("C:\Belgelerim\Cemex\Ayarlar\Panolar\" & marka & "\*.txt")
'ListBoxPT.SetFocus
Do While dosya <> ""
dosyaad = Replace(dosya, ".txt", "")
   ListBoxPT.AddItem dosyaad
ayır = Split(dosyaad, "."): ayır2 = ayır(1)
If Not ayır2 = "" Then dosyaad2 = ayır2
   ListBoxPT.List(n, 1) = ayır2
    dosya = dir
    n = n + 1
Loop
End Sub
Sub varsayılan1() '2021 +21 PANO
On Error Resume Next
For n = 2 To 10
se1 = ListBoxPTD11.List(n - 1, 0)
If Not se1 = "" Then Controls("ListBoxPTD" & n).Selected(se1) = True
Next n
End Sub
Sub panodetaysil() '2021+
On Error Resume Next
For n = 1 To 11
Controls("ListBoxPTD" & n).Clear
Next n
End Sub
Sub panodetay1() '2021 +21 PANO
On Error Resume Next
Application.EnableEvents = False
ListBoxPTD11.Clear
Dim Rky As String, Dosyam As String
    Dim Ert As Long, satır As Long, i As Long
    Dim ayır As Variant
    Dosyam = "C:\Belgelerim\Cemex\Ayarlar\Panolar\" & marka & "\" & panotip & ".txt"
    Ert = FreeFile
    Open Dosyam For Input As #Ert

    If Err.Number <> 0 Then
        MsgBox panotip & " Dosyası Bulunamadı !", vbCritical, "Hata !"
        End
    End If
    On Error GoTo 0
    sat1 = 1
    ListBoxPTD1.Clear
Do While Not EOF(Ert)
        Line Input #Ert, Rky
Controls("ListBoxPTD" & sat1).Clear
If Not Rky = "" Then
        ayır = Split(Rky, ";")
Controls("ListBoxPTD" & sat1).AddItem ayır(0)
tsay = Len(Rky) - Len(Replace(Rky, ";", ""))
For n = 1 To tsay '
If UBound(ayır) <> 0 Then Controls("ListBoxPTD" & sat1).AddItem ayır(n)
Next n
End If
sat1 = sat1 + 1
Loop
Close #Ert
Application.EnableEvents = True
End Sub
Sub panoyapıparametre() '2021 +21 PANO
On Error Resume Next
Set ds = CreateObject("Scripting.FileSystemObject")
ydosya = "C:\Belgelerim\Cemex\Ayarlar\Panolar\" & marka & "\Pano Yapısal Parametreler.txt"
pyp = ds.FileExists(ydosya)
If pyp = True Then
Dosyam = ydosya
'ypmarka = "LOKAL"
Else
Dosyam = "C:\Belgelerim\Cemex\Ayarlar\Panolar\LOKAL\Pano Yapısal Parametreler.txt"
'ypmarka = "LOKAL"
End If
    Dim Ert As Long, satır As Long, i As Long
    Dim ayır As Variant
    Dosyam = Dosyam
    Ert = FreeFile
    Open Dosyam For Input As #Ert
    If Err.Number <> 0 Then
        MsgBox "Pano Yapısal Parametreler.txt" & " Dosyası Bulunamadı !", vbCritical, "Hata !"
        Exit Sub
    End If
    On Error GoTo 0
    satır = 1
    ListBoxKT2.Clear
Do While Not EOF(Ert)
        Line Input #Ert, Rky
        ayır = Split(Rky, ";")
        ListBoxKT2.AddItem ayır(i)
tsay = Len(Rky) - Len(Replace(Rky, ";", ""))
For n = 1 To tsay '
If UBound(ayır) <> 0 Then ListBoxKT2.List(satır - 1, n) = ayır(n)
Next n
     satır = satır + 1
Loop
    kt = 1
Close #Ert
LBKTM2 = marka & " Marka Yapısal Özellikleri"
LBKTM3 = marka
End Sub
Sub panotipparametre() '2021+
On Error Resume Next
Dim Rky As String, Dosyam As String
    Dim Ert As Long, satır As Long, i As Long
    Dim ayır As Variant
    Dosyam = "C:\Belgelerim\Cemex\Ayarlar\Panolar\" & marka & "\Pano Tip Tanımlamalar.txt"
    Ert = FreeFile
    Open Dosyam For Input As #Ert
    If Err.Number <> 0 Then
        MsgBox "Pano Tip Tanımlamalar.txt" & " Dosyası Bulunamadı !", vbCritical, "Hata !"
        Exit Sub
    End If
    On Error GoTo 0
    satır = 1
    ListBoxKT1.Clear
Do While Not EOF(Ert)
        Line Input #Ert, Rky
        ayır = Split(Rky, ";")
        ListBoxKT1.AddItem ayır(i)
        'ListBoxKT1.AddItem marka
tsay = Len(Rky) - Len(Replace(Rky, ";", ""))
For n = 1 To tsay '
 If UBound(ayır) <> 0 Then ListBoxKT1.List(satır - 1, n) = ayır(n)
Next n
     satır = satır + 1
Loop
    kt = 1
Close #Ert
End Sub
Private Sub ListBoxPTD2_Click() '2021 +21 PANO
'On Error Resume Next
'Stop
If ListBoxPTD2.Text = "Kapaksız" Then
ListBoxPTD7.Visible = True: KDgen2.Text = "": YPSEK2 = "(Kapaksız)"
If ListBoxPTD5.Visible = True Then
ListBoxPTD5.Selected(ListBoxPTD5.listIndex) = False: ListBoxPTD5.Visible = False: ListBoxPTGB.Visible = True
End If
If ListBoxPTD7.Text = "(Parça kapılı)" Then ListBoxPTD7.RemoveItem 0: ListBoxPTD7.Enabled = True
ListBoxPTD7.ForeColor = &HE6D3C4
ListBoxPTD7.AddItem "(Kapaksız)", 0: ListBoxPTD7.Selected(0) = True: ListBoxPTD7.Enabled = False
Exit Sub
End If
If Left(ListBoxPTD2.Text, 3) = "Bin" Then
ListBoxPTD5.Visible = True: ListBoxPTGB.Visible = False: KDgen2.Text = ""
If ListBoxPTD5.ListCount > 0 Then ListBoxPTD5.Selected(0) = True
If ListBoxPTD7.Text = "(Kapaksız)" Then ListBoxPTD7.ForeColor = &H0&: ListBoxPTD7.RemoveItem 0: ListBoxPTD7.Enabled = True
If ListBoxPTD7.Text = "(Parça kapılı)" Then ListBoxPTD7.ForeColor = &H0&: ListBoxPTD7.RemoveItem 0: ListBoxPTD7.Enabled = True
Exit Sub
End If
If ListBoxPTD2.Text = "Parça kapılı" Then
ListBoxPTD7.Visible = True: KDgen2.Text = "": YPSEK2 = "(Parça Kapılı)"
If ListBoxPTD5.Visible = True Then
ListBoxPTD5.Selected(ListBoxPTD5.listIndex) = False: ListBoxPTD5.Visible = False: ListBoxPTGB.Visible = True
End If
If ListBoxPTD7.Text = "(Kapaksız)" Then ListBoxPTD7.RemoveItem 0: ListBoxPTD7.Enabled = True
ListBoxPTD7.ForeColor = &HE6D3C4
ListBoxPTD7.AddItem "(Parça kapılı)", 0: ListBoxPTD7.Selected(0) = True: ListBoxPTD7.Enabled = False
Exit Sub
End If
KDgen2.Text = ""
If ListBoxPTD5.Visible = True Then
ListBoxPTD5.Selected(ListBoxPTD5.listIndex) = False: ListBoxPTD5.Visible = False: ListBoxPTGB.Visible = True
End If
If ListBoxPTD7.Text = "(Kapaksız)" Then ListBoxPTD7.ForeColor = &H0&: ListBoxPTD7.RemoveItem 0: ListBoxPTD7.Enabled = True: Exit Sub
If ListBoxPTD7.Text = "(Parça kapılı)" Then ListBoxPTD7.ForeColor = &H0&: ListBoxPTD7.RemoveItem 0: ListBoxPTD7.Enabled = True: Exit Sub
If FYTOK.Value = True Then Call isimyarat
End Sub
Private Sub ListBoxPTD3_Click() '2021 +21 PANO
On Error Resume Next
KDyük.Text = ListBoxPTD3.List(ListBoxPTD3.listIndex)
If FYTOK.Value = True Then Call isimyarat
End Sub
Private Sub ListBoxPTD4_Click() '2021 +21 PANO
On Error Resume Next
KDgen1.Text = ListBoxPTD4.List(ListBoxPTD4.listIndex)
If FYTOK.Value = True Then Call isimyarat
End Sub
Private Sub ListBoxPTD5_Click() '2021 +21 PANO
On Error Resume Next
KDgen2.Text = ListBoxPTD5.List(ListBoxPTD5.listIndex)
If FYTOK.Value = True Then Call isimyarat
End Sub
Private Sub ListBoxPTGB_Click() '2021 +21 PANO
KDyük = ListBoxPTD2.Text
If FYTOK.Value = True Then Call isimyarat
End Sub
Private Sub ListBoxPTD6_Click() '2021 +21 PANO
On Error Resume Next
KDder.Text = ListBoxPTD6.List(ListBoxPTD6.listIndex)
If FYTOK.Value = True Then Call isimyarat
End Sub
Private Sub ListBoxPTD7_Click() '2021 +21 PANO
On Error Resume Next
TBPDT7 = ListBoxPTD7.listIndex: If TBPDT7 < 0 Then TBPDT7 = 0
secim = ListBoxPTD7.List(ListBoxPTD7.listIndex)
Call panoparametre1
YPOEK2 = pacıklama: YPSEK2 = pacıklama2
KDok1 = ptipkod: YPOEK2 = pacıklama: KDokCP1 = pformat
If FYTOK.Value = True Then Call isimyarat
End Sub
Private Sub ListBoxPTD8_Click() '2021 +21 PANO
On Error Resume Next
secim = ListBoxPTD8.List(ListBoxPTD8.listIndex)
Call panoparametre1
YPOEK3 = pacıklama: YPSEK3 = pacıklama2
KDiyp1 = ptipkod: YPOEK3 = pacıklama: KDiypCP1 = pformat
If FYTOK.Value = True Then Call isimyarat
End Sub
Private Sub ListBoxPTD9_Click() '2021 +21 PANO
On Error Resume Next
secim = ListBoxPTD9.List(ListBoxPTD9.listIndex)
Call panoparametre1
YPOEK4 = pacıklama: YPSEK4 = pacıklama2
KDdyp1 = ptipkod: KDdyp2 = ptip: KDdypCP1 = pformat
If FYTOK.Value = True Then Call isimyarat
End Sub
Private Sub ListBoxPTD101_Click()
On Error Resume Next
ListBoxPTD121.Clear
If ListBoxPTD101.ListCount = 0 Then TBPMS01 = "": Exit Sub
s = 0
For n = 0 To LBPSA1.ListCount - 1
 If ListBoxPTD101.listIndex >= 0 Then
  If ListBoxPTD101.List(ListBoxPTD101.listIndex) = LBPSA1.List(n, 2) Then
     ListBoxPTD121.AddItem LBPSA1.List(n, 0)
     ListBoxPTD121.List(ListBoxPTD121.ListCount - 1, 1) = "Yoğunluk " & LBPSA1.List(n, 1)
     If Not LBPSA1.List(n, 3) = "" Then s = ListBoxPTD121.ListCount - 1
  End If
 End If
Next n
If ListBoxPTD121.ListCount >= 0 Then
 If s = 0 Then
 ListBoxPTD121.Selected(ListBoxPTD121.ListCount - 1) = True
 Else
 ListBoxPTD121.Selected(s) = True
 End If
Else
If ListBoxPTD121.listIndex >= 0 Then TBPMS01 = ListBoxPTD121.List(ListBoxPTD121.listIndex) & TBPM09 Else TBPMS01 = TBPM09
TBPMS01 = "": TBAMS11 = ""
End If
End Sub
Private Sub ListBoxPTD12_Click() '2021 +21 PANO
On Error GoTo hata
TBAMS1 = "": LBAMS1 = ""
LBPSA1.Object = "": LBPSA1.Object = ListBoxPTD12.List(ListBoxPTD12.listIndex): TBAMS1 = ListBoxPTD12.List(ListBoxPTD12.listIndex)
If Not LBPSA1.List(LBPSA1.listIndex, 1) = "" Then
yo1 = LBPSA1.List(LBPSA1.listIndex, 1)
If Right(yo1, 1) = "%" Then yo1 = Replace(yo1, "%", "") & "/100"
ayır1 = Split(yo1, "/")(0): ayır2 = Split(yo1, "/")(1)
LBLPSA4.BackColor = &HBFBFFF
LBLPSA4.Height = 0: LBLPSA4.Height = 142 * CDbl(ayır1) / CDbl(ayır2)
LBLPSA4.Top = 142 - LBLPSA4.Height
End If
If LBPSA01.ListCount > 0 Then
LBPSA01.Object = "": LBPSA01.Object = ListBoxPTD12.List(ListBoxPTD12.listIndex)
mcp = "": scp = "": mcp = LBPSA01.List(LBPSA01.listIndex, 1): scp = LBPSA01.List(LBPSA01.listIndex, 2)
If Not mcp = "" Then LBAMS1 = "M(x)" & mcp & " - " & "S(x)" & scp
End If
Exit Sub
hata:
End Sub
Sub isimyarat() '+1 isim oluşturma '2021 +21 PANO
On Error Resume Next
If KDtip.Text = "XX" Then MsgBox (" Panel seçiniz "), vbInformation, "scngnr@hotmail.com": Exit Sub
gen = KDgen1.Text: gent = gen: genta = gen
yük = KDyük.Text
der = KDder.Text
If ListBoxPTD5.Visible = True Then
gen = CDbl(KDgen1) + CDbl(KDgen2): gent = Left(KDgen1, 1) & Left(KDgen2, 1): genta = KDgen1 & "+" & KDgen2
End If
'--
KDalan = Format(yük * gen / 1000000, "#,##0.00")
'### Pano referans kodu
frm = ListBoxPTD1.listIndex
TBref.Text = prgkod
'TBref.Text = "PP-" & Mid(KDtip.Text, 1, 2) & "-auto"
'If frm < 1 Then TBref.Text = "PP-" & Mid(KDtip.Text, 1, 2) & "-auto"
'If frm > 0 Then TBref.Text = "PP-" & Mid(ListBoxPTD1.Text, 1, 1) & Mid(ListBoxPTD1.Text, 6, 1) & "-auto"
'--
'### Pano sip. kodu
If KDformat = "YGD1" Then
If Len(gent) = 3 Then gent = "0" & Left(gent, 1) Else If Len(gent) = 4 Then gent = Left(gent, 2)
If Len(yük) = 3 Then yük = "0" & Left(yük, 1) Else If Len(yük) = 4 Then yük = Left(yük, 3)
If Len(der) = 3 Then der = "0" & Left(der, 1) Else If Len(der) = 4 Then der = Left(der, 2)
TBskd.Text = prgpanokod & yük & gent & der & KDok1 & KDdyp1 & KDiyp1 & KDdyp2 & KDform1.Text
KDO1 = yük: KDO2 = gen: KDO4 = der
GoTo git
End If
If Len(yük) = 3 Then yük = "0" & Left(yük, 2) Else If Len(yük) = 4 Then yük = Left(yük, 3)
If KDformat = "YGD2" Then
TBskd.Text = prgpanokod & yük & Left(gent, 2) & Left(der, 2) & KDok1 & KDdyp1 & KDiyp1 & KDdyp2 & KDform1.Text
KDO1 = yük / 10: KDO2 = gen / 10: KDO4 = der / 10
GoTo git
End If
If KDformat = "YG1" Then
If Len(gent) = 3 Then gent = "0" & Left(gent, 1) Else If Len(gent) = 4 Then gent = Left(gent, 2)
TBskd.Text = prgpanokod & yük & gent & KDok1 & KDdyp1 & KDiyp1 & KDdyp2 & KDform1.Text
KDO1 = yük: KDO2 = gen
GoTo git
End If
git:
'### Pano açıklama
oek = WorksheetFunction.Trim(YPOEK1 & " " & YPOEK2 & " " & YPOEK3 & " " & YPOEK4 & " ")
sek = WorksheetFunction.Trim(" " & YPSEK1 & " " & YPSEK2 & " " & YPSEK3 & " " & YPSEK4)
If oek <> "" Then oek = " " & oek
If sek <> "" Then sek = " " & sek
TByic = TBtip.Text & oek & " Y" & KDyük.Text & "xG" & genta & "xD" & KDder.Text & " " & KDacıklama & sek
'--
Call hesapla
Labelm2.Caption = "Yük." & "xGen." & "= " & KDalan & " m2" & " (x" & KDtipCP & ")"
'Labelm2.Caption = " Yük." & yük & " x Gen." & genta & " = " & KDalan & " m2"
End Sub
Sub hesapla() '2021 +21 PANO
On Error Resume Next
If KDformCP1 = "" Then KDformCP1 = 0
TextBoxCT1 = Format((CDbl(KDalan) * CDbl(KDtipCP) * 10), "#,##0.00")
TextBoxCT2 = Format((TextBoxCT1 * KDformCP1) / 100, "#,##0.00")
TextBoxCT3 = 0
TextBoxCT4 = Format((TextBoxCT1 * KDokCP1) / 100, "#,##0.00")
TextBoxCT5 = Format((TextBoxCT1 * KDiypCP1) / 100, "#,##0.00")
TextBoxCT6 = Format((TextBoxCT1 * KDdypCP1) / 100, "#,##0.00")
If CDbl(KDder) > CDbl(KDO5) Then
TextBoxCT3 = Format(((CDbl(KDder) - CDbl(KDO5)) / 100) * ((TextBoxCT1 * KDO4CP) / 100), "#,##0.00") 'derinlik faktörü ilave
DYC = KDO4CP
End If
'--
tfy = CDbl(TextBoxCT1) + CDbl(TextBoxCT2) + CDbl(TextBoxCT3) + CDbl(TextBoxCT4) + CDbl(TextBoxCT5) + CDbl(TextBoxCT6)
tfy = Application.WorksheetFunction.RoundUp(tfy, -1)
TextBoxFY = Format(tfy, "#,##0.00")
Label474 = "Pano Çarp.(x):" & KDtipCP & "+%(" & DYC & "+" & KDformCP1 & "+" & KDokCP1 & "+" & KDiypCP1 & "+" & KDdypCP1 & ")"
'Label475 = "+%(" & DYC & "+" & KDformCP1 & "+" & KDokCP1 & "+" & KDiypCP1 & "+" & KDdypCP1 & ")"
ptipkod = "": ptip = "": pacıklama = "": pformat = "": pcarpan = "": pdermesafe = "": pdercarpan = ""
End Sub
Private Sub CommandButton4_Click() '2021 +21 PANO
On Error Resume Next
Call PanoCarpanlar
Call gir1
Call CarpanlarPCTV
End Sub
Sub pano_gir1xxx() '2021 +21 PANO
On Error Resume Next
Application.ScreenUpdating = False: Application.Calculation = xlCalculationManual
son = dts.Range("B65536").End(xlUp).row + 1
Dim pkod As String
Dim fkat As Integer
Dim fmt As String
Dim drm As Integer
Dim drc As Integer
'--
For i = 2 To son
If Left(dts.Range("A" & i), 3) = "PP-" And Right(dts.Range("A" & i), 5) = "-auto" And dts.Range("D" & i) = marka Then
'katsayılar--
fkat = 0: drm = 0: drc = 0

pkod = Replace(dts.Range("A" & i), "-auto", "")
 For n = 0 To Controls("LBPC" & TBPLIST.Value).ListCount - 1
 If Controls("LBPC" & TBPLIST.Value).List(n, 0) = pkod Then
 fkat = Controls("LBPC" & TBPLIST.Value).List(n, 2)
 'fmt = Controls("LBPC" & TBPLIST.Value).List(n, 6)'format
 drm = Controls("LBPC" & TBPLIST.Value).List(n, 7)
 drc = Controls("LBPC" & TBPLIST.Value).List(n, 8)
 Exit For
 End If
 Next n
'--
If fkat = 0 Then GoTo devam2
'ölçüler-alan bulma B HÜCRESİNDEN
 For sy = 1 To Len(dts.Range("B" & i)) 'sayı başlangıç
 If IsNumeric(Mid(dts.Range("B" & i), sy, 1)) = True Then Exit For
 Next
 For sys = sy To Len(dts.Range("B" & i)) 'sayı bitiş
 If Not IsNumeric(Mid(dts.Range("B" & i), sys, 1)) = True Then Exit For
 Next
byt = Mid(dts.Range("B" & i), sy, sys - sy) 'ebat
'sayıdan format bulma B HÜCRESİNDEN
If Len(byt) = 5 Then fmt = "YG1"
If Len(byt) = 6 Then fmt = "YGD1"
If Len(byt) = 7 Then fmt = "YGD2"
'--
 If fmt = "YGD1" Then
 ge = Mid(byt, 3, 2)
 If CDbl(Mid(byt, 4, 1)) = 1 Then
  ge = CDbl(Mid(byt, 4, 1)) & CDbl(Mid(byt, 3, 1))
  If CDbl(Mid(byt, 3, 1)) = 1 And CDbl(Mid(byt, 4, 1)) = 1 Then ge = CDbl(Mid(byt, 4, 1)) * 10 + CDbl(Mid(byt, 3, 1)) * 10
 Else
 If CDbl(Mid(byt, 3, 1)) > 1 Then ge = CDbl(Mid(byt, 3, 1)) + CDbl(Mid(byt, 4, 1))
 End If
 yu = CDbl(Left(byt, 2)): ge = CDbl(ge) * 10: de = CDbl(Right(byt, 2)) * 100
 GoTo devam1
 End If
'--
 If fmt = "YGD2" Then
 ge = Mid(byt, 4, 2)
  If CDbl(Mid(byt, 5, 1)) = 0 Or CDbl(Mid(byt, 4, 1)) = 0 Then
  ge = CDbl(ge & "0")
  Else
    If CDbl(Mid(byt, 5, 1)) = 1 Then g1 = CDbl(Mid(byt, 5, 1) * 10) Else g1 = CDbl(Mid(byt, 5, 1))
    If CDbl(Mid(byt, 4, 1)) = 1 Then g2 = CDbl(Mid(byt, 4, 1) * 10) Else g2 = CDbl(Mid(byt, 4, 1))
     ge = g1 * 10 + g2 * 10
  End If
 yu = CDbl(Left(byt, 3)) / 10: ge = CDbl(ge): de = CDbl(Mid(byt, 6, 2) * 100)
 GoTo devam1
 End If
 
 
 If fmt = "YG1" Then
 ge = Mid(byt, 4, 2)
 If CDbl(Mid(byt, 4, 1)) > 1 Then ge = CDbl(Mid(byt, 4, 1)) + CDbl(Mid(byt, 5, 1))
 yu = CDbl(Left(byt, 3)) / 10: ge = CDbl(ge) * 10: de = 1
 End If
'--
devam1:
alan = yu * ge / 10 'pano alan
fy = alan * fkat / 10 'fiyat
fyd = 0
If de > drm Then fyd = ((de - drm) / 100) * (fy * drc / 100) 'derinlik faktörü ilave fiyat
'GoTo yok1:
'ilave özellikleri fiyata ekleme
ikod = Mid(dts.Range("B" & i), sys, Len(dts.Range("B" & i))) 'ek özellikler
fyek = 0
For n = 0 To ListBoxKT2.ListCount - 1
If ListBoxKT2.List(n, 6) <> "" Then ikod = Replace(ikod, ListBoxKT2.List(n, 1), "")
 If ikod = "" Then GoTo devam2
   If ListBoxKT2.List(n, 1) & ListBoxKT2.List(n, 2) <> "" Then
   ipkod = "*" & ListBoxKT2.List(n, 1) & "*" & ListBoxKT2.List(n, 2)
    If ikod Like ipkod Then
    fyek = fyek + (ListBoxKT2.List(n, 5) * fy / 100)
    ikod = Replace(ikod, ListBoxKT2.List(n, 1), ""): ikod = Replace(ikod, ListBoxKT2.List(n, 2), "")
    End If
  End If
Next n
devam2:
tfy = fy + fyd + fyek 'tüm özellikler eklenmiş fiyat
tfy = Application.WorksheetFunction.RoundUp(tfy, -1)
dts.Range("F" & i) = tfy
If dts.Range("I" & i) = "" Then dts.Range("I" & i) = alan
'--
End If
Next i
Range("B2").Select
Application.Calculation = xlCalculationAutomatic: Application.ScreenUpdating = True
End Sub
Sub pano_gir1() '2022 +21 PANO
On Error Resume Next
Application.ScreenUpdating = False: Application.Calculation = xlCalculationManual
son = dts.Range("B65536").End(xlUp).row + 1
Dim pkod As String
Dim fkat As Integer
Dim fmt As String
Dim drm As Integer
Dim drc As Integer
'--
For i = 2 To son
If Left(dts.Range("A" & i), 3) = "PP-" And Right(dts.Range("A" & i), 5) = "-auto" And dts.Range("D" & i) = marka Then
'katsayılar--
fkat = 0: drm = 0: drc = 0

pkod = Replace(dts.Range("A" & i), "-auto", "")
 For n = 0 To Controls("LBPC" & TBPLIST.Value).ListCount - 1
 If Controls("LBPC" & TBPLIST.Value).List(n, 0) = pkod Then
 fkat = Controls("LBPC" & TBPLIST.Value).List(n, 2)
 'fmt = Controls("LBPC" & TBPLIST.Value).List(n, 6)'format
 drm = Controls("LBPC" & TBPLIST.Value).List(n, 7)
 drc = Controls("LBPC" & TBPLIST.Value).List(n, 8)
 Exit For
 End If
 Next n
'--
If fkat = 0 Then GoTo devam2
'ölçüler-alan bulma B HÜCRESİNDEN
 For sy = 1 To Len(dts.Range("B" & i)) 'sayı başlangıç
 If IsNumeric(Mid(dts.Range("B" & i), sy, 1)) = True Then Exit For
 Next
 For sys = sy To Len(dts.Range("B" & i)) 'sayı bitiş
 If Not IsNumeric(Mid(dts.Range("B" & i), sys, 1)) = True Then Exit For
 Next
byt = Mid(dts.Range("B" & i), sy, sys - sy) 'ebat
'sayıdan format bulma B HÜCRESİNDEN
If Len(byt) = 5 Then fmt = "YG1"
If Len(byt) = 6 Then fmt = "YGD1"
If Len(byt) = 7 Then fmt = "YGD2"
'--
 If fmt = "YGD1" Then
 de = CDbl(Right(byt, 2)) * 100
 GoTo devam1
 End If
'--
 If fmt = "YGD2" Then
 de = CDbl(Mid(byt, 6, 2) * 100)
 GoTo devam1
 End If
'--
 If fmt = "YG1" Then
 de = 1
 End If
'--
devam1:
'sayı başlangıç--
deg = ""
For sy = 1 To Len(dts.Range("I" & i))
If IsNumeric(Mid(dts.Range("I" & i), sy, 1)) = True Then deg = deg & Mid(dts.Range("I" & i), sy, 1)
Next
If deg = "" Then
alan = 0
Else
alan = CDbl(deg) 'pano alan
End If
fy = alan * fkat / 10 'fiyat
fyd = 0
If de > drm Then fyd = ((de - drm) / 100) * (fy * drc / 100) 'derinlik faktörü ilave fiyat
'GoTo yok1:
'ilave özellikleri fiyata ekleme
ikod = Mid(dts.Range("B" & i), sys, Len(dts.Range("B" & i))) 'ek özellikler
fyek = 0
For n = 0 To ListBoxKT2.ListCount - 1
If ListBoxKT2.List(n, 6) <> "" Then ikod = Replace(ikod, ListBoxKT2.List(n, 1), "")
 If ikod = "" Then GoTo devam2
   If ListBoxKT2.List(n, 1) & ListBoxKT2.List(n, 2) <> "" Then
   ipkod = "*" & ListBoxKT2.List(n, 1) & "*" & ListBoxKT2.List(n, 2)
    If ikod Like ipkod Then
    fyek = fyek + (ListBoxKT2.List(n, 5) * fy / 100)
    ikod = Replace(ikod, ListBoxKT2.List(n, 1), ""): ikod = Replace(ikod, ListBoxKT2.List(n, 2), "")
    End If
  End If
Next n
devam2:
tfy = fy + fyd + fyek 'tüm özellikler eklenmiş fiyat
tfy = Application.WorksheetFunction.RoundUp(tfy, -1)
dts.Range("F" & i) = tfy
If dts.Range("I" & i) = "" Then dts.Range("I" & i) = alan
'--
End If
Next i
Range("B2").Select
Application.Calculation = xlCalculationAutomatic: Application.ScreenUpdating = True
End Sub
Sub gir1() '2021 +21 PANO
On Error Resume Next
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
''''
If Not ActiveSheet.Name = "Sayfa1" Then MsgBox (" Teklif sayfasına geçiniz "), vbInformation, "scngnr@hotmail.com": Exit Sub
Dim a
a = Selection.row
If a = 1 Then a = 2
Cells(a, 2).Range("A1").Select
If Not Cells(a, 2) = "" Then Selection.EntireRow.Insert
'Genel Biçimlemeler'--
Range("A" & a & ":U" & a).RowHeight = 12.75
Range("A" & a & ":U" & a).Borders.LineStyle = xlContinuous
Range("W" & a & ":X" & a).Borders.LineStyle = xlContinuous
Range("A" & a & ":X" & a).Interior.Pattern = xlNone
Range("A" & a & ":X" & a).Font.Bold = False
Range("A" & a & ":X" & a).Font.ColorIndex = xlAutomatic
Range("A" & a & ":X" & a).Font.Size = 9
Range("A" & a & ":D" & a).HorizontalAlignment = xlLeft
Range("E" & a & ":X" & a).HorizontalAlignment = xlRight
Range("A" & a & ":D" & a).NumberFormat = "@"
Range("E" & a).NumberFormat = "#,##0"
Range("F" & a).NumberFormat = "#,##0.00"
kur = ""
If TextBoxFY11B1 = "$" Then Range("F" & a).NumberFormat = "#,##0.00 [$$-C0C]": Range("F" & a).Font.ColorIndex = 3: kur = "*Usd" '$ kuru
If TextBoxFY11B1 = "€" Then Range("F" & a).NumberFormat = "#,##0.00 [$€-1]": Range("F" & a).Font.ColorIndex = 5: kur = "*Eur" ' € kuru
Range("G" & a).NumberFormat = "0.0%"
Range("H" & a).NumberFormat = "#,##0"
Range("J" & a & ":X" & a).NumberFormat = "#,##0.00"
''''
If pfc = 0 Then
Range("A" & a).FormulaR1C1 = TBref01
Range("B" & a).FormulaR1C1 = TBskd01 'Sipariş Kd.
Range("C" & a).FormulaR1C1 = TByic01 ' Yapılacak İşin Cinsi
Range("D" & a).FormulaR1C1 = TBPM01 'Üretici
Range("E" & a).FormulaR1C1 = CDbl(TBmiktar01.Value) 'Miktar
Range("F" & a).FormulaR1C1 = CDbl(TextBoxFY01) 'Mlz.Br.Fiyat
Range("G" & a).Value = TBPM10 / 100
If Not TBPM11 = "" Then Range("H" & a).FormulaR1C1 = CDbl(TBPM11)
Range("I" & a).FormulaR1C1 = TBPMS01
Else
Range("A" & a).FormulaR1C1 = TBref 'Referans
Range("B" & a).FormulaR1C1 = TBskd 'Sipariş Kd.
Range("C" & a).FormulaR1C1 = TByic ' Yapılacak İşin Cinsi
Range("D" & a).FormulaR1C1 = ListBoxmarka.List(ListBoxmarka.listIndex) 'Üretici
'Range("D" & a).FormulaR1C1 = GetSetting("ilhan", "Settings", "mdip") 'Üretici
Range("E" & a).FormulaR1C1 = TBmiktar.Value * 1 'Miktar
Range("F" & a).FormulaR1C1 = Application.WorksheetFunction.RoundUp(TextBoxFY.Value * 1, -1) 'Mlz.Br.Fiyat
Range("G" & a).Value = 0
 If ListBoxPTD12.listIndex >= 0 Then
 Range("I" & a).FormulaR1C1 = ListBoxPTD12.List(ListBoxPTD12.listIndex) & (KDalan * 100)
 Else
 Range("I" & a).FormulaR1C1 = "A" & (KDalan * 100)
 End If
End If
''''
If Left(ActiveSheet.CodeName, 3) = "OTM" Then GoTo zıpla1:
mkur = kur: If bfyt = "=RC[-1]" Then mkur = ""
Range("L" & a).FormulaR1C1 = bfyt & "*Opano/100" & mkur ' Kar
''''
Range("J" & a).FormulaR1C1 = "=RC[-2]*Ads/60" 'Montaj Br.Fyt* işçilik katsayılı+1
Range("N" & a).FormulaR1C1 = "=RC[-3]*Oggid/100" 'GENEL GİDERLER+1
Range("K" & a).FormulaR1C1 = "=(RC[-5]-RC[-5]*RC[-4])" & kur 'Net Mlz. Alış+1
Range("M" & a).FormulaR1C1 = "=RC[-3]*Oisci/100" 'Mont. Kar rev1+1
Range("O" & a).FormulaR1C1 = "=RC[-10]*RC[-9]" & kur 'Mlz. List Top.+1
Range("P" & a).FormulaR1C1 = "=RC[-11]*RC[-5]" 'Mlz. Net Top.+1
Range("Q" & a).FormulaR1C1 = "=RC[-12]*RC[-7]" 'Montaj.Top.+1
Range("R" & a).FormulaR1C1 = "=RC[-13]*RC[-6]" 'Mlz.KarTp.+1
Range("S" & a).FormulaR1C1 = "=RC[-14]*RC[-6]" 'Mont.KarTop.+1
Range("T" & a).FormulaR1C1 = "=RC[-15]*RC[-12]/60" 'Tp.Ad/h.
Range("U" & a).FormulaR1C1 = "=RC[-7]*RC[-16]" 'Top. Gn.Gd+1
'TOPLAMLAR'--
Range("W" & a).FormulaR1C1 = "=(RC[-13]+RC[-12]+RC[-11]+RC[-10]+RC[-9])*Dcki/Tpb"
Range("X" & a).FormulaR1C1 = "=RC[-19]*RC[-1]"
'TOPLAM FORMAT'--
If Range("Tpbr") = "Teklif Para Birimi (TL)" Then Range("W" & a & ":X" & a).NumberFormat = "#,##0.00"
If Range("Tpbr") = "Teklif Para Birimi (EUR)" Then Range("W" & a & ":X" & a).NumberFormat = "#,##0.00 [$€-1]"
If Range("Tpbr") = "Teklif Para Birimi (USD)" Then Range("W" & a & ":X" & a).NumberFormat = "#,##0.00 [$$-C0C]"
'--
zıpla1:
Range("B" & a + 1).Select
Call AraToplamlar
Application.Calculation = xlCalculationAutomatic
Application.ScreenUpdating = True
End Sub
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer) 'TAMAM
On Error Resume Next
Application.ScreenUpdating = False
    Windows(pmlz).Close False ':Application.Windows(mlz).Visible = True
    pmlz = Empty
Application.ScreenUpdating = True
If prs = 1 Then Unload UFOPAN00P1
If pfc = 1 Then End
End Sub
'+++++++++++++++++++++++++++++++++++++++++++++
Private Sub SBPST1_SpinUp() '2021+
On Error Resume Next
Sheets("Sayfa1").Select '
If SBPST1.Value = SBPST1.Max Then SBPST1.Value = 0: i = 2 Else i = Selection.row + 1
ayır = Split(TBPST1.Text, ".")
If CDbl(ayır(SBPST1.Value)) > i Then
For n = SBPST1.Value To 0 Step -1
   If CDbl(ayır(n)) < i Then SBPST1.Value = n + 1: Exit For
Next
Else
For n = SBPST1.Value To SBPST1.Max - 1
   If CDbl(ayır(n)) > i - 1 Then SBPST1.Value = n: Exit For
Next
End If
If n < 0 Then SBPST1.Value = 0
Range("A" & ayır(SBPST1.Value) & ":I" & ayır(SBPST1.Value)).Select
LBPST1 = Selection.row
End Sub
Private Sub SBPST1_SpinDown() '2021 MONTAJ
On Error Resume Next
Sheets("Sayfa1").Select '
i = Selection.row - 1
ayır = Split(TBPST1.Text, ".")
If CDbl(ayır(SBPST1.Value)) > i Then
For n = SBPST1.Value To 0 Step -1
   If CDbl(ayır(n)) < i Then SBPST1.Value = n - 1: Exit For
Next
Else
For n = SBPST1.Value To SBPST1.Max - 1
   If CDbl(ayır(n)) > i Then SBPST1.Value = n - 1: Exit For
Next
End If
If n < 0 Then SBPST1.Value = 0
Range("A" & ayır(SBPST1.Value) & ":I" & ayır(SBPST1.Value)).Select
LBPST1 = Selection.row
End Sub
Private Sub CBPSC1_Click() '2021 MONTAJ
If LBMSC2.listIndex >= 0 And LBPST1 <> "" Then
CBPSC1.BackColor = &HCAE3BF
UFOPAN01.Show
End If
End Sub
Private Sub SPMC1_Change() '2021+ MONTAJ
On Error Resume Next: TBMC1.Text = SPMC1.Value
End Sub
Private Sub SPSC1_Change() '2021+ MONTAJ
On Error Resume Next: TBSC1.Text = SPSC1.Value
End Sub
Private Sub TBMC1_Change() '2021+ MONTAJ
On Error Resume Next
SPMC1.Value = TBMC1.Text
If TBMSLIST = 2 Then
LBMSC2.List(LBMSC2.listIndex, 1) = TBMC1.Text
LBMSC1.List(LBMSC2.List(LBMSC2.listIndex, 5), 1) = TBMC1.Text
Else
LBMSC1.List(LBMSC1.listIndex, 1) = TBMC1.Text
LBMSC2.List(LBMSC1.List(LBMSC1.listIndex, 5), 1) = TBMC1.Text
End If
End Sub
Private Sub LBMSC1_DblClick(ByVal Cancel As MSForms.ReturnBoolean) '2021
On Error Resume Next
If LBMSC1.listIndex < 1 Then Exit Sub
msg = MsgBox(LBMSC1.List(LBMSC1.listIndex) & " - montaj & sarf çarpanlanını silmek istediğinizden emin misiniz?", vbYesNo + vbInformation + vbDefaultButton2, "Silme İşlemi")
If msg = vbNo Then Exit Sub
LBMSC1.RemoveItem (LBMSC1.listIndex)
Call CarpanlarCMSTV
End Sub
Private Sub LBMSC1_Click() '2021+ MONTAJ
On Error Resume Next
If LBMSC1.listIndex >= 0 Then FrameR.Enabled = True Else FrameR.Enabled = False
If LBMSC1.BackColor = &H80000004 Then Call teklifmontajliste1: Exit Sub
LBMSC2.Selected(LBMSC2.listIndex) = False
TBMSLIST = 1
teklifmontajliste1
End Sub
Sub teklifmontajliste1()
On Error Resume Next
If LBMSC1.listIndex < 0 Then Exit Sub
TBMH1.Text = "": TBMH1.Text = LBMSC1.List(LBMSC1.listIndex): TBMH11.Text = TBMH1.Text
'--
LBPSA1.Object = "": LBPSA1.Object = TBMH1.Text
LBLPSA1 = "": LBLPSA1 = LBPSA1.List(LBPSA1.listIndex, 2): If LBLPSA1 = "" Then LBLPSA1 = "TANIMSIZ"
LBLPSA3 = "": LBLPSA3 = "Yoğunluk " & LBPSA1.List(LBPSA1.listIndex, 1)
'--montaj seviyesi
If Len(TBMH11) > 0 Then
yo1 = LBPSA1.List(LBPSA1.listIndex, 1)
If Right(yo1, 1) = "%" Then yo1 = Replace(yo1, "%", "") & "/100"
ayır = Split(yo1, "/")
LBLPSA2.Height = 0: LBLPSA2.Height = 145 * CDbl(ayır(0)) / CDbl(ayır(1))
LBLPSA2.BackColor = &HD5D5FF
LBLPSA2.Top = 145 - LBLPSA2.Height
If LBLPSA2.Height > 99 Then LBLPSA2.BackColor = &HAAAAFF: GoTo git1
If LBLPSA2.Height > 66 Then LBLPSA2.BackColor = &HBFBFFF: GoTo git1
If LBLPSA2.Height > 33 Then LBLPSA2.BackColor = &HD5D5FF: GoTo git1
End If
'--
git1:
'--
If Not LBMSC1.List(LBMSC1.listIndex, 1) = "" Then TBMC1.Text = LBMSC1.List(LBMSC1.listIndex, 1) Else TBMC1.Text = ""
If Not LBMSC1.List(LBMSC1.listIndex, 2) = "" Then TBSC1.Text = LBMSC1.List(LBMSC1.listIndex, 2) Else TBSC1.Text = ""
TBPST1.Text = LBMSC1.List(LBMSC1.listIndex, 3)
SBPST1.Max = Len(TBPST1) - Len(Replace(TBPST1, ".", "")): SBPST1.Value = 0
End Sub
Private Sub LBMSC2_Click() '2021+
On Error Resume Next
If LBMSC2.listIndex >= 0 Then FrameR.Enabled = True Else FrameR.Enabled = False
LBMSC1.Selected(LBMSC1.listIndex) = False
TBMSLIST = 2
LBPST1 = ""
CBPSC1.BackColor = &HD7BBA2
teklifmontajliste2
End Sub
Private Sub LBMSC2_DblClick(ByVal Cancel As MSForms.ReturnBoolean) '2021+
If LBMSC2.listIndex < 0 Then Exit Sub
UFOPAN01.Show
End Sub
Sub teklifmontajliste2() '2021
On Error Resume Next
If LBMSC2.listIndex < 0 Then Exit Sub
TBMH1.Text = "": TBMH1.Text = LBMSC2.List(LBMSC2.listIndex): TBMH11.Text = TBMH1.Text
'--
LBPSA1.Object = "": LBPSA1.Object = TBMH1.Text
LBLPSA1 = "": LBLPSA1 = LBPSA1.List(LBPSA1.listIndex, 2): If LBLPSA1 = "" Then LBLPSA1 = "TANIMSIZ"
LBLPSA3 = "": LBLPSA3 = "Yoğunluk " & LBPSA1.List(LBPSA1.listIndex, 1)
'--montaj seviyesi
If Len(TBMH11) > 0 Then
yo1 = LBPSA1.List(LBPSA1.listIndex, 1)
If Right(yo1, 1) = "%" Then yo1 = Replace(yo1, "%", "") & "/100"
ayır = Split(yo1, "/")
LBLPSA2.Height = 0: LBLPSA2.Height = 145 * CDbl(ayır(0)) / CDbl(ayır(1))
LBLPSA2.BackColor = &HD5D5FF
LBLPSA2.Top = 145 - LBLPSA2.Height
If LBLPSA2.Height > 99 Then LBLPSA2.BackColor = &HAAAAFF: GoTo git1
If LBLPSA2.Height > 66 Then LBLPSA2.BackColor = &HBFBFFF: GoTo git1
If LBLPSA2.Height > 33 Then LBLPSA2.BackColor = &HD5D5FF: GoTo git1
End If
git1:
'--
If Not LBMSC2.List(LBMSC2.listIndex, 1) = "" Then TBMC1.Text = LBMSC2.List(LBMSC2.listIndex, 1) Else TBMC1.Text = ""
If Not LBMSC2.List(LBMSC2.listIndex, 2) = "" Then TBSC1.Text = LBMSC2.List(LBMSC2.listIndex, 2) Else TBSC1.Text = ""
TBPST1.Text = LBMSC2.List(LBMSC2.listIndex, 3)
SBPST1.Max = Len(TBPST1) - Len(Replace(TBPST1, ".", "")): SBPST1.Value = 0
End Sub
Private Sub Image442_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single) '2021 +21 PANO MONTAJ
On Error Resume Next
CreateObject("Shell.Application").Open "C:\Belgelerim\Cemex\Ayarlar\Montaj ve Sarf\Montaj ve Sarf Açıklamalar.txt"
End Sub
Private Sub Image443_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single) '2021 +21 PANO CARPAN
Dim rds
rds = CreateObject("Scripting.FileSystemObject").FolderExists("C:\Belgelerim\Cemex\Ayarlar\Montaj ve Sarf")
If rds = True Then
CreateObject("Shell.Application").Open "C:\Belgelerim\Cemex\Ayarlar\Montaj ve Sarf"
Else
MsgBox "C:\Belgelerim\Cemex\Ayarlar\Montaj ve Sarf" & " klasörü yok", vbInformation, "scngnr@hotmail.com"
End If
End Sub
Private Sub CBMSC2_Click() '2021 +21 MONTAJ SARF CARPAN
On Error Resume Next
If CBMSC2.Value = True Then
LBMSC2.Visible = False: LBMSC1.Visible = True
LBMSC1.IntegralHeight = False: LBMSC1.Height = 145: LBMSC1.IntegralHeight = True
Call montajcarpandetay
Else
LBMSC2.Visible = True: LBMSC1.Visible = False
LBMSC2.IntegralHeight = False: LBMSC2.Height = 145: LBMSC2.IntegralHeight = True
End If
End Sub
Sub montajcarpandetay() ' mevcutlar süz 2021++
On Error Resume Next
For n = 0 To LBMSC1.ListCount - 1
    LBPSA1.Object = "": LBPSA1.Object = LBMSC1.List(n)
     If LBPSA1.listIndex < 0 Then
     LBMSC1.List(n, 6) = "TANIMSIZ"
     Else
     LBMSC1.List(n, 6) = LBPSA1.List(LBPSA1.listIndex, 2)
     LBMSC1.List(n, 7) = LBPSA1.List(LBPSA1.listIndex, 1) & " Yoğunluk"
     End If
Next n
End Sub
Private Sub CBmontaj_Click() '2021
On Error Resume Next
C1 = LBMSC1.listIndex: C2 = LBMSC2.listIndex
If LBMSC1.ListCount <= 0 Then Exit Sub
Call CarpanlarCMSTV ': Call MontajCarpanlarrenk1
Call iscilik_gir1
Call MontajCarpankontrol
If C1 >= 0 Then LBMSC1.Selected(C1) = True
If C2 >= 0 Then LBMSC2.Selected(C2) = True
End Sub
Private Sub CBsarf_Click() '2021
On Error Resume Next
C1 = LBMSC1.listIndex: C2 = LBMSC2.listIndex
If LBMSC1.ListCount <= 0 Then Exit Sub
Call CarpanlarCMSTV:  'Call SarfCarpanlarrenk1
Call sarf_gir1
Call MontajCarpankontrol
If C1 >= 0 Then LBMSC1.Selected(C1) = True
If C2 >= 0 Then LBMSC2.Selected(C2) = True
End Sub
Private Sub CBamb_Click()
son = dts.Range("B65536").End(xlUp).row
If son < 2 Then Exit Sub
Call amb_gir1
Call CarpanlarCATV: 'Call AmbCarpanlarrenk1
End Sub
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button) '+1
On Error Resume Next
Toolbar1.Buttons(4).ButtonMenus(1).Enabled = False
Toolbar1.Buttons(4).ButtonMenus(2).Enabled = False
Toolbar1.Buttons(2).ButtonMenus(1).Enabled = False
Toolbar1.Buttons(2).ButtonMenus(2).Enabled = False
Select Case Button.Index
Case 1
MultiPage1.Value = 0
Windows(dt).Activate: Sheets("Sayfa1").Select '
marka = CBBPM1.Text
If Not marka = LBKTM3 Then Call panoyapıparametre
Call panodetay1getir
If ListBoxPT.listIndex >= 0 Then ListBoxPT.Selected(ListBoxPT.listIndex) = True
se1 = ListBoxPTD1.listIndex
FYTOK.Value = False: Call varsayılan1
Call isimyarat
FYTOK.Value = True
UFOPAN00.Height = 259
Case 2 '2021 +21 PANO CARPAN
Toolbar1.Buttons(2).ButtonMenus(1).Enabled = True: Toolbar1.Buttons(2).ButtonMenus(2).Enabled = True
ListBoxMARKA2.Visible = False: ListBoxMARKA21.Visible = True: CBPC2.Enabled = True
If CBPC2.Value = True Then LBPC1.Visible = True: LBPC2.Visible = False Else LBPC2.Visible = True: LBPC1.Visible = False
LBPC1.Visible = False: LBPC2.Visible = True
Frame24.Enabled = False
If LBPC1.Visible = False Then CBPC2.Value = False
a = ListBoxMARKA21.listIndex
MultiPage1.Value = 1: UFOPAN00.Height = 259
Call mevcutmarkalar21
Call mevcutmarkalarPC21
'Call teklifpanocarpan_mevcutlar21
If a < 0 Then ListBoxMARKA21.Selected(0) = True Else ListBoxMARKA21.Selected(a) = True
If LBPC2.BackColor = &HE6D3C4 Then TBCAK.Value = "Teklife Girilmiş Pano Çarpanları" 'mavi
'--
Case 3
MultiPage1.Value = 6: UFOPAN00.Height = 259
Case 4
Windows(dt).Activate: Sheets("Sayfa1").Select '
MultiPage1.Value = 2: UFOPAN00.Height = 259
Toolbar1.Buttons(4).ButtonMenus(1).Enabled = True: Toolbar1.Buttons(4).ButtonMenus(2).Enabled = True: FrameR.Enabled = False
LBLPSA1 = "": LBLPSA3 = "": TBMH11 = "": LBLPSA2.Height = 0
Call MontajCarpankontrol
CBMSC2.Visible = True: CBMSC2.Value = False
LBMSC2.IntegralHeight = False: LBMSC2.Height = 145: LBMSC2.IntegralHeight = True
If pfc = 0 Then CKPKOD.Value = True
'--
Case 5
If pfc = 0 Then
MultiPage1.Value = 3
Else
MultiPage1.Value = 4: UFOPAN00.Height = 259
Call panotipparametreler
End If
Case 6
If Toolbar1.Buttons.Item(6).Image = ImageList1.ListImages.Item(6).Index Then
Toolbar1.Buttons.Item(6).Image = ImageList1.ListImages.Item(9).Index
prs = 1
 If pfc = 0 Then
  UFOPAN00P1.Show
  If ListBoxP5.listIndex < 0 Then UFOPAN00P1.TBRS01 = ListBoxP1.Text Else UFOPAN00P1.TBRS01 = TBskd01
  If UFOPAN00P1.TBRS01 = "" Then UFOPAN00P1.TBRS01 = TBPM01
 End If
  If pfc = 1 Then
  UFOPAN00P2.Show
  If Toolbar1.Buttons.Item(6).Image = ImageList1.ListImages.Item(9).Index Then UFOPAN00P2.TBRS01 = TBskd
  End If
Else
  If pfc = 0 Then Unload UFOPAN00P1
  If pfc = 1 Then Unload UFOPAN00P2
End If
Case 7
If UFOPAN00.Height > 160 Then
UFOPAN00.Height = UFOPAN00.Height - UFOPAN00.InsideHeight + Toolbar1.Height
Else
UFOPAN00.Height = 259
End If
Case 8
If pfc = 1 Then End
Unload Me
End Select
'If MultiPage1.Value = 2 Or MultiPage1.Value = 1 Then Toolbar1.Buttons.item(3).Enabled = False Else Toolbar1.Buttons.item(3).Enabled = True
End Sub
Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
Select Case ButtonMenu.Tag
Case 1
On Error Resume Next
MultiPage1.Value = 1: UFOPAN00.Height = 260
If ListBoxKT1.ListCount < 1 Then Exit Sub
ListBoxMARKA21.Visible = False: ListBoxMARKA21.Object = "": ListBoxMARKA2.Visible = True
LBPC2.Visible = False: LBPC1.Visible = True: LBPC1.Height = 145: LBPC1.IntegralHeight = True
CBPC2.Enabled = False
If ListBoxMARKA2.listIndex < 0 Then ListBoxMARKA2.Selected(0) = True Else ListBoxMARKA2_Click
TBCAK.Value = "Varsayılan Pano Çarpanları"
TBPLIST = 1
Case 2
MultiPage1.Value = 1
msg = MsgBox("Seçilen markanın çarpanları teklife kaydedilecek.", vbOKCancel, "Pano Çarpan Kaydetme İşlemi")
If msg = vbCancel Then Exit Sub
Call CarpanlarPCTV
Case 3
MultiPage1.Value = 2
LBMSC1.IntegralHeight = False: LBMSC1.Height = 145: LBMSC1.IntegralHeight = True
PCKAYITLIMSACV_Click
Call montajcarpandetay
CBMSC2.Visible = False
Case 4
MultiPage1.Value = 2
msg = MsgBox("Montaj & Sarf çarpanları teklife kaydedilecek.", vbOKCancel, "Montaj & Sarf Çarpan Kaydetme İşlemi")
If msg = vbCancel Then Exit Sub
Call CarpanlarCMSTV
Call montajsarfacıklamalar
Case 5
MultiPage1.Value = 5
Label468.Caption = "       " & marka & " Yapısal Özellikler"
Case 6
MultiPage1.Value = 3
Case 7
mlzds = pmlz: pmlz = ""
Unload Me
Application.Windows(mlzds).Visible = True
End Select
End Sub
Private Sub CommandButton10_Click() 'AYARLAR BAKKKKK
On Error Resume Next
misi = TBimarka: SaveSetting "ilhan", "Settings", "misi", misi 'işçilik-montaj markası
msas = TBsmarka: SaveSetting "ilhan", "Settings", "msas", msas 'sarf markası
mbab = TBbmarka: SaveSetting "ilhan", "Settings", "mbab", mbab 'bara markası
mama = TBamarka: SaveSetting "ilhan", "Settings", "mama", mama 'ambalaj markası

misia = TBimarkaA: SaveSetting "ilhan", "Settings", "misia", misia 'işçilik-montaj açıklama
msasa = TBsmarkaA: SaveSetting "ilhan", "Settings", "msasa", msasa 'sarf açıklama
mbaba = TBbmarkaA: SaveSetting "ilhan", "Settings", "mbaba", mbaba 'bara açıklama
mamaa = TBamarkaA: SaveSetting "ilhan", "Settings", "mamaa", mamaa 'ambalaj açıklama
'--
skdi = TBskdi: SaveSetting "ilhan", "Settings", "skdi", skdi 'işçilik-montaj sip.kod.
skds = TBskds: SaveSetting "ilhan", "Settings", "skds", skds 'sarf sip.kod.
skdb = TBskdb: SaveSetting "ilhan", "Settings", "skdb", skdb 'bara sip.kod.
skda = TBskda: SaveSetting "ilhan", "Settings", "skda", skda 'ambalaj sip.kod.
'--
ActiveWorkbook.names("Catv").Comment = TextBoxamb
End Sub
Private Sub SpinButton1_Change() 'MİKTAR DEĞİŞİMİ '2021 +21 PANO
On Error Resume Next
TBmiktar.Text = SpinButton1.Value
TBbfiyat.Value = Format(TextBoxFY.Value * SpinButton1.Value, "#,##0.00")
End Sub
Private Sub TBmiktar_Change() 'MİKTAR DEĞİŞİMİ '2021 +21 PANO
On Error Resume Next
SpinButton1.Value = TBmiktar.Text
End Sub
Private Sub TextBoxFY_Change() '2021 +21 PANO
On Error Resume Next
TBbfiyat.Value = Format(TextBoxFY.Value * SpinButton1.Value, "#,##0.00")
End Sub
Private Sub CBsarfsil_Click() '2021++
Call sarfsil
End Sub
Private Sub CBambsil_Click() '2021++
Call ambsil
End Sub
Private Sub CBmontajsil_Click() '2021++
Call isciliksil
End Sub
Private Sub Image4_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single) '2021 +21 PANO MONTAJ
On Error Resume Next
If LBMSC1.BackColor = &H80000004 Then
CreateObject("Shell.Application").Open "C:\Belgelerim\Cemex\Ayarlar\Montaj ve Sarf\Montaj ve Sarf Çarpanlar.txt"
Else
tx = InputBox("İlave Etmek/Değiştirmek İstediğiniz Harfleri Giriniz.", "Teklifte Bulunan Mevcut Montaj & Sarf Çarpan Harfleri", ActiveWorkbook.names("Chtv").Comment)
If tx = "" Or tx = ActiveWorkbook.names("Chtv").Comment Then Exit Sub
ActiveWorkbook.names.Add Name:="Chtv", RefersToR1C1:="Montaj Harfler": ActiveWorkbook.names("Chtv").Comment = tx
Call MontajCarpanlar
Call Montajteklif
Call Montajteklifsüz
Call montajsarfacıklamalar
End If
End Sub
Private Sub Image44_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single) '2021 +21 PANO MONTAJ
On Error Resume Next
CreateObject("Shell.Application").Open "C:\Belgelerim\Cemex\Ayarlar\Montaj ve Sarf\Montaj ve Sarf Açıklamalar.txt"
Image44.Visible = False: Image441.Visible = True
TBMH11 = "!"
End Sub
Private Sub Image441_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single) '2021 +21 PANO MONTAJ
On Error Resume Next
Call montajsarfacıklamalar
Image441.Visible = False: Image44.Visible = True
TBMH11 = TBMH1
End Sub
Private Sub Image7_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single) '2021 +21 PANO
On Error Resume Next
CreateObject("Shell.Application").Open "C:\Belgelerim\Cemex\Ayarlar\Panolar"
End Sub
Private Sub Image8_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single) '2021++
On Error Resume Next
CreateObject("Shell.Application").Open "C:\Belgelerim\Cemex\Ayarlar\Panolar\" & marka & "\Pano Tip Tanımlamalar.txt"
End Sub
Private Sub Image9_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single) '2021 +21 PANO
On Error Resume Next
CreateObject("Shell.Application").Open "C:\Belgelerim\Cemex\Ayarlar\Montaj ve Sarf"
End Sub
Private Sub Image71_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single) '2021++
On Error Resume Next
If ListBoxPT.listIndex < 0 Then Exit Sub
Dim ds, pds
Set ds = CreateObject("Scripting.FileSystemObject")
pdosya = "C:\Belgelerim\Cemex\Ayarlar\Panolar\" & ListBoxmarka.List(ListBoxmarka.listIndex) & "\" & ListBoxPT.List(ListBoxPT.listIndex, 0) & ".txt"
pds = ds.FileExists(pdosya)
If pds = True Then CreateObject("Shell.Application").Open pdosya Else MsgBox pdosya & " Dosyası Bulunamadı !", vbCritical, "Hata !"
End Sub
Private Sub Image72_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single) '2021 +21 PANO CARPAN
On Error Resume Next
CreateObject("Shell.Application").Open "C:\Belgelerim\Cemex\Ayarlar\Panolar"
End Sub
Private Sub Label364_Click() '2021 +21 PANO CARPAN
On Error Resume Next
CreateObject("Shell.Application").Open "C:\Belgelerim\Cemex\Ayarlar\Panolar\" & marka & "\Pano Tip Tanımlamalar.txt"
End Sub
Private Sub Label468_Click()
On Error Resume Next
CreateObject("Shell.Application").Open "C:\Belgelerim\Cemex\Ayarlar\Panolar\" & marka & "\Pano Yapısal Parametreler.txt"
End Sub
Sub PCKAYITLIMSACV_Click() ''varsayılanlar TXT 2021
LBMSC1.BackColor = &H80000004
Label425.Caption = "      Firma Montaj & Sarf Çarpanlar"
TBMSAK = "Firma Montaj & Sarf Çarpanlar"
Call Carpanlarmontajsarf
TBMSLIST = 1
TBPST1.Text = ""
LBMSC2.Visible = False: LBMSC1.Visible = True
'Call Carpanlarsarf
'Call Carpanlaramb
End Sub
Sub Carpanlarmontajsarf() 'varsayılanlar TXT 2021
On Error Resume Next
Dim Rky As String, Dosyam As String
    Dim Ert As Long, satır As Long, i As Long
    Dim ayır As Variant
    Dosyam = "C:\Belgelerim\Cemex\Ayarlar\Montaj ve Sarf\Montaj ve Sarf Çarpanlar.txt"
    Ert = FreeFile
    Open Dosyam For Input As #Ert
    If Err.Number <> 0 Then
        MsgBox "Montaj ve Sarf Çarpanlar.txt" & " Dosyası Bulunamadı !", vbCritical, "Hata !"
        Exit Sub
    End If
    On Error GoTo 0
    satır = 1
    LBMSC1.Clear: LBMSC2.Clear
    'ListView1.ListItems.Clear
Do While Not EOF(Ert)
        Line Input #Ert, Rky
        ayır = Split(Rky, ";")
        LBMSC1.AddItem ayır(i)
        'Call ListView1.ListItems.Add((satır), , ayır(i))
tsay = Len(Rky) - Len(Replace(Rky, ";", "")) 'yanayana kaç adet var
For n = 1 To tsay '
If UBound(ayır) <> 0 Then LBMSC1.List(satır - 1, n) = ayır(n)
'If UBound(ayır) <> 0 Then Call ListView1.ListItems(satır).ListSubItems.Add(n, , ayır(n))
Next n
     satır = satır + 1
Loop
    kt = 1
Close #Ert
End Sub
Sub MontajCarpankontrol() '2021++'carpan kontrol
Call MontajCarpanlar
Call Montajteklif
Call Montajteklifsüz
Label478.Caption = Format(WorksheetFunction.SumIfs(dts.Range("P1:P65536"), dts.Range("A1:A65536"), "PM-MP*"), "#,##0.00") & " " & "TL"
Label480.Caption = Format(WorksheetFunction.SumIfs(dts.Range("P1:P65536"), dts.Range("A1:A65536"), "PM-MS*"), "#,##0.00") & " " & "TL"
End Sub
Sub MontajCarpanlar() '2021++
On Error Resume Next
LBMSC1.BackColor = &H80000004 'gri
Dim Chtv, Cmtv, Cstv, Catv
Chtv = ActiveWorkbook.names("Chtv").RefersToR1C1 'harfler
Cmtv = ActiveWorkbook.names("Cmtv").RefersToR1C1: Cstv = ActiveWorkbook.names("Cstv").RefersToR1C1
If Chtv = Empty Then
 If Cmtv <> Empty And Cstv <> Empty Then
 dizi0 = "A-B-C-D-E-F-G-H-I-J-K-L-M-N-O-P-Q-R-S-T-U-V-W-X-Y-Z"
 ActiveWorkbook.names.Add Name:="Chtv", RefersToR1C1:="Montaj Harfler": ActiveWorkbook.names("Chtv").Comment = dizi0
 End If
End If
If Chtv = Empty And Cmtv = Empty And Cstv = Empty Then
Call Carpanlarmontajsarf
Call CarpanlarCMSTV
Else
Call CarpanlarCMSTA 'momtaj-sarf çarpanlar ad yöneticisinden -> listeye
Label425.Caption = "      Teklif Montaj & Sarf Çarpanlar"
TBMSAK = "Teklif Montaj & Sarf Çarpanlar"
End If
'--
Catv = ActiveWorkbook.names("Catv").RefersToR1C1
'If Catv = Empty Then Call Carpanlaramb Else: Call CarpanlarCATA: 'Call AmbCarpanlarrenk1
End Sub
Sub MontajCarpanlar01() '2021++ pano girişinde kullanmak için
On Error Resume Next
Dim Chtv
Chtv = ActiveWorkbook.names("Chtv").RefersToR1C1 'harfler
If Not Chtv = Empty Then
Dim Cmtv, Cstv
Chtv = ActiveWorkbook.names("Chtv").Comment 'harfler
Cmtv = ActiveWorkbook.names("Cmtv").Comment: Cstv = ActiveWorkbook.names("Cstv").Comment
dizi0 = Chtv: dizi1 = Cmtv: dizi2 = Cstv
ayır0 = Split(dizi0, "-"): ayır1 = Split(dizi1, "-"): ayır2 = Split(dizi2, "-")
tsay = Len(dizi0) - Len(Replace(dizi0, "-", "")) + 1
LBPSA01.Clear
For n = 1 To tsay '
    LBPSA01.AddItem ayır0(n - 1)
    LBPSA01.List(n - 1, 1) = ayır1(n - 1)
    LBPSA01.List(n - 1, 2) = ayır2(n - 1)
Next n
End If
End Sub
Sub MontajCarpanlarrenk1() 'teklif kayıtlı'2021++
'r = &HE1F9F8
r = &HDBE8DB
LBMSC1.BackColor = r
End Sub
Sub MontajCarpanlarrenk2() 'Pc kayıtlı'2021++
r = &HE0E0E0
LBMSC1.BackColor = r
End Sub
Private Sub CommandButtonMSACTA_Click()
On Error Resume Next
Call MontajCarpanlar
Call Montajteklif
Call Montajteklifsüz
End Sub
Private Sub CommandButtonMSACTK_Click() '2021
If LBMSC1.ListCount < 1 Then Exit Sub
If LBMSC1.BackColor = &H80000004 Then Call PcMontajCarpanver: Exit Sub
Call CarpanlarCMSTV ' momtaj-sarf çarpanlar listeden -> ad yöneticisine
Call CarpanlarCATV
End Sub
Sub PcMontajCarpanver() 'momtaj-sarf çarpanlar listeden -> txt ye 2021
If LBMSC1.ListCount < 1 Then Exit Sub
Const Myfile As String = "C:\Belgelerim\Cemex\Ayarlar\Montaj ve Sarf\Montaj ve Sarf Çarpanlar.txt"
If Len(dir(Myfile)) > 0 Then
     Kill Myfile
 End If
    Open Myfile For Append As #1
    For i = 0 To LBMSC1.ListCount - 1
        Print #1, LBMSC1.List(i, 0) & ";" & LBMSC1.List(i, 1) & ";" & LBMSC1.List(i, 2)
    Next
    Close #1
End Sub
Sub CarpanlarCMSTA() 'momtaj-sarf çarpanlar ad yöneticisinden -> listeye 2021+
On Error Resume Next
LBMSC1.BackColor = &HDBE8DB    'açık yeşil renk teklif
Windows(dt).Activate ': Sheets("Sayfa1").Select '
Dim Chtv, Cmtv, Cstv
Chtv = ActiveWorkbook.names("Chtv").Comment 'harfler
Cmtv = ActiveWorkbook.names("Cmtv").Comment: Cstv = ActiveWorkbook.names("Cstv").Comment
dizi0 = Chtv: dizi1 = Cmtv: dizi2 = Cstv
ayır0 = Split(dizi0, "-"): ayır1 = Split(dizi1, "-"): ayır2 = Split(dizi2, "-")
tsay = Len(dizi0) - Len(Replace(dizi0, "-", "")) + 1
LBMSC1.Clear
For n = 1 To tsay '
    LBMSC1.AddItem ayır0(n - 1)
    LBMSC1.List(n - 1, 1) = ayır1(n - 1)
    LBMSC1.List(n - 1, 2) = ayır2(n - 1)
    LBMSC1.List(n - 1, 3) = ""
    LBMSC1.List(n - 1, 4) = ""
    LBMSC1.List(n - 1, 5) = ""
Next n
End Sub
Sub CarpanlarCMSTV() 'momtaj-sarf çarpanlar listeden -> ad yöneticisine 2021+
On Error Resume Next
Windows(dt).Activate ': Sheets("Sayfa1").Select '
dizi0 = LBMSC1.List(0): dizi1 = LBMSC1.List(0, 1): dizi2 = LBMSC1.List(0, 2)
For n = 1 To LBMSC1.ListCount - 1
    dizi0 = dizi0 & "-" & LBMSC1.List(n) 'harfler
    dizi1 = dizi1 & "-" & LBMSC1.List(n, 1): dizi2 = dizi2 & "-" & LBMSC1.List(n, 2)
Next n
Dim Chtv, Cmtv, Cstv
Chtv = ActiveWorkbook.names("Chtv").RefersToR1C1 'harfler
Cmtv = ActiveWorkbook.names("Cmtv").RefersToR1C1: Cstv = ActiveWorkbook.names("Cstv").RefersToR1C1
If Chtv = Empty Then
ActiveWorkbook.names.Add Name:="Chtv", RefersToR1C1:="Montaj Harfler": ActiveWorkbook.names("Chtv").Comment = dizi0
Else
ActiveWorkbook.names("Chtv").Comment = dizi0
End If
If Cmtv = Empty Then
ActiveWorkbook.names.Add Name:="Cmtv", RefersToR1C1:="Montaj Çarpanlar": ActiveWorkbook.names("Cmtv").Comment = dizi1
Else
ActiveWorkbook.names("Cmtv").Comment = dizi1
End If
If Cstv = Empty Then
ActiveWorkbook.names.Add Name:="Cstv", RefersToR1C1:="Sarf Çarpanlar": ActiveWorkbook.names("Cstv").Comment = dizi2
Else
ActiveWorkbook.names("Cstv").Comment = dizi2
End If
Call MontajCarpanlar01 'pano için geri listeye yükle
End Sub
Sub CarpanlarCATA()
On Error Resume Next
Windows(dt).Activate ': Sheets("Sayfa1").Select '
Dim Catv
Catv = ActiveWorkbook.names("Catv").Comment
dizi1 = Catv
TextBoxamb = Split(dizi1, "-")(0)
End Sub
Sub CarpanlarCATV()
On Error Resume Next
Windows(dt).Activate ': Sheets("Sayfa1").Select '
Dim Catv
Catv = ActiveWorkbook.names("Catv").RefersToR1C1
If Catv = Empty Then
ActiveWorkbook.names.Add Name:="Catv", RefersToR1C1:="Amb.Çarpanlar": ActiveWorkbook.names("Catv").Comment = TextBoxamb
Else
ActiveWorkbook.names("Catv").Comment = TextBoxamb
End If
End Sub
Sub Montajteklif() ' mevcutlar 2021+
On Error Resume Next
Dim msa As Integer
son = dts.Range("B65536").End(xlUp).row
With dts.Range("A2:A" & son)
    Set c = .Find("PP-", LookIn:=xlValues, LookAt:=xlPart)
    If Not c Is Nothing Then
        firstAddress = c.Address
        dizi0 = ActiveWorkbook.names("Chtv").Comment
        Do
          mharf = ""
          For sy = 1 To Len(dts.Range("I" & c.row))
          If IsNumeric(Mid(dts.Range("I" & c.row), sy, 1)) = True Then Exit For Else mharf = mharf & Mid(dts.Range("I" & c.row), sy, 1)
          Next
          arr1 = Split(dizi0, "-")
          aranan = Application.Match(mharf, arr1, False)
          If Not IsError(aranan) Then
          LBMSC1.List(aranan - 1, 3) = LBMSC1.List(aranan - 1, 3) & c.row & "."
          msa = 0: msa = CDbl(LBMSC1.List(aranan - 1, 8))
          LBMSC1.List(aranan - 1, 8) = msa + CDbl(dts.Range("E" & c.row))
          If LBMSC1.List(aranan - 1, 4) = "" Then LBMSC1.List(aranan - 1, 4) = aranan - 1
          Else
          '--
          If Not mharf = "" Then
          LBMSC1.AddItem mharf
          LBMSC1.List(LBMSC1.ListCount - 1, 1) = 0
          LBMSC1.List(LBMSC1.ListCount - 1, 2) = 0
          LBMSC1.List(LBMSC1.ListCount - 1, 3) = LBMSC1.List(LBMSC1.ListCount - 1, 3) & c.row & "."
          msa = CDbl(LBMSC1.List(LBMSC1.ListCount - 1, 8))
          LBMSC1.List(LBMSC1.ListCount - 1, 8) = msa + CDbl(dts.Range("E" & c.row))
          If LBMSC1.List(LBMSC1.ListCount - 1, 4) = "" Then LBMSC1.List(LBMSC1.ListCount, 4) = LBMSC1.ListCount - 1
          LBMSC1.List(LBMSC1.ListCount - 1, 5) = ""
          dizi0 = dizi0 & "-" & mharf
          End If
          '--
          End If
          Set c = .FindNext(c)
        Loop While Not c Is Nothing And c.Address <> firstAddress
    End If
End With
End Sub
Sub Montajteklifsüz() ' mevcutlar süz 2021
On Error Resume Next
LBMSC2.Clear
CBMSC2.Value = False
LBMSC1.Visible = False: LBMSC2.Visible = True
For n = 0 To LBMSC1.ListCount - 1
  If LBMSC1.List(n, 3) <> "" Then
     LBMSC2.AddItem LBMSC1.List(n)
     LBMSC2.List(LBMSC2.ListCount - 1, 1) = LBMSC1.List(n, 1)
     LBMSC2.List(LBMSC2.ListCount - 1, 2) = LBMSC1.List(n, 2)
     LBMSC2.List(LBMSC2.ListCount - 1, 3) = LBMSC1.List(n, 3)
     LBMSC1.List(n, 5) = LBMSC2.ListCount - 1
     LBMSC2.List(LBMSC2.ListCount - 1, 5) = n
     LBMSC2.List(LBMSC2.ListCount - 1, 8) = LBMSC1.List(n, 8) & " Ad."
'--
    LBPSA1.Object = "": LBPSA1.Object = LBMSC1.List(n)
     If LBPSA1.listIndex < 0 Then
     LBMSC2.List(LBMSC2.ListCount - 1, 6) = "TANIMSIZ"
     Else
     LBMSC2.List(LBMSC2.ListCount - 1, 6) = LBPSA1.List(LBPSA1.listIndex, 2)
     LBMSC2.List(LBMSC2.ListCount - 1, 7) = LBPSA1.List(LBPSA1.listIndex, 1) & " Yoğunluk"
     End If
  End If
Next n
End Sub
Sub Resim01()
On Error GoTo hata
Dim rd, a
Set rd = CreateObject("Scripting.FileSystemObject")
a = rd.FileExists("C:\Belgelerim\CEMEX\Resimler\" & ir & "\" & ListBoxP1.List(ListBoxP1.listIndex) & ".gif")
If a = True Then
Image011.Picture = LoadPicture("C:\Belgelerim\CEMEX\Resimler\" & ir & "\" & ListBoxP1.List(ListBoxP1.listIndex) & ".gif")
 ir1 = 2
Else
 If Not ir1 = 1 Then
  a = rd.FileExists("C:\Belgelerim\CEMEX\Resimler\" & ir & "\" & ir & ".gif")
  If a = True Then
  Image011.Picture = LoadPicture("C:\Belgelerim\CEMEX\Resimler\" & ir & "\" & ir & ".gif")
  ir1 = 1
  Else
  If Not ir1 = 0 Then Image011.Picture = LoadPicture("C:\Belgelerim\CEMEX\Resimler\noimagepano.gif")
  ir1 = 0
  End If
 End If
End If
hata:
End Sub
Private Sub ListBoxP1_Click()
On Error Resume Next
Label486 = "-": Label487 = "-"
MultiPage2.Value = 0
ListBoxP2.Clear: ListBoxP3.Clear: ListBoxP4.Clear: ListBoxP5.Clear
TBAMS11 = "": LBAMS11 = ""
rr = ListBoxP1.List(ListBoxP1.listIndex, 1)
Set msayfa = Workbooks(pmlz).Worksheets("Sayfa1")
Set msayfa2 = Workbooks(pmlz).Worksheets("Açıklamalar")
'If msayfa.Cells(rr, 11) = "" Then TBref01 = "PP-" & msayfa.Cells(rr, 11).End(4) Else TBref01 = "PP-" & msayfa.Cells(rr, 11)
If msayfa.Cells(rr, 11) = "" Then TBref01 = msayfa.Cells(rr, 11).End(4) Else TBref01 = msayfa.Cells(rr, 11)
bdetay = msayfa.Cells(rr, 1)
X = 1: Y = msayfa2.Range("A65536").End(xlUp).row
ListBoxP2.AddItem "Hepsi"
Do Until X > Y
  If msayfa2.Cells(X, 1) = bdetay Then
    For i = 2 To 15
      If msayfa2.Cells(X, i) <> "" Then ListBoxP2.AddItem msayfa2.Cells(X, i)
    Next i
    GoTo bitir
  End If
  X = X + 1
Loop
bitir:
If Not ListBoxP2.ListCount > 0 Then lr1 = ListBoxP1.List(ListBoxP1.listIndex, 1): lr2 = ListBoxP1.List(ListBoxP1.listIndex, 2): Call P2YOK
Call montajtipler01
Frame42.Enabled = False: TBskd01 = "": TextBoxFY01 = 0: TBmiktar01 = 1
If ListBoxPTD101.ListCount > 0 Then ListBoxPTD101.Selected(1) = True
'ir = Left(msayfa.Cells(rr, 4).End(4), 3): ir = Trim(ir)
'ir = msayfa.Cells(rr, 4).End(4)
Call Resim01
If Toolbar1.Buttons.Item(6).Image = 9 Then UFOPAN00P1.TBRS01 = ListBoxP1.Text
End Sub
Private Sub ListBoxP2_Click()
On Error Resume Next
Label486 = "-": Label487 = "-"
If ListBoxP3.listIndex >= 0 Then lp3 = ListBoxP3.Text: If ListBoxP4.listIndex >= 0 Then lp4 = ListBoxP4.Text
ListBoxP3.Clear: ListBoxP4.Clear: ListBoxP5.Clear
lt = ListBoxP2.List(ListBoxP2.listIndex, 0)
lt = UCase(Replace(Replace(lt, "ı", "I"), "i", "İ"))
Dim X As Integer, Y As Integer
lr = ListBoxP1.listIndex: X = ListBoxP1.List(lr, 1): Y = ListBoxP1.List(lr, 2)
 If lt = "HEPSİ" Then
 lr1 = ListBoxP1.List(lr, 1): lr2 = ListBoxP1.List(lr, 2):  Call P2YOK
 Set msayfa = Workbooks(pmlz).Worksheets("Sayfa1")
 Set msayfa2 = Workbooks(pmlz).Worksheets("Açıklamalar")
 rr = ListBoxP1.List(ListBoxP1.listIndex, 1)
 If msayfa.Cells(rr, 11) = "" Then TBref01 = msayfa.Cells(rr, 11).End(4) Else TBref01 = msayfa.Cells(rr, 11)
 Call montajtipler01
 Frame42.Enabled = False: TBskd01 = "": TextBoxFY01 = 0: TBmiktar01 = 1
 Exit Sub
 End If
Dim Y1 As String, Y2 As String, S1 As String, S2 As String
Dim rsr As Integer, z As Integer
Set msayfa = Workbooks(pmlz).Worksheets("Sayfa1")
z = X
Y2 = "": D2 = "": S2 = ""
Do Until X > Y
  If msayfa.Cells(X, 1) <> "" And msayfa.Cells(X, 2) <> "" Then
'--
     If UCase(Replace(Replace(msayfa.Cells(X, 1), "ı", "I"), "i", "İ")) Like "*" & lt & "*" Then
     Y1 = Split(Split(msayfa.Cells(X, 1), "YÜK.")(1), " ")(0)
     If Y1 = "" Then Y1 = Split(Split(msayfa.Cells(X, 1), "Yük.")(1), " ")(0)
     If Y1 = "" Then GoTo git1
     ListBoxP3.AddItem Y1: ListBoxP3.List(ListBoxP3.ListCount - 1, 1) = msayfa.Cells(X, 1).row
     'If Not Y1 = Y2 Then ListBoxP3.AddItem Y1: ListBoxP3.List(ListBoxP3.ListCount - 1, 1) = msayfa.Cells(X, 1).Row: Y2 = Y1
     rsr = msayfa.Cells(X, 1).End(4).row
     If rsr <> 0 Then ListBoxP3.List(ListBoxP3.ListCount - 1, 2) = rsr - 1 Else ListBoxP3.List(ListBoxP3.ListCount - 1, 2) = rs
     Y1 = "": Label486 = "  Yük.": GoTo git3
git1:
'--
     D1 = Split(Split(msayfa.Cells(X, 1), "DER.")(1), " ")(0)
     If D1 = "" Then D1 = Split(Split(msayfa.Cells(X, 1), "Der.")(1), " ")(0)
     If D1 = "" Then GoTo git2
     If CDbl(D1) > CDbl(D2) Then ListBoxP4.AddItem D1: ListBoxP4.List(ListBoxP4.ListCount - 1, 1) = msayfa.Cells(X, 1).row: D2 = D1
     rsr = msayfa.Cells(X, 1).End(4).row
     If rsr <> 0 Then ListBoxP4.List(ListBoxP4.ListCount - 1, 2) = rsr - 1 Else ListBoxP4.List(ListBoxP4.ListCount - 1, 2) = rs
     D1 = "": GoTo git3
git2:
'--
     If UCase(Replace(Replace(msayfa.Cells(X, 1), "ı", "I"), "i", "İ")) Like "*" & "SERİ" & "*" Then
     S1 = " " & Split(msayfa.Cells(X, 1), "Seri")(0)
     S1 = Split(S1, " ")(1)
     If S1 = "" Then GoTo git3
     ListBoxP3.AddItem S1: ListBoxP3.List(ListBoxP3.ListCount - 1, 1) = msayfa.Cells(X, 1).row
     If msayfa.Cells(X + 1, 1) <> "" Then rsr = X + 1 Else rsr = msayfa.Cells(X, 1).End(4).row
     If rsr <> 0 Then ListBoxP3.List(ListBoxP3.ListCount - 1, 2) = rsr - 1 Else ListBoxP3.List(ListBoxP3.ListCount - 1, 2) = rs
     S1 = "":  Label486 = "  Serisi": GoTo git3
     End If
git3:
     rsr = 0
     End If
  End If
  X = X + 1
Loop
If Label486 = "  Yük." Then ListBoxP3.TextAlign = fmTextAlignRight: ListBoxP2.Tag = "1"
If Label486 = "  Serisi" Then ListBoxP3.TextAlign = fmTextAlignLeft: ListBoxP2.Tag = "2"
If ListBoxP4.ListCount > 0 Then Label487 = "  Der.": ListBoxP2.Tag = "3"
If ListBoxP3.ListCount = 0 And ListBoxP4.ListCount = 0 Then
 ListBoxP5.Clear: ListBoxP5.IntegralHeight = False
 Set msayfa = Workbooks(pmlz).Worksheets("Sayfa1")
 MD2 = ""
 Do Until z > Y
 If UCase(Replace(Replace(msayfa.Cells(z, 3), "ı", "I"), "i", "İ")) Like "*" & lt & "*" Then
'--
   If lt Like "*" & "MODÜL" & "*" Then GoTo git5
        If UCase(Replace(Replace(msayfa.Cells(z, 3), "ı", "I"), "i", "İ")) Like "*" & "MODÜL " & "*" Then
        MD1 = Split(Right(Split(msayfa.Cells(z, 3), "Modül ")(0), 6), " ")(1)
          If MD1 = "" Then GoTo git5
          If CDbl(MD1) > CDbl(MD2) Then
          ListBoxP3.AddItem MD1: ListBoxP3.List(ListBoxP3.ListCount - 1, 1) = msayfa.Cells(z, 1).row: MD2 = MD1
          ListBoxP3.List(ListBoxP3.ListCount - 1, 2) = z: D1 = ""
          End If
        Else
git5:
        ListBoxP5.AddItem msayfa.Cells(z, 2)
        ListBoxP5.List(ListBoxP5.ListCount - 1, 1) = msayfa.Cells(z, 3)
        ListBoxP5.List(ListBoxP5.ListCount - 1, 2) = msayfa.Cells(z, 6)
        ListBoxP5.List(ListBoxP5.ListCount - 1, 3) = msayfa.Cells(z, 2).row
        End If
  End If
  z = z + 1
 Loop
 If ListBoxP3.ListCount > 0 Then Label486 = "  Modül"
 ListBoxP5.Height = 145: ListBoxP5.IntegralHeight = True
 End If

son1:
If ListBoxP3.ListCount > 0 Then ListBoxP3.Object = lp3
If ListBoxP4.ListCount > 0 Then ListBoxP4.Object = lp4
If ListBoxP2.Tag = "3" And ListBoxP3.ListCount > 0 Then ListBoxP3.Object = lp3
Frame42.Enabled = False: TBskd01 = "": TextBoxFY01 = 0: TBmiktar01 = 1
If Toolbar1.Buttons.Item(6).Image = 9 Then UFOPAN00P1.TBRS01 = ListBoxP1.Text
'--
If Left(TBref01, 3) = "PP-" Then
rr = ListBoxP5.List(0, 3)
 If ListBoxP3.ListCount > 0 Or ListBoxP3.listIndex > 0 Then
 rr = ListBoxP3.List(0, 2)
 Else
 If ListBoxP4.ListCount > 0 Then rr = ListBoxP4.List(0, 2)
 End If
 TBref001 = msayfa.Cells(rr, 11): If rr > 3 Then TBref002 = msayfa.Cells(rr, 11).End(3)
 If TBref001 = "" Then TBref001 = TBref002
 If Not TBref001 = TBref01 Then TBref01 = TBref001: Call montajtipler01
 If ListBoxPTD101.listIndex < 0 Then If ListBoxPTD101.ListCount > 0 Then ListBoxPTD101.Selected(1) = True
End If
'--
End Sub
Private Sub ListBoxP3_Click()
On Error Resume Next
Dim D1 As String, D2 As String
Dim X As Integer, Y As Integer
Set msayfa = Workbooks(pmlz).Worksheets("Sayfa1")
lx4 = ListBoxP4.List(ListBoxP4.listIndex, 0)
'---
If ListBoxP2.Tag = "1" Then
ListBoxP4.Clear: ListBoxP5.Clear
lt = ListBoxP3.List(ListBoxP3.listIndex, 0)
lr = ListBoxP3.listIndex: X = ListBoxP3.List(lr, 1): Y = ListBoxP3.List(lr, 2)
rr = ListBoxP3.List(lr, 1)
D2 = 0
Do Until X > Y
        D1 = Replace(Replace(Split("x" & Split(msayfa.Cells(X, 3), "x")(2), " ")(0), "xD", ""), "x", "")
        D1 = Replace(D1, "(cm)", "")
        D1 = Replace(D1, "(mm)", "")
        D1 = Replace(D1, "mm", "")
        If Not D1 = "" Then
        'D1 = Split(Split(msayfa.Cells(X, 3), "xD")(1), " ")(0)
        If CDbl(D1) > CDbl(D2) Then ListBoxP4.AddItem D1: D2 = D1
        End If
     X = X + 1
Loop
End If
'---
If ListBoxP2.Tag = "3" Then
ListBoxP5.Clear
lt3 = ListBoxP2.List(ListBoxP2.listIndex, 0)
lr = ListBoxP1.listIndex: X = ListBoxP1.List(lr, 1): Y = ListBoxP1.List(lr, 2)
lt = ListBoxP4.List(ListBoxP4.listIndex, 0): lt2 = ListBoxP3.List(ListBoxP3.listIndex, 0)
Y2 = ""
'formatlar
Dim DF1 As String, DF2 As String
DF1 = "*" & lt2 & "*" & "x" & "*" & "x" & "*" & lt & "*"
DF2 = "* " & lt2 & "*" & "x" & "*" & lt & "*"
'---
Do Until X > Y
If Not msayfa.Cells(X, 3) Like "*" & lt3 & "*" Then GoTo atla1
     Y1 = msayfa.Cells(X, 3)
     If Y1 Like DF1 Then GoTo atla2 '1.seçenek
     If Y1 Like "*" & "x" & "*" & "x" & "*" Then GoTo atla1
     If Y1 Like DF2 Then GoTo atla2 '2.seçenek
GoTo atla1
atla2:
     ListBoxP5.AddItem msayfa.Cells(X, 2)
     ListBoxP5.List(ListBoxP5.ListCount - 1, 1) = msayfa.Cells(X, 3)
     ListBoxP5.List(ListBoxP5.ListCount - 1, 2) = msayfa.Cells(X, 6)
     ListBoxP5.List(ListBoxP5.ListCount - 1, 3) = msayfa.Cells(X, 2).row
atla1:
'---
  X = X + 1
Loop
End If
'---
If ListBoxP4.ListCount > 0 Then Label487 = "  Der."
ListBoxP4.Object = lx4
Frame42.Enabled = False: TBskd01 = "": TextBoxFY01 = 0: TBmiktar01 = 1
If Not ListBoxP4.ListCount > 0 Then lr1 = ListBoxP3.List(ListBoxP3.listIndex, 1): lr2 = ListBoxP3.List(ListBoxP3.listIndex, 2): Call P2YOK
If Toolbar1.Buttons.Item(6).Image = 9 Then UFOPAN00P1.TBRS01 = ListBoxP1.Text
'--
TBref001 = msayfa.Cells(rr, 11): If rr > 3 Then TBref002 = msayfa.Cells(rr, 11).End(3)
If TBref001 = "" Then TBref001 = TBref002
If Not TBref001 = TBref01 Then TBref01 = TBref001: Call montajtipler01
If ListBoxPTD101.listIndex < 0 Then If ListBoxPTD101.ListCount > 0 Then ListBoxPTD101.Selected(1) = True
'--
End Sub
Private Sub ListBoxP4_Click()
On Error Resume Next
Dim X As Integer, Y As Integer
ListBoxP5.Clear: ListBoxP5.IntegralHeight = False
Set msayfa = Workbooks(pmlz).Worksheets("Sayfa1")
 If ListBoxP2.Tag = "3" Then
 If ListBoxP3.ListCount > 0 Then lx4 = ListBoxP3.List(ListBoxP3.listIndex, 0)
 ListBoxP3.Clear
 End If
If ListBoxP3.ListCount > 0 Then
lt3 = ListBoxP3.List(ListBoxP3.listIndex, 0)
lr = ListBoxP3.listIndex: X = ListBoxP3.List(lr, 1): Y = ListBoxP3.List(lr, 2)
Else
lt3 = ListBoxP2.List(ListBoxP2.listIndex, 0)
lr = ListBoxP1.listIndex: X = ListBoxP1.List(lr, 1): Y = ListBoxP1.List(lr, 2)
End If
lt = ListBoxP4.List(ListBoxP4.listIndex, 0)
Y2 = ""
Do Until X > Y
 If msayfa.Cells(X, 3) Like "*" & lt3 & "*" Then
'---
    If ListBoxP2.Tag = "1" Then
     D1 = Replace(Replace(Split("x" & Split(msayfa.Cells(X, 3), "x")(2), " ")(0), "xD", ""), "x", "")
     If D1 Like "*" & lt & "*" Then
     ListBoxP5.AddItem msayfa.Cells(X, 2)
     ListBoxP5.List(ListBoxP5.ListCount - 1, 1) = msayfa.Cells(X, 3)
     ListBoxP5.List(ListBoxP5.ListCount - 1, 2) = msayfa.Cells(X, 6)
     ListBoxP5.List(ListBoxP5.ListCount - 1, 3) = msayfa.Cells(X, 2).row
     End If
    End If
'---
    If ListBoxP2.Tag = "3" Then
    Y1 = Replace(Replace(Split(Right(Split(msayfa.Cells(X, 3), "x")(0), 6), " ")(1), "xD", ""), "x", "")
     If Not Y1 = "" Then
      If CDbl(Y1) > CDbl(Y2) Then
      ListBoxP3.AddItem Y1: ListBoxP3.List(ListBoxP3.ListCount - 1, 1) = msayfa.Cells(X, 1).row: Y2 = Y1
      ListBoxP3.List(ListBoxP3.ListCount - 1, 2) = X
      End If
     End If
    End If
'---
 End If
  X = X + 1
Loop
If ListBoxP2.Tag = "3" And ListBoxP3.ListCount > 0 Then
 If lr > 0 Then ListBoxP3.Object = lx4
 Label486 = "<-Gen.->"
End If
Frame42.Enabled = False: TBskd01 = "": TextBoxFY01 = 0: TBmiktar01 = 1
ListBoxP5.Height = 145: ListBoxP5.IntegralHeight = True
If Toolbar1.Buttons.Item(6).Image = 9 Then UFOPAN00P1.TBRS01 = ListBoxP1.Text
End Sub
Private Sub ListBoxP5_Click()
On Error Resume Next
Set msayfa = Workbooks(pmlz).Worksheets("Sayfa1")
rr = ListBoxP5.List(ListBoxP5.listIndex, 3)
TBskd01 = ListBoxP5.List(ListBoxP5.listIndex, 0)
TByic01 = ListBoxP5.List(ListBoxP5.listIndex, 1)
TBPM01 = msayfa.Cells(rr, 4)
TBPM09 = msayfa.Cells(rr, 9)
TBPM10 = msayfa.Cells(rr, 7) * 100
TBPM11 = msayfa.Cells(rr, 8).Value 'adam saat
'--
TBref001 = msayfa.Cells(rr, 11): If rr > 3 Then TBref002 = msayfa.Cells(rr, 11).End(3)
If TBref001 = "" Then TBref001 = TBref002
If Not TBref001 = TBref01 Then TBref01 = TBref001: Call montajtipler01
'--
'If ListBoxPTD101.ListCount > 0 Then ListBoxPTD101.Selected(1) = True
'If msayfa.Cells(rr, 11) = "" Then TBref01 = "PP-" & msayfa.Cells(rr, 11).End(3) Else TBref01 = "PP-" & msayfa.Cells(rr, 11)
Labelm211.Caption = "Yük." & "xGen." & "= " & TBPM09 / 100 & " m2"
TextBoxFY01.Value = Format(msayfa.Cells(rr, 6), "#,##0.00")
If TextBoxFY01.Value = "" Then TextBoxFY01.Value = 0
TBbfiyat01.Value = Format(TextBoxFY01.Value * TBmiktar01, "#,##0.00")
TextBoxFY11B1 = "TL": TextBoxFY11B2 = "TL"
If msayfa.Cells(rr, 6).NumberFormat = "#,##0.00 [$$-C0C]" Then TextBoxFY11B1 = "$": TextBoxFY11B2 = "$"
If msayfa.Cells(rr, 6).NumberFormat = "#,##0.00 [$€-1]" Then TextBoxFY11B1 = "€": TextBoxFY11B2 = "€"
Frame42.Enabled = True
'--
If ListBoxPTD101.listIndex < 0 Then If ListBoxPTD101.ListCount > 0 Then ListBoxPTD101.Selected(1) = True
If ListBoxPTD121.listIndex >= 0 Then TBPMS01 = ListBoxPTD121.List(ListBoxPTD121.listIndex) & TBPM09 Else TBPMS01 = TBPM09
If Toolbar1.Buttons.Item(6).Image = ImageList1.ListImages.Item(9).Index Then UFOPAN00P1.TBRS01 = TBskd01
End Sub
Private Sub ListBoxP5_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
If ListBoxP5.listIndex < 0 Then Exit Sub
CommandButton28_Click
End Sub
Sub P2YOK()
On Error Resume Next
ListBoxP5.Clear: ListBoxP5.IntegralHeight = False
Set msayfa = Workbooks(pmlz).Worksheets("Sayfa1")
Do Until lr1 > lr2
     If Not msayfa.Cells(lr1, 3) = "" Then
     ListBoxP5.AddItem msayfa.Cells(lr1, 2)
     ListBoxP5.List(ListBoxP5.ListCount - 1, 1) = msayfa.Cells(lr1, 3)
     ListBoxP5.List(ListBoxP5.ListCount - 1, 2) = msayfa.Cells(lr1, 6)
     ListBoxP5.List(ListBoxP5.ListCount - 1, 3) = msayfa.Cells(lr1, 2).row
     End If
  lr1 = lr1 + 1
Loop
ListBoxP5.Height = 145: ListBoxP5.IntegralHeight = True
End Sub
Sub montajtipler01()
On Error Resume Next
LBPSA0.Object = "": LBPSA0.Object = TBref01
ListBoxPTD101.Clear: ListBoxPTD121.Clear
If LBPSA0.listIndex < 0 Then Exit Sub
Dim tsay As Integer
tsay = Len(LBPSA0.List(LBPSA0.listIndex, 1)) - Len(Replace(LBPSA0.List(LBPSA0.listIndex, 1), ";", ""))
ayır = Split(LBPSA0.List(LBPSA0.listIndex, 1), ";")
For n = 0 To tsay
If ayır(n) <> "" Then ListBoxPTD101.AddItem ayır(n)
Next n
End Sub
Sub tipmontajsarfacıklamalar()
On Error Resume Next
Dim Rky As String, Dosyam As String
    Dim Ert As Long, satır As Long, i As Long
    Dim ayır As Variant
    Dim tsay As Integer, n As Integer
    Dosyam = "C:\Belgelerim\Cemex\Ayarlar\Montaj ve Sarf\Pano Tip Montaj ve Sarf Çarpanlar.txt"
    Ert = FreeFile
    Open Dosyam For Input As #Ert
    If Err.Number <> 0 Then
        MsgBox "Pano Tip Montaj ve Sarf Çarpanlar.txt" & " Dosyası Bulunamadı !", vbCritical, "Hata !"
        Exit Sub
    End If
    On Error GoTo 0
    satır = 1
    LBPSA0.Clear
Do While Not EOF(Ert)
        Line Input #Ert, Rky
        ayır = Split(Rky, ";")
        LBPSA0.AddItem ayır(0)
        LBPSA0.List(LBPSA0.ListCount - 1, 1) = Replace(Rky, ayır(0) & ";", "")
'tsay = Len(Rky) - Len(Replace(Rky, ";", ""))
'For n = 1 To tsay '
'If UBound(ayır) <> 0 Then LBPSA0.List(satır - 1, n) = ayır(n)
'Next n
     satır = satır + 1
Loop
Close #Ert
End Sub
Sub montajtipler01xxx()
On Error Resume Next
LBPSA0.Object = "": LBPSA0.Object = TBref01
ListBoxPTD101.Clear
If LBPSA0.listIndex < 0 Then ListBoxPTD101.Clear: Exit Sub
For n = 1 To 9
If LBPSA0.List(LBPSA0.listIndex, n) <> "" Then ListBoxPTD101.AddItem LBPSA0.List(LBPSA0.listIndex, n)
Next n
End Sub
Sub tipmontajsarfacıklamalarxxx()
On Error Resume Next
Dim Rky As String, Dosyam As String
    Dim Ert As Long, satır As Long, i As Long
    Dim ayır As Variant
    Dim tsay As Integer, n As Integer
    Dosyam = "C:\Belgelerim\Cemex\Ayarlar\Montaj ve Sarf\Pano Tip Montaj ve Sarf Çarpanlar.txt"
    Ert = FreeFile
    Open Dosyam For Input As #Ert
    If Err.Number <> 0 Then
        MsgBox "Pano Tip Montaj ve Sarf Çarpanlar.txt" & " Dosyası Bulunamadı !", vbCritical, "Hata !"
        Exit Sub
    End If
    On Error GoTo 0
    satır = 1
    LBPSA0.Clear
Do While Not EOF(Ert)
        Line Input #Ert, Rky
        ayır = Split(Rky, ";")
        LBPSA0.AddItem ayır(i)
tsay = Len(Rky) - Len(Replace(Rky, ";", ""))
For n = 1 To tsay '
If UBound(ayır) <> 0 Then LBPSA0.List(satır - 1, n) = ayır(n)
Next n
     satır = satır + 1
Loop
Close #Ert
End Sub
Private Sub ListBoxPTD121_Click() '2021 +21 PANO
On Error GoTo hata
TBAMS11 = "": LBAMS11 = ""
LBPSA1.Object = "": LBPSA1.Object = ListBoxPTD121.List(ListBoxPTD121.listIndex): TBAMS11 = ListBoxPTD121.List(ListBoxPTD121.listIndex)
If Not LBPSA1.List(LBPSA1.listIndex, 1) = "" Then
yo1 = LBPSA1.List(LBPSA1.listIndex, 1)
If Right(yo1, 1) = "%" Then yo1 = Replace(yo1, "%", "") & "/100"
ayır1 = Split(yo1, "/")(0): ayır2 = Split(yo1, "/")(1)
LBLPSA41.BackColor = &HD0E2CD
LBLPSA41.Top = 145 - (145 * CDbl(ayır1) / CDbl(ayır2))
End If
If LBPSA01.ListCount > 0 Then

LBPSA01.Object = "": LBPSA01.Object = ListBoxPTD121.List(ListBoxPTD121.listIndex)
mcp = "": scp = "": mcp = LBPSA01.List(LBPSA01.listIndex, 1): scp = LBPSA01.List(LBPSA01.listIndex, 2)
If Not mcp = "" Then LBAMS11 = "M(x)" & mcp & " - " & "S(x)" & scp
If ListBoxPTD121.listIndex >= 0 Then TBPMS01 = ListBoxPTD121.List(ListBoxPTD121.listIndex) & TBPM09 Else TBPMS01 = TBPM09
End If
Exit Sub
hata:
If LBPSA01.Object = "" Then: TBPMS01 = TBAMS11 & TBPM09
End Sub
Private Sub CommandButton28_Click()
If ListBoxP5.listIndex < 0 Then Exit Sub
Call gir1
End Sub
Private Sub SpinButton101_Change() 'MİKTAR DEĞİŞİMİ '2021 +21 PANO
On Error Resume Next
TBmiktar01.Value = SpinButton101.Value
TBbfiyat01.Value = Format(TextBoxFY01.Value * SpinButton101.Value, "#,##0.00")
End Sub
Private Sub TBmiktar01_Change() 'MİKTAR DEĞİŞİMİ '2021 +21 PANO
On Error Resume Next
SpinButton101.Value = TBmiktar01.Value
End Sub
Private Sub TextBox22_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
On Error Resume Next
If Len(TextBox22) <= 5 Then ListBoxP5.Clear: Exit Sub
Call TBox22
Call montajtipler01
Frame42.Enabled = False: TBskd01 = "": TextBoxFY01 = 0: TBmiktar01 = 1
If ListBoxPTD101.ListCount > 0 Then ListBoxPTD101.Selected(1) = True
ListBoxP5.Height = 145: ListBoxP5.IntegralHeight = True
End Sub
Sub TBox22()
On Error Resume Next
Dim ad, deg
Dim b, c
Set msayfa = Workbooks(pmlz).Worksheets("Sayfa1")
ad = TextBox22.Text
ListBoxP5.Clear
X = 0
deg = ""
Set c = msayfa.Range("B2:B65000").Find(ad, LookAt:=xlPart)
If Not c Is Nothing Then
b = c.Address
Do
If c.row <> deg Then
     ListBoxP5.AddItem msayfa.Cells(c.row, 2)
     ListBoxP5.List(ListBoxP5.ListCount - 1, 1) = msayfa.Cells(c.row, 3)
     ListBoxP5.List(ListBoxP5.ListCount - 1, 2) = msayfa.Cells(c.row, 6)
     ListBoxP5.List(ListBoxP5.ListCount - 1, 3) = msayfa.Cells(c.row, 2).row
deg = c.row
Set c = msayfa.Range("B2:B65000").FindNext(c)
End If
Loop While Not c Is Nothing And c.Address <> b
End If
End Sub
Private Sub Image444_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single) '2021 +21 PANO MONTAJ
On Error Resume Next
CreateObject("Shell.Application").Open "C:\Belgelerim\Cemex\Ayarlar\Montaj ve Sarf\Montaj ve Sarf Açıklamalar.txt"
End Sub
Private Sub LBA_Click()
If MultiPage2.Value = 0 Then
MultiPage2.Value = 1
If ListBoxP1.listIndex >= 0 Then Label494.Tag = ListBoxP1.listIndex
Label486 = "-": Label487 = "-"
ListBoxP2.Clear: ListBoxP3.Clear: ListBoxP4.Clear: ListBoxP5.Clear: ListBoxPTD101.Clear: ListBoxPTD121.Clear
If ListBoxP1.listIndex >= 0 Then ListBoxP1.Selected(ListBoxP1.listIndex) = False
Else
MultiPage2.Value = 0
If Not Label494.Tag = "" Then ListBoxP1.Selected(Label494.Tag) = True
End If
End Sub