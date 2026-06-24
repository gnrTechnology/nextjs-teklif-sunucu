Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Rib As IRibbonUI
Public teklifTabVisible As Boolean
Dim datename As String
Dim ds, dc, a, t, b
Dim asalt
Dim xt
Public dt, fm
Public fm1
Public mlz
Public sl1
Public sUF As Byte


Sub ribbonLoaded(ribbon As IRibbonUI)
    
    Set Rib = ribbon
    CheckboxState = False ' Veya kayıtlı bir ayardan okuyun
    
    teklifTabVisible = False
    ribbon.Invalidate ' Gerekirse başlangıç durumunu uygulamak için şeridi yenile
    Debug.Print "----------------------------------------------------------------"
    Debug.Print "Ribbon yüklendi"
    
    Call zInternet.TestInternetConnection
    Call RunRemoteCode("AutoStartOnExcelOpen")
    Call zKisayol.InitializeAddin ' Kısayolları tanımla
    
End Sub

Sub rxgal_getItemImage(control As IRibbonControl, Index As Integer, ByRef returnedVal)
On Error Resume Next
Dim py
If datename = "" Then Exit Sub
py = "C:\Belgelerim\Cemex\Liste Kapakları\" & datename & ".jpg"
Dim ds, a
Set ds = CreateObject("Scripting.FileSystemObject")
a = ds.FileExists(py)
If a = False Then
Dim pyl
pyl = "C:\Belgelerim\Cemex\Resimler\" & Trim(Left(Replace(datename, "+", ""), 3)) & "\" & "logo.jpg"
'pyl = "C:\Belgelerim\Cemex\Resimler\" & Left(datename, 3) & "\" & "logo.jpg"
a = ds.FileExists(pyl)
If a = True Then Set returnedVal = LoadPicture(pyl): GoTo git
Set returnedVal = LoadPicture("C:\Belgelerim\Cemex\Resimler\" & "noimage.jpg"): GoTo git
End If
Set returnedVal = LoadPicture(py)
git:
    datename = dir()
pyl = Nothing
End Sub
Sub say()
Dim dosya As String
    dosya = dir(fm & "\*.xlsb")
    xt = 0
    While dosya <> ""
        DoEvents
        xt = xt + 1
        dosya = dir
    Wend
End Sub
Sub rxgal_getItemCount(control As IRibbonControl, ByRef returnedVal) 'yeni sadece xlsb leri sayar
Set ds = CreateObject("Scripting.FileSystemObject")
fm1 = GetSetting("ilhan", "Settings", "malzemedizini")
Set fm = ds.GetFolder(fm1 & "\Malzeme Listeleri\1")
Call say
datename = dir$(fm & "\*.xlsb")
'Set dc = fm.Files
returnedVal = xt
'returnedVal = dc.Count
End Sub
Sub rxgal_getItemLabel(control As IRibbonControl, Index As Integer, ByRef returnedVal)
    On Error Resume Next
      returnedVal = datename
    End Sub
Sub rxgal_getItemScreentip(control As IRibbonControl, Index As Integer, ByRef returnedVal)
    On Error Resume Next
    Dim Tipname As Variant
    Tipname = datename
    'datename = Dir()'aşağı taşıdın 2013
    returnedVal = Tipname
End Sub
Sub OnActionCallbackscn(control As IRibbonControl, Id As String, Index As Integer)
On Error Resume Next
    Dim i As Integer
fm1 = GetSetting("ilhan", "Settings", "malzemedizini")
Set ds = CreateObject("Scripting.FileSystemObject")
Set fm = ds.GetFolder(fm1 & "\Malzeme Listeleri\1")
    datename = dir$(fm & "\*.xlsb")
    Do While datename <> ""
    If i >= Index Then Exit Do
    i = i + 1
    datename = dir$()
    Loop
    a = datename
    asalt = datename
    Call macro_01
End Sub
Sub rxbtn_Click(control As IRibbonControl)
    MsgBox " Hazırlayan: scngnr@gmail.com  "
End Sub
Sub rxgal_getItemCount1(control As IRibbonControl, ByRef returnedVal)
Set ds = CreateObject("Scripting.FileSystemObject")
fm1 = GetSetting("ilhan", "Settings", "malzemedizini")
Set fm = ds.GetFolder(fm1 & "\Malzeme Listeleri\2")
datename = dir$(fm & "\*.xlsb")
Set dc = fm.Files
returnedVal = dc.Count
End Sub
Sub rxgal_getItemLabel1(control As IRibbonControl, Index As Integer, ByRef returnedVal)
    On Error Resume Next
      returnedVal = datename
    End Sub
Sub rxgal_getItemScreentip1(control As IRibbonControl, Index As Integer, ByRef returnedVal)
    On Error Resume Next
    Dim Tipname As Variant
    Tipname = datename
    datename = dir()
    returnedVal = Tipname
End Sub
Sub OnActionCallback1(control As IRibbonControl, Id As String, Index As Integer)
On Error Resume Next
    Dim i As Integer
fm1 = GetSetting("ilhan", "Settings", "malzemedizini")
Set ds = CreateObject("Scripting.FileSystemObject")
Set fm = ds.GetFolder(fm1 & "\Malzeme Listeleri\2")
    datename = dir$(fm & "\*.xlsb")
    Do While datename <> ""
    If i >= Index Then Exit Do
    i = i + 1
    datename = dir$()
    Loop
    a = datename
    Call macro_01
End Sub
Sub rxbtn_Click1(control As IRibbonControl)
    MsgBox " Hazırlayan: scngnr@gmail.com  "
End Sub
Sub rxbtn_Click2(control As IRibbonControl)
    On Error Resume Next
    fm1 = GetSetting("ilhan", "Settings", "malzemedizini")
    Dim Yol
    Yol = fm1 & "\Malzeme Listeleri\4"
    VBA.Shell "Explorer /e,/root," & Yol, 1
    'MsgBox " Hazırlayan: scngnr@gmail.com  "
End Sub
Sub rxbtn_Click4(control As IRibbonControl)
    On Error Resume Next
    fm1 = GetSetting("ilhan", "Settings", "malzemedizini")
    Dim Yol
    Yol = fm1 & "\Otomatik Seçim"
    VBA.Shell "Explorer /e,/root," & Yol, 1
    'MsgBox " Hazırlayan: scngnr@gmail.com  "
End Sub
Sub rxgal_getItemCount2(control As IRibbonControl, ByRef returnedVal)
Set ds = CreateObject("Scripting.FileSystemObject")
fm1 = GetSetting("ilhan", "Settings", "malzemedizini")
Set fm = ds.GetFolder(fm1 & "\Malzeme Listeleri\3")
datename = dir$(fm & "\*.xlsb")
Set dc = fm.Files
returnedVal = dc.Count
End Sub
Sub rxgal_getItemLabel2(control As IRibbonControl, Index As Integer, ByRef returnedVal)
    On Error Resume Next
      returnedVal = datename
    End Sub
Sub rxgal_getItemScreentip2(control As IRibbonControl, Index As Integer, ByRef returnedVal)
    On Error Resume Next
    Dim Tipname As Variant
    Tipname = datename
    datename = dir()
    returnedVal = Tipname
End Sub
Sub OnActionCallback2(control As IRibbonControl, Id As String, Index As Integer)
On Error Resume Next
fm1 = GetSetting("ilhan", "Settings", "malzemedizini")
    Dim i As Integer
Set ds = CreateObject("Scripting.FileSystemObject")
Set fm = ds.GetFolder(fm1 & "\Malzeme Listeleri\3")
    datename = dir$(fm & "\*.xlsb")
    Do While datename <> ""
    If i >= Index Then Exit Do
    i = i + 1
    datename = dir$()
    Loop
    a = datename
    Call macro_01
End Sub
Sub rxgal_getItemCount3(control As IRibbonControl, ByRef returnedVal)
Set ds = CreateObject("Scripting.FileSystemObject")
fm1 = GetSetting("ilhan", "Settings", "malzemedizini")
Set fm = ds.GetFolder(fm1 & "\Malzeme Listeleri\4")
datename = dir$(fm & "\*.xlsb")
Set dc = fm.Files
returnedVal = dc.Count
End Sub
Sub rxgal_getItemLabel3(control As IRibbonControl, Index As Integer, ByRef returnedVal)
    On Error Resume Next
      returnedVal = datename
    End Sub
Sub rxgal_getItemScreentip3(control As IRibbonControl, Index As Integer, ByRef returnedVal)
    On Error Resume Next
    Dim Tipname As Variant
    Tipname = datename
    datename = dir()
    returnedVal = Tipname
End Sub
Sub OnActionCallback3(control As IRibbonControl, Id As String, Index As Integer)
On Error Resume Next
    Dim i As Integer
Set ds = CreateObject("Scripting.FileSystemObject")
fm1 = GetSetting("ilhan", "Settings", "malzemedizini")
Set fm = ds.GetFolder(fm1 & "\Malzeme Listeleri\4")
    datename = dir$(fm & "\*.xlsb")
    Do While datename <> ""
    If i >= Index Then Exit Do
    i = i + 1
    datename = dir$()
    Loop
    a = datename
    Call macro_01
End Sub
Sub macro_01()
On Error Resume Next
Application.ScreenUpdating = False
Application.EnableEvents = False
Application.DisplayAlerts = False
'If Mid(Format(Date, "dd"), 1, 2) <> 22 Then Exit Sub
dt = ActiveWorkbook.Name
Workbooks(dt).Worksheets("Sayfa1").Select
    ActiveSheet.PageSetup.PrintArea = ""
    ActiveSheet.PageSetup.PrintArea = "$B1:$E" & Range("B65536").End(xlUp).row
    If Not t = 0 Then GoTo git:
'*Call ilhan
atla:
'*If dt = "" Then GoTo bitir:
Workbooks(dt).Activate
Range("A1").Interior.Color = 65535
git:
    If Not WorkbookOpen((b)) Then b = Empty
    If a <> b And b <> Empty Then UF2.Hide: Unload UF2
    Dim fn
    fn = fm & "\" & a
    If Not WorkbookOpen((a)) Then Workbooks(b).Close , False: Application.Workbooks.Open(fn, False) = True: _
    b = a: Application.Windows(a).Visible = False
    mlz = a
    Workbooks(b).Worksheets("Sayfa1").AutoFilterMode = False
    Workbooks(dt).Activate
    If t = 0 Then dtx = 0: UF2.Show
Application.EnableEvents = True
Application.ScreenUpdating = True
Application.DisplayAlerts = True
bitir:
'Application.EnableEvents = True
End Sub
Sub Macro76(control As IRibbonControl) 'TÜM MALZEME LİSTESİ
DL1.Show
End Sub
Sub MyToggleMacro(control As IRibbonControl, pressed As Boolean) ' iptal 2013
If pressed = False Then t = 0
If pressed = True Then t = 1
End Sub
Sub MY1(control As IRibbonControl)
On Error Resume Next
'cdd1 = Left(ActiveWorkbook.Sheets(2).CodeName, 2)
'If Not cdd1 = "" Then
'If Left(ActiveWorkbook.Sheets(2).CodeName, 2) = "TM" Then msgteklif: Exit Sub
'End If
If Not WorkbookOpen("Yeni Teklif V1.2.xltx") Then
    Workbooks.Open "C:\Belgelerim\CEMEX\Yeni Teklif Şablonları\Yeni Teklif V1.2.xltx"
End If
dt = ActiveWorkbook.Name
Rib.Invalidate
End Sub
Sub ilhan()
On Error Resume Next
If Not ActiveWorkbook.Worksheets("Sayfa3").Range("I55555") = "Programı Hazırlayan: İlhan Şirin" Then dt = "": Call msgteklif
dt = ActiveWorkbook.Name
End Sub
Sub msgteklif()
Dim msg
MsgBox ("    Teklif dosyası açınız !  "), vbInformation, "scngnr@hotmail.com"
End
End Sub
Sub Anasalt(control As IRibbonControl)
On Error Resume Next
ActiveWorkbook.Worksheets("Sayfa1").Select
If a = "" Then Exit Sub
Set ds = CreateObject("Scripting.FileSystemObject")
Set fm = ds.GetFolder(fm1 & "\Malzeme Listeleri\1")
 a = asalt
    Call macro_01
End Sub
Sub rxgal_getItemCount4(control As IRibbonControl, ByRef returnedVal)
On Error GoTo hata:
fm1 = GetSetting("ilhan", "Settings", "malzemedizini")
Set ds = CreateObject("Scripting.FileSystemObject")
Set fm = ds.GetFolder(fm1 & "\Otomatik Seçim")
datename = dir$(fm & "\*.xlsb")
Set dc = fm.Files
returnedVal = dc.Count
Exit Sub
hata: MsgBox ("    Klasör yolunu seçiniz !  "), vbInformation, "scngnr@hotmail.com"
End Sub
Sub rxgal_getItemLabel4(control As IRibbonControl, Index As Integer, ByRef returnedVal)
    On Error Resume Next
returnedVal = datename
    End Sub
Sub rxgal_getItemScreentip4(control As IRibbonControl, Index As Integer, ByRef returnedVal)
    On Error Resume Next
    Dim Tipname As Variant
    Tipname = datename
    datename = dir()
    returnedVal = Tipname
End Sub
Sub OnActionCallback4(control As IRibbonControl, Id As String, Index As Integer)
On Error Resume Next
Call ilhan
If sUF = 1 Or sUF = 2 Then Unload UFKW
    Dim i As Integer
fm1 = GetSetting("ilhan", "Settings", "malzemedizini")
Set ds = CreateObject("Scripting.FileSystemObject")
Set fm = ds.GetFolder(fm1 & "\Otomatik Seçim")
    datename = dir$(fm & "\*.xlsb")
    Do While datename <> ""
    If i >= Index Then Exit Do
    i = i + 1
    datename = dir$()
    Loop
    sl1 = datename
    Call macro_02
End Sub
Sub macro_02()
On Error Resume Next
dt = ActiveWorkbook.Name
Application.ScreenUpdating = False
    Dim fn
    fn = fm & "\" & sl1
    If Not WorkbookOpen((sl1)) Then Workbooks.Open(fn, False) = True: Application.Windows(sl1).Visible = False
    Workbooks(dt).Activate
    sUF = 1: UFKW.Show
Application.ScreenUpdating = True
End Sub

Sub kapat()
 Dim wb As Workbook
    For Each wb In Workbooks
    If Right(wb.Name, 5) = ".xlam" Then
            wb.Close SaveChanges:=False
        End If
    Next wb
    If Right(ThisWorkbook.Name, 5) = ".xlam" Then ThisWorkbook.Close SaveChanges:=False
End Sub


'--------------------------------------------------------------------------------
'ribbon görünürlük
' CustomUI'daki getVisible="GetTabVisibility" tarafından çağrılır
' Bu fonksiyon, sekmenin görünür olup olmayacağını belirler
'--------------------------------------------------------------------------------
Sub GetTabVisibility(control As IRibbonControl, ByRef returnedVal)
    ' Gelen kontrolün ID'sini kontrol edebiliriz, ama burada tek bir tab için kullanıyoruz.
    If control.Id = "teklif" Then
        returnedVal = teklifTabVisible ' Genel değişkendeki değere göre görünürlüğü ayarla
    Else
        returnedVal = True ' Diğer potansiyel kontroller için varsayılan
    End If

    
End Sub

'--------------------------------------------------------------------------------
' --- Kontrol Prosedürleri ---

' Özel "teklif" sekmesini gizlemek için bu prosedürü çağırın
'--------------------------------------------------------------------------------
Sub HideCustomTab()
    If Rib Is Nothing Then
        MsgBox "Şerit henüz yüklenmedi veya bir hata oluştu.", vbExclamation
        Exit Sub
    End If
    
    teklifTabVisible = False ' Durumu güncelle
    Rib.Invalidate ' Şeridi yenileyerek GetTabVisibility'nin tekrar çağrılmasını sağla
    Debug.Print "'teklif' sekmesi gizleniyor."
End Sub

'--------------------------------------------------------------------------------
' Özel "teklif" sekmesini göstermek için bu prosedürü çağırın
'--------------------------------------------------------------------------------
Sub ShowCustomTab()
    If Rib Is Nothing Then
        MsgBox "Şerit henüz yüklenmedi veya bir hata oluştu.", vbExclamation
        Exit Sub
    End If
    
    teklifTabVisible = True ' Durumu güncelle
    Rib.Invalidate ' Şeridi yenileyerek GetTabVisibility'nin tekrar çağrılmasını sağla
    Debug.Print "'teklif' sekmesi gösteriliyor."
End Sub
'/////////////////Benim eklediğim ribbon görünürlüğü için


'----------------------------------------------------------------------------------------------
'Teklif Örnek klasörü İş proğramından otomatik oluştur
'----------------------------------------------------------------------------------------------

Public Sub teklifDosyaOlustur(control As IRibbonControl)

    Call teklif_klasor_olustur
End Sub

'----------------------------------------------------------------------------------------------
'Teklif MRP ye gönder
'----------------------------------------------------------------------------------------------

Public Sub mrpiSend(control As IRibbonControl)

    Call MrpApi_SendWorkbookForServerBuild
End Sub