'--- Standart Modül ---
Option Explicit

' Class nesnesini tutmak için Public değişken (Makro dosyası açık olduğu sürece yaşar)
Public SayfaIzleyiciNesnesi As clsSayfaIzleyici

'şerit üzerinden tetklenir. B sütunundaki boş malzeme satırını otomatik doldurur.
'#şerit üzerine eklenen checkbox ile tetiklenir. açıldığında
'B sütunda yazılan kodlar malzeme listelerinde bulunuyor ise fiyat açıklama, iskonto gibi bilgileri getirir

Sub Checkbox_OnAction(control As IRibbonControl, pressed As Boolean)
    CheckboxState = pressed ' Yeni durumu değişkende sakla
    
    Select Case control.Id
        Case "chkYeniSecenek"
            If pressed Then
                Call Otomatik_Dolum_open
                'MsgBox "Özel Seçenek Aktif!"
                ' Checkbox işaretlendiğinde yapılacak işlemler
            Else
                Call SayfaIzlemeyiDurdur
                'MsgBox "Özel Seçenek Pasif!"
                ' Checkbox işareti kaldırıldığında yapılacak işlemler
            End If
    End Select
    
    ' Belki başka kontrolleri etkinleştirmek/devre dışı bırakmak istersiniz
    ' myRibbon.InvalidateControl "btnYeni1" ' Örn: btnYeni1'in durumunu güncelle
End Sub

Sub Otomatik_Dolum_open()
    Dim wbAcilanSablon As Workbook
    Dim sablonDosyaYolu As String
    Dim izlenecekSayfaAdi As String
    
    'aktif dosya yolunu al
    Call GetActiveWorkbookPath
    
    sablonDosyaYolu = GetSetting("sercan", "fileOpenWorkBooks", "nowOpenPropsFile")
    izlenecekSayfaAdi = "Sayfa1" ' Veya şablondaki ilgili sayfanın adı

    ' ... (Gerekirse diğer kodlar) ...

    ' Şablonu aç
    Set wbAcilanSablon = Workbooks.Open(sablonDosyaYolu)

    ' Eğer başarıyla açıldıysa izlemeyi başlat
    If Not wbAcilanSablon Is Nothing Then
       Call SayfaIzlemeyiBaslat(wbAcilanSablon, izlenecekSayfaAdi)
    End If
End Sub


Sub SayfaIzlemeyiBaslat(ByVal wbTemplate As Workbook, ByVal sheetNameToWatch As String)
    ' wbTemplate: Olayları izlenecek şablon çalışma kitabı
    ' sheetNameToWatch: Olayları izlenecek sayfanın adı (örn. "Sayfa1" veya aktif sayfanın adı)

    Dim wsToWatch As Worksheet

    On Error Resume Next ' Sayfa bulunamazsa hatayı yakala
    Set wsToWatch = wbTemplate.Worksheets(sheetNameToWatch)
    On Error GoTo 0 ' Normal hata işlemeye dön

    If wsToWatch Is Nothing Then
        MsgBox "'" & sheetNameToWatch & "' sayfası '" & wbTemplate.Name & "' kitabında bulunamadı. İzleme başlatılamadı."
        Exit Sub
    End If

    ' Mevcut bir izleyici varsa temizle (güvenlik için)
    Set SayfaIzleyiciNesnesi = Nothing

    ' Yeni Class nesnesi oluştur ve hedef sayfayı ata
    Set SayfaIzleyiciNesnesi = New clsSayfaIzleyici
    SayfaIzleyiciNesnesi.Init wsToWatch

    Debug.Print "'" & wsToWatch.Name & "' sayfası için izleme başlatıldı."

End Sub

Sub SayfaIzlemeyiDurdur()
    ' Şablon kitabı kapatıldığında veya izleme durdurulmak istendiğinde çağrılır
    Set SayfaIzleyiciNesnesi = Nothing
    Debug.Print "Sayfa izleme durduruldu."
End Sub


Sub ExcelMalzemeListeleriBirlestir()

    Dim fso As Object          ' FileSystemObject
    Dim anaKlasor As Object
    Dim altKlasor As Object
    Dim dosya As Object
    Dim wbKaynak As Workbook
    Dim wsKaynak As Worksheet
    Dim wbHedef As Workbook
    Dim wsHedef As Worksheet
    ' Dim rngKaynak As Range ' Kaldırıldı
    ' Dim sonSatirHedef As Long ' Kaldırıldı - rHedef kullanılacak
    Dim sonSatirKaynak As Long
    Dim rKaynak As Long        ' Kaynak satır sayacı
    Dim rHedef As Long         ' Hedef satır sayacı (Genel sayaç olacak)
    Dim startRowKaynak As Long ' Kaynakta veri başlangıç satırı (başlıkları atlamak için)
    Dim anaKlasorYolu As String
    Dim hedefDosyaYolu As String
    Dim hedefSayfaAdi As String
    Dim kaynakSayfaAdi As String
    Dim currentReferans As Variant ' K sütunundaki son geçerli değeri tutmak için

    ' --- Ayarlar ---
    anaKlasorYolu = "C:\Belgelerim\cemex\Malzeme Listeleri" ' Ana klasörünüzün yolu
    hedefDosyaYolu = anaKlasorYolu & "\TümExcelListeleri.xlsb"
    hedefSayfaAdi = "Sayfa1" ' Hedef dosyada verilerin ekleneceği sayfa adı
    kaynakSayfaAdi = "Sayfa1" ' Kaynak dosyalarda verilerin alınacağı sayfa adı
    startRowKaynak = 2       ' Kaynak dosyalarda verilerin başladığı satır (Görseldeki gibi başlık varsayıldı)
    ' -------------

    On Error GoTo HataYonetimi

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False

    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FolderExists(anaKlasorYolu) Then
        MsgBox "Ana klasör bulunamadı: " & anaKlasorYolu, vbCritical, "Hata"
        GoTo Cikis
    End If
    Set anaKlasor = fso.GetFolder(anaKlasorYolu)

    ' Hedef dosya yönetimi
    If Not fso.FileExists(hedefDosyaYolu) Then
        Set wbHedef = Workbooks.Add
        On Error Resume Next
        wbHedef.Worksheets(1).Name = hedefSayfaAdi
        On Error GoTo HataYonetimi
        wbHedef.SaveAs fileName:=hedefDosyaYolu, FileFormat:=xlExcel12
        If Err.Number <> 0 Then
             MsgBox "Hedef dosya oluşturulamadı veya kaydedilemedi: " & hedefDosyaYolu & vbCrLf & Err.Description, vbCritical
             wbHedef.Close False
             GoTo Cikis
        End If
        Debug.Print "Hedef dosya oluşturuldu: " & hedefDosyaYolu
    Else
        On Error Resume Next
        Set wbHedef = Workbooks.Open(hedefDosyaYolu)
        If Err.Number <> 0 Then
            MsgBox "Hedef dosya açılamadı: " & hedefDosyaYolu & vbCrLf & "Lütfen dosyanın başka bir yerde açık olmadığından emin olun.", vbCritical
            GoTo Cikis
        End If
        Debug.Print "Hedef dosya açıldı: " & hedefDosyaYolu
        On Error GoTo HataYonetimi
    End If

    On Error Resume Next
    Set wsHedef = wbHedef.Worksheets(hedefSayfaAdi)
    If Err.Number <> 0 Then
        MsgBox "Hedef dosyada '" & hedefSayfaAdi & "' sayfası bulunamadı!", vbCritical
        wbHedef.Close False
        GoTo Cikis
    End If
    On Error GoTo HataYonetimi

    ' *** YENİ ADIM: Hedef sayfayı temizle ***
    wsHedef.Cells.Clear
    Debug.Print "Hedef sayfa (" & wsHedef.Name & ") temizlendi."

    ' *** YENİ ADIM: Başlıkları tekrar yaz ***
    With wsHedef
        .Cells(1, "A").Value = "REFERANS"     ' Kaynak K sütunu başlığı
        .Cells(1, "B").Value = "SİPARİŞ KODU" ' Kaynak B sütunu başlığı
        .Cells(1, "C").Value = "AÇIKLAMA"      ' Kaynak C sütunu başlığı
        .Cells(1, "D").Value = "ÜRETİCİ"       ' Kaynak D sütunu başlığı
        .Cells(1, "E").Value = "BİRİM"         ' Kaynak E sütunu başlığı
        .Cells(1, "F").Value = "BR. FİYAT"     ' Kaynak F sütunu başlığı
        .Cells(1, "G").Value = "İSKONTO"       ' Kaynak G sütunu başlığı
        .Cells(1, "H").Value = "ADAM/DK"       ' Kaynak H sütunu başlığı
        .Cells(1, "I").Value = "BOYUT"         ' Kaynak I sütunu başlığı
        ' Gerekirse J, K, L vb. başlıkları ekle
        .Rows(1).Font.Bold = True ' Başlıkları kalın yap
    End With
    Debug.Print "Hedef sayfaya başlıklar yazıldı."
    ' *** BAŞLIK YAZMA SONU ***

    rHedef = 2 ' Veri yazılacak ilk satır her zaman 2 olacak (başlıklardan sonra)

    ' --- Alt Klasörleri ve Dosyaları Tara ---
    Debug.Print "Alt klasörler taranıyor..."
    For Each altKlasor In anaKlasor.SubFolders
        Debug.Print "İşlenen Alt Klasör: " & altKlasor.Name
        For Each dosya In altKlasor.Files
            If LCase(fso.GetExtensionName(dosya.Name)) = "xlsb" And _
               LCase(dosya.path) <> LCase(hedefDosyaYolu) Then

                Debug.Print "  İşlenen Kaynak Dosya: " & dosya.Name
                currentReferans = "" ' Her yeni kaynak dosya için referansı sıfırla

                On Error Resume Next
                Set wbKaynak = Workbooks.Open(dosya.path, UpdateLinks:=0, ReadOnly:=True)
                Application.Windows(wbKaynak.Name).Visible = False
                If Err.Number <> 0 Then
                    Debug.Print "    Hata: Kaynak dosya açılamadı - " & dosya.path & " - Hata: " & Err.Description
                    Err.Clear
                    GoTo SonrakiDosya
                End If
                On Error GoTo HataYonetimi

                On Error Resume Next
                Set wsKaynak = wbKaynak.Worksheets(kaynakSayfaAdi)
                If Err.Number <> 0 Then
                    Debug.Print "    Uyarı: '" & kaynakSayfaAdi & "' sayfası bulunamadı - " & dosya.path
                    wbKaynak.Close False
                    GoTo SonrakiDosya
                End If
                On Error GoTo HataYonetimi

                sonSatirKaynak = wsKaynak.Cells(wsKaynak.Rows.Count, "B").End(xlUp).row
                Debug.Print "    Veri kopyalanıyor. Hedef Başlangıç Satırı: " & rHedef & ", Kaynak Son Satır: " & sonSatirKaynak

                If sonSatirKaynak >= startRowKaynak Then
                    For rKaynak = startRowKaynak To sonSatirKaynak
                        ' A Sütunu (Referans) Mantığı
                        If Not IsEmpty(wsKaynak.Cells(rKaynak, "K").Value) Then
                            currentReferans = wsKaynak.Cells(rKaynak, "K").Value
                        End If
                        wsHedef.Cells(rHedef, "A").Value = currentReferans

                        ' B-I Sütunlarını Kopyala
                        wsHedef.Cells(rHedef, "B").Resize(1, 8).Value = wsKaynak.Cells(rKaynak, "B").Resize(1, 8).Value

                        ' Hedef satır sayacını artır
                        rHedef = rHedef + 1
                    Next rKaynak
                Else
                     Debug.Print "    Uyarı: Kaynak sayfada işlenecek veri bulunamadı (Satır " & startRowKaynak & "'dan sonra) - " & dosya.path
                End If

                wbKaynak.Close False
                Set wbKaynak = Nothing
                Set wsKaynak = Nothing

            End If
SonrakiDosya:
        Next dosya
    Next altKlasor

    ' Hedef dosyadaki son verinin yazıldığı satırdan sonrasını temizle (İsteğe bağlı, eski verilerden kalıntı olmaması için)
    ' On Error Resume Next ' Satır yoksa hata vermesin
    ' wsHedef.Rows(rHedef & ":" & wsHedef.Rows.Count).ClearContents
    ' On Error GoTo HataYonetimi

    wbHedef.Save
    MsgBox "İşlem tamamlandı!" & vbCrLf & "Veriler '" & hedefDosyaYolu & "' dosyasına eklendi.", vbInformation

Cikis:
    On Error Resume Next
    Set wsKaynak = Nothing
    If Not wbKaynak Is Nothing Then
        If wbKaynak.Name <> ThisWorkbook.Name Then wbKaynak.Close False
    End If
    Set wbKaynak = Nothing
    Set wsHedef = Nothing
    If Not wbHedef Is Nothing Then
         If wbHedef.Name <> ThisWorkbook.Name Then wbHedef.Close True ' Hedefi kaydet ve kapat
    End If
    Set wbHedef = Nothing
    Set dosya = Nothing
    Set altKlasor = Nothing
    Set anaKlasor = Nothing
    Set fso = Nothing

    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Exit Sub

HataYonetimi:
    MsgBox "Beklenmeyen bir hata oluştu:" & vbCrLf & _
           "Hata No: " & Err.Number & vbCrLf & _
           "Açıklama: " & Err.Description, vbCritical, "Makro Hatası"
    GoTo Cikis

End Sub

Sub MlzVeriEkleMacrosu3()

    Dim wsKaynak As Worksheet
    Dim wbMaster As Workbook
    Dim wsMaster As Worksheet
    Dim masterFilePath As String
    Dim masterSheetName As String
    Dim sonSatirKaynak As Long, sonSatirMaster As Long
    Dim i As Long, j As Long
    
    Dim dictMaster As Object
    Dim arrMasterValues As Variant
    Dim arrMasterFormatsF As Variant
    Dim arrMasterFormatsG As Variant
    Dim key As Variant
    Dim itemData(1 To 9) As Variant
    Dim retrievedData As Variant
    
    Dim bulunanSayac As Long
    Dim bfyt As String, kur As String, mkur As String
    Dim Ckar As Variant, nameo As Variant
    Dim cellBValue As String
    
    ' --- YENİ: Toplu Biçimlendirme için Range Nesneleri ---
    Dim rngProcessed As Range      ' İşlenen tüm satırlar (Genel formatlama)
    Dim rngFormatUSD As Range      ' F sütunu $ olacak satırlar
    Dim rngFormatEUR As Range      ' F sütunu € olacak satırlar
    Dim rngOTM As Range            ' OTM sayfa formatı uygulanacaklar
    Dim rngNonOTM As Range         ' OTM olmayan (formüllü) sayfa formatı uygulanacaklar
    Dim rngNonOTM_JX As Range      ' OTM olmayan J-X sütunları
    Dim rngNonOTM_WX As Range      ' OTM olmayan W-X sütunları (Tpbr için)
    ' --- YENİ SON ---
    
    Set wsKaynak = ActiveSheet
    masterFilePath = "C:\Belgelerim\cemex\Malzeme Listeleri\TümExcelListeleri.xlsb"
    masterSheetName = "Sayfa1"
    
    Set dictMaster = CreateObject("Scripting.Dictionary")
    dictMaster.CompareMode = vbTextCompare

    ' --- bfyt ayarı (Değişiklik yok) ---
    On Error Resume Next
    Ckar = ThisWorkbook.names("CkarO").RefersToR1C1
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo HataYonetimi_Makro
        If MsgBox("'CkarO' adlı ad tanımlı değil. Varsayılan ('Liste Fiyatı') oluşturulsun mu?", vbYesNo + vbQuestion) = vbYes Then
            ThisWorkbook.names.Add Name:="CkarO", RefersToR1C1:="=""Liste Fiyatı"""
            bfyt = "=RC[-6]"
        Else
            MsgBox "İşlem iptal edildi.", vbInformation
            Exit Sub
        End If
    Else
        On Error GoTo HataYonetimi_Makro
        If Ckar = "=""Net Fiyatı""" Then bfyt = "=RC[-1]" Else bfyt = "=RC[-6]"
    End If
    ' --- bfyt ayarlandı ---

    On Error GoTo HataYonetimi_Makro
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.StatusBar = "Ana malzeme listesi hafızaya yükleniyor..."

    If dir(masterFilePath) = "" Then GoTo DosyaYok
    
    On Error Resume Next
    Set wbMaster = Workbooks(dir(masterFilePath))
    If wbMaster Is Nothing Then
        Set wbMaster = Workbooks.Open(masterFilePath, ReadOnly:=True, UpdateLinks:=0)
        If wbMaster Is Nothing Then GoTo DosyaAcilamadi
    End If
    On Error GoTo HataYonetimi_Makro

    On Error Resume Next
    Set wsMaster = wbMaster.Worksheets(masterSheetName)
    If wsMaster Is Nothing Then GoTo SayfaYok
    On Error GoTo HataYonetimi_Makro

    ' === Ana Listeyi Oku ve Dictionary'ye Yükle (Değişiklik yok) ===
    sonSatirMaster = wsMaster.Cells(wsMaster.Rows.Count, "B").End(xlUp).row
    If sonSatirMaster < 2 Then GoTo VeriYok
    
    arrMasterValues = wsMaster.Range("A1:I" & sonSatirMaster).Value
    arrMasterFormatsF = wsMaster.Range("F1:F" & sonSatirMaster).NumberFormat
    arrMasterFormatsG = wsMaster.Range("G1:G" & sonSatirMaster).NumberFormat
    
    wbMaster.Close False
    Set wbMaster = Nothing
    Set wsMaster = Nothing
    
    For j = 2 To sonSatirMaster
        key = arrMasterValues(j, 2)
        If Not IsEmpty(key) And Not dictMaster.exists(key) Then
            itemData(1) = arrMasterValues(j, 1) ' A
            itemData(2) = arrMasterValues(j, 3) ' C
            itemData(3) = arrMasterValues(j, 4) ' D
            itemData(4) = arrMasterValues(j, 6) ' F
            itemData(5) = arrMasterValues(j, 7) ' G
            itemData(6) = arrMasterValues(j, 8) ' H
            itemData(7) = arrMasterValues(j, 9) ' I
            itemData(8) = arrMasterFormatsF(j, 1) ' F Format
            itemData(9) = arrMasterFormatsG(j, 1) ' G Format
            dictMaster.Add key, itemData
        End If
    Next j
    
    Erase arrMasterValues
    Erase arrMasterFormatsF
    Erase arrMasterFormatsG
    
    ' === Dictionary Yüklendi ===

    sonSatirKaynak = wsKaynak.Cells(wsKaynak.Rows.Count, "B").End(xlUp).row
    bulunanSayac = 0

    ' === ANA DÖNGÜ: SADECE VERİ VE FORMÜL YAZMA ===
    For i = 2 To sonSatirKaynak
        
        ' Hız için: Her 100 satırda bir durum çubuğunu güncelle
        If i Mod 100 = 0 Then
            Application.StatusBar = "Satır " & i & "/" & sonSatirKaynak & " işleniyor..."
        End If

        If Not IsEmpty(wsKaynak.Cells(i, "B").Value) And IsEmpty(wsKaynak.Cells(i, "A").Value) Then
            cellBValue = Trim(UCase(CStr(wsKaynak.Cells(i, "B").Value)))
            
            If cellBValue <> "BÖLÜM TOPLAMI:" And cellBValue <> "BÖLÜM ADI/NO:" Then
                lookupValue = wsKaynak.Cells(i, "B").Value

                If dictMaster.exists(lookupValue) Then
                    retrievedData = dictMaster(lookupValue)
                    bulunanSayac = bulunanSayac + 1
                    
                    ' --- BİÇİMLENDİRME NESNELERİNE EKLEME ---
                    If rngProcessed Is Nothing Then
                        Set rngProcessed = wsKaynak.Rows(i)
                    Else
                        Set rngProcessed = Union(rngProcessed, wsKaynak.Rows(i))
                    End If
                    ' ---

                    ' 1. VERİLERİ YAZ (Biçimlendirmesiz)
                    wsKaynak.Cells(i, "A").Value = retrievedData(1)
                    wsKaynak.Cells(i, "C").Value = retrievedData(2)
                    wsKaynak.Cells(i, "D").Value = retrievedData(3)
                    wsKaynak.Cells(i, "F").Value = retrievedData(4)
                    wsKaynak.Cells(i, "G").Value = retrievedData(5)
                    wsKaynak.Cells(i, "H").Value = retrievedData(6)
                    wsKaynak.Cells(i, "I").Value = retrievedData(7)
                    
                    ' 2. KUR'u belirle (Formatı henüz UYGULAMA, sadece grupla)
                    kur = ""
                    If retrievedData(8) = "#,##0.00 [$$-C0C]" Then
                        kur = "*Usd"
                        If rngFormatUSD Is Nothing Then Set rngFormatUSD = wsKaynak.Cells(i, "F") Else Set rngFormatUSD = Union(rngFormatUSD, wsKaynak.Cells(i, "F"))
                    ElseIf retrievedData(8) = "#,##0.00 [$€-1]" Then
                        kur = "*Eur"
                        If rngFormatEUR Is Nothing Then Set rngFormatEUR = wsKaynak.Cells(i, "F") Else Set rngFormatEUR = Union(rngFormatEUR, wsKaynak.Cells(i, "F"))
                    End If
                     
                    ' 3. FORMÜLLERİ YAZ
                    If Left(wsKaynak.CodeName, 3) <> "OTM" Then
                        ' OTM Değilse: Formülleri yaz ve OTM-değil grubuna ekle
                        If rngNonOTM Is Nothing Then Set rngNonOTM = wsKaynak.Rows(i) Else Set rngNonOTM = Union(rngNonOTM, wsKaynak.Rows(i))
                        
                        mkur = kur: If bfyt = "=RC[-1]" Then mkur = ""

                        Dim referansA As String
                        referansA = CStr(retrievedData(1)) ' A sütununa yazdığımız değeri kullan

                        If Left(referansA, 5) = "PM-MP" Then wsKaynak.Cells(i, "L").FormulaR1C1 = bfyt & "*Oisci/100" & mkur
                        ElseIf Left(referansA, 5) = "PM-MS" Then wsKaynak.Cells(i, "L").FormulaR1C1 = bfyt & "*Osarf/100" & mkur
                        ElseIf Left(referansA, 5) = "PM-MA" Then wsKaynak.Cells(i, "L").FormulaR1C1 = bfyt & "*Oamb/100" & mkur
                        ElseIf Left(referansA, 5) = "PM-MN" Then wsKaynak.Cells(i, "L").FormulaR1C1 = bfyt & "*Onak/100" & mkur
                        ElseIf Left(referansA, 5) = "PM-MB" Then wsKaynak.Cells(i, "L").FormulaR1C1 = bfyt & "*Obara/100" & mkur
                        ElseIf Left(referansA, 3) = "PP-" Then wsKaynak.Cells(i, "L").FormulaR1C1 = bfyt & "*Opano/100" & mkur
                        ElseIf Left(referansA, 3) = "PS-" Then
                            On Error Resume Next
                            nameo = ThisWorkbook.names("Opsac").RefersToR1C1
                            If Err.Number = 0 And Not IsEmpty(nameo) Then wsKaynak.Cells(i, "L").FormulaR1C1 = bfyt & "*Opsac/100" & mkur Else wsKaynak.Cells(i, "L").FormulaR1C1 = bfyt & "*Opano/100" & mkur
                            Err.Clear
                            On Error GoTo HataYonetimi_Makro
                        Else
                            wsKaynak.Cells(i, "L").FormulaR1C1 = bfyt & "*Osalt/100" & mkur ' Varsayılan
                        End If

                        ' Diğer Formüller (J, K, M-X)
                        wsKaynak.Cells(i, "J").FormulaR1C1 = "=RC[-2]*Ads/60"
                        wsKaynak.Cells(i, "N").FormulaR1C1 = "=RC[-3]*Oggid/100"
                        wsKaynak.Cells(i, "K").FormulaR1C1 = "=(RC[-5]-RC[-5]*RC[-4])" & kur
                        wsKaynak.Cells(i, "M").FormulaR1C1 = "=RC[-3]*Oisci/100"
                        wsKaynak.Cells(i, "O").FormulaR1C1 = "=RC[-10]*RC[-9]" & kur
                        wsKaynak.Cells(i, "P").FormulaR1C1 = "=RC[-11]*RC[-5]"
                        wsKaynak.Cells(i, "Q").FormulaR1C1 = "=RC[-12]*RC[-7]"
                        wsKaynak.Cells(i, "R").FormulaR1C1 = "=RC[-13]*RC[-6]"
                        wsKaynak.Cells(i, "S").FormulaR1C1 = "=RC[-14]*RC[-6]"
                        wsKaynak.Cells(i, "T").FormulaR1C1 = "=RC[-15]*RC[-12]/60"
                        wsKaynak.Cells(i, "U").FormulaR1C1 = "=RC[-7]*RC[-16]"
                        wsKaynak.Cells(i, "W").FormulaR1C1 = "=(RC[-13]+RC[-12]+RC[-11]+RC[-10]+RC[-9])*Dcki/Tpb"
                        wsKaynak.Cells(i, "X").FormulaR1C1 = "=RC[-19]*RC[-1]"
                    Else
                        ' OTM sayfasıysa: Formülleri temizle ve OTM grubuna ekle
                        If rngOTM Is Nothing Then Set rngOTM = wsKaynak.Range("J" & i & ":N" & i & ",O" & i & ":X" & i) Else Set rngOTM = Union(rngOTM, wsKaynak.Range("J" & i & ":N" & i & ",O" & i & ":X" & i))
                    End If
                    
                    ' --- DÖNGÜ İÇİNDEKİ TÜM BİÇİMLENDİRME KODLARI KALDIRILDI ---
                    
                End If
            End If
        End If
    Next i
    ' === ANA DÖNGÜ SONA ERDİ ===
    
    Application.StatusBar = "Toplu biçimlendirme uygulanıyor..."

    ' === TOPLU BİÇİMLENDİRME BAŞLANGICI ===
    
    ' Eğer hiç satır bulunamadıysa, formatlamayı atla
    If rngProcessed Is Nothing Then
        MsgBox "İşlenecek veri bulunamadı veya tüm veriler güncel.", vbInformation
        GoTo Cikis_Makro
    End If

    ' 1. OTM: J-X arasını temizle
    If Not rngOTM Is Nothing Then
        rngOTM.ClearContents
    End If

    ' 2. Genel Biçimlendirme (Tüm işlenen satırlar için)
    With rngProcessed.Range("A1:U1,W1:X1") ' V sütunu hariç
        .Borders.LineStyle = xlContinuous
        .Font.Bold = False
        .Font.ColorIndex = xlAutomatic
        .Font.Size = 9
        .Font.Name = "Arial"
    End With
    
    ' 3. Hizalama
    With rngProcessed.Range("A1:D1")
        .HorizontalAlignment = xlLeft
        .NumberFormat = "@" ' Metin
    End With
    With rngProcessed.Range("E1:U1,W1:X1") ' V hariç
        .HorizontalAlignment = xlRight
    End With
    
    ' 4. Özel Sayı Biçimleri
    rngProcessed.Range("E1").NumberFormat = "#,##0" ' Miktar
    rngProcessed.Range("F1").NumberFormat = "#,##0.00" ' Fiyat (Varsayılan TL)
    rngProcessed.Range("G1").NumberFormat = "0.0%"    ' İskonto
    rngProcessed.Range("H1").NumberFormat = "#,##0"    ' Adam/dk
    
    ' 5. Kur Renkleri
    If Not rngFormatUSD Is Nothing Then
        rngFormatUSD.Font.ColorIndex = 3
        ' İsteğe bağlı: Para birimi formatını da burada toplu ata
        ' rngFormatUSD.NumberFormat = "#,##0.00 [$$-C0C]"
    End If
    If Not rngFormatEUR Is Nothing Then
        rngFormatEUR.Font.ColorIndex = 5
        ' İsteğe bağlı: Para birimi formatını da burada toplu ata
        ' rngFormatEUR.NumberFormat = "#,##0.00 [$€-1]"
    End If
    
    ' 6. OTM Olmayan Satırların Biçimlendirmesi (J-X)
    If Not rngNonOTM Is Nothing Then
        ' J'den X'e kadar varsayılan sayı formatı
        rngNonOTM.Range("J1:X1").NumberFormat = "#,##0.00"
        
        ' W ve X Sütunlarını Tpbr'ye göre özel formatla
        On Error Resume Next
        Dim tpbrValue As String
        tpbrValue = ThisWorkbook.names("Tpbr").RefersToRange.Value
        If Err.Number = 0 And Not IsEmpty(tpbrValue) Then
            Select Case tpbrValue
                Case "Teklif Para Birimi (TL)"
                    ' Zaten #,##0.00 yapıldı
                Case "Teklif Para Birimi (EUR)"
                    rngNonOTM.Range("W1,X1").NumberFormat = "#,##0.00 [$€-1]"
                Case "Teklif Para Birimi (USD)"
                    rngNonOTM.Range("W1,X1").NumberFormat = "#,##0.00 [$$-C0C]"
            End Select
        End If
        Err.Clear
        On Error GoTo HataYonetimi_Makro
    End If
    ' === TOPLU BİÇİMLENDİRME SONU ===
    
    
    Application.StatusBar = False
    MsgBox bulunanSayac & " adet satır için veri ve formatlama ana listeden güncellendi.", vbInformation
    GoTo Cikis_Makro

' === HATA YÖNETİMİ ETİKETLERİ ===
DosyaYok:
    MsgBox "Ana malzeme listesi dosyası bulunamadı: " & masterFilePath, vbCritical
    GoTo Cikis_Makro
DosyaAcilamadi:
    MsgBox "Ana malzeme listesi dosyası açılamadı: " & masterFilePath, vbCritical
    GoTo Cikis_Makro
SayfaYok:
    MsgBox "'" & masterSheetName & "' sayfası ana listede bulunamadı!", vbExclamation
    If Not wbMaster Is Nothing Then wbMaster.Close False
    GoTo Cikis_Makro
VeriYok:
    MsgBox "Ana listede veri bulunamadı.", vbExclamation
    If Not wbMaster Is Nothing Then wbMaster.Close False
    GoTo Cikis_Makro

Cikis_Makro:
    On Error Resume Next
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.StatusBar = False
    Set wbMaster = Nothing
    Set wsMaster = Nothing
    Set wsKaynak = Nothing
    Set dictMaster = Nothing
    Set rngProcessed = Nothing
    Set rngFormatUSD = Nothing
    Set rngFormatEUR = Nothing
    Set rngOTM = Nothing
    Set rngNonOTM = Nothing
    Exit Sub

HataYonetimi_Makro:
    MsgBox "MlzVeriEkleMacrosu içinde hata oluştu:" & vbCrLf & _
            "Hata No: " & Err.Number & vbCrLf & _
            "Açıklama: " & Err.Description & vbCrLf & _
            "İşlenen Satır (Yaklaşık): " & i, vbCritical, "Makro Hatası"
    GoTo Cikis_Makro

End Sub

Sub MlzVeriEkleMacrosu31()

    Dim wsKaynak As Worksheet ' Verinin ekleneceği sayfa (Aktif Sayfa)
    Dim wbMaster As Workbook
    Dim wsMaster As Worksheet
    Dim masterFilePath As String
    Dim masterSheetName As String
    Dim sonSatirKaynak As Long
    Dim i As Long ' Döngü için satır sayacı (verilergir'deki 'Y' ve 'k' yerine)
    Dim lookupValue As Variant
    Dim rngFound As Range
    Dim masterRow As Long
    Dim bulunanSayac As Long
    Dim bfyt As String ' bfyt değişkenini burada da tanımla
    Dim kur As String  ' kur değişkenini burada da tanımla
    Dim mkur As String ' mkur değişkeni
    Dim Ckar As Variant ' CkarO kontrolü için
    Dim nameo As Variant ' Opsac kontrolü için
    Dim cellBValue As String ' B sütunundaki değeri kontrol etmek için

    ' Ayarlar
    Set wsKaynak = ActiveSheet ' Aktif olan çalışma sayfasını referans al
    masterFilePath = "C:\Belgelerim\cemex\Malzeme Listeleri\TümExcelListeleri.xlsb" ' Ana listenin tam yolu
    masterSheetName = "Sayfa1" ' Ana listedeki sayfa adı

    ' --- bfyt'yi ayarla (UserForm_Initialize'daki gibi) ---
    On Error Resume Next
    Ckar = ThisWorkbook.names("CkarO").RefersToR1C1
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo HataYonetimi_Makro
        If MsgBox("'CkarO' adlı ad tanımlı değil. Varsayılan ('Liste Fiyatı') oluşturulsun mu?", vbYesNo + vbQuestion) = vbYes Then
             ThisWorkbook.names.Add Name:="CkarO", RefersToR1C1:="=""Liste Fiyatı"""
             bfyt = "=RC[-6]"
        Else
             MsgBox "İşlem iptal edildi.", vbInformation
             Exit Sub
        End If
    Else
        On Error GoTo HataYonetimi_Makro
        If Ckar = "=""Net Fiyatı""" Then bfyt = "=RC[-1]" Else bfyt = "=RC[-6]"
    End If
    ' --- bfyt ayarlandı ---

    On Error GoTo HataYonetimi_Makro

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual ' Hesaplamayı manuel yap
    Application.EnableEvents = False ' Olayları durdur

    ' Ana dosya var mı kontrol et
    If dir(masterFilePath) = "" Then
        MsgBox "Ana malzeme listesi dosyası bulunamadı: " & masterFilePath, vbCritical
        GoTo Cikis_Makro
    End If

    ' Ana dosyayı aç (Henüz açık değilse)
    On Error Resume Next
    Set wbMaster = Workbooks(dir(masterFilePath))
    If wbMaster Is Nothing Then
        Set wbMaster = Workbooks.Open(masterFilePath, ReadOnly:=True, UpdateLinks:=0)
        If wbMaster Is Nothing Then
            MsgBox "Ana malzeme listesi dosyası açılamadı: " & masterFilePath & vbCrLf & "Lütfen dosyanın başka bir yerde açık olmadığından emin olun.", vbCritical
            GoTo Cikis_Makro
        End If
        Application.Windows(wbMaster.Name).Visible = False
    End If
    On Error GoTo HataYonetimi_Makro

    ' Ana çalışma sayfasını ayarla
    On Error Resume Next
    Set wsMaster = wbMaster.Worksheets(masterSheetName)
    If wsMaster Is Nothing Then
        MsgBox "'" & masterSheetName & "' sayfası ana listede bulunamadı!", vbExclamation
        If Workbooks.Count > 1 Then wbMaster.Close False ' Sadece bu makro açtıysa kapat
        GoTo Cikis_Makro
    End If
    On Error GoTo HataYonetimi_Makro

    ' Kaynak sayfadaki son dolu satırı bul (B sütununa göre)
    sonSatirKaynak = wsKaynak.Cells(wsKaynak.Rows.Count, "B").End(xlUp).row
    bulunanSayac = 0

    ' Kaynak sayfadaki satırları tara
    For i = 2 To sonSatirKaynak ' 1. satır başlık varsayıldı
        Application.StatusBar = "Satır " & i & "/" & sonSatirKaynak & " işleniyor..."

        ' Eğer B sütunu dolu ve A sütunu boş ise
        If Not IsEmpty(wsKaynak.Cells(i, "B").Value) And IsEmpty(wsKaynak.Cells(i, "A").Value) Then

            ' *** YENİ KONTROL BAŞLANGICI ***
            ' B sütunundaki değeri al (boşlukları temizle ve büyük harfe çevirerek kontrol et)
            cellBValue = Trim(UCase(CStr(wsKaynak.Cells(i, "B").Value)))

            ' Eğer B sütunundaki değer istenmeyen metinlerden BİRİ DEĞİLSE işlem yap
            If cellBValue <> "BÖLÜM TOPLAMI:" And cellBValue <> "BÖLÜM ADI/NO:" Then
            ' *** YENİ KONTROL SONU ***

                lookupValue = wsKaynak.Cells(i, "B").Value ' Arama değerini burada ayarla

                ' Ana listede değeri bul
                Set rngFound = wsMaster.Columns("B").Find(What:=lookupValue, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)

                ' Eğer bulunduysa
                If Not rngFound Is Nothing Then
                    masterRow = rngFound.row
                    bulunanSayac = bulunanSayac + 1

                    ' ---- verigir Makrosundan Uyarlanan Kod Başlangıcı ----

                    ' 1. GENEL VERİLERİ Ana Listeden Çek ve Hedefe Yaz
                    wsKaynak.Cells(i, "A").Value = wsMaster.Cells(masterRow, "A").Value ' Referans (ug1 olmadan)
                    ' wsKaynak.Cells(i, "B").Value = lookupValue ' B zaten dolu
                    wsKaynak.Cells(i, "C").Value = wsMaster.Cells(masterRow, "C").Value ' Yapılacak İşin Cinsi
                    wsKaynak.Cells(i, "D").Value = wsMaster.Cells(masterRow, "D").Value ' Üretici
                    'wsKaynak.Cells(i, "E").Value = 1 ' Miktar (Varsayılan 1)
                    wsKaynak.Cells(i, "F").Value = wsMaster.Cells(masterRow, "F").Value ' Mlz. Br. Fiyat
                    wsKaynak.Cells(i, "F").NumberFormat = wsMaster.Cells(masterRow, "F").NumberFormat ' Formatı al
                    wsKaynak.Cells(i, "G").Value = wsMaster.Cells(masterRow, "G").Value ' mlz.isk.
                    wsKaynak.Cells(i, "G").NumberFormat = wsMaster.Cells(masterRow, "G").NumberFormat ' Formatı al
                    wsKaynak.Cells(i, "H").Value = wsMaster.Cells(masterRow, "H").Value ' Adam/dk
                    wsKaynak.Cells(i, "I").Value = wsMaster.Cells(masterRow, "I").Value ' Boyut

                    ' 2. KUR'u belirle (Yazılan F sütunu formatına göre)
                    kur = ""
                    If wsKaynak.Cells(i, "F").NumberFormat = "#,##0.00 [$$-C0C]" Then wsKaynak.Cells(i, "F").Font.ColorIndex = 3: kur = "*Usd"
                    If wsKaynak.Cells(i, "F").NumberFormat = "#,##0.00 [$€-1]" Then wsKaynak.Cells(i, "F").Font.ColorIndex = 5: kur = "*Eur"

                    ' 3. ÜRÜN GRUPLARI ve Formülleri Uygula (L Sütunu ve J, K, M-X)
                    If Left(wsKaynak.CodeName, 3) <> "OTM" Then ' Sadece OTM olmayan sayfalarda formül uygula
                        mkur = kur: If bfyt = "=RC[-1]" Then mkur = "" ' Net fiyatsa kur ekleme (L hariç formüllerde)

                        ' L Sütunu Formülü (A sütunundaki değere göre)
                        Dim referansA As String
                        referansA = CStr(wsKaynak.Cells(i, "A").Value)

                        If Left(referansA, 5) = "PM-MP" Then wsKaynak.Cells(i, "L").FormulaR1C1 = bfyt & "*Oisci/100" & mkur
                        ElseIf Left(referansA, 5) = "PM-MS" Then wsKaynak.Cells(i, "L").FormulaR1C1 = bfyt & "*Osarf/100" & mkur
                        ElseIf Left(referansA, 5) = "PM-MA" Then wsKaynak.Cells(i, "L").FormulaR1C1 = bfyt & "*Oamb/100" & mkur
                        ElseIf Left(referansA, 5) = "PM-MN" Then wsKaynak.Cells(i, "L").FormulaR1C1 = bfyt & "*Onak/100" & mkur
                        ElseIf Left(referansA, 5) = "PM-MB" Then wsKaynak.Cells(i, "L").FormulaR1C1 = bfyt & "*Obara/100" & mkur
                        ElseIf Left(referansA, 3) = "PP-" Then wsKaynak.Cells(i, "L").FormulaR1C1 = bfyt & "*Opano/100" & mkur
                        ElseIf Left(referansA, 3) = "PS-" Then
                            On Error Resume Next
                            nameo = ThisWorkbook.names("Opsac").RefersToR1C1
                            If Err.Number = 0 And Not IsEmpty(nameo) Then wsKaynak.Cells(i, "L").FormulaR1C1 = bfyt & "*Opsac/100" & mkur Else wsKaynak.Cells(i, "L").FormulaR1C1 = bfyt & "*Opano/100" & mkur
                            Err.Clear
                            On Error GoTo HataYonetimi_Makro
                        Else
                            wsKaynak.Cells(i, "L").FormulaR1C1 = bfyt & "*Osalt/100" & mkur ' Varsayılan
                        End If

                        ' Diğer Formüller (J, K, M-X)
                        wsKaynak.Cells(i, "J").FormulaR1C1 = "=RC[-2]*Ads/60"
                        wsKaynak.Cells(i, "N").FormulaR1C1 = "=RC[-3]*Oggid/100"
                        wsKaynak.Cells(i, "K").FormulaR1C1 = "=(RC[-5]-RC[-5]*RC[-4])" & kur
                        wsKaynak.Cells(i, "L").FormulaR1C1 = "=(RC[-6])* Osalt/100 "
                        wsKaynak.Cells(i, "M").FormulaR1C1 = "=RC[-3]*Oisci/100"
                        wsKaynak.Cells(i, "O").FormulaR1C1 = "=RC[-10]*RC[-9]" & kur
                        wsKaynak.Cells(i, "P").FormulaR1C1 = "=RC[-11]*RC[-5]"
                        wsKaynak.Cells(i, "Q").FormulaR1C1 = "=RC[-12]*RC[-7]"
                        wsKaynak.Cells(i, "R").FormulaR1C1 = "=RC[-13]*RC[-6]"
                        wsKaynak.Cells(i, "S").FormulaR1C1 = "=RC[-14]*RC[-6]"
                        wsKaynak.Cells(i, "T").FormulaR1C1 = "=RC[-15]*RC[-12]/60"
                        wsKaynak.Cells(i, "U").FormulaR1C1 = "=RC[-7]*RC[-16]"
                        wsKaynak.Cells(i, "W").FormulaR1C1 = "=(RC[-13]+RC[-12]+RC[-11]+RC[-10]+RC[-9])*Dcki/Tpb"
                        wsKaynak.Cells(i, "X").FormulaR1C1 = "=RC[-19]*RC[-1]"
                    Else
                         ' OTM sayfasıysa formülleri temizle
                         wsKaynak.Range("J" & i & ":N" & i & ",O" & i & ":X" & i).ClearContents
                    End If ' OTM kontrolü sonu

                    ' 4. Biçimlendirmeleri Uygula
                    ' *** YENİ/GÜNCELLENMİŞ: Sayısal Biçimlendirmeleri AÇIKÇA Ayarla (İstenen formata) ***
                    With wsKaynak.Rows(i) ' İşlenen satır için
                        ' F Sütunu (Br. Fiyat)
                        .Cells(1, "F").NumberFormat = "#,##0.00"  ' Virgül binlik, Nokta ondalık

                        ' G Sütunu (İskonto)
                        .Cells(1, "G").NumberFormat = "0.0%"     ' Yüzde formatı

                        ' H Sütunu (Adam/dk) - Görseldeki gibi ondalıksız, binlik ayırıcılı
                        .Cells(1, "H").NumberFormat = "#,##0"

                        ' Diğer Sayısal Sütunlar (J-X) - Eğer OTM değilse
                        If Left(wsKaynak.CodeName, 3) <> "OTM" Then
                             'J'den X'e kadar biçimlendirme (X'ten önceki satırda yapılıyordu, buraya taşıdık)
                             .Range("J" & i & ":X" & i).NumberFormat = "#,##0.00" ' J'den X'e kadar (Nokta ondalık)

                             ' W ve X Sütunlarını Tpbr'ye göre özel formatla
                             On Error Resume Next
                             Dim tpbrValue As String
                             tpbrValue = ThisWorkbook.names("Tpbr").RefersToRange.Value
                             If Err.Number = 0 And Not IsEmpty(tpbrValue) Then
                                 With .Range("W" & i & ",X" & i) ' W ve X hücreleri için
                                     Select Case tpbrValue
                                         Case "Teklif Para Birimi (TL)"
                                             .NumberFormat = "#,##0.00"
                                         Case "Teklif Para Birimi (EUR)"
                                             .NumberFormat = "#,##0.00 [$€-1]"
                                         Case "Teklif Para Birimi (USD)"
                                             .NumberFormat = "#,##0.00 [$$-C0C]"
                                         Case Else
                                             .NumberFormat = "#,##0.00" ' Varsayılan
                                     End Select
                                 End With
                             Else
                                 .Range("W" & i & ",X" & i).NumberFormat = "#,##0.00" ' Tpbr yoksa TL varsay
                                 Err.Clear
                             End If
                             On Error GoTo HataYonetimi_Makro
                        End If
                    End With
                    ' *** Sayısal Biçimlendirme Sonu ***

                    ' Diğer Biçimlendirmeler (Kenarlık, Font, Hizalama vb.)
                    With wsKaynak.Range("A" & i & ":U" & i & ",W" & i & ":X" & i) ' V sütunu hariç
                        .Borders.LineStyle = xlContinuous
                        .Font.Bold = False
                        .Font.ColorIndex = xlAutomatic
                        .Font.Size = 9
                        .Font.Name = "Arial"
                    End With
                    With wsKaynak.Range("A" & i & ":D" & i)
                        .HorizontalAlignment = xlLeft
                        .NumberFormat = "@" ' Metin olarak formatla
                    End With
                    With wsKaynak.Range("E" & i & ":U" & i & ",W" & i & ":X" & i) ' V sütunu hariç
                        .HorizontalAlignment = xlRight
                    End With
                    wsKaynak.Cells(i, "E").NumberFormat = "#,##0" ' Miktar formatı (ondalıksız)

                    ' Fiyat sütunlarındaki kur renklerini tekrar uygula (format değiştiği için kaybolabilir)
                    If kur = "*Usd" Then wsKaynak.Cells(i, "F").Font.ColorIndex = 3
                    If kur = "*Eur" Then wsKaynak.Cells(i, "F").Font.ColorIndex = 5

                    ' ---- verigir Makrosundan Uyarlanan Kod Sonu ----

                End If ' Found
                Set rngFound = Nothing

            ' *** YENİ KONTROL için End If ***
            End If ' cellBValue <> "BÖLÜM TOPLAMI:" And cellBValue <> "BÖLÜM ADI/NO:"

        'End If ' B dolu, A boş
    Next i ' Sonraki satır

    ' Ana dosyayı kapat
    If Not wbMaster Is Nothing Then
        ' Eğer bu makro dışında başka Excel dosyası açık değilse ana dosyayı kapatma
        ' Sadece bu makro tarafından açıldıysa ve başka workbook yoksa kapatılır.
        ' Ancak birden fazla workbook açıksa, hangisinin ana dosya olduğunu bilmek zor.
        ' En güvenli yol, eğer başka workbook'lar açıksa kullanıcıya sormak veya her zaman açık bırakmak olabilir.
        ' Şimdilik, sadece 1'den fazla workbook açıksa kapatıyoruz (bu, en azından Excel'in kendisinin açık kalmasını sağlar).
        If Workbooks.Count > 1 Then
             ' Daha güvenli kapatma: Sadece bu makronun açtığı dosyayı kapattığından emin ol
             Dim wb As Workbook
             Dim openedByMacro As Boolean
             openedByMacro = False ' Başlangıçta false
             For Each wb In Workbooks
                 If wb.Name = dir(masterFilePath) Then
                     ' Eğer dosya adı eşleşiyorsa ve görünür değilse, muhtemelen bu makro açmıştır
                     If Not Application.Windows(wb.Name).Visible Then
                         openedByMacro = True
                     End If
                     Exit For
                 End If
             Next wb

             If openedByMacro Then
                 wbMaster.Close False
             End If
        End If
    End If


    Application.StatusBar = False
    ' İsteğe bağlı: Bulunan sayaç mesajını tekrar aktif edebilirsiniz.
    ' MsgBox bulunanSayac & " adet satır için veri ve formatlama ana listeden güncellendi.", vbInformation

Cikis_Makro:
    On Error Resume Next ' Hata yönetimi çıkışta sorun çıkarmasın
    Application.Calculation = xlCalculationAutomatic ' Hesaplamayı otomatik yap
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.StatusBar = False
    Set wbMaster = Nothing
    Set wsMaster = Nothing
    Set rngFound = Nothing
    Set wsKaynak = Nothing
    Exit Sub

HataYonetimi_Makro:
    MsgBox "MlzVeriEkleMacrosu içinde hata oluştu:" & vbCrLf & _
           "Hata No: " & Err.Number & vbCrLf & _
           "Açıklama: " & Err.Description & vbCrLf & _
           "İşlenen Satır (Yaklaşık): " & i, vbCritical, "Makro Hatası"
    GoTo Cikis_Makro ' Hata durumunda temizleyerek çık

End Sub