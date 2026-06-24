'----------------------------------------------------------------------------------------------
'Teklif Örnek klasörü İş proğramından otomatik oluştur
'
' MasaÜstünde bir klasör açılır
' Gerekli Klasör ve Teklif Şablonu eklenir
' İş Proğramı Bilgilerine göre İçeriği Hazırlanır
'
'----------------------------------------------------------------------------------------------


Sub teklif_klasor_olustur()

Dim ws As Worksheet
    Dim fso As Object ' Scripting.FileSystemObject
    Dim sonDoluSatirA As Long
    Dim ilkBosSatirA As Long
    Dim sonDoluSatirF As Long
    Dim klasorAdi As String
    Dim oncekiID As String
    Dim yeniSiraNo As Long
    Dim yeniID As String, projeAdi As String, firmaAdi As String, teklifIlgilisi As String
    Dim masaustuYolu As String
    Dim anaKlasorYolu As String
    Dim altKlasorler As Variant
    Dim altKlasor As Variant
    Dim hedefKlasorYolu As String
    Dim hedefDosyaAdi As String
    Dim hedefTamYol As String
    Const SABLON_YOLU As String = "C:\Belgelerim\Cemex\Yeni Teklif Şablonları\sablon.xlsm" ' KENDİ YOLUNUZLA GÜNCELLEYİN!
    'Const SABLON_YOLU As String = "C:\Belgelerim\Cemex\Yeni Teklif Şablonları\sablon.xlsx" ' KENDİ YOLUNUZLA GÜNCELLEYİN!
    Const GEREKLI_DOSYA_ADI_KISMI As String = "TEKLİF VE İŞ PROGRAMI"

    On Error GoTo HataYonetimi
    Application.ScreenUpdating = False ' Ekran güncellemeyi kapat

    ' --- Adım 1: Aktif Dosya Adını Kontrol Et ---
    If InStr(1, UCase(ActiveWorkbook.Name), UCase(GEREKLI_DOSYA_ADI_KISMI)) = 0 Then
        MsgBox "Bu makro yalnızca adı '" & GEREKLI_DOSYA_ADI_KISMI & "' içeren Excel dosyasında çalıştırılabilir.", vbExclamation, "Yanlış Dosya"
        GoTo Cleanup
    End If
    Debug.Print "Adım 1: Dosya adı kontrolü başarılı."

    ' Aktif sayfayı ayarla (Gerekirse belirli bir sayfa adı kullanın)
    Set ws = ActiveSheet
    ' Set ws = ThisWorkbook.Sheets("Sayfa1") ' Örnek

    ' FileSystemObject oluştur
    Set fso = CreateObject("Scripting.FileSystemObject")

    ' --- Adım 2: Yeni ID Oluştur (A Sütunu) ---
    sonDoluSatirA = ws.Cells(ws.Rows.Count, "A").End(xlUp).row
    ilkBosSatirA = sonDoluSatirA + 1

    ' İlk ID mi yoksa devam mı?
    If sonDoluSatirA < 2 Then ' Başlık satırı olduğunu varsayarsak
        yeniSiraNo = 1
        Debug.Print "A sütununda ilk ID oluşturuluyor (Sıra No: 1)."
    Else
        oncekiID = ws.Cells(sonDoluSatirA, "A").Value
        ' Önceki ID'nin formatını basitçe kontrol et (sonda 4 haneli sayı var mı?)
        If InStrRev(oncekiID, "-") > 0 And Len(oncekiID) - InStrRev(oncekiID, "-") = 4 And IsNumeric(Right(oncekiID, 4)) Then
            yeniSiraNo = CLng(Right(oncekiID, 4)) + 1
             Debug.Print "Önceki ID: " & oncekiID & ", Yeni Sıra No: " & yeniSiraNo
        Else
            MsgBox "A sütunundaki son değer (" & sonDoluSatirA & ". satır: '" & oncekiID & "') beklenen ID formatında değil (örn: EPR-ddmmyy-0001)." & vbCrLf & _
                   "Sıra numarası 1'den başlatılacak.", vbInformation, "ID Format Uyarısı"
            yeniSiraNo = 1 ' Hatalı format durumunda 1'den başla
        End If
    End If

    yeniID = "EPR-" & Format(Date, "ddmmyy") & "-" & Format(yeniSiraNo, "0000")
    ws.Cells(ilkBosSatirA, "A").Value = yeniID
    ws.Cells(ilkBosSatirA, "A").NumberFormat = "@" ' Metin olarak biçimlendir
    Debug.Print "Adım 2: Yeni ID '" & yeniID & "' oluşturuldu ve " & ilkBosSatirA & ". satıra yazıldı."
    
    ' --- Proje adı, firma adını değişkene kayıt et
    firmaAdi = ws.Cells(ilkBosSatirA, "c").Value
    projeAdi = ws.Cells(ilkBosSatirA, "e").Value
    teklifIlgilisi = ws.Cells(ilkBosSatirA, "d").Value
    
    ' --- Adım 3: Klasör Adını Al (F Sütunu) ---
    sonDoluSatirF = ws.Cells(ilkBosSatirA, "F").End(xlUp).row

    klasorAdi = Trim(ws.Cells(ilkBosSatirA, "F").Value)

    If klasorAdi = "" Then
        MsgBox "F sütunundaki son dolu hücre (" & sonDoluSatirF & ". satır) boş. Geçerli bir klasör adı girilmelidir.", vbCritical, "Eksik Veri (F Sütunu)"
        GoTo Cleanup
    End If

    ' Klasör adı için geçersiz karakterleri kontrol et (Basit kontrol)
    Dim invalidChars As String: invalidChars = "\/:*?""<>|"
    Dim i As Integer
    For i = 1 To Len(invalidChars)
        If InStr(klasorAdi, Mid(invalidChars, i, 1)) > 0 Then
             MsgBox "F sütunundaki değer '" & klasorAdi & "' geçersiz karakter içeriyor (" & Mid(invalidChars, i, 1) & "). Geçerli bir klasör adı girin.", vbCritical, "Geçersiz Klasör Adı"
            GoTo Cleanup
        End If
    Next i
    Debug.Print "Adım 3: Klasör adı '" & klasorAdi & "' F" & ilkBosSatirA & " hücresinden alındı."


    ' --- Adım 4: Klasör Yapısını Oluştur ---
    masaustuYolu = CreateObject("Wscript.Shell").SpecialFolders("Desktop")
    If masaustuYolu = "" Then
        MsgBox "Masaüstü yolu alınamadı. Klasörler oluşturulamıyor.", vbCritical, "Sistem Hatası"
        GoTo Cleanup
    End If
    Debug.Print "Masaüstü Yolu: " & masaustuYolu

    anaKlasorYolu = fso.BuildPath(masaustuYolu, klasorAdi)
    Debug.Print "Oluşturulacak Ana Klasör Yolu: " & anaKlasorYolu

    ' Ana klasör var mı kontrol et
    If fso.FolderExists(anaKlasorYolu) Then
        MsgBox "'" & anaKlasorYolu & "' klasörü zaten mevcut." & vbCrLf & "İşlem durduruldu.", vbExclamation, "Klasör Mevcut"
        GoTo Cleanup
    Else
        ' Ana klasörü oluştur
        On Error Resume Next ' Oluşturma hatasını yakala
        fso.CreateFolder anaKlasorYolu
        If Err.Number <> 0 Then
            MsgBox "Ana klasör oluşturulamadı!" & vbCrLf & anaKlasorYolu & vbCrLf & "Hata: " & Err.Description, vbCritical, "Klasör Oluşturma Hatası"
            On Error GoTo HataYonetimi ' Normal hata yönetimine dön
            GoTo Cleanup
        End If
        On Error GoTo HataYonetimi ' Normal hata yönetimine dön
        Debug.Print "Ana klasör başarıyla oluşturuldu."
    End If

    ' Alt klasörleri oluştur
    altKlasorler = Array("1-İş Emri", "2-Projeler", "3-Malzeme Listesi", "4-Teklif", "5-Resimler", "6-Test Raporu", "7-Analiz")
    Dim altKlasorYolu As String
    For Each altKlasor In altKlasorler
        altKlasorYolu = fso.BuildPath(anaKlasorYolu, CStr(altKlasor))
         If Not fso.FolderExists(altKlasorYolu) Then
            On Error Resume Next ' Oluşturma hatasını yakala
            fso.CreateFolder altKlasorYolu
            If Err.Number <> 0 Then
                Debug.Print "  ! Uyarı: Alt klasör oluşturulamadı: " & altKlasorYolu & " Hata: " & Err.Description
                Err.Clear ' Hatayı logla ama devam etmeyi dene
            Else
                Debug.Print "  + Alt klasör oluşturuldu: " & altKlasorYolu
            End If
            On Error GoTo HataYonetimi
        Else
             Debug.Print "  * Alt klasör zaten mevcut: " & altKlasorYolu
        End If
    Next altKlasor
    Debug.Print "Adım 4: Klasör yapısı oluşturuldu/kontrol edildi."


    ' --- Adım 5: Şablon Dosyasını Kopyala ve Yeniden Adlandır ---
    ' Kaynak şablon dosyasının varlığını kontrol et
    If Not fso.FileExists(SABLON_YOLU) Then
        MsgBox "Kaynak şablon dosyası bulunamadı:" & vbCrLf & SABLON_YOLU, vbCritical, "Şablon Eksik"
        GoTo Cleanup
    End If

    ' Hedef klasör yolunu belirle
    hedefKlasorYolu = fso.BuildPath(anaKlasorYolu, "4-Teklif")

    ' Hedef klasörün varlığını kontrol et (normalde adım 4'te oluşmuş olmalı)
    If Not fso.FolderExists(hedefKlasorYolu) Then
        MsgBox "Hedef teklif klasörü bulunamadı!" & vbCrLf & hedefKlasorYolu & vbCrLf & "Klasör yapısı oluşturulamamış olabilir.", vbCritical, "Hedef Klasör Eksik"
        GoTo Cleanup
    End If

    ' Hedef dosya adını oluştur (F sütunundaki adı kullanarak ve uzantıyı .xlsx olarak değiştirerek)
    hedefDosyaAdi = klasorAdi & ".xlsm" ' Şablonu normal Excel dosyası olarak kaydet
    hedefTamYol = fso.BuildPath(hedefKlasorYolu, hedefDosyaAdi)
    Debug.Print "Oluşturulacak Teklif Dosyası: " & hedefTamYol

    ' Hedef dosya zaten var mı kontrol et
    If fso.FileExists(hedefTamYol) Then
        MsgBox "Hedef teklif dosyası zaten mevcut:" & vbCrLf & hedefTamYol & vbCrLf & "Kopyalama işlemi atlandı.", vbExclamation, "Dosya Mevcut"
    Else
        ' Dosyayı kopyala
        On Error Resume Next ' Kopyalama hatasını yakala
        FileCopy SABLON_YOLU, hedefTamYol
        If Err.Number <> 0 Then
            MsgBox "Şablon dosyası kopyalanamadı!" & vbCrLf & SABLON_YOLU & " -> " & hedefTamYol & vbCrLf & "Hata: " & Err.Description, vbCritical, "Kopyalama Hatası"
            Err.Clear
            ' Hata olsa bile diğer adımlar tamamlandığı için çıkmayabiliriz.
        Else
            Debug.Print "Adım 5: Şablon başarıyla '" & hedefDosyaAdi & "' olarak kopyalandı."
        End If
        On Error GoTo HataYonetimi
    End If
    
        ' Oluşturduğumuz klasör içerisinde Teklif klasörüne kopyalanan exceli aç
    Dim hedefKitap As Workbook
    Dim hedefSayfa As Worksheet
    
    Set hedefKitap = Workbooks.Open(hedefTamYol)
    Set hedefSayfa = hedefKitap.Sheets("sayfa3")
    
    'b2 Sütununa Proje Numarasını Gir
    hedefSayfa.Range("c3").Value = projeAdi '"PROJE ADI"
    hedefSayfa.Range("c4").Value = "PROJE KISA ADI"
    hedefSayfa.Range("c5").Value = firmaAdi '"İŞVEREN"
    hedefSayfa.Range("c7").Value = yeniID '"TEKLİF ID"
    hedefSayfa.Range("c8").Value = Date '"İLK VERİLİŞ TARİHİ"
    hedefSayfa.Range("c10").Value = "SERCAN GÜNGÖR" '"HAZIRLAYAN"
    hedefSayfa.Range("c11").Value = teklifIlgilisi '"TEKLİFLE İLGİLİ KİŞİ"
    
    
    'hedefKitap.Close SaveChanges:=True
    ' --- İşlem Sonu ---
    'MsgBox "Teklif klasör yapısı ve başlangıç dosyası başarıyla oluşturuldu:" & vbCrLf & anaKlasorYolu, vbInformation, "İşlem Başarılı"


Cleanup:
    Application.ScreenUpdating = True ' Ekran güncellemeyi aç
    ' Nesneleri temizle
    Set ws = Nothing
    Set fso = Nothing
    Exit Sub ' Normal çıkış

HataYonetimi:
    MsgBox "Beklenmeyen bir hata oluştu!" & vbCrLf & vbCrLf & _
           "Hata No: " & Err.Number & vbCrLf & _
           "Açıklama: " & Err.Description & vbCrLf & _
           "Kaynak: " & Err.Source, vbCritical, "Genel Hata"
    Resume Cleanup ' Hata durumunda Cleanup bölümüne git ve çık

End Sub