Option Explicit ' Modülün başına eklenmesi ÖNERİLİR - tüm değişkenlerin tanımlanmasını zorunlu kılar

' --- Global Değişkenler (Dikkatli Kullanın!) ---
Dim wsHedef As Worksheet
Dim secilenDosyaYolu As Variant ' Kaynak dosya yolu için
Dim g_wbKaynak As Workbook      ' Global variable for the source workbook

' --- InitializeAddin (Degisiklik Yok) ---
Sub InitializeAddin()
    ' Bu Sub, lisans geçerli oldugunda Workbook_Open tarafindan çagrilir.
    ' Eklentinin normal baslatma islemlerini (kisayollar vb.) yapar.
    Debug.Print "--------------------------------------------"
    Debug.Print "InitializeAddin: Eklenti baslatiliyor..."
    On Error Resume Next ' Hata olursa devam et
    Application.OnKey "+^w", "TestKaynakVeriAl_Tip1"
    Application.OnKey "+^d", "TestKaynakVeriAl_Tip2"
    Application.OnKey "+^s", "sarf_iscilik_bakir"
    Application.OnKey "+^f", "sinyal_ekle"
    Application.OnKey "+^q", "macro6"
    If Err.Number <> 0 Then
        Debug.Print "InitializeAddin: Kisayol atanirken hata olustu: " & Err.Description
        Err.Clear
    End If
    On Error GoTo 0
    Debug.Print "InitializeAddin: Özel kisayollar atandi ve baslatma tamamlandi."
End Sub

' --- Workbook_BeforeClose (Degisiklik Yok) ---
Sub DeleteInitializeAddin()
    Debug.Print "Workbook_BeforeClose: Özel kisayollar kaldiriliyor..."
    On Error Resume Next ' Hata olursa devam et
    Application.OnKey "+^w"
    Application.OnKey "+^d"
    Application.OnKey "+^f"
    Application.OnKey "+^s"
    Application.OnKey "+^q"
    If Err.Number <> 0 Then
        Debug.Print "Workbook_BeforeClose: Kisayol kaldirilirken hata olustu: " & Err.Description
        Err.Clear
    End If
    On Error GoTo 0
    Debug.Print "Workbook_BeforeClose: Özel kisayollar kaldirildi."
End Sub

'###################### Kısayollar buradadır #####################################

Sub macro6()
    Set wsHedef = Nothing
End Sub

' --- Ana Fonksiyon: Kaynak Dosyadan Kullanıcının YENİ SEÇİMİNİ (Do While ile) Alır ---
Function KullanicidanKaynakVeriAl(ByVal SecimTipi As Integer) As Variant
    ' Amaç: Kaynak dosyayı (gerekirse) açar/aktive eder, mevcut seçimi temizler,
    '       kullanıcının YENİ BİR SEÇİM (tek hücre dışında) yapmasını Do While
    '       döngüsüyle bekler ve belirtilen tipe göre veriyi döndürür.
    '       EĞER KAYNAK KİTAP ZATEN AÇIKSA (g_wbKaynak global değişkeni), ONU KULLANIR.
    '       Seçim sonrası HEDEF KİTABI/SAYFAYI AKTİVE EDER.
    '       Kaynak dosyayı açık tutar (g_wbKaynak).
    ' Döndürür: Başarılı ise Variant (tek değer veya dizi), değilse CVErr(xlErrNA).

    Dim hedefDosyaYolu As String
    Dim wbHedef As Workbook
    Dim userSelection As Range ' Kullanıcının aktif seçimi için
    Dim bulundu As Boolean
    Dim wb As Workbook
    Dim dataToReturn As Variant ' Dönecek veriyi tutmak için geçici değişken
    Dim sourcePathToUse As Variant
    Dim workbookName As String
    Dim isKaynakStillOpen As Boolean

    ' --- Hata İşleyici ---
    On Error GoTo ErrorHandler

    ' --- Varsayılan dönüş değeri: Hata ---
    dataToReturn = CVErr(xlErrNA) ' Hata veya iptal durumunda #N/A! döner

    ' --- Obje değişkenlerini başlangıçta Nothing yap (local ones) ---
    Set wbHedef = Nothing
    Set userSelection = Nothing
    Set wb = Nothing

    Application.ScreenUpdating = False ' Başlangıçta ekran güncellemeyi kapat

    ' --- 1. Hedef Çalışma Kitabını Ayarla (GetSetting ile) ---
    ' (Bu kısım aynı kalıyor)
    hedefDosyaYolu = GetSetting("sercan", "fileOpenWorkBooks", "nowOpenPropsFile", "")
    If hedefDosyaYolu = "" Then
        MsgBox "Hedef çalışma kitabı yolu kayıt defterinde bulunamadı veya boş.", vbCritical, "Hata"
        GoTo FonksiyonCikis
    End If
    bulundu = False
    For Each wb In Application.Workbooks
        If StrComp(wb.fullName, hedefDosyaYolu, vbTextCompare) = 0 Then
            Set wbHedef = wb
            bulundu = True
            Exit For
        End If
    Next wb
    If Not bulundu Then
        MsgBox "Belirtilen hedef çalışma kitabı açık değil:" & vbCrLf & hedefDosyaYolu, vbCritical, "Hata"
        GoTo FonksiyonCikis
    End If

    ' --- 2. Hedef Çalışma Sayfasını Ayarla (Global wsHedef'i kullan) ---
    ' (Bu kısım aynı kalıyor)
    If wsHedef Is Nothing Then
        On Error Resume Next
        Set wsHedef = wbHedef.ActiveSheet
        On Error GoTo ErrorHandler
        If wsHedef Is Nothing Then
            MsgBox "Hedef çalışma kitabında geçerli bir çalışma sayfası aktif değil veya global wsHedef değişkeni ayarlanmamış.", vbCritical, "Hata"
            GoTo FonksiyonCikis
        End If
    Else
        If Not wsHedef.Parent Is wbHedef Then
            MsgBox "Global wsHedef değişkeni, kayıt defterinden alınan çalışma kitabına (" & wbHedef.Name & ") ait değil!", vbCritical, "Hata"
            GoTo FonksiyonCikis
        End If
    End If
    If wbHedef Is Nothing Or wsHedef Is Nothing Then
         MsgBox "Hedef kitap veya sayfa nesnesi ayarlanamadı!", vbCritical
         GoTo FonksiyonCikis
    End If

    ' --- 3 & 4. Kaynak Dosyayı Yönet (Aç/Bul/Aktive Et) ---
    Debug.Print "--- Kaynak Dosya Kontrolü Başladı ---"
    isKaynakStillOpen = False
    If Not g_wbKaynak Is Nothing Then
        ' g_wbKaynak global değişkeni boş değil, hala açık mı kontrol et
        Debug.Print "Global 'g_wbKaynak' nesnesi mevcut. Açıklık durumu kontrol ediliyor..."
        On Error Resume Next ' Kapalıysa .Name hatası verir
        workbookName = g_wbKaynak.Name
        If Err.Number = 0 Then
             isKaynakStillOpen = True
             Debug.Print "'g_wbKaynak' (" & workbookName & ") hala açık ve geçerli."
        Else
             Debug.Print "'g_wbKaynak' nesnesi vardı ama kitap artık kapalı veya geçersiz. Nesne sıfırlanıyor."
             Set g_wbKaynak = Nothing: Err.Clear
             isKaynakStillOpen = False ' Açık değil olarak işaretle
        End If
        On Error GoTo ErrorHandler
    Else
         Debug.Print "Global 'g_wbKaynak' nesnesi boş (Nothing)."
    End If

    ' Kaynak kitap zaten açıksa ve global değişkende varsa, onu kullan
    If isKaynakStillOpen Then
        Debug.Print "Mevcut açık 'g_wbKaynak' (" & g_wbKaynak.Name & ") kullanılacak."
        ' Global yol değişkenini de açık kitapla senkronize et (gerekirse)
        If IsEmpty(secilenDosyaYolu) Or secilenDosyaYolu = False Or secilenDosyaYolu = "" Then
             secilenDosyaYolu = g_wbKaynak.fullName
             Debug.Print "Global 'secilenDosyaYolu' boştu, açık kitaptan ayarlandı: " & secilenDosyaYolu
        ElseIf StrComp(g_wbKaynak.fullName, secilenDosyaYolu, vbTextCompare) <> 0 Then
             Debug.Print "Uyarı: Global yol (" & secilenDosyaYolu & ") açık olan farklı bir kaynak kitaptan (" & g_wbKaynak.fullName & "). Açık olan kullanılacak ve yol güncellenecek."
             secilenDosyaYolu = g_wbKaynak.fullName
        End If
        sourcePathToUse = secilenDosyaYolu ' Kullanılacak yolu belirle
    Else
        ' Kaynak kitap açık değil veya global değişken geçersizdi.
        Debug.Print "Yeni kaynak kitap açma/bulma işlemi gerekiyor."
        Set g_wbKaynak = Nothing ' Emin olmak için tekrar Nothing yap
        sourcePathToUse = secilenDosyaYolu ' Önce globaldeki yolu dene

        ' Eğer global yol boşsa veya geçersizse, kullanıcıdan iste
         If IsEmpty(sourcePathToUse) Or sourcePathToUse = False Or sourcePathToUse = "" Then
             Debug.Print "Global 'secilenDosyaYolu' boş. Kullanıcıdan dosya seçmesi istenecek."
             sourcePathToUse = Application.GetOpenFilename("Excel Dosyaları (*.xls*), *.xls*", , "Lütfen Verilerin Alınacağı Kaynak Excel Dosyasını Seçin")
             If sourcePathToUse = False Then
                Debug.Print "Kullanıcı dosya seçimini iptal etti."
                GoTo FonksiyonCikis ' Kullanıcı iptal etti
             End If
             secilenDosyaYolu = sourcePathToUse ' Seçilen yolu globale kaydet
             Debug.Print "Kullanıcı dosya seçti: " & secilenDosyaYolu
         Else
            Debug.Print "Global 'secilenDosyaYolu' kullanılıyor: " & sourcePathToUse
         End If

        ' Belirlenen yoldaki kitap zaten açık mı diye kontrol et (başka bir yerden açılmış olabilir)
        For Each wb In Application.Workbooks
            If StrComp(wb.fullName, sourcePathToUse, vbTextCompare) = 0 Then
                Set g_wbKaynak = wb ' Açık olanı bulduk, globale ata
                Debug.Print "Kaynak kitap zaten ('" & wb.Name & "' adıyla) açık bulundu."
                Exit For
            End If
        Next wb

        ' Eğer hala açık bulunamadıysa, dosyayı açmayı dene
        If g_wbKaynak Is Nothing Then
             Debug.Print "Kaynak kitap açık değil, şimdi açılıyor: " & sourcePathToUse
             On Error Resume Next ' Açma işlemi hata verebilir
             Set g_wbKaynak = Workbooks.Open(fileName:=sourcePathToUse, UpdateLinks:=0, ReadOnly:=True)
             On Error GoTo ErrorHandler ' Hata yakalamayı normale döndür

             If g_wbKaynak Is Nothing Then
                  MsgBox "Kaynak dosya açılamadı veya bulunamadı:" & vbCrLf & sourcePathToUse, vbCritical, "Hata"
                  secilenDosyaYolu = Empty ' Hatalı yolu temizle
                  GoTo FonksiyonCikis
             Else
                  Debug.Print "Kaynak kitap başarıyla açıldı ve 'g_wbKaynak' olarak ayarlandı."
             End If
        End If
    End If
    Debug.Print "--- Kaynak Dosya Kontrolü Bitti ---"

    If g_wbKaynak Is Nothing Then GoTo FonksiyonCikis ' Eğer hala bir kitap yoksa çık

    ' --- 5. KAYNAK KİTABI AKTİVE ET, SEÇİMİ TEMİZLE VE YENİ SEÇİM İÇİN BEKLE (Do While ile) ---
    g_wbKaynak.Activate
    
    
    ' Açılıştaki olası seçimi temizlemek için A1 hücresini seç
    ' ################################
    '       mevcut seçimi iptal etmek için kullan
    ' ###############################
    
    'On Error Resume Next ' Sayfa korumalıysa vs. hata vermesin
    'If Not g_wbKaynak.ActiveSheet Is Nothing Then
    '    g_wbKaynak.ActiveSheet.Range("A1").Select
    'End If
    'On Error GoTo ErrorHandler ' Hata yakalamayı normale döndür

    Application.ScreenUpdating = True ' Kullanıcı etkileşimi için ekranı aç

    ' Kullanıcıya ne yapacağını bildir (InputBox yerine)
    'MsgBox "Kaynak dosya '" & g_wbKaynak.Name & "' aktif." & vbCrLf & _
    '       "Lütfen şimdi istediğiniz hücreleri seçin." & vbCrLf & _
    '       "(Tek bir hücre SEÇMEYİN, birden fazla hücre veya bir aralık seçin)." & vbCrLf & _
    '       "Kod, siz seçimi yapana kadar bekleyecektir.", vbInformation, "Seçim Bekleniyor"

    ' Kullanıcının birden fazla hücre seçmesini bekle (Selection.Count 1 iken döngüde kal)
    Do While Application.Selection.Count = 1
        DoEvents ' Excel'in kullanıcı eylemlerini işlemesine izin ver

        ' Kaynak kitabın hala açık olup olmadığını kontrol et (Döngü içinde tekrar kontrol önemli)
        isKaynakStillOpen = False: On Error Resume Next
        workbookName = g_wbKaynak.Name ' Kitap kapandıysa burada hata verir
        If Err.Number <> 0 Then
             isKaynakStillOpen = False
             Err.Clear
        Else
             isKaynakStillOpen = True
        End If
        On Error GoTo ErrorHandler ' Hata yakalamayı normale döndür

        If Not isKaynakStillOpen Then
             MsgBox "Kaynak kitap seçim sırasında kapatıldı. İşlem iptal ediliyor.", vbExclamation
             Set g_wbKaynak = Nothing: secilenDosyaYolu = Empty
             GoTo FonksiyonCikis ' Temizle ve çık
        End If
    Loop

    ' Döngü bittiğinde (Selection.Count != 1) seçimi al
    Set userSelection = Application.Selection
    Application.ScreenUpdating = False ' İşleme devam etmeden önce ekranı kapat

    ' --- Seçimin Geçerliliğini Kontrol Et ---
    ' (Bu kısım aynı kalıyor)
     If userSelection Is Nothing Or Not TypeOf userSelection Is Range Then
         MsgBox "Geçerli bir hücre seçimi yapılamadı. Lütfen tekrar deneyin.", vbExclamation
         GoTo FonksiyonCikis
     End If
    If Not userSelection.Worksheet.Parent Is g_wbKaynak Then
        MsgBox "Hata: Seçim beklenen kaynak kitapta (" & g_wbKaynak.Name & ") yapılmadı!" & vbCrLf & _
               "Seçilen Kitap: " & userSelection.Worksheet.Parent.Name, vbCritical
         GoTo FonksiyonCikis
    End If
    If userSelection.Cells.CountLarge = 1 Then
         MsgBox "Tek bir hücre seçilmemesi gerekiyor. İşlem iptal edildi.", vbExclamation
         GoTo FonksiyonCikis
    End If

    ' --- 6. Seçimi SeçimTipi'ne göre işle ---
    ' (Bu kısım aynı kalıyor)
    Select Case SecimTipi
        Case 1 ' İlk Hücrenin Değeri
            dataToReturn = userSelection.Cells(1, 1).Value
        Case 2 ' Aralık Değerleri Dizisi
             dataToReturn = userSelection.Value2
        Case Else ' Geçersiz Tip
             MsgBox "Fonksiyona geçersiz seçim tipi (" & SecimTipi & ") gönderildi. 1 veya 2 olmalı.", vbCritical
             GoTo FonksiyonCikis
    End Select

    ' --- *** SEÇİM BAŞARILI - HEDEF KİTABI AKTİVE ET *** ---
    ' (Bu kısım aynı kalıyor)
    On Error Resume Next
    Application.ScreenUpdating = True
    'wbHedef.Activate
    If Err.Number <> 0 Then
        MsgBox "Uyarı: Hedef kitap (" & hedefDosyaYolu & ") aktifleştirilemedi. Kapatılmış olabilir.", vbExclamation
        Err.Clear
    Else
        wsHedef.Activate
        If Err.Number <> 0 Then
            MsgBox "Uyarı: Hedef sayfa (" & wsHedef.Name & ") aktifleştirilemedi.", vbExclamation
            Err.Clear
        End If
    End If
    Application.ScreenUpdating = False
    On Error GoTo ErrorHandler

    GoTo FonksiyonCikis

ErrorHandler:
    MsgBox "Fonksiyonda bir hata oluştu!" & vbCrLf & _
           "Hata Numarası: " & Err.Number & vbCrLf & _
           "Açıklama: " & Err.Description, vbCritical, "Fonksiyon Hatası"

FonksiyonCikis:
    Application.ScreenUpdating = True
    KullanicidanKaynakVeriAl = dataToReturn
    Set wbHedef = Nothing
    Set wb = Nothing
    Set userSelection = Nothing
    On Error GoTo 0

End Function
' --- Test Prosedürü 1: Tek Hücre Değeri Al ve "Bölüm Adı" Olarak Biçimlendir (E Sütunu Dahil) ---
Sub TestKaynakVeriAl_Tip1()
Dim secilenData As Variant ' Dönen değer (tek) veya hata
Dim Y As Long             ' Yeni eklenen satırın numarası
Dim hedefSayfa As Worksheet ' İşlem yapılacak sayfa
Dim bicimlenecekAlan As Range ' Biçimlendirme uygulanacak hücreler
Dim testAdetDegeri As Long ' E sütunu için test değeri

' --- Hata Yönetimi ---
On Error GoTo HataYonetimi

' --- 1. Hedef Sayfayı Belirle ---
If wsHedef Is Nothing Then
    Set hedefSayfa = ActiveSheet
    If hedefSayfa Is Nothing Then
        MsgBox "Aktif bir çalışma sayfası bulunamadı!", vbCritical, "Hata"
        Exit Sub
    End If
    Set wsHedef = hedefSayfa
    Debug.Print "Test için wsHedef aktif sayfaya ayarlandı: " & wsHedef.Name
Else
    Set hedefSayfa = wsHedef
    On Error Resume Next
    Dim testAd As String
    testAd = hedefSayfa.Name
    If Err.Number <> 0 Then
         MsgBox "Global olarak ayarlanan hedef sayfa (" & wsHedef.Name & ") artık geçerli değil!", vbCritical
         Set wsHedef = Nothing
         On Error GoTo HataYonetimi
         Exit Sub
    End If
    On Error GoTo HataYonetimi
End If

Application.ScreenUpdating = True
hedefSayfa.Activate
Application.ScreenUpdating = False

' --- 2. Yeni Satır Ekle (Aktif hücrenin üzerine) ---
If TypeName(Selection) <> "Range" Then
    MsgBox "Lütfen önce veri eklenecek konumda bir hücre seçin.", vbExclamation
    Exit Sub
End If
Selection.EntireRow.Insert Shift:=xlDown
Y = Selection.row

' --- 3. "Bölüm Adı" Biçimlendirmelerini Uygula ---
Set bicimlenecekAlan = hedefSayfa.Range("A" & Y & ",B" & Y & ":E" & Y & ",F" & Y & ":U" & Y & ",W" & Y & ":X" & Y)

bicimlenecekAlan.Borders.LineStyle = xlContinuous
hedefSayfa.Range("F" & Y & ":U" & Y & ",W" & Y & ":X" & Y).Borders(xlInsideVertical).LineStyle = xlNone
hedefSayfa.Range("B" & Y & ":E" & Y).Borders(xlInsideVertical).LineStyle = xlNone

With bicimlenecekAlan
    .Interior.Pattern = xlNone
    .Font.Size = 9
    .Font.Bold = True
    .Font.ColorIndex = 11 ' Bölüm Adı için Renk İndeksi (Koyu Mavi)
End With

'aktif sayfayı kayıt et
Call zActiveWb.GetActiveWorkbookPath

hedefSayfa.Rows(Y).RowHeight = 12.75
hedefSayfa.Range("A" & Y & ":D" & Y).NumberFormat = "@"

' --- 4. Sabit Değerleri ve E Sütunu Test Değerini Gir ---
hedefSayfa.Range("B" & Y).Value = "BÖLÜM ADI/NO:"

' E Sütunu için Test Değeri (TextBox20 simülasyonu)
' Formdaki gibi bir kontrol olmadığı için direkt değeri atıyoruz.
testAdetDegeri = 1 ' TextBox20'den gelmesi gereken değer için sabit bir test değeri
hedefSayfa.Range("E" & Y).Value = testAdetDegeri ' .Value kullanmak genellikle daha güvenlidir
hedefSayfa.Range("E" & Y).NumberFormat = "#,##0 ""Adet""" ' Sayı biçimini uygula

' --- 5. Fonksiyonu Çağır (Kaynak Dosyadan Veri Al) ---
'MsgBox "Kaynak veri alma fonksiyonu (Tip 1: Tek Değer) çağrılıyor..." & vbCrLf & _
'       "Kaynak dosya açık değilse açılacak/seçtirilecek." & vbCrLf & _
'       "Lütfen kaynak kitaptaki BİRDEN FAZLA hücre seçin (ilk hücrenin değeri alınacak).", vbInformation

Application.ScreenUpdating = False
secilenData = KullanicidanKaynakVeriAl(1) ' SecimTipi = 1
Application.ScreenUpdating = True

On Error Resume Next
hedefSayfa.Activate
On Error GoTo HataYonetimi

' --- 6. Sonucu Kontrol Et ve C Hücresine Yaz ---
If Not IsError(secilenData) Then
    If Not IsEmpty(secilenData) And CStr(secilenData) <> "" Then
        hedefSayfa.Range("C" & Y).Value = secilenData
        'MsgBox "İşlem başarıyla tamamlandı." & vbCrLf & _
        '       "'" & hedefSayfa.Name & "' sayfasına " & Y & ". satır eklendi ve biçimlendirildi." & vbCrLf & _
        '       "E sütununa test 'Adet' değeri (" & testAdetDegeri & ") girildi." & vbCrLf & _
        '       "Kaynak dosyadan seçilen ilk hücre değeri '" & CStr(secilenData) & "' C sütununa yazıldı.", _
        '       vbInformation, "İşlem Başarılı"
        Debug.Print "Test Tip 1 Başarılı. Veri: " & CStr(secilenData) & ". Satır: " & Y & ". Kaynak Kitap: " & g_wbKaynak.Name & " (Açık Kalmalı)"
    Else
        hedefSayfa.Range("C" & Y).Value = ""
         MsgBox "Kaynak dosyadan geçerli bir veri seçilmedi (Seçilen hücre boş olabilir)." & vbCrLf & _
               "'" & hedefSayfa.Name & "' sayfasındaki " & Y & ". satırın C sütunu boş bırakıldı." & vbCrLf & _
               "(E sütununa test 'Adet' değeri girildi.)", vbExclamation, "Veri Boş"
    End If
Else
    hedefSayfa.Range("C" & Y).Value = ""
    MsgBox "Kaynak dosyadan veri alınamadı (İptal edildi veya hata oluştu)." & vbCrLf & _
           "Dönen Hata Kodu: " & CStr(secilenData) & vbCrLf & _
           "'" & hedefSayfa.Name & "' sayfasındaki " & Y & ". satırın C sütunu boş bırakıldı." & vbCrLf & _
               "(E sütununa test 'Adet' değeri girildi.)", vbExclamation, "Veri Alımı Başarısız"
    ' İsteğe bağlı: Hata durumunda satırı silme kodu eklenebilir.
End If

' --- Test Sonrası Temizlik ---
'MsgBox "Test (Tip 1 - Bölüm Adı Formatıyla) Tamamlandı." & vbCrLf & "Kaynak kitap açık bırakıldı.", vbInformation
GoTo TemizCikis
HataYonetimi:
MsgBox "Test prosedüründe bir hata oluştu!" & vbCrLf & _
"Hata Numarası: " & Err.Number & vbCrLf & _
"Açıklama: " & Err.Description, vbCritical, "Test Hatası"
Set wsHedef = Nothing

TemizCikis:
Application.ScreenUpdating = True
Set hedefSayfa = Nothing
Set bicimlenecekAlan = Nothing
On Error GoTo 0

End Sub

' --- Test Prosedürü 2: Aralık Değerleri Dizisi Al, İşle ve Hedefe Yaz (Aralara Boşluk Ekleyerek) ---
Sub TestKaynakVeriAl_Tip2()
    Dim secilenData As Variant      ' Fonksiyondan dönen değer (dizi) veya hata
    Dim dict As Object              ' Benzersiz metinleri ve adetleri tutmak için Dictionary
    Dim i As Long, j As Long        ' Dizi indisleri için
    Dim lRow As Long, uRow As Long
    Dim lCol As Long, uCol As Long
    Dim birlesikMetin As String     ' Satırdaki elemanları birleştirmek için geçici değişken
    Dim key As Variant              ' Dictionary anahtarlarında döngü için
    Dim hedefDosyaYolu As String    ' GetSetting'den alınacak hedef dosya yolu
    Dim wbHedefGercek As Workbook   ' Sonuçların yazılacağı hedef çalışma kitabı
    Dim wsHedefGercek As Worksheet  ' Sonuçların yazılacağı hedef çalışma sayfası ("Sayfa1")
    Dim hedefSatir As Long          ' Hedef sayfada yazılacak başlangıç satırı
    Dim yazilanAdet As Long         ' Hedefe kaç satır yazıldığını saymak için

    ' --- Hata Yönetimi ---
    On Error GoTo HataYonetimi

    ' --- Dictionary Nesnesini Oluştur ---
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare ' Büyük/küçük harf duyarsız karşılaştırma

    ' --- 1. Fonksiyonu Çağır (Tip 2: Dizi Al) ---
    If wsHedef Is Nothing Then
        If Not ActiveSheet Is Nothing Then
            Set wsHedef = ActiveSheet
            Debug.Print "Test için global wsHedef aktif sayfaya ayarlandı: " & wsHedef.Name
        Else
             Debug.Print "UYARI: Fonksiyon çağrısı için aktif bir sayfa bulunamadı!"
             GoTo TemizCikis
        End If
    End If

    Application.ScreenUpdating = False
    secilenData = KullanicidanKaynakVeriAl(2) ' SecimTipi = 2
    Application.ScreenUpdating = True

    ' --- 2. Dönen Veriyi Doğrula ---
    If IsError(secilenData) Then
        Debug.Print "Fonksiyon (Tip 2) veri alamadı. Hata: " & CStr(secilenData)
        GoTo TemizCikis
    End If
    If Not IsArray(secilenData) Then
        Debug.Print "Hata: Fonksiyon (Tip 2) bir dizi döndürmedi!"
        GoTo TemizCikis
    End If
    On Error Resume Next
    lRow = LBound(secilenData, 1): uRow = UBound(secilenData, 1)
    lCol = LBound(secilenData, 2): uCol = UBound(secilenData, 2)
    If Err.Number <> 0 Then
        Debug.Print "Dönen veri bir dizi gibi görünse de boyutları okunamadı. Hata: " & Err.Description
        Err.Clear
        On Error GoTo HataYonetimi
        GoTo TemizCikis
    End If
    On Error GoTo HataYonetimi

    ' --- 3. Diziyi İşle: Satırları Birleştir (Araya Boşluk Koyarak) ve Say ---
    For i = lRow To uRow
        birlesikMetin = ""
        For j = lCol To uCol
            On Error Resume Next
            Dim hucreDegeriStr As String
            hucreDegeriStr = CStr(secilenData(i, j))
            If Err.Number <> 0 Then hucreDegeriStr = "": Err.Clear
            On Error GoTo HataYonetimi

            If birlesikMetin <> "" Then
                birlesikMetin = birlesikMetin & " " & hucreDegeriStr
            Else
                birlesikMetin = hucreDegeriStr
            End If
        Next j

        If Trim(birlesikMetin) <> "" Then
            If dict.exists(birlesikMetin) Then
                dict(birlesikMetin) = dict(birlesikMetin) + 1
            Else
                dict.Add birlesikMetin, 1
            End If
        End If
    Next i

    ' --- 4. İşlenecek Veri Var mı Kontrol Et ---
     If dict.Count = 0 Then
        Debug.Print "Kaynak seçimde işlenecek (boş olmayan) birleştirilmiş veri bulunamadı."
        GoTo TemizCikis
    End If

    ' --- 5. Hedef Çalışma Kitabını ve Sayfasını Ayarla ---
    hedefDosyaYolu = GetSetting("sercan", "fileOpenWorkBooks", "nowOpenPropsFile", "")
    If hedefDosyaYolu = "" Then
        Debug.Print "Hata: Hedef çalışma kitabı yolu kayıt defterinde bulunamadı veya boş."
        GoTo TemizCikis
    End If

    Set wbHedefGercek = Nothing
    On Error Resume Next
    Set wbHedefGercek = Workbooks(Mid(hedefDosyaYolu, InStrRev(hedefDosyaYolu, "\") + 1))
    On Error GoTo HataYonetimi
    If wbHedefGercek Is Nothing Then
        Debug.Print "Hata: Belirtilen hedef çalışma kitabı açık değil: " & hedefDosyaYolu
        GoTo TemizCikis
    End If

    Set wsHedefGercek = Nothing
    On Error Resume Next
    Set wsHedefGercek = wbHedefGercek.Worksheets("Sayfa1")
    On Error GoTo HataYonetimi
    If wsHedefGercek Is Nothing Then
        Debug.Print "Hata: '" & wbHedefGercek.Name & "' kitabında 'Sayfa1' isimli sayfa bulunamadı!"
        GoTo TemizCikis
    End If

    ' --- 6. Hedefteki Son Satırı Bul ve Yazma Satırını Belirle --- <<< BU BÖLÜM GÜNCELLENDİ
    ' Daha basit ve güvenilir yöntem: B sütununda veri varsa son satırın bir altını bul,
    ' yoksa (tamamen boşsa) 1. satırdan başla.
    If wsHedefGercek Is Nothing Then
        ' MsgBox "Hedef sayfa nesnesi geçerli değil!", vbCritical ' Yoruma alındı
        Debug.Print "Hata: Hedef sayfa nesnesi (wsHedefGercek) geçerli değil!"
        GoTo TemizCikis
    End If

    If Application.WorksheetFunction.CountA(wsHedefGercek.Columns("B")) = 0 Then
        ' B sütununda hiç veri yoksa, ilk satır 1'dir.
        hedefSatir = 1
    Else
        ' B sütununda veri varsa, son dolu satırı bul ve bir ekle.
        hedefSatir = wsHedefGercek.Cells(wsHedefGercek.Rows.Count, "B").End(xlUp).row + 1
    End If
    ' Debug.Print "Hedef yazma satırı belirlendi: " & hedefSatir ' Kontrol için eklenebilir
    ' --- GÜNCELLEME BİTTİ ---

    ' --- 7. Sonuçları Hedef Sayfaya Yaz ---
    Application.ScreenUpdating = False
    yazilanAdet = 0
    For Each key In dict.keys
        wsHedefGercek.Cells(hedefSatir, "B").Value = key
        wsHedefGercek.Cells(hedefSatir, "E").Value = dict(key)
        wsHedefGercek.Cells(hedefSatir, "E").NumberFormat = "#,##0"
        hedefSatir = hedefSatir + 1
        yazilanAdet = yazilanAdet + 1
    Next key
    Application.ScreenUpdating = True

    Debug.Print yazilanAdet & " adet benzersiz grup '" & wbHedefGercek.Name & "'::'" & wsHedefGercek.Name & "' sayfasına yazıldı."

    GoTo TemizCikis

HataYonetimi:
    MsgBox "Test prosedürü 2'de bir hata oluştu!" & vbCrLf & _
           "Hata Numarası: " & Err.Number & vbCrLf & _
           "Açıklama: " & Err.Description, vbCritical, "Test Hatası"
    Application.ScreenUpdating = True

TemizCikis:
    Set dict = Nothing
    Set wbHedefGercek = Nothing
    Set wsHedefGercek = Nothing
    On Error GoTo 0

End Sub

'###################################################
'kaynak sayfaya sinyal sigortası ve değer ekle
'####################################################
Sub sinyal_ekle()

    Dim hedefDosyaYolu As String
    Dim wbHedef As Workbook
    Dim wsHedef As Worksheet
    Dim hedefSonSatir As Long
    Dim wb As Workbook ' Açık çalışma kitapları arasında döngü için

    ' --- Hata Yönetimi ---
    On Error GoTo HataYonetimi

    ' --- 1. Hedef Çalışma Kitabının Yolunu GetSetting ile Al ---
    hedefDosyaYolu = GetSetting("sercan", "fileOpenWorkBooks", "nowOpenPropsFile", "")
    If hedefDosyaYolu = "" Then
        Debug.Print "Hata: Hedef çalışma kitabı yolu kayıt defterinde bulunamadı veya boş."
        MsgBox "Hedef çalışma kitabı yolu kayıt defterinde bulunamadı veya boş.", vbCritical, "Hata"
        GoTo TemizCikis
    End If
    Debug.Print "Hedef dosya yolu kayıt defterinden alındı: " & hedefDosyaYolu

    ' --- 2. Hedef Çalışma Kitabını Açık Olanlar Arasında Bul ---
    Set wbHedef = Nothing ' Başlangıçta boşalt
    On Error Resume Next ' Kapalı kitap varsa hata vermesin diye değil, daha çok döngü için genel önlem
    For Each wb In Application.Workbooks
        ' Tam dosya yolu ile karşılaştır (büyük/küçük harf duyarsız)
        If StrComp(wb.fullName, hedefDosyaYolu, vbTextCompare) = 0 Then
            Set wbHedef = wb
            Exit For ' Bulunca döngüden çık
        End If
    Next wb
    On Error GoTo HataYonetimi ' Hata yakalamayı normale döndür

    ' Hedef çalışma kitabı bulunamadıysa hata ver ve çık
    If wbHedef Is Nothing Then
        Debug.Print "Hata: Belirtilen hedef çalışma kitabı şu anda açık değil: " & hedefDosyaYolu
        MsgBox "GetSetting ile belirtilen hedef çalışma kitabı açık değil:" & vbCrLf & hedefDosyaYolu, vbCritical, "Hata: Kitap Açık Değil"
        GoTo TemizCikis
    End If
    Debug.Print "Hedef çalışma kitabı bulundu: " & wbHedef.Name

    ' --- 3. Hedef Çalışma Sayfasını Ayarla ("Sayfa1") ---
    Set wsHedef = Nothing ' Başlangıçta boşalt
    On Error Resume Next ' Sayfa yoksa hata vermesin
    Set wsHedef = wbHedef.Worksheets("Sayfa1")
    On Error GoTo HataYonetimi ' Hata yakalamayı normale döndür

    ' Hedef sayfa bulunamadıysa hata ver ve çık
    If wsHedef Is Nothing Then
        Debug.Print "Hata: '" & wbHedef.Name & "' kitabında 'Sayfa1' isimli sayfa bulunamadı!"
        MsgBox "'" & wbHedef.Name & "' çalışma kitabında 'Sayfa1' isimli sayfa bulunamadı!", vbCritical, "Hata: Sayfa Bulunamadı"
        GoTo TemizCikis
    End If
    Debug.Print "Hedef çalışma sayfası bulundu: " & wsHedef.Name

    ' --- 4. Hedef Sayfada B Sütununa Göre Sonraki Boş Satırı Bul ---
    ' B sütununda hiç veri yoksa 1. satırdan başla, varsa son dolu satırın bir altını bul.
    If Application.WorksheetFunction.CountA(wsHedef.Columns("B")) = 0 Then
        hedefSonSatir = 1
        Debug.Print "Hedef sayfada B sütunu boş. Yazılacak ilk satır: 1"
    Else
        hedefSonSatir = wsHedef.Cells(wsHedef.Rows.Count, "B").End(xlUp).row + 1
        Debug.Print "Hedef sayfada B sütunundaki son dolu satırdan sonraki satır: " & hedefSonSatir
    End If

    ' --- 5. Verileri Hedef Satırlara Yaz ---
    Application.ScreenUpdating = False ' Ekran titremesini önle

    ' İlk veri seti (Sinyal Sigortası)
    wsHedef.Cells(hedefSonSatir, "B").Value = "B 1x6 A.O.S"
    wsHedef.Cells(hedefSonSatir, "E").Value = 3 ' Adet
    Debug.Print hedefSonSatir & ". satıra yazıldı: B=" & wsHedef.Cells(hedefSonSatir, "B").Value & ", E=" & wsHedef.Cells(hedefSonSatir, "E").Value

    ' İkinci veri seti (Sinyal Ledi) - bir sonraki satıra
    wsHedef.Cells(hedefSonSatir + 1, "B").Value = "6601"
    wsHedef.Cells(hedefSonSatir + 1, "E").Value = 3 ' Adet
    Debug.Print (hedefSonSatir + 1) & ". satıra yazıldı: B=" & wsHedef.Cells(hedefSonSatir + 1, "B").Value & ", E=" & wsHedef.Cells(hedefSonSatir + 1, "E").Value

    Application.ScreenUpdating = True

    ' --- Başarı Mesajı (Opsiyonel) ---
    ' MsgBox "Veriler '" & wsHedef.Name & "' sayfasına başarıyla eklendi.", vbInformation, "İşlem Tamamlandı"
    Debug.Print "Sabit veriler hedef sayfaya başarıyla eklendi."

    GoTo TemizCikis ' Hata etiketini atla

HataYonetimi:
    ' Hata oluştuğunda kullanıcıyı bilgilendir
    MsgBox "Veri ekleme işlemi sırasında bir hata oluştu!" & vbCrLf & vbCrLf & _
           "Hata Numarası: " & Err.Number & vbCrLf & _
           "Açıklama: " & Err.Description, vbCritical, "İşlem Hatası"
    ' Hata durumunda ekran güncellemesini açmayı dene (zaten açıksa sorun olmaz)
    On Error Resume Next
    Application.ScreenUpdating = True
    On Error GoTo 0 ' Hata işleyiciyi sıfırla

TemizCikis:
    ' --- Kullanılan Nesneleri Temizle ---
    Set wsHedef = Nothing
    Set wbHedef = Nothing
    Set wb = Nothing ' Döngü değişkenini de temizle
    ' --- Hata Yakalamayı Sıfırla ---
    On Error GoTo 0

End Sub

'--- Kaynak Kitabı Kapatmak İçin Yardımcı Sub ---
Sub CloseGlobalSourceWorkbook()
    If Not g_wbKaynak Is Nothing Then
        On Error Resume Next ' Zaten kapalıysa veya başka sorun varsa
        Dim workbookName As String
        workbookName = g_wbKaynak.Name ' İsim almayı dene
        If Err.Number = 0 Then
            If Not ThisWorkbook Is g_wbKaynak Then ' Kendimizi kapatmayalım
                Debug.Print "Global kaynak kitap kapatılıyor: " & workbookName
                g_wbKaynak.Close SaveChanges:=False
                If Err.Number <> 0 Then
                    MsgBox "Kaynak kitap kapatılırken hata oluştu: " & Err.Description, vbExclamation
                Else
                    MsgBox "'" & workbookName & "' kaynak kitabı kapatıldı.", vbInformation
                End If
            Else
                 MsgBox "Global kaynak kitap olarak mevcut çalışma kitabı ayarlanmış, kapatılmadı.", vbExclamation
            End If
        Else
            MsgBox "Global kaynak kitap nesnesi vardı ancak kitap bulunamadı/kapatılamadı.", vbExclamation
        End If
         Err.Clear
         On Error GoTo 0
    Else
        MsgBox "Kapatılacak global kaynak kitap ayarlanmamış.", vbInformation
    End If

    ' Global değişkenleri sıfırla
    Set g_wbKaynak = Nothing
    secilenDosyaYolu = Empty
    ' wsHedef'i de sıfırlamak isteyebilirsiniz:
    ' Set wsHedef = Nothing
End Sub


' =========================================================================
' Fonksiyon: FindNextEmptyRowInColumn
' Amaç    : Belirtilen çalışma sayfasının belirtilen sütunundaki ilk boş
'           satırın numarasını bulur.
' Parametreler:
'   TargetSheet As Worksheet : Boş satırın aranacağı çalışma sayfası.
'   ColumnIdentifier As Variant : Boş satırın aranacağı sütun.
'                                 Bu bir harf (örn. "B", "C", "AA") veya
'                                 bir sayı (örn. 2, 3, 27) olabilir.
' Dönüş Değeri:
'   Long : Belirtilen sütundaki ilk boş satırın numarası.
'          Eğer TargetSheet geçersizse, ColumnIdentifier geçersizse veya
'          bir hata oluşursa 0 döner.
' Notlar  : Sadece görünür değere göre işlem yapar. Formül sonucu "" olan
'           hücreler dolu sayılmaz (End(xlUp) davranışı).
' =========================================================================
Function FindNextEmptyRowInColumn(TargetSheet As Worksheet, ColumnIdentifier As Variant) As Long
    Dim lastUsedRow As Long
    Dim NextRow As Long
    Dim colNum As Long ' Sütun numarasını tutmak için

    ' --- Giriş Doğrulaması ---
    If TargetSheet Is Nothing Then
        FindNextEmptyRowInColumn = 0 ' Geçersiz sayfa
        Exit Function
    End If

    ' ColumnIdentifier'ın geçerli olup olmadığını kontrol et (basit kontrol)
    If IsEmpty(ColumnIdentifier) Or IsNull(ColumnIdentifier) Then
         FindNextEmptyRowInColumn = 0 ' Geçersiz sütun tanımlayıcı
         Exit Function
    End If
    
    ' Sütun tanımlayıcısını kullanarak Cells'e erişmeyi dene
    On Error Resume Next ' Hatalı sütun adı/numarası hatasını yakala
    
    ' 1. Son Dolu Satırı Bul
    lastUsedRow = TargetSheet.Cells(TargetSheet.Rows.Count, ColumnIdentifier).End(xlUp).row
    If Err.Number <> 0 Then ' Cells veya End(xlUp) hatası oluştu mu?
        FindNextEmptyRowInColumn = 0 ' Hata durumunda 0 döndür
        On Error GoTo 0          ' Hata yönetimini sıfırla
        Exit Function
    End If

    ' 2. İlk Hücreyi Kontrol Et (Eğer son satır 1 ise)
    Dim firstCellIsEmpty As Boolean
    firstCellIsEmpty = IsEmpty(TargetSheet.Cells(1, ColumnIdentifier).Value)
    If Err.Number <> 0 Then ' İlk hücreye erişimde hata oldu mu?
        FindNextEmptyRowInColumn = 0 ' Hata durumunda 0 döndür
        On Error GoTo 0          ' Hata yönetimini sıfırla
        Exit Function
    End If
    
    On Error GoTo 0 ' Normal hata yönetimine geri dön

    ' --- Sonraki Boş Satırı Belirle ---
    If lastUsedRow = 1 And firstCellIsEmpty Then
        NextRow = 1 ' Sütun tamamen boş veya ilk hücre boş
    Else
        NextRow = lastUsedRow + 1 ' Son dolu hücrenin bir altındaki satır
    End If

    FindNextEmptyRowInColumn = NextRow ' Hesaplanan satır numarasını döndür

End Function


' =========================================================================
' Ana Prosedür: sarf_iscilik_bakir
' =========================================================================
Sub sarf_iscilik_bakir()

    Dim hedefDosyaYolu As String
    Dim wbHedef As Workbook
    Dim wsHedef As Worksheet
    Dim hedefSatir As Long
    Dim hedefHucre As Range
    Dim wb As Workbook ' Çalışma kitaplarında döngü için
    Const TARGET_COLUMN As String = "B" ' Hedef sütunu sabit olarak tanımla

    ' --- Hata Yönetimi ve Ekran Güncelleme ---
    On Error GoTo HataYonetimi
    Application.ScreenUpdating = False ' Ekran güncellemeyi kapat

    ' --- 1. Hedef Dosya Yolunu Kayıt Defterinden Al ---
    hedefDosyaYolu = GetSetting("sercan", "fileOpenWorkBooks", "nowOpenPropsFile", "")
    If hedefDosyaYolu = "" Then
        MsgBox "Hedef çalışma kitabı yolu kayıt defterinde bulunamadı veya boş." & vbCrLf & _
               "(Anahtar: sercan\fileOpenWorkBooks\nowOpenPropsFile)", vbCritical, "Hata: Kayıt Defteri Yolu Eksik"
        GoTo TemizCikis
    End If

    ' --- 2. Hedef Çalışma Kitabını Bul (Açık Olanlar Arasında) ---
    Set wbHedef = Nothing
    For Each wb In Application.Workbooks
        If StrComp(wb.fullName, hedefDosyaYolu, vbTextCompare) = 0 Then
            Set wbHedef = wb
            Exit For
        End If
    Next wb
    If wbHedef Is Nothing Then
        MsgBox "Belirtilen hedef çalışma kitabı şu anda açık değil:" & vbCrLf & hedefDosyaYolu, vbCritical, "Hata: Kitap Açık Değil"
        GoTo TemizCikis
    End If

    ' --- 3. Hedef Çalışma Sayfasını ("Sayfa1") Bul ---
    Set wsHedef = Nothing
    On Error Resume Next
    Set wsHedef = wbHedef.Worksheets("Sayfa1")
    On Error GoTo HataYonetimi ' Normal hata yönetimine geri dön
    If wsHedef Is Nothing Then
        MsgBox "'" & wbHedef.Name & "' çalışma kitabında 'Sayfa1' isimli sayfa bulunamadı!", vbCritical, "Hata: Sayfa Bulunamadı"
        GoTo TemizCikis
    End If

    ' --- 4. Hedef Sütundaki İlk Boş Satırı Bul (Fonksiyon Kullanarak) ---
    hedefSatir = FindNextEmptyRowInColumn(wsHedef, TARGET_COLUMN) ' Fonksiyona sütunu da gönder
    If hedefSatir = 0 Then ' Fonksiyon hata döndürdüyse
        MsgBox "Hedef sayfada ('" & wsHedef.Name & "') '" & TARGET_COLUMN & "' sütunundaki ilk boş satır belirlenemedi." & vbCrLf & _
               "Sütun adı/numarası geçerli mi?", vbCritical, "Hata: Satır Bulunamadı"
        GoTo TemizCikis
    End If

    ' --- 5. Hedef Hücreyi Aktif Et ---
    wbHedef.Activate
    wsHedef.Activate
    Set hedefHucre = wsHedef.Cells(hedefSatir, TARGET_COLUMN) ' Hedef hücreyi belirle
    hedefHucre.Activate                                      ' Aktif et

    ' --- 6. Veri Girişlerini Yap ---
    ' yeni satir belirle ve bakir ekle
    'hedefSatir = FindNextEmptyRowInColumn(wsHedef, TARGET_COLUMN) ' Tekrar fonksiyonu çağır
    'hedefHucre.Activate                                      ' Yeni boş hücreyi aktif et
    Call bakir_gir
    
    'isçilik satırı gir
    hedefSatir = FindNextEmptyRowInColumn(wsHedef, TARGET_COLUMN) ' Tekrar fonksiyonu çağır
    hedefHucre.Activate
    Call isciliksatır_gir1
    
    'sarfı işçilik altına denk gelecek şekilde ekle. ( formül sebebiyle )
    hedefSatir = FindNextEmptyRowInColumn(wsHedef, TARGET_COLUMN) ' Tekrar fonksiyonu çağır
    hedefSatir = hedefSatir + 1
    hedefHucre.Activate                                      ' Yeni boş hücreyi aktif et
    Call sarfsatır_gir1
    

    ' --- 7. Yeni Satır Ekle ve Nokta Koy ---
    ' Not: Satır ekleme işlemi, mevcut hedefSatir'in *altına* ekler.
    wsHedef.Rows(hedefSatir).Insert Shift:=xlDown
    wsHedef.Cells(hedefSatir, TARGET_COLUMN).Value = "."  ' Eklenen yeni satırın hedef sütununa nokta koy

    ' --- 8. Yeni Eklenen Satırdan Sonraki Boş Satırı Bul ve Aktif Et ---
    hedefSatir = FindNextEmptyRowInColumn(wsHedef, TARGET_COLUMN) ' Tekrar fonksiyonu çağır

    Set hedefHucre = wsHedef.Cells(hedefSatir, TARGET_COLUMN) ' Yeni hedef hücre
    hedefHucre.Activate                                      ' Yeni boş hücreyi aktif et

    ' --- 9. Toplam Hesaplamasını Çağır ---
    Call toplam1(Nothing) ' Eğer toplam1 argüman alıyorsa uygun şekilde düzenleyin
    
    'Pano işçilik tutarını hesapla ve ilgili satıra yaz
    'Call PanoHesaplaVePmMpSatirinaYaz  ' bölüm adı ve toplamı arasındakipano kalemlerini hesaplayıp işçiliğe değer olarak atar. işçilik baremi sayfa3-ads değeridir
    Call SumRonFormulasToPmMpRow   ' bölüm adı ve toplamı arasındakipano kalemlerini hesaplayıp işçiliğe formül olarak atar.
    
    ' --- Başarı Mesajı (İsteğe Bağlı) ---
    ' MsgBox "İşlemler tamamlandı. Sonraki giriş hücresi: " & TARGET_COLUMN & hedefSatir, vbInformation

' --- Temizlik ve Çıkış ---
TemizCikis:
    Application.ScreenUpdating = True ' Ekran güncellemeyi tekrar aç
    Set wsHedef = Nothing
    Set wbHedef = Nothing
    Set hedefHucre = Nothing
    Set wb = Nothing
    Exit Sub

' --- Hata Yönetimi Bloğu ---
HataYonetimi:
    MsgBox "Beklenmeyen bir hata oluştu!" & vbCrLf & vbCrLf & _
           "Hata Numarası: " & Err.Number & vbCrLf & _
           "Açıklama: " & Err.Description, vbCritical, "İşlem Hatası"
    Resume TemizCikis ' Hata durumunda temizliğe git ve çık

End Sub

'/////////////////////////////////////////////////////////////////////////////////////////////////
' =========================================================================
' Yardımcı Fonksiyon: ExtractNumberFromString
' Amaç    : Belirtilen metin içinde, verilen başlangıç karakterinden
'           sonraki 3 veya 4 haneli sayıyı (arada 'x' olup olmadığına
'           göre) ayıklar. Excel formülündeki mantığı taklit eder.
' ... (Fonksiyon aynı kalıyor, önceki cevaplarda olduğu gibi) ...
' =========================================================================
Function ExtractNumberFromString(sourceText As String, startChar As String) As Double
    Dim startPos As Long
    Dim searchStr As String
    Dim numPart As String
    Dim xPos As Long
    Dim extractedValue As Double

    extractedValue = 0 ' Varsayılan değer hata durumunda 0

    On Error Resume Next ' Metin işleme hatalarını yakalamak için

    startPos = InStr(1, sourceText, startChar, vbBinaryCompare)

    If startPos > 0 Then
        If Len(sourceText) >= startPos + 4 Then
            searchStr = Mid$(sourceText, startPos + 1, 4)
        ElseIf Len(sourceText) >= startPos + 1 Then
            searchStr = Mid$(sourceText, startPos + 1)
        Else
            searchStr = ""
        End If

        If Len(searchStr) > 0 Then
            xPos = InStr(1, searchStr, "x", vbBinaryCompare)

            If xPos > 0 Then
                numPart = Left$(searchStr, 3)
            Else
                numPart = searchStr
            End If

            If IsNumeric(numPart) Then
                extractedValue = val(numPart)
            End If
        End If
    End If

    On Error GoTo 0
    ExtractNumberFromString = extractedValue
End Function

' =========================================================================
' Ana Prosedür: HesaplaRonVePmMpSatirinaYaz
' Amaç    : Seçili hücrenin bulunduğu bölümdeki "ron" ile başlayan
'           satırlar için özel bir hesaplama yapar (adetler dahil)
'           ve sonucu bölüm içindeki ilk "PM-MP" satırının F sütununa yazar.
' =========================================================================
Sub PanoHesaplaVePmMpSatirinaYaz()

    Dim ws As Worksheet
    Dim currentRow As Range
    Dim startRow As Long
    Dim endRow As Long
    Dim scanRow As Long
    Dim cellB_Value As String
    Dim adsValue As Double
    Dim val1 As Double
    Dim val2 As Double
    Dim adetValue As Double
    Dim cellE_Value As Variant
    Dim calculatedValue As Double
    Dim totalSum As Double
    Dim foundStart As Boolean
    Dim foundEnd As Boolean
    Dim pmMpTargetRow As Long ' "PM-MP" satırının numarasını tutacak
    Dim i As Long

    On Error GoTo HataYonetimi
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual ' Hesaplamayı manuel yap

    ' --- Aktif Sayfa ve Seçili Hücreyi Al ---
    Set ws = ActiveSheet
    Set currentRow = Selection.Cells(1, 1)

    ' --- "Ads" Değerini Al ---
    On Error Resume Next
    adsValue = ws.Evaluate("Ads")
    If Err.Number <> 0 Or Not IsNumeric(adsValue) Then
        MsgBox "'Ads' adlı aralık bulunamadı veya sayısal bir değer içermiyor.", vbCritical, "Hata: 'Ads' Değeri Eksik/Hatalı"
        GoTo TemizCikis
    End If
    On Error GoTo HataYonetimi

    ' --- Başlangıç Satırını ("BÖLÜM ADI/NO:") Bul ---
    foundStart = False
    For i = currentRow.row - 1 To 1 Step -1
        If StrComp(Trim(ws.Cells(i, "B").Text), "BÖLÜM ADI/NO:", vbTextCompare) = 0 Then
            startRow = i
            foundStart = True
            Exit For
        End If
    Next i

    ' --- Bitiş Satırını ("BÖLÜM TOPLAMI:") Bul ---
    foundEnd = False
    Dim lastRow As Long, lastRowE As Long
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).row
    lastRowE = ws.Cells(ws.Rows.Count, "E").End(xlUp).row
    If lastRowE > lastRow Then lastRow = lastRowE

    For i = currentRow.row To lastRow
        If StrComp(Trim(ws.Cells(i, "B").Text), "BÖLÜM TOPLAMI:", vbTextCompare) = 0 Then
            endRow = i
            foundEnd = True
            Exit For
        End If
    Next i

    ' --- Sınırlar Bulundu mu Kontrol Et ---
    If Not foundStart Then
        MsgBox "Seçili hücreden önce 'BÖLÜM ADI/NO:' bulunamadı.", vbExclamation, "Sınır Bulunamadı"
        GoTo TemizCikis
    End If
    If Not foundEnd Then
        MsgBox "Seçili hücreden sonra 'BÖLÜM TOPLAMI:' bulunamadı.", vbExclamation, "Sınır Bulunamadı"
        GoTo TemizCikis
    End If
     If startRow >= endRow Then
         MsgBox "'BÖLÜM ADI/NO:' satırı ('" & startRow & "') 'BÖLÜM TOPLAMI:' satırından ('" & endRow & "') sonra veya aynı satırda olamaz.", vbCritical, "Geçersiz Aralık"
         GoTo TemizCikis
    End If

    ' --- Satırları Tara, Hesapla ve PM-MP Satırını Bul ---
    totalSum = 0
    pmMpTargetRow = 0 ' Hedef satırı başlangıçta sıfırla

    For scanRow = startRow + 1 To endRow - 1 ' Başlangıç ve bitiş hariç
        cellB_Value = Trim(ws.Cells(scanRow, "B").Text)

        ' 1. PM-MP satırını ara (henüz bulunmadıysa)
        If pmMpTargetRow = 0 Then ' Henüz hedef satır bulunmadıysa ara
            If StrComp(cellB_Value, "PM-MP", vbTextCompare) = 0 Then
                pmMpTargetRow = scanRow ' Hedef satır numarasını kaydet
                ' İlk bulunan yeterli olduğu için aramayı bırakmaya gerek yok,
                ' aynı döngüde 'ron' kontrolüne devam edebiliriz.
            End If
        End If

        ' 2. "ron" ile başlayan satırları hesapla
        If LCase$(Left$(cellB_Value, 3)) = "ron" Then
            val1 = ExtractNumberFromString(cellB_Value, " ")
            val2 = ExtractNumberFromString(cellB_Value, "x")
            cellE_Value = ws.Cells(scanRow, "E").Value
            adetValue = 0

            If IsNumeric(cellE_Value) Then
                 If Not IsEmpty(cellE_Value) Then
                    adetValue = CDbl(cellE_Value)
                 End If
            End If

            If val1 <> 0 Or val2 <> 0 Then
                 calculatedValue = (val1 * val2 * (adsValue / 100#) / 100#) * adetValue
                 totalSum = totalSum + calculatedValue
            End If
        End If ' End If "ron" kontrolü
    Next scanRow

    ' --- Sonucu Hedef Satıra Yaz ---
    If pmMpTargetRow > 0 Then ' PM-MP satırı bulunduysa
        ws.Cells(pmMpTargetRow, "F").Value = totalSum
        ' İsteğe bağlı: Hücreye format uygulayabilirsiniz
        ' ws.Cells(pmMpTargetRow, "F").NumberFormat = "#,##0.0000"
    Else
        ' PM-MP satırı bulunamadıysa kullanıcıyı bilgilendir
        MsgBox "Hesaplama tamamlandı (Toplam: " & Format$(totalSum, "#,##0.0000") & "). Ancak bölüm içinde 'PM-MP' içeren bir satır bulunamadığı için sonuç yazılamadı.", vbExclamation, "Hedef Satır Bulunamadı"
    End If

    ' --- MsgBox'ı yoruma al ---
    ' MsgBox "Bulunan 'ron' satırları için (adetler dahil) hesaplanan toplam değer:" & vbCrLf & vbCrLf & _
    '        Format$(totalSum, "#,##0.0000"), vbInformation, "Adetli Hesaplama Sonucu"

' --- Temizlik ve Çıkış ---
TemizCikis:
    Application.Calculation = xlCalculationAutomatic ' Hesaplamayı otomatiğe al
    Application.ScreenUpdating = True
    Set ws = Nothing
    Set currentRow = Nothing
    Exit Sub

' --- Hata Yönetimi Bloğu ---
HataYonetimi:
    MsgBox "Beklenmeyen bir hata oluştu!" & vbCrLf & vbCrLf & _
           "Hata Numarası: " & Err.Number & vbCrLf & _
           "Açıklama: " & Err.Description & vbCrLf & _
           "Hata Oluşan Satır (Tahmini): " & Erl, vbCritical, "Makro Hatası"
    ' Hata durumunda hesaplamayı otomatiğe almayı unutma
    Application.Calculation = xlCalculationAutomatic
    Resume TemizCikis

End Sub

' =========================================================================
' Ana Prosedür: SumRonFormulasToPmMpRow
' Amaç    : Seçili hücrenin bulunduğu bölümdeki her "ron" satırı için
'           işçilik formülünü temsil eden bir metin oluşturur. Bu metinleri
'           toplayarak ("+" ile birleştirerek) tek bir SUM formülü haline getirir
'           ve bu formülü bölüm içindeki ilk "PM-MP" satırının F sütununa yazar.
' =========================================================================
Sub SumRonFormulasToPmMpRow()

    Dim ws As Worksheet
    Dim currentRow As Range
    Dim startRow As Long
    Dim endRow As Long
    Dim scanRow As Long
    Dim cellB_Value As String
    Dim adsExists As Boolean
    Dim formulaParts As Object ' Collection kullanacağız
    Dim individualCalcString As String ' Her bir ron satırı için formül parçası
    Dim finalFormula As String
    Dim foundStart As Boolean
    Dim foundEnd As Boolean
    Dim pmMpTargetRow As Long ' "PM-MP" satırının numarasını tutacak
    Dim i As Long
    Dim partIndex As Long ' Dizi için

    On Error GoTo HataYonetimi
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual ' Hesaplamayı manuel yap

    ' --- Aktif Sayfa ve Seçili Hücreyi Al ---
    Set ws = ActiveSheet
    Set currentRow = Selection.Cells(1, 1)
    Set formulaParts = CreateObject("System.Collections.ArrayList") ' Daha esnek ArrayList

    ' --- "Ads" Adlı Aralığın Varlığını Kontrol Et ---
    adsExists = False
    On Error Resume Next
    If Not IsEmpty(ActiveWorkbook.names("Ads").Name) Then
        adsExists = True
    End If
    On Error GoTo HataYonetimi
    If Not adsExists Then
        MsgBox "'Ads' adlı aralık (Named Range) bu çalışma kitabında tanımlı değil." & vbCrLf & _
               "Formül '#AD?' hatası verebilir.", vbExclamation, "Uyarı: 'Ads' Tanımlı Değil"
    End If

    ' --- Başlangıç Satırını ("BÖLÜM ADI/NO:") Bul ---
    foundStart = False
    For i = currentRow.row - 1 To 1 Step -1
        If StrComp(Trim(ws.Cells(i, "B").Text), "BÖLÜM ADI/NO:", vbTextCompare) = 0 Then
            startRow = i
            foundStart = True
            Exit For
        End If
    Next i

    ' --- Bitiş Satırını ("BÖLÜM TOPLAMI:") Bul ---
    foundEnd = False
    Dim lastRow As Long, lastRowE As Long
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).row
    lastRowE = ws.Cells(ws.Rows.Count, "E").End(xlUp).row
    If lastRowE > lastRow Then lastRow = lastRowE

    For i = currentRow.row To lastRow
        If StrComp(Trim(ws.Cells(i, "B").Text), "BÖLÜM TOPLAMI:", vbTextCompare) = 0 Then
            endRow = i
            foundEnd = True
            Exit For
        End If
    Next i

    ' --- Sınırlar Bulundu mu Kontrol Et ---
    If Not foundStart Then MsgBox "Başlangıç sınırı bulunamadı.", vbCritical: GoTo TemizCikis
    If Not foundEnd Then MsgBox "Bitiş sınırı bulunamadı.", vbCritical: GoTo TemizCikis
    If startRow >= endRow Then MsgBox "Başlangıç ve bitiş sınırları geçersiz.", vbCritical: GoTo TemizCikis

    ' --- Satırları Tara, Formül Parçalarını Topla ve PM-MP Satırını Bul ---
    pmMpTargetRow = 0

    For scanRow = startRow + 1 To endRow - 1
        cellB_Value = Trim(ws.Cells(scanRow, "B").Text)

        ' 1. PM-MP satırını ara (henüz bulunmadıysa)
        If pmMpTargetRow = 0 Then
            If StrComp(cellB_Value, "PM-MP", vbTextCompare) = 0 Then
                pmMpTargetRow = scanRow
            End If
        End If

        ' 2. "ron" ile başlayan satırların formül parçalarını oluştur
        If LCase$(Left$(cellB_Value, 3)) = "ron" Then
            ' Her bir ron satırı için hesaplama formülünü (adet dahil) metin olarak oluştur:
            individualCalcString = "IFERROR((IFERROR(IF(FIND(""x"",MID(B" & scanRow & ",FIND("" "",B" & scanRow & ")+1,4)),MID(B" & scanRow & ",FIND("" "",B" & scanRow & ")+1,3),0),MID(B" & scanRow & ",FIND("" "",B" & scanRow & ")+1,4)) * IFERROR(IF(FIND(""x"",MID(B" & scanRow & ",FIND(""x"",B" & scanRow & ")+1,4)),MID(B" & scanRow & ",FIND(""x"",B" & scanRow & ")+1,3),0),MID(B" & scanRow & ",FIND(""x"",B" & scanRow & ")+1,4)) * (Ads/100) / 100) * N(E" & scanRow & "), 0)"
            ' Oluşturulan formül parçasını listeye ekle
            formulaParts.Add individualCalcString
        End If
    Next scanRow

    ' --- Sonucu Hedef Satıra Yaz ---
    If pmMpTargetRow = 0 Then
        MsgBox "Bölüm içinde 'PM-MP' içeren hedef satır bulunamadı. İşlem yapılamadı.", vbExclamation, "Hedef Satır Eksik"
        GoTo TemizCikis
    End If

    ' Formül parçalarını birleştirerek nihai formülü oluştur
    If formulaParts.Count > 0 Then
        ' ArrayList'i Join ile birleştirmek için önce bir Variant diziye atamamız gerekir
         Dim partsArray As Variant
         partsArray = formulaParts.ToArray() ' ArrayList'i diziye çevir
         finalFormula = "=" & Join(partsArray, "+") ' Parçaları "+" ile birleştir
         
         ' Alternatif (Eğer SUM kullanmak isterseniz, ama + daha mantıklı):
         ' finalFormula = "=SUM(" & Join(partsArray, ",") & ")"

         ' Çok uzun formül kontrolü (Excel'in formül karakter limiti ~8192)
         If Len(finalFormula) > 8000 Then ' Limite yakınsa uyar
             MsgBox "Uyarı: Oluşturulan formül çok uzun (" & Len(finalFormula) & " karakter)." & vbCrLf & _
                    "Excel formül limiti aşılmış olabilir ve hata verebilir.", vbExclamation, "Formül Uzunluğu Uyarısı"
         End If

    Else
        ' Hiç "ron" satırı bulunamadıysa PM-MP satırına 0 yaz
        finalFormula = "=0"
    End If

    ' Nihai formülü PM-MP satırının F sütununa yaz
    On Error Resume Next ' Yazma hatasını yakala
    ws.Cells(pmMpTargetRow, "F").Formula = finalFormula
    If Err.Number <> 0 Then
        MsgBox "'PM-MP' satırının (" & pmMpTargetRow & ") F sütununa formül yazılırken hata oluştu: " & Err.Description & vbCrLf & vbCrLf & _
               "Oluşturulan Formül (ilk 500 kr): " & Left(finalFormula, 500), vbCritical, "Yazma Hatası"
        Err.Clear
    Else
        ' Başarı mesajı (isteğe bağlı)
        ' MsgBox formulaParts.Count & " adet 'ron' satırının toplam formülü 'PM-MP' satırının F sütununa başarıyla yazıldı.", vbInformation, "İşlem Tamamlandı"
    End If
    On Error GoTo HataYonetimi ' Normal hata yönetimine dön

' --- Temizlik ve Çıkış ---
TemizCikis:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Set ws = Nothing
    Set currentRow = Nothing
    Set formulaParts = Nothing
    Exit Sub

' --- Hata Yönetimi Bloğu ---
HataYonetimi:
    MsgBox "Beklenmeyen bir hata oluştu!" & vbCrLf & vbCrLf & _
           "Hata Numarası: " & Err.Number & vbCrLf & _
           "Açıklama: " & Err.Description & vbCrLf & _
           "Hata Oluşan Satır (Tahmini): " & Erl, vbCritical, "Makro Hatası"
    If Application.Calculation <> xlCalculationAutomatic Then Application.Calculation = xlCalculationAutomatic
    Resume TemizCikis

End Sub