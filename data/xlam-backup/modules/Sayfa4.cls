Public DegerleriEsleVeYazCalisiyor As Boolean

Sub BenzersizVerileriListele()

    Dim wsKaynak As Worksheet
    Dim wsHedef As Worksheet
    Dim sonSatirKaynak As Long
    Dim veri As Variant
    Dim benzersizVeriler As Object
    Dim anahtar As Variant
    Dim i As Long

    ' DegerleriEsleVeYaz çalışıyorsa çık
    If DegerleriEsleVeYazCalisiyor Then Exit Sub

    ' Sayfa nesnelerini tanımla
    Set wsKaynak = ThisWorkbook.Sheets("Sayfa1")
    Set wsHedef = ThisWorkbook.Sheets("Ayar")

    ' Kaynak sayfadaki son dolu satırı bul
    sonSatirKaynak = wsKaynak.Cells(wsKaynak.Rows.Count, "B").End(xlUp).row

    ' Benzersiz verileri saklamak için bir Dictionary nesnesi oluştur
    Set benzersizVeriler = CreateObject("Scripting.Dictionary")

    ' Kaynak sayfadaki verileri döngüyle kontrol et
    For i = 1 To sonSatirKaynak
        veri = wsKaynak.Cells(i, "B").Value

        ' Belirli koşulları sağlayan verileri filtrele
        If veri <> "BÖLÜM ADI/NO:" And veri <> "BÖLÜM TOPLAMI:" And veri <> "siparis" And Left(veri, 1) <> "P" And Right(veri, 1) <> "." Then
            ' Veriyi Dictionary'ye ekle (benzersiz olacak şekilde)
            benzersizVeriler(veri) = 1
        End If
    Next i

    ' Hedef sayfadaki mevcut verileri temizle
    wsHedef.Range("J4:J" & wsHedef.Rows.Count).ClearContents

    ' Benzersiz verileri hedef sayfaya yaz
    i = 4
    For Each anahtar In benzersizVeriler.keys
        wsHedef.Cells(i, "J").Value = anahtar
        i = i + 1
    Next anahtar

    ' Belleği temizle
    Set benzersizVeriler = Nothing
    Set wsKaynak = Nothing
    Set wsHedef = Nothing

    'MsgBox "Benzersiz veriler Ayar sayfasına J4'ten itibaren listelendi.", vbInformation

End Sub

Sub DegerleriEsleVeYaz()

    Dim wsKaynak As Worksheet
    Dim wsAyar As Worksheet
    Dim sonSatirAyar As Long
    Dim sonSatirKaynak As Long
    Dim i As Long
    Dim j As Long
    Dim deger As Variant
    Dim eslesmeBulundu As Boolean

    ' Prosedür çalışıyor
    DegerleriEsleVeYazCalisiyor = True

    ' Çalışma sayfalarını tanımla
    Set wsKaynak = ThisWorkbook.Sheets("Sayfa1")
    Set wsAyar = ThisWorkbook.Sheets("Ayar")

    ' Ayar sayfasındaki son dolu satırı bul
    sonSatirAyar = wsAyar.Cells(wsAyar.Rows.Count, "J").End(xlUp).row

    ' Sayfa1'deki son dolu satırı bul
    sonSatirKaynak = wsKaynak.Cells(wsKaynak.Rows.Count, "B").End(xlUp).row

    ' Ayar sayfasındaki benzersiz değerleri döngüyle kontrol et
    For i = 4 To sonSatirAyar
        deger = wsAyar.Cells(i, "J").Value
        eslesmeBulundu = False ' Eşleşme bulunup bulunmadığını kontrol etmek için değişken

        ' Sayfa1'deki değerleri döngüyle kontrol et
        For j = 1 To sonSatirKaynak
            If wsKaynak.Cells(j, "B").Value = deger Then
                ' K sütunundaki değer boş değilse Sayfa1'de B sütununa yaz
                If wsAyar.Cells(i, "K").Value <> "" Then
                    wsKaynak.Cells(j, "B").Value = wsAyar.Cells(i, "K").Value
                    eslesmeBulundu = True ' Eşleşme bulundu
                End If
            End If
        Next j

        ' Eşleşme bulunduysa K sütunundaki değeri sil
        If eslesmeBulundu Then
            wsAyar.Cells(i, "K").ClearContents
        End If
    Next i

    ' Prosedür tamamlandı
    DegerleriEsleVeYazCalisiyor = False

    MsgBox "Değerler eşleştirildi ve Sayfa1'e yazıldı.", vbInformation

End Sub

' otomatik olarak belirlenen ifadeler verilen kodlara dönüştürülecek
Sub DegerleriEsleVeYazNM()

    Dim wsKaynak As Worksheet
    Dim wsAyar As Worksheet
    Dim sonSatirAyar As Long
    Dim sonSatirKaynak As Long
    Dim i As Long
    Dim j As Long
    Dim deger As Variant

    ' Çalışma sayfalarını tanımla
    Set wsKaynak = ThisWorkbook.Sheets("Sayfa1")
    Set wsAyar = ThisWorkbook.Sheets("ayar")
    
    
    ' Ayar sayfasındaki son dolu satırı bul
    sonSatirAyar = wsAyar.Cells(wsAyar.Rows.Count, "N").End(xlUp).row

    ' Sayfa1'deki son dolu satırı bul
    sonSatirKaynak = wsKaynak.Cells(wsKaynak.Rows.Count, "B").End(xlUp).row

    ' Ayar sayfasındaki benzersiz değerleri döngüyle kontrol et
    For i = 4 To sonSatirAyar
        deger = wsAyar.Cells(i, "N").Value
        
        ' Sayfa1'deki değerleri döngüyle kontrol et
        For j = 1 To sonSatirKaynak
            
            If wsKaynak.Cells(j, "B").Value = deger Then
        
                ' O sütunundaki değer boş değilse Sayfa1'de B sütununa yaz
                If wsAyar.Cells(i, "O").Value <> "" Then
                    wsKaynak.Cells(j, "B").Value = wsAyar.Cells(i, "O").Value
                End If
            End If
        Next j
    Next i

    'MsgBox "Değerler eşleştirildi ve Sayfa1'e yazıldı.", vbInformation

End Sub
Sub otomatikDegerDegistir()

  Dim wsAyar As Worksheet
  Dim wsSayfa1 As Worksheet
  Dim sonSatirAyar As Long
  Dim sonSatirSayfa1 As Long
  Dim i As Long
  Dim j As Long
  Dim aranan As String
  Dim arananDegerMiktar As String
  Dim ilkBosSatir As Long
  Dim k As Long
  Dim bulundu As Boolean ' Yeni bir bayrak değişkeni

  ' Çalışma sayfalarını tanımla
  Set wsAyar = ThisWorkbook.Sheets("Ayar") ' Ayar sayfasının adını girin
  Set wsSayfa1 = ThisWorkbook.Sheets("Sayfa1") ' Sayfa1 sayfasının adını girin

  ' Ayar ve Sayfa1 sayfalarındaki son satırları bul
  sonSatirAyar = wsAyar.Cells(wsAyar.Rows.Count, "R").End(xlUp).row
  sonSatirSayfa1 = wsSayfa1.Cells(wsSayfa1.Rows.Count, "B").End(xlUp).row

  ' Sayfa1 sayfasının B sütunundaki dolu hücreler için döngü
  i = 1
  Do While i <= sonSatirSayfa1 ' Başlık satırı da dahil 1. satırdan başla
    If wsSayfa1.Cells(i, "B").Value <> "" Then ' Hücre doluysa

      aranan = wsSayfa1.Cells(i, "B").Value ' Aranan değeri al
      arananDegerMiktar = wsSayfa1.Cells(i, "E").Value ' Aranan değeri al

      ' Ayar sayfasının R sütununda arama yap
      For j = 1 To sonSatirAyar ' Başlık satırı da dahil 1. satırdan başla

        If wsAyar.Cells(j, "R").Value = aranan Then ' Aranan değer bulunduysa

          ilkBosSatir = j ' Başlangıç satırının bir altından aramaya başla
          bulundu = False ' Bayrağı sıfırla

          ' Aşağı yöndeki ilk boş satırı bul
          Do While wsAyar.Cells(ilkBosSatir, "S").Value <> ""
            ilkBosSatir = ilkBosSatir + 1 ' Bir sonraki satıra geç
            bulundu = True ' Bayrağı ayarla
          Loop

          ' Ayar sayfasının S sütunundan ilk değeri Sayfa1'e yaz ve Miktar değerini işle
          If bulundu Then ' Bayrak kontrolü
            If arananDegerMiktar <> "" Then ' arananDegerMiktar boş değilse
              Call AyarSutunDegerleriniIsle(wsAyar, wsSayfa1, i, j, arananDegerMiktar, ilkBosSatir)
            End If
          End If

        End If

      Next j

    End If
    i = i + 1 ' Bir sonraki satıra geç
  Loop

End Sub

Function AyarSutunDegerleriniIsle(wsAyar As Worksheet, wsSayfa1 As Worksheet, i As Long, j As Long, arananDegerMiktar As String, ilkBosSatir As Long)

  Dim k As Long

  wsSayfa1.Cells(i, "B").Value = wsAyar.Cells(j, "S").Value ' Sayfa1'in B sütununa Ayar'ın S sütunundaki değeri yaz

  'ayar sayfa t de miktar var ise işle
  If wsAyar.Cells(j, "T").Value <> "" Then
    'Ayar sayfa T boş değilse Sayfa1 * Ayar sayfası T sütunu
    wsSayfa1.Cells(i, "E").Value = arananDegerMiktar * wsAyar.Cells(j, "T").Value
  Else
    'Ayar sayfa T boş ise Sayfa1 deki adeti gir
    wsSayfa1.Cells(i, "E").Value = arananDegerMiktar ' İlgili kalemdeki Miktar değerini işle
  End If

  ' j ve ilkBosSatir arasındaki değerleri Sayfa1'e ekle
  For k = j + 1 To ilkBosSatir - 1 ' j+1'den ilkBosSatir-1'e kadar olan satırlar için döngü
    wsSayfa1.Rows(i + 1).Insert ' Yeni satır ekle
    wsSayfa1.Cells(i + 1, "B").Value = wsAyar.Cells(k, "S").Value ' Sayfa1'in B sütununa Ayar'ın S sütunundaki değeri yaz

    'ayar sayfa t de miktar var ise işle
    If wsAyar.Cells(k, "T").Value <> "" Then
      'Ayar sayfa T boş değilse Sayfa1 * Ayar sayfası T sütunu
      wsSayfa1.Cells(i + 1, "E").Value = arananDegerMiktar * wsAyar.Cells(k, "T").Value
    Else
      'Ayar sayfa T boş ise Sayfa1 deki adeti gir
      wsSayfa1.Cells(i + 1, "E").Value = arananDegerMiktar ' İlgili kalemdeki Miktar değerini işle
    End If

    i = i + 1 ' Bir sonraki satıra geç
    sonSatirSayfa1 = wsSayfa1.Cells(wsSayfa1.Rows.Count, "B").End(xlUp).row ' sonSatirSayfa1'i güncelle
  Next k

End Function