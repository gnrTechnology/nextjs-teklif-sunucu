'----------------------------------------------------------------------------------------------
' Teklif Durumu ALındı olarak işaretlenince
'   1.Sipariş arşive teknik dosyalar - sipariş pdf'i
'   2.Uretim ve imalata ise çizimi gönder
'
'----------------------------------------------------------------------------------------------

Sub SiparisOtomasyonu()
    Dim Dosyalar As Object
    Set Dosyalar = CreateObject("Scripting.Dictionary")
    
    ' 1. Bilgileri Topla
    Dim SevkTarihi As String: SevkTarihi = InputBox("Sevk Tarihi:", "Bilgi Girişi", "10.11.2025")
    Dim Odeme As String: Odeme = InputBox("Ödeme Koşulu:", "Bilgi Girişi", "60 gün vadeli")
    Dim NakliyeSecim As String: NakliyeSecim = InputBox("Nakliye Durumu:" & vbCrLf & "0 - Tarafımıza ait" & vbCrLf & "1 - Müşteriye ait", "Nakliye Bilgisi", "0")
    
    If SevkTarihi = "" Or Odeme = "" Or NakliyeSecim = "" Then Exit Sub
    
    Dim NakliyeMetni As String: NakliyeMetni = IIf(NakliyeSecim = "0", "Nakliye tarafımıza aittir.", "Nakliye müşteriye aittir.")

    ' 2. Dosya Seçimleri
    ' Sipariş PDF Seçimi (ZORUNLU - Sadece Arşiv Mailine gidecek)
    Dosyalar("SiparisPDF") = DosyaSec("ZORUNLU: Sipariş PDF Dosyasını Seçin", "*.pdf")
    
    If Dosyalar("SiparisPDF") = "" Then
        MsgBox "Sipariş PDF'i seçilmesi zorunludur! İşlem iptal edildi.", vbCritical, "Eksik Dosya"
        Exit Sub
    End If
    
    ' AutoCAD Seçimi (OPSİYONEL - Eklenirse Her İki Maile de gidecek)
    Dosyalar("AutoCAD") = DosyaSec("AutoCAD Çizim Dosyasını Seçin (Opsiyonel)", "*.dwg; *.dxf")
    
    ' 3. BİRİNCİ MAİL: Arşiv Birimi (siparisarsiv@epron.com.tr)
    ' İçerik: Sipariş PDF + (Varsa) AutoCAD
    
    Call MailOlustur( _
        Alıcı:="siparisarsiv@epron.com.tr", _
        CC:="", _
        Konu:=UzantisizIsim(ActiveWorkbook.Name) & " - TEKLİF HK.", _
        Metin:="Konu sipariş kapsamında ;" & vbCrLf & _
               "- Ödeme koşulu " & Odeme & " anlaşılmıştır." & vbCrLf & _
               "- " & NakliyeMetni & vbCrLf & _
               "- İlgili teknik dokümanlar ektedir." & vbCrLf & _
               "- İlgili sipariş evrakı ektedir." & vbCrLf & _
               "- Sevk tarihi : " & SevkTarihi, _
        Ekler:=Array(Dosyalar("SiparisPDF"), Dosyalar("AutoCAD")) _
    )

    ' 4. İKİNCİ MAİL: Üretim ve İmalat (uretim@epron.com.tr)
    ' İçerik: Sadece (Varsa) AutoCAD
    Call MailOlustur( _
        Alıcı:="uretim@epron.com.tr", _
        CC:="imalat@epron.com.tr", _
        Konu:=UzantisizIsim(ActiveWorkbook.Name) & " - TEKLİF HK.", _
        Metin:="Konu sipariş kapsamında ;" & vbCrLf & _
               "- " & NakliyeMetni & vbCrLf & _
               "- İlgili teknik dokümanlar ektedir." & vbCrLf & _
               "- Sevk tarihi : " & SevkTarihi, _
        Ekler:=Array(Dosyalar("AutoCAD")) _
    )
    
    MsgBox "İşlem başarıyla tamamlandı.", vbInformation, "Onay"
End Sub

' --- YARDIMCI FONKSİYONLAR ---

Function DosyaSec(Baslik As String, Filtre As String) As String
    Dim fd As Object
    Set fd = Application.FileDialog(3)
    With fd
        .Title = Baslik
        .Filters.Clear
        .Filters.Add "Dosyalar", Filtre
        If .Show = -1 Then DosyaSec = .SelectedItems(1)
    End With
End Function

Sub MailOlustur(Alıcı As String, CC As String, Konu As String, Metin As String, Ekler As Variant)
    Dim OutApp As Object, OutMail As Object
    Dim i As Integer
    
    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)
    
    With OutMail
        .To = Alıcı
        .CC = CC
        .Subject = Konu
        .HTMLBody = "<html><body style='font-family: Calibri, sans-serif; font-size: 11pt;'>" & _
                    "Merhaba ,<br><br><br>" & _
                    Replace(Metin, vbCrLf, "<br>") & _
                    "<br><br>İyi çalışmalar dileriz.</body></html>"
        
        ' Ekler listesini kontrol et (Sadece seçilen dosyalar eklenir)
        For i = LBound(Ekler) To UBound(Ekler)
            If Ekler(i) <> "" Then .Attachments.Add Ekler(i)
        Next i
        
        .Display
    End With
End Sub

Function UzantisizIsim(Yol As String) As String
    Dim fName As String
    fName = Mid(Yol, InStrRev(Yol, "\") + 1)
    If InStr(fName, ".") > 0 Then
        UzantisizIsim = Left(fName, InStrRev(fName, ".") - 1)
    Else
        UzantisizIsim = fName
    End If
End Function

' Kullanımı:
'