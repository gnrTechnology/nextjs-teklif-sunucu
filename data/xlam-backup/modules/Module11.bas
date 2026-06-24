Option Explicit

' Gerekli VBE (Visual Basic Extensibility) sabitlerini elle tanımlıyoruz.
Private Const vbext_ct_StdModule As Long = 1
Private Const vbext_ct_ClassModule As Long = 2
Private Const vbext_ct_MSForm As Long = 3
Private Const vbext_ct_Document As Long = 100

Sub TumKodSatirlariniSay()
    
    Dim VBProj As Object ' VBProject
    Dim vbComp As Object ' VBComponent
    Dim satirSayisi As Long
    Dim toplamSatir As Long
    
    ' Kategori bazlı sayaçlar
    Dim modulSatir As Long
    Dim formSatir As Long
    Dim sinifSatir As Long
    Dim belgeSatir As Long
    
    Dim raporMetni As String
    
    On Error GoTo HataYonetimi
    
    ' --- DEĞİŞİKLİK BURADA ---
    ' Aktif olan kitabı değil, BU kodun içinde bulunduğu kitabı say.
    Set VBProj = ThisWorkbook.VBProject
    ' -------------------------
    
    ' Projedeki her bir bileşeni (modül, form, sayfa kodu vb.) döngüye al
    For Each vbComp In VBProj.VBComponents
        
        ' Bileşenin kod modülündeki satır sayısını al
        satirSayisi = vbComp.CodeModule.CountOfLines
        
        ' Toplam satıra ekle
        toplamSatir = toplamSatir + satirSayisi
        
        ' Türüne göre kategorize et
        Select Case vbComp.Type
            Case vbext_ct_StdModule
                modulSatir = modulSatir + satirSayisi
            Case vbext_ct_ClassModule
                sinifSatir = sinifSatir + satirSayisi
            Case vbext_ct_MSForm
                formSatir = formSatir + satirSayisi
            Case vbext_ct_Document
                belgeSatir = belgeSatir + satirSayisi
        End Select
        
    Next vbComp
    
                 
    raporMetni = raporMetni & "Genel Toplam Kod Satırı: " & toplamSatir & vbCrLf & _
                 "------------------------------------------------" & vbCrLf
                 
    raporMetni = raporMetni & "Standart Modüller (bas): " & modulSatir & " satır" & vbCrLf
    raporMetni = raporMetni & "Form Modülleri (frm): " & formSatir & " satır" & vbCrLf
    raporMetni = raporMetni & "Sınıf Modülleri (cls): " & sinifSatir & " satır" & vbCrLf
    raporMetni = raporMetni & "Belge Modülleri (Sayfa/Workbook): " & belgeSatir & " satır"
                 
    MsgBox raporMetni, vbInformation, "Proje Raporu"
    
    ' Nesneleri temizle
    Set vbComp = Nothing
    Set VBProj = Nothing
    Exit Sub

HataYonetimi:
    ' Hata durumunda en olası nedeni kullanıcıya bildir
    If Err.Number = 1004 Then
        MsgBox "Koda erişim hatası!" & vbCrLf & vbCrLf & _
               "Lütfen 'Dosya > Seçenekler > Güven Merkezi > Güven Merkezi Ayarları > Makro Ayarları' yolunu izleyerek," & vbCrLf & _
               "'VBA proje nesne modeline erişime güven' seçeneğini işaretleyin.", _
               vbCritical, "Erişim Reddedildi"
    Else
        MsgBox "Beklenmeyen bir hata oluştu:" & vbCrLf & _
               "Hata No: " & Err.Number & vbCrLf & _
               "Açıklama: " & Err.Description, vbCritical, "Hata"
    End If
    
    Set vbComp = Nothing
    Set VBProj = Nothing
End Sub