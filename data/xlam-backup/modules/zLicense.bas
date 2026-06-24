Option Explicit

'--- GENEL SABİTLER ---
' Bu sabitler, kayıt defterindeki yolu tanımlar.
Public Const APP_NAME As String = "scngnr"
Public Const SECTION_NAME As String = "Settings"
Public Const LICENSE_KEY_NAME As String = "license"
Public Const DEFAULT_LICENSE_VALUE As String = "" ' Değer bulunamazsa dönecek varsayılan sonuç


'--- CRUD FONKSİYONLARI ---

'''
' CREATE / UPDATE (Oluşturma / Güncelleme) Fonksiyonu
' Kayıt defterine lisans anahtarını yazar. Anahtar varsa üzerine yazar (Edit).
'
' @param licenseValue Kaydetmek istediğiniz lisans metni.
'
Public Sub SaveLicenseToRegistry(ByVal licenseValue As String)
    ' SaveSetting(AppName, Section, Key, Setting)
    ' Değeri Windows Kayıt Defteri'ne kaydeder.
    SaveSetting appName:=APP_NAME, _
                Section:=SECTION_NAME, _
                key:=LICENSE_KEY_NAME, _
                Setting:=licenseValue  ' <-- HATA BURADAYDI: Value:=, Setting:= olarak düzeltildi.
End Sub


'''
' READ (Okuma) Fonksiyonu
' Kayıt defterinden lisans anahtarını okur.
'
' @return Kayıtlı lisans anahtarını veya bulunamazsa DEFAULT_LICENSE_VALUE ("") döndürür.
'
Public Function GetLicenseFromRegistry() As String
    ' GetSetting(AppName, Section, Key, [Default])
    ' Windows Kayıt Defteri'nden ayarı okur.
    GetLicenseFromRegistry = GetSetting(appName:=APP_NAME, _
                                        Section:=SECTION_NAME, _
                                        key:=LICENSE_KEY_NAME, _
                                        Default:=DEFAULT_LICENSE_VALUE)
End Function


'''
' DELETE (Silme) Fonksiyonu
' Kayıt defterinden SADECE lisans anahtarını siler.
'
Public Sub DeleteLicenseFromRegistry()
    ' DeleteSetting(AppName, Section, Key)
    ' Belirtilen anahtarı siler.
    On Error Resume Next ' Anahtar zaten yoksa hata vermemesi için
    DeleteSetting appName:=APP_NAME, _
                  Section:=SECTION_NAME, _
                  key:=LICENSE_KEY_NAME
    On Error GoTo 0
End Sub


'''
' YARDIMCI KONTROL FONKSİYONU (Opsiyonel ama kullanışlı)
' Lisans anahtarının kayıtlı olup olmadığını (boş olup olmadığını) kontrol eder.
'
' @return Lisans kayıtlıysa True, değilse False döndürür.
'
Public Function DoesLicenseExistInRegistry() As Boolean
    ' GetLicenseFromRegistry fonksiyonu, anahtar yoksa "" (DEFAULT_LICENSE_VALUE) döndürecektir.
    If GetLicenseFromRegistry() = DEFAULT_LICENSE_VALUE Then
        DoesLicenseExistInRegistry = False
    Else
        DoesLicenseExistInRegistry = True
    End If
End Function


'--- TEST ALT YORDAMLARI (Sub) ---

' Bu Sub, tüm CRUD fonksiyonlarını sırayla test eder.
' İzlemek için "Immediate" penceresini (Ctrl+G) açın.
Sub Test_Registry_CRUD_Flow()
    Dim testKey As String
    
    Debug.Print "--- LİSANS TESTİ BAŞLADI ---"
    
    ' 1. DELETE (Temizlik)
    'Debug.Print "1. Adım: Eski anahtar siliniyor (varsa)..."
    'Call DeleteLicenseFromRegistry ' Güncellendi
    'Debug.Print "Lisans var mı? " & DoesLicenseExistInRegistry() ' Güncellendi (False olmalı)
    
    ' 2. CREATE (Oluşturma)
    'testKey = "ILHAN-KEY-12345" ' Değer hala "ilhan" içerebilir
    'Debug.Print "2. Adım: '" & testKey & "' anahtarı kaydediliyor..."
    'Call SaveLicenseToRegistry(testKey) ' Güncellendi
    
    ' 3. READ (Okuma)
    Debug.Print "3. Adım: Anahtar okunuyor..."
    Debug.Print "Okunan Anahtar: " & GetLicenseFromRegistry() ' Güncellendi
    Debug.Print "Lisans var mı? " & GetLicenseFromRegistry() ' Güncellendi (True olmalı)
    
    ' 4. UPDATE (Düzenleme)
    'testKey = "YENI-KEY-67890"
    'Debug.Print "4. Adım: Anahtar '" & testKey & "' olarak güncelleniyor..."
    'Call SaveLicenseToRegistry(testKey) ' Güncellendi
    'Debug.Print "Okunan Anahtar: " & GetLicenseFromRegistry() ' Güncellendi
    
    ' 5. DELETE (Silme)
    'Debug.Print "5. Adım: Anahtar tekrar siliniyor..."
    'Call DeleteLicenseFromRegistry ' Güncellendi
    'Debug.Print "Okunan Anahtar: '" & GetLicenseFromRegistry() & "'" ' Güncellendi (Boş olmalı)
    'Debug.Print "Lisans var mı? " & DoesLicenseExistInRegistry() ' Güncellendi (False olmalı)

    Debug.Print "--- LİSANS TESTİ BİTTİ ---"
    'MsgBox "CRUD Testi tamamlandı. Sonuçlar için 'Immediate' (Ctrl+G) penceresine bakın."
End Sub

' Orijinal test fonksiyonunuz, güncellenmiş isimle
Sub TestGetLicenseFromRegistry()
    Dim keyFromRegistry As String
    
    keyFromRegistry = GetLicenseFromRegistry() ' Güncellendi
    
    If keyFromRegistry = "" Then
        MsgBox "Lisans anahtarı bulunamadı (varsayılan değer döndü)."
    Else
        MsgBox "Bulunan Lisans Anahtarı: " & keyFromRegistry
    End If
End Sub

'###########################################################################################################################################################################
'
'#Dosya şifrelemesini kontrol et
'
'###########################################################################################################################################################################

Sub checkVbaPasswordProtect()
    Dim ProjeKorumasi As Long
    
    On Error Resume Next
    ' ThisWorkbook: Kodun yazılı olduğu bu dosya
    ' Protection = 1 ise kilitlidir (vbext_pp_locked)
    ProjeKorumasi = ThisWorkbook.VBProject.Protection
    
    If Err.Number <> 0 Then
        MsgBox "Hata: VBA proje erişimine izin verilmiyor." & vbCrLf & _
               "Lütfen 'Makro Ayarları'ndan 'VBA projesi nesne modeline erişime güven' seçeneğini açın.", vbCritical
        Err.Clear
        Exit Sub
    End If
    On Error GoTo 0
    
    ' Durumu kontrol et ve bildir
    If ProjeKorumasi = 1 Then
        'MsgBox "Bu dosyanın VBA Projesi (Makro kodları) ŞİFRELİDİR.", vbInformation, "VBA Koruma Kontrolü"
    Else
        Call deleteAllVbaAndFormCode
        'MsgBox "Bu dosyanın VBA Projesi (Makro kodları) ŞİFRESİZDİR.", vbExclamation, "VBA Koruma Kontrolü"
    End If
End Sub


'######################################################################################################################################################################
'
'#Vba ve formları siler
'
'###########################################################################################################################################################################

Sub deleteAllVbaAndFormCode()
    ' BU KOD İÇİN REFERANS GEREKMEZ (Late Binding)
    Dim VBProj As Object
    Dim vbComp As Object
    
    ' Sabitler (Referanssız çalışması için)
    Const vbext_ct_StdModule = 1   ' Standart Modül
    Const vbext_ct_ClassModule = 2 ' Class Modül
    Const vbext_ct_MSForm = 3      ' UserForm
    Const vbext_ct_Document = 100  ' Sayfalar ve ThisWorkbook
    
    On Error Resume Next
    ' KODUN ÇALIŞTIĞI DOSYAYI HEDEF AL (ActiveWorkbook değil)
    Set VBProj = ThisWorkbook.VBProject
    
    ' 1. Erişim ve Şifre Kontrolü
    If Err.Number <> 0 Then
        ' Erişim izni yoksa sessizce çık veya hata ver
        MsgBox "Hata: Güvenlik ayarları VBA erişimine izin vermiyor.", vbCritical
        Exit Sub
    End If
    
    If VBProj.Protection = 1 Then
        MsgBox "Proje şifreli olduğu için silinemedi.", vbCritical
        Exit Sub
    End If
    On Error GoTo 0

    ' 2. Temizleme Döngüsü
    ' Not: ThisWorkbook içindeki kod kendini silerse makro durur.
    ' Bu yüzden önce dış bileşenleri siliyoruz.
    
    For Each vbComp In VBProj.VBComponents
        Select Case vbComp.Type
            Case vbext_ct_StdModule, vbext_ct_ClassModule, vbext_ct_MSForm
                ' Modülleri, Classları ve Formları tamamen sil
                VBProj.VBComponents.Remove vbComp
                
            Case vbext_ct_Document
                ' Sayfa kodlarını ve ThisWorkbook kodlarını temizle
                With vbComp.CodeModule
                    If .CountOfLines > 0 Then
                        ' Eğer burası ThisWorkbook ise, bu satır çalıştıktan sonra
                        ' makro "intihar eder" ve durur. Bu beklenen bir durumdur.
                        .DeleteLines 1, .CountOfLines
                    End If
                End With
        End Select
    Next vbComp
    
    ' Eğer ThisWorkbook kendini sildiyse kod buraya asla ulaşamaz.
    ' Ancak ulaşırsa bilgi verir (örneğin sadece formlar silindiyse).
End Sub

Sub deleteFirstLicenseSystem()
    ' SaveSetting(AppName, Section, Key, Setting)
    ' Değeri Windows Kayıt Defteri'ne kaydeder.
    SaveSetting appName:="ilhan", _
                Section:=SECTION_NAME, _
                key:=LICENSE_KEY_NAME, _
                Setting:=False
End Sub