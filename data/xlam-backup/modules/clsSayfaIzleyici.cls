'--- Class Module: clsSayfaIzleyici ---
Option Explicit

' Olaylarını yakalamak istediğimiz Worksheet nesnesi
Public WithEvents hedefSayfa As Worksheet
Attribute hedefSayfa.VB_VarHelpID = -1

' Bu Class nesnesi oluşturulduğunda hedef sayfayı ayarlamak için
Public Sub Init(ByVal ws As Worksheet)
    Set hedefSayfa = ws
End Sub

' Hedef sayfada değişiklik olduğunda tetiklenecek olay
Private Sub HedefSayfa_Change(ByVal Target As Range)

    ' Hedef sayfa hala geçerli mi kontrol et (kapatılmış olabilir)
    If hedefSayfa Is Nothing Then Exit Sub

    Dim KeyCells As Range

    ' --- Kontroller ---
    ' 1. Sayfa3 var mı? (Aynı çalışma kitabında)
    If Not SheetExists("Sayfa3", hedefSayfa.Parent) Then Exit Sub ' HedefSayfa.Parent: Sayfanın ait olduğu Workbook

    ' 2. Değişiklik B sütununda mı ve tek bir hücre mi?
    Set KeyCells = hedefSayfa.Columns("B")
    If Intersect(Target, KeyCells) Is Nothing Or Target.Cells.CountLarge > 1 Then Exit Sub

    ' 3. B sütununa girilen değer boş mu?
    If IsEmpty(Target.Value) Then Exit Sub

    ' 4. Aynı satırda A sütunu boş mu?
    If Not IsEmpty(hedefSayfa.Cells(Target.row, "A").Value) Then Exit Sub
    ' --- Kontroller Bitti ---

    ' --- Tetikleme ---
    On Error GoTo Hata_Change

    Debug.Print "clsSayfaIzleyici: HedefSayfa_Change tetiklendi: " & Target.Address & " değeri: " & Target.Value

    ' Önemli: Olayları doğrudan burada kapatıp açmak yerine,
    ' çağrılan makronun içinde yönetmek daha güvenli olabilir.
    ' Ancak şimdilik burada bırakalım.
    Application.EnableEvents = False

    ' İlgili makroyu çağır (Standart modülde olmalı)
    ' Makro artık hangi sayfada çalışacağını bilmeli. Aktif sayfayı kullanabilir.
    Call MlzVeriEkleMacrosu3 ' Bu makro ActiveSheet üzerinde çalışacak şekilde ayarlanmıştı

    Debug.Print "clsSayfaIzleyici: MlzVeriEkleMacrosu3 çağrıldı."

Cikis_Change:
    On Error Resume Next
    Application.EnableEvents = True
    Exit Sub

Hata_Change:
    MsgBox "clsSayfaIzleyici olayında hata oluştu: " & Err.Description, vbCritical
    Resume Cikis_Change

End Sub

' Class nesnesi yok edildiğinde HedefSayfa referansını temizle
Private Sub Class_Terminate()
    Set hedefSayfa = Nothing
    Debug.Print "clsSayfaIzleyici Terminate"
End Sub

' Yardımcı Fonksiyon (Class içinde de olabilir veya Standart Modülde)
Private Function SheetExists(sheetName As String, Optional wb As Workbook) As Boolean
    Dim ws As Worksheet
    If wb Is Nothing Then Set wb = ThisWorkbook
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    SheetExists = (Err.Number = 0)
    On Error GoTo 0
End Function