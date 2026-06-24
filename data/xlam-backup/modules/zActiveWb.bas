Sub GetActiveWorkbookPath()

    Dim activeWB As Workbook
    Dim wbPath As String

    ' Hata yakalamayı etkinleştir (eğer hiçbir kitap açık değilse veya başka bir sorun varsa)
    On Error Resume Next
    Set activeWB = Application.ActiveWorkbook
    On Error GoTo 0 ' Hata yakalamayı normale döndür

    ' Aktif bir çalışma kitabı olup olmadığını kontrol et
    If activeWB Is Nothing Then
        MsgBox "Şu anda aktif bir Excel çalışma kitabı bulunmuyor.", vbExclamation, "Hata"
        Exit Sub
    End If

    ' Çalışma kitabının kaydedilip kaydedilmediğini kontrol et
    ' Eğer kitap henüz kaydedilmemişse, Path özelliği boş olacaktır.
    If activeWB.path = "" Then
        MsgBox "Aktif çalışma kitabı (" & activeWB.Name & ") henüz kaydedilmemiş." & vbCrLf & _
               "Yolu alabilmek için lütfen önce dosyayı kaydedin.", vbInformation, "Bilgi"
        ' Opsiyonel: Kaydedilmemişse sadece adını alabilirsiniz: wbPath = activeWB.Name
        ' Veya işlemi burada sonlandırabilirsiniz.
        Exit Sub
    Else
        ' Çalışma kitabı kaydedilmişse tam yolunu al
        wbPath = activeWB.fullName
        
        SaveSetting "sercan", "fileOpenWorkBooks", "nowOpenPropsFile", wbPath
        
        ' Sonucu göster (veya bu değişkeni başka bir yerde kullan)
        'MsgBox "Aktif çalışma kitabının tam yolu:" & vbCrLf & wbPath, vbInformation, "Dosya Yolu"
        
        ' Bu 'wbPath' değişkenini kodunuzun geri kalanında kullanabilirsiniz.
        ' Örneğin: Call BaskaBirSubRutine(wbPath)
    End If

End Sub