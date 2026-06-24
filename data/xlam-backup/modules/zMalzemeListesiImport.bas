'================================================================================
' Malzeme listesi Excel dosyalarını Perfex'e aktarır (MrpApi_ProductCreateBulk).
' Sunucu tek istekte en fazla ~150 ürün kabul eder; WAF 429 için küçük paket + dosya/sayfa arası bekleme kullanılır.
'
' Beklenen sütunlar (1. satır başlık, veri 2. satırdan):
'   B = stok kodu (commodity_code)
'   C = ürün adı (description)
'   D = marka -> ürün grubu (group_name, yoksa otomatik oluşturulur)
'   F = fiyat (rate)
'
' Birleştirilmiş hücreler: Excel ekranda metni tüm satırlarda gösterir; VBA .Value alt
' satırlarda boş döner. CellValueMerged ile birleşik alanın sol üst hücresi okunur.
' Sabit vergi: tax_rate 20 (%20 KDV - sistemde bu orana sahip vergi kaydı olmalı)
'
' Klasör: Alt klasörlerdeki Excel dosyaları taranır.
' "4" adlı alt klasördeki tüm ürünler API'ye "manufacturable": true ile gider (üretilebilir).
'
' Kurulum:
'   1. MrpApi modülünü aynı VBA projesine ekleyin.
'   2. Aşağıdaki MALZEME_ROOT ve API sabitlerini doldurun.
'   3. Makrolar: MalzemeImport_Run (tüm klasör) veya MalzemeImport_ActiveSheet (açık sayfa)
'================================================================================

Option Explicit

#If VBA7 Then
Private Declare PtrSafe Sub MalzemeSleepMs Lib "kernel32" Alias "Sleep" (ByVal dwMilliseconds As Long)
#Else
Private Declare Sub MalzemeSleepMs Lib "kernel32" Alias "Sleep" (ByVal dwMilliseconds As Long)
#End If

' --- Kendi ortamınıza göre düzenleyin ---
Private Const MALZEME_ROOT As String = "C:\Belgelerim\Cemex\Malzeme Listeleri"
Private Const API_BASE_URL As String = "https://mrp.cangungor.tr"
Private Const API_JWT As String = "eyJ0eXAiOiJKV1QiLCJhbGciOiJIUzI1NiJ9.eyJ1c2VyIjoibXJwIiwibmFtZSI6InNlcmNhbiIsIkFQSV9USU1FIjoxNzc1Mjk0MTg0fQ._80qqKXalOBUCQgIfD2eoiDQOg59d3VqgRAwk2PVnjM"

Private Const DATA_FIRST_ROW As Long = 2
Private Const COL_STOK As Long = 2      ' B
Private Const COL_AD As Long = 3       ' C
Private Const COL_MARKA As Long = 4    ' D
Private Const COL_FIYAT As Long = 6    ' F
Private Const TAX_RATE As Long = 20
' Malzeme Listeleri\4\ — tamamı üretilebilir ürün (API: manufacturable)
Private Const URETILEBILIR_KLASOR_ADI As String = "4"
' WAF (429 Scanner activity): küçük paket + uzun ara. Sunucu max 150; 25–50 güvenli aralık.
Private Const MALZEME_BATCH_SIZE As Long = 35
' Başarılı toplu istek sonrası bekleme (ms)
Private Const MALZEME_BATCH_PAUSE_MS As Long = 4500
' Her Excel dosyası bittikten sonra (bir sonraki listeye geçmeden önce, ms)
Private Const MALZEME_FILE_PAUSE_MS As Long = 15000
' Aynı dosyada çalışma sayfaları arası (ms); 0 = kapalı
Private Const MALZEME_SHEET_PAUSE_MS As Long = 3000
' 429 alındığında en fazla kaç kez yeniden dene (toplam gönderim sayısı)
Private Const MALZEME_429_MAX_TRIES As Long = 8

'Tüm alt klasörlerdeki Excel dosyalarını işler
Public Sub MalzemeImport_Run()
    Dim fso As Object
    Dim root As Object
    Dim sf As Object
    Dim ok As Long
    Dim Fail As Long
    Dim skip As Long

    MrpApi_Configure API_BASE_URL, API_JWT

    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(MALZEME_ROOT) Then
        MsgBox "Klasör yok: " & MALZEME_ROOT, vbExclamation
        Exit Sub
    End If

    Set root = fso.GetFolder(MALZEME_ROOT)
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    ok = 0: Fail = 0: skip = 0

    ProcessFolderFiles fso, root, ok, Fail, skip, False

    For Each sf In root.SubFolders
        ProcessFolderFiles fso, sf, ok, Fail, skip, IsUretilebilirKlasor(sf.Name)
    Next sf

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    MsgBox "Bitti." & vbCrLf & "OK: " & ok & vbCrLf & "Hata: " & Fail & vbCrLf & "Atlanan satır: " & skip, vbInformation
End Sub

'Açık çalışma kitabının aktif sayfasını işler (tek liste test için)
Public Sub MalzemeImport_ActiveSheet()
    Dim ok As Long
    Dim Fail As Long
    Dim skip As Long
    MrpApi_Configure API_BASE_URL, API_JWT
    ok = 0: Fail = 0: skip = 0
    Application.ScreenUpdating = False
    On Error GoTo Clean
    ImportSheet ActiveWorkbook.ActiveSheet, True, ok, Fail, skip, False
Clean:
    Application.ScreenUpdating = True
End Sub

Private Function IsUretilebilirKlasor(ByVal klasorAdi As String) As Boolean
    IsUretilebilirKlasor = (StrComp(Trim$(LCase$(klasorAdi)), LCase$(URETILEBILIR_KLASOR_ADI), vbTextCompare) = 0)
End Function

Private Sub ProcessFolderFiles(ByVal fso As Object, ByVal fol As Object, ByRef ok As Long, ByRef Fail As Long, ByRef skip As Long, ByVal uretilebilir As Boolean)
    Dim f As Object
    Dim ext As String
    Dim wb As Workbook

    For Each f In fol.Files
        ext = LCase$(fso.GetExtensionName(f.Name))
        If ext = "xlsb" Or ext = "xlsm" Then
            On Error Resume Next
            Set wb = Workbooks.Open(f.path, ReadOnly:=True, UpdateLinks:=False)
            On Error GoTo 0
            If Not wb Is Nothing Then
                ImportWorkbook wb, ok, Fail, skip, uretilebilir
                wb.Close SaveChanges:=False
                Set wb = Nothing
                If MALZEME_FILE_PAUSE_MS > 0 Then MalzemeSleepMs MALZEME_FILE_PAUSE_MS
            End If
        End If
    Next f
End Sub

Private Sub ImportWorkbook(ByVal wb As Workbook, ByRef ok As Long, ByRef Fail As Long, ByRef skip As Long, ByVal uretilebilir As Boolean)
    Dim ws As Worksheet
    Dim first As Boolean
    first = True
    For Each ws In wb.Worksheets
        If Not first And MALZEME_SHEET_PAUSE_MS > 0 Then MalzemeSleepMs MALZEME_SHEET_PAUSE_MS
        first = False
        ImportSheet ws, False, ok, Fail, skip, uretilebilir
    Next ws
End Sub

Private Sub ImportSheet(ByVal ws As Worksheet, ByVal silent As Boolean, ByRef ok As Long, ByRef Fail As Long, ByRef skip As Long, ByVal uretilebilir As Boolean)
    Dim r As Long
    Dim lastRow As Long
    Dim code As String
    Dim ad As String
    Dim marka As String
    Dim js As String
    Dim batchJson As String
    Dim batchCount As Long

    lastRow = ws.Cells(ws.Rows.Count, COL_STOK).End(xlUp).row
    If lastRow < DATA_FIRST_ROW Then Exit Sub

    batchJson = ""
    batchCount = 0

    For r = DATA_FIRST_ROW To lastRow
        code = CellValueMerged(ws, r, COL_STOK)
        ad = CellValueMerged(ws, r, COL_AD)
        marka = CellValueMerged(ws, r, COL_MARKA)

        If Len(code) = 0 And Len(ad) = 0 Then GoTo NextRow

        If Len(ad) = 0 Then
            skip = skip + 1
            Debug.Print "Satir "; r; " atlandi (aciklama bos): "; code
            GoTo NextRow
        End If

        If Len(code) = 0 Then
            skip = skip + 1
            Debug.Print "Satir "; r; " atlandi (stok bos): "; ad
            GoTo NextRow
        End If

        js = BuildProductJson(code, ad, marka, CellValueMergedRaw(ws, r, COL_FIYAT), uretilebilir)
        If Len(batchJson) > 0 Then batchJson = batchJson & ","
        batchJson = batchJson & js
        batchCount = batchCount + 1

        If batchCount >= MALZEME_BATCH_SIZE Then
            FlushProductBatch batchJson, batchCount, ok, Fail, skip
            batchJson = ""
            batchCount = 0
        End If

        DoEvents
NextRow:
    Next r

    If batchCount > 0 Then
        FlushProductBatch batchJson, batchCount, ok, Fail, skip
    End If

    If silent Then
        MsgBox "Sayfa: " & ws.Name & vbCrLf & "OK: " & ok & " Hata: " & Fail & " Atlanan: " & skip, vbInformation
    End If
End Sub

Private Sub FlushProductBatch(ByVal batchObjects As String, ByVal rowCount As Long, ByRef ok As Long, ByRef Fail As Long, ByRef skip As Long)
    Dim payload As String
    Dim resp As String
    Dim httpCode As Long
    Dim attempt As Long
    Dim waitMs As Long

    If rowCount <= 0 Or Len(batchObjects) = 0 Then Exit Sub

    payload = "{""products"":[" & batchObjects & "]}"

    For attempt = 1 To MALZEME_429_MAX_TRIES
        resp = MrpApi_ProductCreateBulk(payload)
        httpCode = MrpApi_LastHttpStatus()
        If httpCode <> 429 And httpCode <> 503 Then Exit For
        If attempt >= MALZEME_429_MAX_TRIES Then Exit For
        ' WAF: gittikçe uzayan bekleme (ms), tavan 120 sn
        waitMs = 15000 + CLng(attempt) * 12000
        If waitMs > 120000 Then waitMs = 120000
        MalzemeSleepMs waitMs
    Next attempt

    If httpCode >= 200 And httpCode < 300 And JsonResponseIsOk(resp) Then
        ok = ok + BulkJsonExtractLong(resp, "created")
        skip = skip + BulkJsonExtractLong(resp, "skipped_duplicate")
        Fail = Fail + BulkJsonExtractLong(resp, "failed")
    Else
        Fail = Fail + rowCount
        Debug.Print "Toplu gonderim HTTP "; httpCode; " "; Left$(resp, 800)
    End If

    If MALZEME_BATCH_PAUSE_MS > 0 Then MalzemeSleepMs MALZEME_BATCH_PAUSE_MS
End Sub

Private Function BulkJsonExtractLong(ByVal resp As String, ByVal key As String) As Long
    Dim needle As String
    Dim p As Long
    Dim i As Long
    Dim c As String
    needle = """" & key & """:"
    p = InStr(1, resp, needle, vbTextCompare)
    If p = 0 Then BulkJsonExtractLong = 0: Exit Function
    p = p + Len(needle)
    Do While p <= Len(resp)
        c = Mid$(resp, p, 1)
        If c <> " " And c <> vbTab Then Exit Do
        p = p + 1
    Loop
    i = p
    Do While i <= Len(resp)
        c = Mid$(resp, i, 1)
        If c < "0" Or c > "9" Then Exit Do
        i = i + 1
    Loop
    If i > p Then
        On Error Resume Next
        BulkJsonExtractLong = CLng(Mid$(resp, p, i - p))
        On Error GoTo 0
    Else
        BulkJsonExtractLong = 0
    End If
End Function

' Birleştirilmiş hücrelerde görünen metin için MergeArea sol üst; .Value2 formül/önbellek tutarlılığı.
Private Function CellValueMerged(ByVal ws As Worksheet, ByVal row As Long, ByVal col As Long) As String
    Dim rng As Range
    Dim v As Variant
    On Error Resume Next
    Set rng = ws.Cells(row, col)
    If rng Is Nothing Then CellValueMerged = "": On Error GoTo 0: Exit Function
    If rng.MergeCells Then
        v = rng.MergeArea.Cells(1, 1).Value2
    Else
        v = rng.Value2
    End If
    On Error GoTo 0
    If IsError(v) Or IsEmpty(v) Or IsNull(v) Then
        CellValueMerged = ""
    Else
        CellValueMerged = Trim$(CStr(v))
    End If
End Function

Private Function CellValueMergedRaw(ByVal ws As Worksheet, ByVal row As Long, ByVal col As Long) As Variant
    Dim rng As Range
    On Error Resume Next
    Set rng = ws.Cells(row, col)
    If rng Is Nothing Then CellValueMergedRaw = Empty: On Error GoTo 0: Exit Function
    If rng.MergeCells Then
        CellValueMergedRaw = rng.MergeArea.Cells(1, 1).Value2
    Else
        CellValueMergedRaw = rng.Value2
    End If
    On Error GoTo 0
End Function

Private Function BuildProductJson(ByVal stokKodu As String, ByVal urunAdi As String, ByVal marka As String, ByVal fiyatCell As Variant, ByVal uretilebilir As Boolean) As String
    Dim rateStr As String
    Dim tail As String
    rateStr = JsonNumberFromCell(fiyatCell)

    tail = ""
    If uretilebilir Then
        tail = ",""manufacturable"":true"
    End If

    BuildProductJson = "{" & _
        """commodity_code"":""" & JsonEscape(stokKodu) & """," & _
        """description"":""" & JsonEscape(urunAdi) & """," & _
        """rate"":" & rateStr & "," & _
        """group_name"":""" & JsonEscape(marka) & """," & _
        """tax_rate"":" & CStr(TAX_RATE) & tail & _
        "}"
End Function

Private Function JsonResponseIsOk(ByVal resp As String) As Boolean
    Dim t As String
    t = Replace(resp, " ", "")
    JsonResponseIsOk = (InStr(1, t, """status"":true", vbTextCompare) > 0)
End Function

Private Function JsonEscape(ByVal s As String) As String
    Dim t As String
    t = s
    t = Replace(t, "\", "\\")
    t = Replace(t, """", "\""")
    t = Replace(t, vbCrLf, "\n")
    t = Replace(t, vbLf, "\n")
    t = Replace(t, vbCr, "\n")
    t = Replace(t, Chr$(9), "\t")
    JsonEscape = t
End Function

'Ondalık ayırıcı nokta (JSON)
Private Function JsonNumberFromCell(ByVal v As Variant) As String
    Dim d As Double
    If IsEmpty(v) Or IsNull(v) Then
        JsonNumberFromCell = "0"
        Exit Function
    End If
    On Error Resume Next
    If IsNumeric(v) Then
        d = CDbl(v)
    Else
        d = val(Replace(Replace(Trim$(CStr(v)), ".", ""), ",", "."))
    End If
    On Error GoTo 0
    JsonNumberFromCell = Replace(Format$(d, "0.########"), ",", ".")
End Function