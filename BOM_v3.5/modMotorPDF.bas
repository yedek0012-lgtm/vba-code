Attribute VB_Name = "modMotorPDF"
Option Explicit
Option Private Module

' =========================================================================
' PDF motoru: Tekla / CAD çýktýsý (METÝN tabanlý) PDF'teki malzeme listesini okur.
' Power Query'nin PDF baðlayýcýsý (Excel 365 / 2021, Windows) sayfayý KOORDÝNATA göre
' tabloya çevirir; tablo geçici bir çalýþma kitabýna dökülür ve normal Excel motoru
' (baþlýk algýlama, sýnýflandýrma, mutabakat) aynen çalýþýr.
' Taranmýþ (resim) PDF'ler okunamaz -> AUDIT'e hata yazýlýr.
' v3.5: Okuma arka planda yapýlýr ve PDF_TIMEOUT_SEC saniyede bitmezse iptal edilir
'       (büyük montaj çizimlerinde Excel'in boþ "Tablo Tasarýmý" penceresinde donmasý).
'       Geçici çalýþma kitabý gizli açýlýr. ESC ile de iptal edilebilir.
' =========================================================================

' Bir PDF için en fazla bekleme süresi (saniye). Aþýlýrsa dosya atlanýr, AUDIT'e yazýlýr.
Private Const PDF_TIMEOUT_SEC As Long = 90

Public Sub ProcessPdfFile()
    Dim wb As Workbook, wsQ As Worksheet, wsD As Worksheet, lo As ListObject, qn As String, v As Variant
    Dim t0 As Single, el As Single, stopReason As String, prevCancel As XlEnableCancelKey
    On Error GoTo Fail
    Application.StatusBar = "PDF okunuyor: " & fileName
    Set wb = Application.Workbooks.Add(xlWBATWorksheet)
    ' Boþ geçici pencere ekranda kalmasýn
    wb.Windows(1).Visible = False
    ThisWorkbook.Activate
    Set wsQ = wb.Worksheets(1)
    qn = "BOM_PDF"
    wb.Queries.Add Name:=qn, Formula:=BuildPdfQuery(CStr(vItem))
    Set lo = wsQ.ListObjects.Add(SourceType:=0, Source:="OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=" & qn & ";Extended Properties=""""", _
                                 Destination:=wsQ.Range("$A$1"))
    With lo.QueryTable
        .CommandType = xlCmdSql
        .CommandText = Array("SELECT * FROM [" & qn & "]")
        .BackgroundQuery = True
        .Refresh BackgroundQuery:=True
    End With

    ' Arka planda okunurken bekle; süre dolarsa veya ESC'ye basýlýrsa iptal et
    prevCancel = Application.EnableCancelKey
    Application.EnableCancelKey = xlErrorHandler
    t0 = Timer
    Do While lo.QueryTable.Refreshing
        el = Timer - t0
        If el < 0 Then el = el + 86400!              ' gece yarýsý
        Application.StatusBar = "PDF okunuyor (" & CLng(el) & " / " & PDF_TIMEOUT_SEC & " sn, iptal: ESC): " & fileName
        If el > PDF_TIMEOUT_SEC Then stopReason = "zaman aþýmý (" & PDF_TIMEOUT_SEC & " sn)": Exit Do
        DoEvents
        Application.Wait Now + TimeSerial(0, 0, 1)
    Loop
    Application.EnableCancelKey = prevCancel
    If stopReason = "" And lo.ListRows.Count = 0 Then stopReason = "tablo bulunamadý veya Power Query hata verdi (taranmýþ/resim PDF olabilir)"
    If stopReason <> "" Then GoTo PdfStopped

    v = lo.Range.Value
    Set wsD = wb.Worksheets.Add(After:=wsQ)
    wsD.Name = "PDF"
    If IsArray(v) Then wsD.Range("A1").Resize(UBound(v, 1), UBound(v, 2)).Value = v
    Application.DisplayAlerts = False
    wsQ.Delete

    If Not IsArray(v) Then
        AuditRecord fileName, "PDF", 0, "FILE", CStr(vItem), "FILE", "ERROR", 0, "", "PDF'te tablo bulunamadý (taranmýþ/resim PDF olabilir)."
    Else
        AuditRecord fileName, "PDF", 0, "FILE", CStr(vItem), "FILE", "EXACT", 100, UBound(v, 1) & " satýr okundu", ""
        Set pdfTempWb = wb
        Call ProcessExcelFile
    End If

Cleanup:
    On Error Resume Next
    Set pdfTempWb = Nothing
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    Application.StatusBar = False
    Exit Sub
PdfStopped:
    On Error Resume Next
    If Not lo Is Nothing Then
        If lo.QueryTable.Refreshing Then lo.QueryTable.CancelRefresh
    End If
    On Error GoTo 0
    AuditRecord fileName, "PDF", 0, "FILE", CStr(vItem), "FILE", "ERROR", 0, "", _
                "PDF okunamadý: " & stopReason & ". PDF_BOM_Okuyucu.html ile Excel'e çevirip o dosyayý seçin."
    GoTo Cleanup
Fail:
    If Err.Number = 18 Then                            ' ESC (kullanýcý iptali)
        stopReason = "kullanýcý iptal etti (ESC)"
        Application.EnableCancelKey = prevCancel
        Resume PdfStopped
    End If
    AuditRecord fileName, "PDF", 0, "FILE", CStr(vItem), "FILE", "ERROR", 0, "", _
                "PDF okunamadý: " & Err.Description & " | Gerekenler: Excel 365/2021 (Power Query PDF baðlayýcýsý) ve metin tabanlý PDF."
    Resume Cleanup
End Sub

' Power Query (M) sorgusu: önce PDF'in algýladýðý TABLOLAR, hiç yoksa SAYFA metni.
' Her tablonun sütunlarý C1..Cn yapýlýr, araya boþ satýr konur ve hepsi alt alta birleþtirilir.
Public Function BuildPdfQuery(ByVal pdfPath As String) As String
    Dim q As String, p As String
    p = Replace(pdfPath, """", """""")
    q = "let" & vbCrLf
    q = q & "    Src = Pdf.Tables(File.Contents(""" & p & """))," & vbCrLf
    q = q & "    Tbl = Table.SelectRows(Src, each [Kind] = ""Table"")," & vbCrLf
    q = q & "    Use = if Table.RowCount(Tbl) > 0 then Tbl else Table.SelectRows(Src, each [Kind] = ""Page"")," & vbCrLf
    q = q & "    Norm = (d as table) as table =>" & vbCrLf
    q = q & "        let" & vbCrLf
    q = q & "            nm = Table.ColumnNames(d)," & vbCrLf
    q = q & "            d1 = if List.Count(nm) > 0 and Text.StartsWith(nm{0}, ""Column"") then d else Table.DemoteHeaders(d)," & vbCrLf
    q = q & "            n = Table.ColumnCount(d1)," & vbCrLf
    q = q & "            d2 = Table.RenameColumns(d1, List.Zip({Table.ColumnNames(d1), List.Transform({1..n}, each ""C"" & Text.From(_))}))" & vbCrLf
    q = q & "        in" & vbCrLf
    q = q & "            d2 & #table({""C1""}, {{null}})," & vbCrLf
    q = q & "    All = Table.Combine(List.Transform(Use[Data], Norm))" & vbCrLf
    q = q & "in" & vbCrLf
    q = q & "    All"
    BuildPdfQuery = q
End Function
