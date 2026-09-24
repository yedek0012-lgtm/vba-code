Attribute VB_Name = "modMotorPDF"
Option Explicit
Option Private Module

' =========================================================================
' PDF motoru: Tekla / CAD çýktýsý (METÝN tabanlý) PDF'teki malzeme listesini okur.
' Power Query'nin PDF baðlayýcýsý (Excel 365 / 2021, Windows) sayfayý KOORDÝNATA göre
' tabloya çevirir; tablo geçici bir çalýþma kitabýna dökülür ve normal Excel motoru
' (baþlýk algýlama, sýnýflandýrma, mutabakat) aynen çalýþýr.
' Taranmýþ (resim) PDF'ler okunamaz -> AUDIT'e hata yazýlýr.
' =========================================================================

Public Sub ProcessPdfFile()
    Dim wb As Workbook, wsQ As Worksheet, wsD As Worksheet, lo As ListObject, qn As String, v As Variant
    On Error GoTo Fail
    Application.StatusBar = "PDF okunuyor: " & fileName
    Set wb = Application.Workbooks.Add(xlWBATWorksheet)
    Set wsQ = wb.Worksheets(1)
    qn = "BOM_PDF"
    wb.Queries.Add Name:=qn, Formula:=BuildPdfQuery(CStr(vItem))
    Set lo = wsQ.ListObjects.Add(SourceType:=0, Source:="OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=" & qn & ";Extended Properties=""""", _
                                 Destination:=wsQ.Range("$A$1"))
    With lo.QueryTable
        .CommandType = xlCmdSql
        .CommandText = Array("SELECT * FROM [" & qn & "]")
        .BackgroundQuery = False
        .Refresh BackgroundQuery:=False
    End With

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
Fail:
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
