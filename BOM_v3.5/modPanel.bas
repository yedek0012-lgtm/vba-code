Attribute VB_Name = "modPanel"
Option Explicit
Option Private Module

' =========================================================================
' SONUÇ PANELÝ (SONUC sayfasý): her dosya için tek satýr, tek bakýþta
'   EKSÝK       : dosyadan parça okunamadý, iþlenirken hata, tanýnmayan / çözülemeyen parça, adet farký
'   KONTROL ET  : aðýrlýk farký tolerans dýþýnda, adedi okunamayan satýr, poz çakýþmasý,
'                 ya da hiçbir doðrulama yapýlamadý (listede aðýrlýk yok ve adet sayýlmadý)
'   GÜVENLÝ     : adet ve/veya aðýrlýk tuttu, sorun yok
' Veriler dosya iþlenirken toplanýr (PanelRecordFile), aðýrlýk farký mutabakattan gelir (PanelSetWeight).
' =========================================================================

Private Const P_FILE As Long = 0
Private Const P_SHEET As Long = 1
Private Const P_ROW As Long = 2
Private Const P_PARTS As Long = 3
Private Const P_BAD As Long = 4
Private Const P_CONF As Long = 5
Private Const P_QTYCHK As Long = 6
Private Const P_QTYDIFF As Long = 7
Private Const P_NOQTY As Long = 8
Private Const P_WCHK As Long = 9
Private Const P_WDIFF As Long = 10
Private Const P_ERR As Long = 11
Private Const P_QSRC As Long = 12
Private Const P_QOUT As Long = 13
Private Const P_LAST As Long = 13

Public Const PANEL_SHEET As String = "SONUC"

' Çalýþtýrma baþýnda
Public Sub PanelReset()
    Set dictPanel = Nothing
    pnlBoltFallback = 0
    pnlAngleZero = 0
    pnlBad = 0
    pnlConf = 0
    pnlFileErr = ""
    pnlActive = False
End Sub

' Her dosya baþýnda
Public Sub PanelBeginFile()
    pnlBad = 0
    pnlConf = 0
    pnlFileErr = ""
    pnlActive = True
End Sub

' Dosya iþlendikten sonra (RecordQtyCheck'ten SONRA çaðrýlýr)
Public Sub PanelRecordFile(ByVal sheetName As String, ByVal sumR As Long, ByVal parts As Long)
    Dim a() As Variant, q As Variant
    On Error Resume Next
    pnlActive = False
    If dictPanel Is Nothing Then Set dictPanel = CreateObject("Scripting.Dictionary")
    ReDim a(0 To P_LAST)
    a(P_FILE) = fileName
    a(P_SHEET) = sheetName
    a(P_ROW) = sumR
    a(P_PARTS) = parts
    a(P_BAD) = pnlBad
    a(P_CONF) = pnlConf
    a(P_ERR) = pnlFileErr
    a(P_QTYCHK) = False
    a(P_QTYDIFF) = 0#
    a(P_NOQTY) = qtyNoRows
    a(P_WCHK) = False
    a(P_WDIFF) = 0#
    a(P_QSRC) = 0#
    a(P_QOUT) = 0#
    If Not dictQtyCheck Is Nothing Then
        If dictQtyCheck.Exists(sumR) Then
            q = dictQtyCheck(sumR)
            a(P_QTYCHK) = (CDbl(q(1)) > 0)
            a(P_QTYDIFF) = CDbl(q(5))
            a(P_QSRC) = CDbl(q(1))
            a(P_QOUT) = CDbl(q(2)) + CDbl(q(3)) + CDbl(q(4))
        End If
    End If
    dictPanel(sheetName & "|" & sumR) = a
End Sub

' Aðýrlýk mutabakatý (CollectReconciliation) her SUM satýrý için çaðýrýr
Public Sub PanelSetWeight(ByVal sheetName As String, ByVal sumR As Long, ByVal hasRef As Boolean, ByVal diff As Double)
    Dim k As String, a As Variant
    On Error Resume Next
    If dictPanel Is Nothing Then Exit Sub
    k = sheetName & "|" & sumR
    If Not dictPanel.Exists(k) Then Exit Sub
    a = dictPanel(k)
    a(P_WCHK) = hasRef
    a(P_WDIFF) = diff
    dictPanel(k) = a
End Sub

' 0 = GÜVENLÝ, 1 = KONTROL ET, 2 = EKSÝK; why = neden / ne yapmalý
Private Function PanelStatus(ByVal a As Variant, ByRef why As String) As Long
    Dim st As Long, tol As Double, d As Double
    tol = prmWeightTolPct / 100#
    why = ""
    If CStr(a(P_ERR)) <> "" Then
        st = 2
        why = AddIssue(why, "iþlenirken hata: " & CStr(a(P_ERR)) & " (dosya yarým kalmýþ olabilir)")
    End If
    If CLng(a(P_PARTS)) = 0 Then
        st = 2
        why = AddIssue(why, "dosyadan hiç parça okunamadý (baþlýk / tablo bulunamadý ya da taranmýþ PDF) -> AUDIT")
    End If
    If CLng(a(P_BAD)) > 0 Then
        st = 2
        why = AddIssue(why, CLng(a(P_BAD)) & " parça tanýnmadý / çözülemedi -> OGRENME'de tanýmlayýn")
    End If
    d = CDbl(a(P_QTYDIFF))
    If CBool(a(P_QTYCHK)) And Abs(d) > 0.001 Then
        st = 2
        why = AddIssue(why, IIf(d > 0, NumToText(d) & " adet EKSÝK", NumToText(-d) & " adet FAZLA") & " (SUM L)")
    End If
    If CLng(a(P_NOQTY)) > 0 Then
        If st < 1 Then st = 1
        why = AddIssue(why, CLng(a(P_NOQTY)) & " satýrda adet okunamadý -> kaynakla karþýlaþtýrýn")
    End If
    d = CDbl(a(P_WDIFF))
    If CBool(a(P_WCHK)) And Abs(d) > tol Then
        If st < 1 Then st = 1
        why = AddIssue(why, "aðýrlýk farký " & Format$(d * 100, "+0.0;-0.0") & "% (SUM J)")
    End If
    If CLng(a(P_CONF)) > 0 Then
        If st < 1 Then st = 1
        why = AddIssue(why, CLng(a(P_CONF)) & " poz çakýþmasý (NOT: ÇAKIÞMA!)")
    End If
    If st = 0 Then
        If CBool(a(P_WCHK)) And CBool(a(P_QTYCHK)) Then
            why = "adet ve aðýrlýk tuttu"
        ElseIf CBool(a(P_QTYCHK)) Then
            why = "adet tuttu (listede toplam aðýrlýk yok)"
        ElseIf CBool(a(P_WCHK)) Then
            why = "aðýrlýk tuttu"
        Else
            st = 1
            why = "doðrulama yapýlamadý: listede aðýrlýk yok, adet sayýlamadý -> elle kontrol edin"
        End If
    End If
    PanelStatus = st
End Function

' SONUC sayfasýný yazar; sayýlarý döndürür. ogrN = OGRENME'deki parça sayýsý, capTxt = kapasite uyarýsý
Public Sub WritePanel(ByVal ogrN As Long, ByVal capTxt As String, ByRef nOk As Long, ByRef nChk As Long, ByRef nBad As Long)
    Dim ws As Worksheet, k As Variant, a As Variant, st As Long, why As String
    Dim r As Long, pass As Long, hdr As Long, n As Long, cel As Range
    Dim lbl As Variant, fillC As Variant, fontC As Variant, wsSum As Worksheet

    nOk = 0: nChk = 0: nBad = 0
    If dictPanel Is Nothing Then Exit Sub
    If dictPanel.Count = 0 Then Exit Sub
    On Error Resume Next

    Set ws = ThisWorkbook.Sheets(PANEL_SHEET)
    If ws Is Nothing Then
        Set wsSum = ThisWorkbook.Sheets("SUM")
        If wsSum Is Nothing Then
            Set ws = ThisWorkbook.Sheets.Add(Before:=ThisWorkbook.Sheets(1))
        Else
            Set ws = ThisWorkbook.Sheets.Add(Before:=wsSum)
        End If
        ws.Name = PANEL_SHEET
    End If
    If ws Is Nothing Then Exit Sub
    ws.Hyperlinks.Delete
    ws.Cells.Clear

    lbl = Array(ChrW(&H2714) & " GÜVENLÝ", ChrW(&H26A0) & " KONTROL ET", ChrW(&H2716) & " EKSÝK")
    fillC = Array(RGB(198, 239, 206), RGB(255, 235, 156), RGB(255, 199, 206))
    fontC = Array(RGB(0, 97, 0), RGB(156, 87, 0), RGB(156, 0, 6))

    For Each k In dictPanel.Keys
        Select Case PanelStatus(dictPanel(k), why)
            Case 0: nOk = nOk + 1
            Case 1: nChk = nChk + 1
            Case Else: nBad = nBad + 1
        End Select
    Next k

    ws.Range("A1").Value = "SONUÇ PANELÝ"
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 14
    ws.Range("C1").Value = Format$(Now, "dd.mm.yyyy hh:nn") & "  |  " & dictPanel.Count & " dosya"
    ws.Range("A2").Value = lbl(0) & ": " & nOk
    ws.Range("B2").Value = lbl(1) & ": " & nChk
    ws.Range("C2").Value = lbl(2) & ": " & nBad
    For st = 0 To 2
        With ws.Cells(2, st + 1)
            .Font.Bold = True
            .Interior.Color = fillC(st)
            .Font.Color = fontC(st)
        End With
    Next st

    ' Genel notlar (dosyaya baðlý olmayan)
    r = 4
    If ogrN > 0 Then
        ws.Cells(r, 1).Value = "OGRENME"
        ws.Hyperlinks.Add Anchor:=ws.Cells(r, 1), Address:="", SubAddress:="'OGRENME'!A1", TextToDisplay:="OGRENME"
        ws.Cells(r, 3).Value = ogrN & " tanýnmayan parça bekliyor: TÜR'ü seçip sarý alanlarý doldurun, KÜTÜPHANEYE AKTAR'a basýn (liste kendini yeniler)."
        r = r + 1
    End If
    If pnlBoltFallback > 0 Then
        ws.Cells(r, 1).Value = "BOLTS&WASHER"
        ws.Cells(r, 3).Value = pnlBoltFallback & " cývata / somun / pul BOLT-LIBRARY'de yok: aðýrlýk tahmini (kýrmýzý kalýn). Kütüphaneye ekleyin."
        r = r + 1
    End If
    If pnlAngleZero > 0 Then
        ws.Cells(r, 1).Value = "ANGLE"
        ws.Cells(r, 3).Value = pnlAngleZero & " satýrda aðýrlýk 0 hesaplanýyor (kod kütüphanede yok ya da kg/m boþ). Alt+F8 > Kutuphane_Kontrol."
        r = r + 1
    End If
    If bufOverflowMsg <> "" Or capTxt <> "" Then
        ws.Cells(r, 1).Value = "KAPASÝTE"
        ws.Cells(r, 3).Value = "Bazý satýrlar eksik ya da aðýrlýksýz: " & Replace(bufOverflowMsg & capTxt, vbCrLf, " ") & " Listeyi bölerek iþleyin."
        r = r + 1
    End If
    If r > 4 Then
        ws.Range(ws.Cells(4, 1), ws.Cells(r - 1, 1)).Font.Bold = True
        ws.Range(ws.Cells(4, 1), ws.Cells(r - 1, 3)).Interior.Color = RGB(255, 235, 156)
        r = r + 1
    End If

    hdr = r
    ws.Range(ws.Cells(hdr, 1), ws.Cells(hdr, 6)).Value = Array("DOSYA (týkla: SUM satýrý)", "DURUM", "NEDEN / NE YAPMALI", _
                                                               "PARÇA SATIRI", "ADET (kaynak / aktarýlan+ayrýlan)", "AÐIRLIK FARKI")
    With ws.Range(ws.Cells(hdr, 1), ws.Cells(hdr, 6))
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(0, 51, 102)
        .WrapText = True
        .VerticalAlignment = xlCenter
    End With
    r = hdr + 1

    ' Önce EKSÝK, sonra KONTROL ET, en son GÜVENLÝ
    For pass = 2 To 0 Step -1
        For Each k In dictPanel.Keys
            a = dictPanel(k)
            st = PanelStatus(a, why)
            If st = pass Then
                Set cel = ws.Cells(r, 1)
                cel.Value = CStr(a(P_FILE))
                ws.Hyperlinks.Add Anchor:=cel, Address:="", _
                                  SubAddress:="'" & Replace(CStr(a(P_SHEET)), "'", "''") & "'!A" & CLng(a(P_ROW)), _
                                  TextToDisplay:=CStr(a(P_FILE))
                ws.Cells(r, 2).Value = lbl(st)
                ws.Cells(r, 2).Interior.Color = fillC(st)
                ws.Cells(r, 2).Font.Color = fontC(st)
                ws.Cells(r, 2).Font.Bold = True
                ws.Cells(r, 3).Value = why
                ws.Cells(r, 4).Value = CLng(a(P_PARTS))
                If CBool(a(P_QTYCHK)) Then
                    ws.Cells(r, 5).Value = NumToText(CDbl(a(P_QSRC))) & " / " & NumToText(CDbl(a(P_QOUT)))
                Else
                    ws.Cells(r, 5).Value = "-"
                End If
                If CBool(a(P_WCHK)) Then
                    ws.Cells(r, 6).Value = CDbl(a(P_WDIFF))
                    ws.Cells(r, 6).NumberFormat = "+0.0%;-0.0%;0.0%"
                Else
                    ws.Cells(r, 6).Value = "liste aðýrlýðý yok"
                End If
                n = n + 1
                r = r + 1
            End If
        Next k
    Next pass

    ws.Columns("A").ColumnWidth = 34
    ws.Columns("B").ColumnWidth = 15
    ws.Columns("C").ColumnWidth = 80
    ws.Columns("D").ColumnWidth = 9
    ws.Columns("E").ColumnWidth = 18
    ws.Columns("F").ColumnWidth = 14
    ws.Range(ws.Cells(hdr + 1, 3), ws.Cells(r, 3)).WrapText = True
    ws.Range(ws.Cells(hdr + 1, 4), ws.Cells(r, 6)).HorizontalAlignment = xlCenter
    ws.Range(ws.Cells(hdr, 1), ws.Cells(r - 1, 6)).Borders.LineStyle = 1
    ws.Range(ws.Cells(hdr, 1), ws.Cells(r - 1, 6)).AutoFilter
    ws.Rows(hdr).RowHeight = 30
End Sub

' Son mesajýn ilk satýrý
Public Function PanelSummaryLine(ByVal nOk As Long, ByVal nChk As Long, ByVal nBad As Long) As String
    If nOk + nChk + nBad = 0 Then Exit Function
    If nChk + nBad = 0 Then
        PanelSummaryLine = "SONUÇ: " & nOk & " dosyanýn HEPSÝ GÜVENLÝ (adet / aðýrlýk tuttu)."
    Else
        PanelSummaryLine = "SONUÇ: " & nOk & " GÜVENLÝ  /  " & nChk & " KONTROL ET  /  " & nBad & " EKSÝK" & vbCrLf & _
                           "Hangi dosyada ne yapýlacaðý 'SONUC' sayfasýnda (ilk sekme)."
    End If
End Function
