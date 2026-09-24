Attribute VB_Name = "modAudit"
Option Explicit
Option Private Module

' =========================================================================
' RAPOR: AUDIT, OGRENME, GECMIS, yedek, REV/DATE
' =========================================================================

Public Sub InitAuditSheet()
    On Error GoTo AuditFail
    Set auditWs = Nothing
    Set auditFiles = CreateObject("Scripting.Dictionary")
    auditFiles.CompareMode = vbTextCompare

    On Error Resume Next
    Set auditWs = ThisWorkbook.Sheets("AUDIT")
    On Error GoTo AuditFail

    If auditWs Is Nothing Then
        Set auditWs = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        auditWs.Name = "AUDIT"
    Else
        auditWs.Cells.Clear
End If

    auditNextRow = 2
    auditProcessed = 0
    auditExact = 0
    auditWarning = 0
    auditError = 0
    auditUnmatched = 0
    auditSkipped = 0

    auditWs.Cells(1, 1).Value = "TARÝH"
    auditWs.Cells(1, 2).Value = "DOSYA"
    auditWs.Cells(1, 3).Value = "KAYNAK"
    auditWs.Cells(1, 4).Value = "SATIR"
    auditWs.Cells(1, 5).Value = "BÖLÜM"
    auditWs.Cells(1, 6).Value = "HAM VERÝ"
    auditWs.Cells(1, 7).Value = "TÝP"
    auditWs.Cells(1, 8).Value = "DURUM"
    auditWs.Cells(1, 9).Value = "GÜVEN"
    auditWs.Cells(1, 10).Value = "AYRIÞTIRILAN"
    auditWs.Cells(1, 11).Value = "AÇIKLAMA"
    auditWs.Range("A1:K1").Font.Bold = True
    auditWs.Range("A1:K1").Interior.Color = RGB(0, 51, 102)
    auditWs.Range("A1:K1").Font.Color = RGB(255, 255, 255)
    auditWs.Range("A1:K1").AutoFilter
    Exit Sub
AuditFail:
    Set auditWs = Nothing
End Sub

Public Sub AuditRecord(ByVal fileName As String, ByVal sourceType As String, ByVal sourceRow As Long, ByVal sectionName As String, ByVal rawText As String, ByVal itemType As String, ByVal status As String, ByVal confidence As Long, ByVal parsedText As String, ByVal issueText As String)
    On Error GoTo AuditExit
    auditProcessed = auditProcessed + 1
    Select Case UCase$(status)
        Case "EXACT": auditExact = auditExact + 1
        Case "WARNING": auditWarning = auditWarning + 1
        Case "ERROR": auditError = auditError + 1
        Case "UNMATCHED": auditUnmatched = auditUnmatched + 1
    End Select

    If auditWs Is Nothing Then Exit Sub
    auditWs.Cells(auditNextRow, 1).Value = Now
    auditWs.Cells(auditNextRow, 2).Value = fileName
    auditWs.Cells(auditNextRow, 3).Value = sourceType
    auditWs.Cells(auditNextRow, 4).Value = sourceRow
    auditWs.Cells(auditNextRow, 5).Value = sectionName
    auditWs.Cells(auditNextRow, 6).Value = Left$(rawText, 500)
    auditWs.Cells(auditNextRow, 7).Value = itemType
    auditWs.Cells(auditNextRow, 8).Value = status
    auditWs.Cells(auditNextRow, 9).Value = confidence
    auditWs.Cells(auditNextRow, 10).Value = Left$(parsedText, 500)
    auditWs.Cells(auditNextRow, 11).Value = Left$(issueText, 500)

    Select Case UCase$(status)
        Case "EXACT"
            auditWs.Cells(auditNextRow, 8).Interior.Color = RGB(0, 153, 76)
            auditWs.Cells(auditNextRow, 8).Font.Color = RGB(255, 255, 255)
        Case "WARNING"
            auditWs.Cells(auditNextRow, 8).Interior.Color = RGB(255, 192, 0)
        Case Else
            auditWs.Cells(auditNextRow, 8).Interior.Color = RGB(192, 0, 0)
            auditWs.Cells(auditNextRow, 8).Font.Color = RGB(255, 255, 255)
    End Select
    auditNextRow = auditNextRow + 1

    ' Köprü: SATIR hücresine týklayýnca kaynaða git
    On Error Resume Next
    Dim linkRow As Long, subAddr As String
    linkRow = auditNextRow - 1
    If sourceRow > 0 Then
        If UCase$(sourceType) = "LIBRARY" Then
            subAddr = "'" & Replace(sectionName, "'", "''") & "'!A" & sourceRow
            auditWs.Hyperlinks.Add Anchor:=auditWs.Cells(linkRow, 4), Address:="", SubAddress:=subAddr, TextToDisplay:=CStr(sourceRow)
        ElseIf auditCurrentPath <> "" Then
            subAddr = ""
            If auditCurrentSheet <> "" Then subAddr = "'" & Replace(auditCurrentSheet, "'", "''") & "'!A" & sourceRow
            auditWs.Hyperlinks.Add Anchor:=auditWs.Cells(linkRow, 4), Address:=auditCurrentPath, SubAddress:=subAddr, TextToDisplay:=CStr(sourceRow)
        End If
    End If
    On Error GoTo 0
AuditExit:
End Sub

Public Sub FinalizeAuditSheet()
    On Error GoTo AuditFail
    If auditWs Is Nothing Then Exit Sub
    With auditWs
        .Rows(1).AutoFilter
        .Columns("A:K").EntireColumn.AutoFit
        .Columns("F:F").ColumnWidth = 42
        .Columns("J:K").ColumnWidth = 42
        .Columns("A:A").NumberFormat = "dd.mm.yyyy hh:mm"

        .Cells(1, 13).Value = "ÖZET"
        .Cells(2, 13).Value = "Toplam kayýt"
        .Cells(2, 14).Value = auditProcessed
        .Cells(3, 13).Value = "Exact"
        .Cells(3, 14).Value = auditExact
        .Cells(4, 13).Value = "Warning"
        .Cells(4, 14).Value = auditWarning
        .Cells(5, 13).Value = "Error"
        .Cells(5, 14).Value = auditError
        .Cells(6, 13).Value = "Unmatched"
        .Cells(6, 14).Value = auditUnmatched
        .Cells(7, 13).Value = "Skipped"
        .Cells(7, 14).Value = auditSkipped
        .Range("M1:N7").Font.Bold = True

        .Cells(9, 13).Value = "ÇALIÞMA BÝLGÝSÝ"
        .Cells(9, 13).Font.Bold = True
        .Cells(10, 13).Value = "Kullanýcý": .Cells(10, 14).Value = Application.UserName & " (" & Environ$("USERNAME") & ")"
        .Cells(11, 13).Value = "Bilgisayar": .Cells(11, 14).Value = Environ$("COMPUTERNAME")
        .Cells(12, 13).Value = "Baþlangýç": .Cells(12, 14).Value = Format$(runStart, "dd.mm.yyyy hh:nn:ss")
        .Cells(13, 13).Value = "Bitiþ": .Cells(13, 14).Value = Format$(Now, "dd.mm.yyyy hh:nn:ss")
        .Cells(14, 13).Value = "Mod": .Cells(14, 14).Value = runMode
        .Cells(15, 13).Value = "REV": .Cells(15, 14).Value = runRev
        .Cells(16, 13).Value = "Yedek": .Cells(16, 14).Value = IIf(runBackupPath = "", "-", runBackupPath)
        .Cells(17, 13).Value = "Kod sürümü": .Cells(17, 14).Value = CODE_VERSION
        .Range("M10:M17").Font.Bold = True
        .Columns("M:N").AutoFit

        If auditError > 0 Or auditUnmatched > 0 Then
            .Tab.Color = RGB(192, 0, 0)
        ElseIf auditWarning > 0 Then
            .Tab.Color = RGB(255, 192, 0)
Else
            .Tab.Color = RGB(0, 153, 76)
End If
    End With
    Exit Sub
AuditFail:
End Sub

' =========================================================================
' ÖÐRENEN KÜTÜPHANE - TOPLU LÝSTE (v3.5)
' Tanýnmayan parçalar tek sayfada. Her parça için TÜR seçilir (ANGLE / PLATE / BOLTS&WASHER) ve
' türe göre 4 veri girilir. KÜTÜPHANEYE AKTAR: kütüphaneye yazar + türü OGRENILEN sayfasýna kaydeder
' (bir sonraki iþlemde parça adýndan tanýnýr, doðru sayfaya gider) + son iþlemi yeniden yapar.
' =========================================================================
Public Sub WriteUnknownsSheet(ByVal dictUnk As Object)
    Dim ws As Worksheet, k As Variant, r As Long, out() As Variant, btn As Object
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("OGRENME")
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = "OGRENME"
    End If

    On Error Resume Next
    ws.Buttons.Delete
    ws.Cells.Validation.Delete
    On Error GoTo 0
    ws.Cells.Clear

    ws.Range("A1:H1").Value = Array("TANINMAYAN PARÇA", "KAYNAK (Dosya / Poz)", "TÜR", _
                                    "MALZEME KODU", "MALZEME CÝNSÝ", _
                                    "DEÐER 1" & vbLf & "ANGLE: kg/m" & vbLf & "PLATE: KALINLIK mm" & vbLf & "BOLT: BÝRÝM AÐIRLIK kg", _
                                    "DEÐER 2" & vbLf & "ANGLE: YÜZEY ALAN" & vbLf & "PLATE: GENÝÞLÝK mm" & vbLf & "BOLT: 1000 ADET kg", _
                                    "DURUM")
    With ws.Range("A1:H1")
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(0, 51, 102)
        .WrapText = True
        .VerticalAlignment = xlCenter
        .HorizontalAlignment = xlCenter
    End With
    ws.Rows(1).RowHeight = 62

    ReDim out(1 To dictUnk.Count, 1 To 5)
    r = 0
    For Each k In dictUnk.keys
        r = r + 1
        out(r, 1) = CStr(k)
        out(r, 2) = CStr(dictUnk(k))
        out(r, 3) = GuessUnknownLibrary(CStr(k))
        out(r, 4) = ""
        out(r, 5) = CStr(k)          ' MALZEME CÝNSÝ: tanýnmayan ad ile önceden doldurulur, deðiþtirilebilir
    Next k
    ws.Range("D2").Resize(r, 1).NumberFormat = "@"       ' kodun baþtaki sýfýrý kaybolmasýn
    ws.Range("A2").Resize(r, 5).Value = out
    ws.Range("F2").Resize(r, 2).NumberFormat = "0.000#"
    ws.Range("C2").Resize(r, 5).Interior.Color = RGB(255, 242, 204)
    ws.Range("A1").Resize(r + 1, 8).Borders.LineStyle = 1

    ' Açýlýr liste seçenekleri gizli Z sütununda (virgül / noktalý virgül bölge ayarýndan baðýmsýz)
    ws.Range("Z1").Value = "ANGLE"
    ws.Range("Z2").Value = "PLATE"
    ws.Range("Z3").Value = "BOLTS&WASHER"
    ws.Columns("Z").Hidden = True
    On Error Resume Next
    With ws.Range("C2").Resize(r, 1).Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=$Z$1:$Z$3"
        .IgnoreBlank = True
        .InCellDropdown = True
    End With
    On Error GoTo 0

    ws.Columns("A").ColumnWidth = 30
    ws.Columns("B").ColumnWidth = 40
    ws.Columns("C").ColumnWidth = 15
    ws.Columns("D").ColumnWidth = 16
    ws.Columns("E").ColumnWidth = 24
    ws.Columns("F").ColumnWidth = 18
    ws.Columns("G").ColumnWidth = 18
    ws.Columns("H").ColumnWidth = 34

    ws.Range("J1").Value = "NASIL KULLANILIR"
    ws.Range("J1").Font.Bold = True
    ws.Range("J2").Value = "1) TÜR seçin: ANGLE (profil/köþebent), PLATE (plaka, ýzgara...), BOLTS&WASHER (cývata, pul, baðlantý seti)."
    ws.Range("J3").Value = "2) Sarý alanlarý doldurun:"
    ws.Range("J4").Value = "   ANGLE        -> KOD, CÝNS, kg/m, YÜZEY ALAN      (L-U-I-O-Y-LIBRARY'ye yazýlýr)"
    ws.Range("J5").Value = "   PLATE        -> KALINLIK mm, GENÝÞLÝK mm (kod/cins isteðe baðlý; boy listeden okunur)"
    ws.Range("J6").Value = "   BOLTS&WASHER -> KOD, CÝNS, BÝRÝM AÐIRLIK veya 1000 ADET AÐIRLIÐI  (BOLT-LIBRARY'ye yazýlýr)"
    ws.Range("J7").Value = "3) Tanýmlamak istemediklerinizi boþ býrakýn. Eksik verili satýrlar aktarýlmaz (DURUM'a bakýn)."
    ws.Range("J8").Value = "4) 'KÜTÜPHANEYE AKTAR': parçalarýn türü OGRENILEN sayfasýna kaydedilir ve son iþlem"
    ws.Range("J9").Value = "   ayný dosyalarla otomatik yeniden yapýlýr (dosyalarý tekrar seçmeniz gerekmez)."

    On Error Resume Next
    Set btn = ws.Buttons.Add(ws.Range("J11").Left, ws.Range("J11").Top, 180, 30)
    btn.Caption = "KÜTÜPHANEYE AKTAR"
    btn.OnAction = "Kutuphaneye_Aktar"
    On Error GoTo 0
End Sub

' Tanýnmayan parçanýn türünü tahmin eder (kullanýcý deðiþtirebilir)
Private Function GuessUnknownLibrary(ByVal nm As String) As String
    nm = UCase$(nm)
    If RxTest(nm, "^M\d|BOLT|SCHRAUB|CIVATA|CÝVATA|SCREW|\bNUT\b|MUTTER|SOMUN|WASH|SCHEIBE|RONDELA|\bPUL\b|\bSCH\b|\bMU\b|\bFDRG\b|FURG|FUTTER|BEFESTIG|STEIGB|ANKER") Then
        GuessUnknownLibrary = "BOLTS&WASHER"
    ElseIf RxTest(nm, "GITTER|BLECH|PLATTE|GRATING|^XP\s?\d|^PL\s?\d|^FL\s?\d") Then
        GuessUnknownLibrary = "PLATE"
    Else
        GuessUnknownLibrary = "ANGLE"
    End If
End Function

' OGRENME'deki TÜR metni -> "ANGLE" / "PLATE" / "BOLT" ("" = seçilmemiþ). Eski PROFÝL / CIVATA da kabul edilir.
Public Function OgrenmeTypeCode(ByVal txt As String) As String
    txt = FoldText(txt)
    If Left$(txt, 3) = "ANG" Or Left$(txt, 4) = "PROF" Then
        OgrenmeTypeCode = "ANGLE"
    ElseIf Left$(txt, 3) = "PLA" Then
        OgrenmeTypeCode = "PLATE"
    ElseIf Left$(txt, 3) = "BOL" Or Left$(txt, 3) = "CIV" Then
        OgrenmeTypeCode = "BOLT"
    End If
End Function

' =========================================================================
' ÖÐRENÝLEN PARÇALAR (OGRENILEN sayfasý): parça adý -> tür + kod + deðerler
' Motor her satýrda önce buraya bakar; varsa sýnýflandýrýcýnýn sonucunu ezer.
' =========================================================================
Public Function LearnedKey(ByVal nm As String) As String
    LearnedKey = Replace(FoldText(nm), " ", "")
End Function

Public Sub LoadLearnedItems()
    Dim ws As Worksheet, r As Long, k As String, tp As String
    Set dictLearned = CreateObject("Scripting.Dictionary")
    dictLearned.CompareMode = 1
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("OGRENILEN")
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub
    For r = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        k = LearnedKey(ws.Cells(r, 1).Text)
        tp = UCase$(Trim$(ws.Cells(r, 2).Text))
        If k <> "" And (tp = "ANGLE" Or tp = "PLATE" Or tp = "BOLT") Then
            dictLearned(k) = Array(tp, Trim$(ws.Cells(r, 3).Text), Trim$(ws.Cells(r, 4).Text), _
                                   CellNumber(ws.Cells(r, 5)), CellNumber(ws.Cells(r, 6)))
        End If
    Next r
End Sub

' Parça türünü kaydeder (varsa günceller). Çalýþma kitabý yapýsý korumasýz olmalý (sayfa eklenebilir).
Public Sub SaveLearnedItem(ByVal nm As String, ByVal tp As String, ByVal code As String, ByVal cins As String, _
                           ByVal v1 As Double, ByVal v2 As Double)
    Dim ws As Worksheet, r As Long, lastR As Long, k As String, hit As Long
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("OGRENILEN")
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = "OGRENILEN"
        ws.Range("A1:G1").Value = Array("PARÇA ADI (listedeki gibi)", "TÜR", "MALZEME KODU", "MALZEME CÝNSÝ", "DEÐER 1", "DEÐER 2", "TARÝH")
        With ws.Range("A1:G1")
            .Font.Bold = True
            .Font.Color = RGB(255, 255, 255)
            .Interior.Color = RGB(0, 51, 102)
        End With
        ws.Columns("A").ColumnWidth = 40
        ws.Columns("B:C").ColumnWidth = 14
        ws.Columns("D").ColumnWidth = 30
        ws.Columns("E:F").ColumnWidth = 12
        ws.Columns("G").ColumnWidth = 16
        ws.Columns("C").NumberFormat = "@"
        ws.Range("I1").Value = "Bu sayfadaki parçalar her iþlemde adýndan tanýnýr ve TÜR'üne göre yazýlýr."
        ws.Range("I2").Value = "DEÐER 1 / 2: ANGLE kg/m / yüzey alan, PLATE kalýnlýk / geniþlik (mm), BOLT birim / 1000 adet aðýrlýk."
        ws.Range("I3").Value = "Yanlýþ kaydý silmek için satýrý silin."
    End If
    k = LearnedKey(nm)
    lastR = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For r = 2 To lastR
        If LearnedKey(ws.Cells(r, 1).Text) = k Then hit = r: Exit For
    Next r
    If hit = 0 Then hit = lastR + 1
    If hit < 2 Then hit = 2
    ws.Cells(hit, 1).Value = nm
    ws.Cells(hit, 2).Value = tp
    ws.Cells(hit, 3).Value = code
    ws.Cells(hit, 4).Value = cins
    ws.Cells(hit, 5).Value = v1
    ws.Cells(hit, 6).Value = v2
    ws.Cells(hit, 7).Value = Now
    ws.Cells(hit, 7).NumberFormat = "dd.mm.yyyy hh:mm"
End Sub

' =========================================================================
' SON ÝÞLEM: dosya listesi ve ayarlar gizli SON_CALISMA sayfasýnda saklanýr;
' OGRENME aktarýmýndan sonra ayný dosyalarla otomatik yeniden iþlemek için.
' =========================================================================
Public Sub SaveLastRun(ByRef files() As String, ByVal fileCount As Long, ByVal isAppend As Boolean, _
                       ByVal pct As Double, ByVal merge As Boolean, ByVal backupPath As String)
    Dim ws As Worksheet, i As Long
    On Error GoTo Done
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("SON_CALISMA")
    On Error GoTo Done
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = "SON_CALISMA"
    End If
    ws.Cells.Clear
    ws.Range("A1:A5").Value = Application.WorksheetFunction.Transpose(Array("MOD", "EK YUZDE", "BIRLESTIR", "YEDEK", "TARIH"))
    ws.Range("B1").Value = IIf(isAppend, "USTUNE", "SIFIRDAN")
    ws.Range("B2").Value = pct
    ws.Range("B3").Value = IIf(merge, 1, 0)
    ws.Range("B4").Value = backupPath
    ws.Range("B5").Value = Now
    For i = 1 To fileCount
        ws.Cells(6 + i, 1).Value = files(i)
    Next i
    ws.Visible = xlSheetVeryHidden
Done:
End Sub

' Üstüne ekle modunda yeniden iþlemeden önce çýktý sayfalarýný son iþlemden önceki yedekten geri yükler.
' Yalnýzca veri alaný (4. satýrdan aþaðýsý) ve SUM REV/DATE geri alýnýr; baþlýklar (birleþtirilmiþ hücreler) ayný kalýr.
' Formül METÝNLERÝ kopyalanýr (kopyala-yapýþtýr yedeðe dýþ baðlantý oluþtururdu). Yedeðin makrolarý çalýþtýrýlmaz.
Public Function RestoreOutputsFromBackup(ByVal backupPath As String) As Boolean
    Dim wbB As Workbook, nm As Variant, wsS As Worksheet, wsD As Worksheet, lastR As Long, lastC As Long
    Dim prevSec As Long, prevEvents As Boolean, bLast As Long
    On Error GoTo Fail
    prevEvents = Application.EnableEvents
    prevSec = Application.AutomationSecurity
    Application.EnableEvents = False
    Application.AutomationSecurity = 3              ' msoAutomationSecurityForceDisable: yedeðin makrolarý çalýþmaz
    Set wbB = Application.Workbooks.Open(backupPath, UpdateLinks:=0, ReadOnly:=True, AddToMru:=False)
    Application.AutomationSecurity = prevSec
    wbB.Windows(1).Visible = False
    ThisWorkbook.Activate
    For Each nm In Array("ANGLE", "PLATE", "BOLTS&WASHER", "SUM")
        Set wsS = wbB.Sheets(CStr(nm))
        Set wsD = ThisWorkbook.Sheets(CStr(nm))
        lastR = UsedLastRow(wsS)
        If UsedLastRow(wsD) > lastR Then lastR = UsedLastRow(wsD)
        lastC = UsedLastCol(wsS)
        If UsedLastCol(wsD) > lastC Then lastC = UsedLastCol(wsD)
        If lastR >= 4 Then Call CopyFormulaBlock(wsS, wsD, 4, lastR, lastC)
    Next nm
    ' SUM: REV / DATE
    ThisWorkbook.Sheets("SUM").Range("G1").Value = wbB.Sheets("SUM").Range("G1").Value
    ThisWorkbook.Sheets("SUM").Range("G2").Value = wbB.Sheets("SUM").Range("G2").Value
    ' BOLTS: geri alýnan iþlemin eklediði satýrlardaki makro boyalarý (yeþil set / turuncu tahmin) temizlenir
    bLast = wbB.Sheets("BOLTS&WASHER").Cells(wbB.Sheets("BOLTS&WASHER").Rows.Count, 2).End(xlUp).Row
    If bLast < 4 Then bLast = 4
    With ThisWorkbook.Sheets("BOLTS&WASHER")
        .Range(.Cells(bLast + 1, 2), .Cells(bLast + 2000, 5)).Interior.ColorIndex = xlColorIndexNone
        .Range(.Cells(bLast + 1, 2), .Cells(bLast + 2000, 5)).Font.ColorIndex = xlAutomatic
    End With
    wbB.Close SaveChanges:=False
    Application.EnableEvents = prevEvents
    RestoreOutputsFromBackup = True
    Exit Function
Fail:
    On Error Resume Next
    Application.AutomationSecurity = prevSec
    Application.EnableEvents = prevEvents
    If Not wbB Is Nothing Then wbB.Close SaveChanges:=False
    RestoreOutputsFromBackup = False
End Function

Private Function UsedLastRow(ByVal ws As Worksheet) As Long
    UsedLastRow = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
End Function

Private Function UsedLastCol(ByVal ws As Worksheet) As Long
    UsedLastCol = ws.UsedRange.Column + ws.UsedRange.Columns.Count - 1
End Function

' Formül metinlerini toplu kopyalar; birleþtirilmiþ hücre yüzünden hata olursa satýr satýr, hücre hücre
Private Sub CopyFormulaBlock(ByVal wsS As Worksheet, ByVal wsD As Worksheet, ByVal r1 As Long, ByVal r2 As Long, ByVal lastC As Long)
    Dim r As Long, c As Long, cel As Range
    On Error Resume Next
    Err.Clear
    wsD.Range(wsD.Cells(r1, 1), wsD.Cells(r2, lastC)).Formula = wsS.Range(wsS.Cells(r1, 1), wsS.Cells(r2, lastC)).Formula
    If Err.Number = 0 Then Exit Sub
    For r = r1 To r2
        Err.Clear
        wsD.Range(wsD.Cells(r, 1), wsD.Cells(r, lastC)).Formula = wsS.Range(wsS.Cells(r, 1), wsS.Cells(r, lastC)).Formula
        If Err.Number <> 0 Then
            For c = 1 To lastC
                Set cel = wsD.Cells(r, c)
                If Not cel.MergeCells Or cel.Address = cel.MergeArea.Cells(1, 1).Address Then
                    cel.Formula = wsS.Cells(r, c).Formula
                End If
            Next c
        End If
    Next r
End Sub

' =========================================================================
' v2.3 - TEMÝZLE düðmesi için rapor sayfalarý
' =========================================================================
Public Sub ClearReportSheets()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = Nothing
    Set ws = ThisWorkbook.Sheets("AUDIT")
    If Not ws Is Nothing Then
        ws.Cells.Clear
        ws.Tab.ColorIndex = xlColorIndexNone
    End If
    Set ws = Nothing
    Set ws = ThisWorkbook.Sheets("FEEDBACK")
    If Not ws Is Nothing Then
        ws.Cells.Clear
        ws.Tab.ColorIndex = xlColorIndexNone
    End If
    Set ws = Nothing
    Set ws = ThisWorkbook.Sheets("OGRENME")
    If Not ws Is Nothing Then
        ws.Buttons.Delete
        ws.Cells.Validation.Delete
        ws.Cells.Clear
    End If
    ' Önceki çalýþtýrmadan kalan SUM mutabakat sütunlarý
    Set ws = Nothing
    Set ws = ThisWorkbook.Sheets("SUM")
    If Not ws Is Nothing Then
        ws.Range("I3:K3").ClearContents
        ws.Range("K5:K5000").FormatConditions.Delete
    End If
    On Error GoTo 0
End Sub

' =========================================================================
' v2.3 - REV / DATE ve GEÇMÝÞ
' =========================================================================
Public Function UpdateRevDate(ByVal isAppend As Boolean) As String
    Dim ws As Worksheet, revNo As Long, t As String
    On Error GoTo Fail
    Set ws = ThisWorkbook.Sheets("SUM")
    t = RxFirst(ws.Range("G1").Text, "\d+")
    revNo = CLng(Val(t))
    If isAppend And prmRevAuto Then revNo = revNo + 1
    ws.Range("G1").Value = "REV:       " & revNo
    ws.Range("G2").Value = "DATE: " & Format$(Day(Date), "00") & "." & Format$(Month(Date), "00") & "." & Year(Date)
    UpdateRevDate = CStr(revNo)
    Exit Function
Fail:
    UpdateRevDate = "?"
End Function

' Kalýcý geçmiþ: her çalýþtýrma bir satýr (Temizle düðmesi bu sayfaya dokunmaz)
Public Sub AppendRunHistory(ByRef files() As String, ByVal fileCount As Long, ByVal posCount As Long)
    Dim ws As Worksheet, r As Long, i As Long, fl As String, fso As Object
    On Error GoTo Done
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("GECMIS")
    On Error GoTo Done
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = "GECMIS"
        ws.Range("A1:J1").Value = Array("TARÝH", "KULLANICI", "BÝLGÝSAYAR", "MOD", "REV", "DOSYA SAYISI", "POZ SAYISI", "KOD SÜRÜMÜ", "YEDEK", "DOSYALAR")
        With ws.Range("A1:J1")
            .Font.Bold = True
            .Interior.Color = RGB(0, 51, 102)
            .Font.Color = RGB(255, 255, 255)
        End With
        ws.Columns("A").ColumnWidth = 18
        ws.Columns("B:H").ColumnWidth = 14
        ws.Columns("I").ColumnWidth = 40
        ws.Columns("J").ColumnWidth = 100
    End If
    Set fso = CreateObject("Scripting.FileSystemObject")
    For i = 1 To fileCount
        If fl <> "" Then fl = fl & " ; "
        fl = fl & fso.GetFileName(files(i))
    Next i
    r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    ws.Cells(r, 1).Value = Now
    ws.Cells(r, 1).NumberFormat = "dd.mm.yyyy hh:mm"
    ws.Cells(r, 2).Value = Application.UserName
    ws.Cells(r, 3).Value = Environ$("COMPUTERNAME")
    ws.Cells(r, 4).Value = runMode
    ws.Cells(r, 5).Value = runRev
    ws.Cells(r, 6).Value = fileCount
    ws.Cells(r, 7).Value = posCount
    ws.Cells(r, 8).Value = CODE_VERSION
    ws.Cells(r, 9).Value = IIf(runBackupPath = "", "-", runBackupPath)
    ws.Cells(r, 10).Value = Left$(fl, 32000)
Done:
End Sub

' =========================================================================
' v2.3 - OTOMATÝK YEDEK
' =========================================================================
Public Function MakeBackup() As String
    Dim fso As Object, folder As String, baseName As String, ext As String, target As String
    On Error GoTo Fail
    Set fso = CreateObject("Scripting.FileSystemObject")

    folder = prmBackupFolder
    If folder = "" Then
        If ThisWorkbook.Path <> "" And LCase$(Left$(ThisWorkbook.Path, 4)) <> "http" Then
            folder = ThisWorkbook.Path & "\YEDEK"
        Else
            ' OneDrive/SharePoint (https://...) veya kaydedilmemiþ dosya
            folder = Environ$("USERPROFILE") & "\Documents\BOM_YEDEK"
        End If
    End If
    If Not EnsureFolder(fso, folder) Then GoTo Fail

    baseName = fso.GetBaseName(ThisWorkbook.Name)
    ext = fso.GetExtensionName(ThisWorkbook.Name)
    If ext = "" Then ext = "xlsm"
    target = folder & "\" & baseName & "_" & Format$(Now, "yyyymmdd_hhnnss") & "." & ext
    ThisWorkbook.SaveCopyAs target
    Call PruneBackups(fso, folder, baseName & "_", prmBackupKeep)
    MakeBackup = target
    Exit Function
Fail:
    MakeBackup = ""
    AuditRecord "", "SYSTEM", 0, "YEDEK", folder, "BACKUP", "WARNING", 50, "", "Yedek alýnamadý: " & Err.Description
End Function

Public Function EnsureFolder(ByVal fso As Object, ByVal folder As String) As Boolean
    On Error GoTo Fail
    If fso.FolderExists(folder) Then EnsureFolder = True: Exit Function
    If Not fso.FolderExists(fso.GetParentFolderName(folder)) Then
        If Not EnsureFolder(fso, fso.GetParentFolderName(folder)) Then Exit Function
    End If
    fso.CreateFolder folder
    EnsureFolder = True
    Exit Function
Fail:
    EnsureFolder = False
End Function

' En yeni 'keep' yedeði býrakýr, eskileri siler (sadece bu dosyanýn yedekleri)
Public Sub PruneBackups(ByVal fso As Object, ByVal folder As String, ByVal prefix As String, ByVal keep As Long)
    Dim f As Object, names() As String, dates() As Date, n As Long, i As Long, j As Long
    Dim tN As String, tD As Date
    On Error GoTo Done
    ReDim names(1 To 1000)
    ReDim dates(1 To 1000)
    For Each f In fso.GetFolder(folder).Files
        If LCase$(Left$(f.Name, Len(prefix))) = LCase$(prefix) Then
            n = n + 1
            If n > 1000 Then Exit For
            names(n) = f.Path
            dates(n) = f.DateLastModified
        End If
    Next f
    If n <= keep Then GoTo Done
    ' yeniden eskiye sýrala
    For i = 1 To n - 1
        For j = i + 1 To n
            If dates(j) > dates(i) Then
                tD = dates(i): dates(i) = dates(j): dates(j) = tD
                tN = names(i): names(i) = names(j): names(j) = tN
            End If
        Next j
    Next i
    For i = keep + 1 To n
        fso.DeleteFile names(i), True
    Next i
Done:
End Sub
