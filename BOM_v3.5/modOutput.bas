Attribute VB_Name = "modOutput"
Option Explicit
Option Private Module

' =========================================================================
' ÇIKTI: tamponlar, þablona yazma, temizleme, mutabakat, biçimlendirme, çakýþma kontrolü
' =========================================================================

Public Sub SafeWrite(ws As Worksheet, rw As Long, col As Long, val As Variant, Optional maxRow As Long = 0, Optional maxCol As Long = 0)
    If ws Is Nothing Or rw <= 0 Or col <= 0 Then Exit Sub
    If IsError(val) Or IsArray(val) Then Exit Sub
    
    Dim finalVal As Variant
    finalVal = val
    If VarType(finalVal) = vbString Then
        If IsNumeric(finalVal) And Len(finalVal) > 0 Then
            If Not finalVal Like "*[!0-9]*" Then finalVal = CDbl(finalVal)
        End If
    End If
    
    On Error Resume Next
    If ws Is curWsAngleRef Then
        If rw <= UBound(bufAngle, 1) And col <= UBound(bufAngle, 2) Then bufAngle(rw, col) = finalVal Else Call NoteOverflow(ws, rw)
    ElseIf ws Is curWsPlateRef Then
        If rw <= UBound(bufPlate, 1) And col <= UBound(bufPlate, 2) Then bufPlate(rw, col) = finalVal Else Call NoteOverflow(ws, rw)
    ElseIf ws Is curWsBoltRef Then
        If rw <= UBound(bufBolt, 1) And col <= UBound(bufBolt, 2) Then bufBolt(rw, col) = finalVal Else Call NoteOverflow(ws, rw)
    ElseIf ws Is curWsSumRef Then
        If rw <= UBound(bufSum, 1) And col <= UBound(bufSum, 2) Then bufSum(rw, col) = finalVal Else Call NoteOverflow(ws, rw)
    Else
        ws.Cells(rw, col).Value = finalVal
    End If
    On Error GoTo 0
End Sub

' Tampon sýnýrý aþýldý: satýr yazýlamadý (sayfa baþýna bir kez not edilir, son mesajda gösterilir)
Private Sub NoteOverflow(ByVal ws As Worksheet, ByVal rw As Long)
    On Error Resume Next
    If InStr(bufOverflowMsg, "[" & ws.Name & "]") = 0 Then
        bufOverflowMsg = bufOverflowMsg & "[" & ws.Name & "] " & rw & ". satýrdan itibaren yazýlamadý" & vbCrLf
    End If
End Sub

' Þablondaki formüllü satýr sayýsý yetmiyor mu? (formülsüz satýrýn aðýrlýðý hesaplanmaz)
Public Function TemplateCapacityWarning(ByVal ws As Worksheet, ByVal lastDataRow As Long, ByVal formulaCol As Long) As String
    Dim lastF As Long
    On Error Resume Next
    If lastDataRow < 5 Then Exit Function
    lastF = ws.Cells(ws.Rows.Count, formulaCol).End(xlUp).Row
    If lastF < 5 Then Exit Function
    If Not ws.Cells(lastF, formulaCol).HasFormula Then Exit Function
    If lastDataRow > lastF Then
        TemplateCapacityWarning = "[" & ws.Name & "] þablonda formül " & lastF & ". satýra kadar; " & _
                                  (lastDataRow - lastF) & " satýrýn aðýrlýðý hesaplanmýyor" & vbCrLf
    End If
End Function

Public Sub FlushBufferBlock(ws As Worksheet, ByRef buf() As Variant, ByVal rStart As Long, ByVal rEnd As Long, ByVal cStart As Long, ByVal cEnd As Long)
    If ws Is Nothing Then Exit Sub
    If rEnd < rStart Or cEnd < cStart Then Exit Sub
    
    Dim out() As Variant
    Dim r As Long, c As Long
    Dim numR As Long, numC As Long
    numR = rEnd - rStart + 1
    numC = cEnd - cStart + 1
    
    ReDim out(1 To numR, 1 To numC)
    For r = rStart To rEnd
        For c = cStart To cEnd
            out(r - rStart + 1, c - cStart + 1) = buf(r, c)
        Next c
    Next r
    
    ws.Cells(rStart, cStart).Resize(numR, numC).Value = out
End Sub

Public Sub FlushAllBuffers(wsA As Worksheet, wsP As Worksheet, wsB As Worksheet, wsS As Worksheet, _
                            ByVal lastRowA As Long, ByVal lastRowP As Long, ByVal lastRowB As Long, ByVal lastRowS As Long, _
                            ByVal colA As Long, ByVal colP As Long, ByVal colB As Long)
    On Error Resume Next
    ' 1. ANGLE Sheet Dump
    If Not wsA Is Nothing And lastRowA >= 5 Then
        Call FlushBufferBlock(wsA, bufAngle, 5, lastRowA - 1, 1, 3)      ' A = TYPE (dosya adý)
        Call FlushBufferBlock(wsA, bufAngle, 5, lastRowA - 1, 5, 6)
        If colA > 9 Then Call FlushBufferBlock(wsA, bufAngle, 5, lastRowA - 1, 9, colA - 1)
        Call FlushBufferBlock(wsA, bufAngle, 5, lastRowA - 1, 57, 57)
    End If

    ' 2. PLATE Sheet Dump
    If Not wsP Is Nothing And lastRowP >= 5 Then
        Call FlushBufferBlock(wsP, bufPlate, 5, lastRowP - 1, 1, 6)      ' A = TYPE (dosya adý)
        If colP > 8 Then Call FlushBufferBlock(wsP, bufPlate, 5, lastRowP - 1, 8, colP - 1)
        Call FlushBufferBlock(wsP, bufPlate, 5, lastRowP - 1, 56, 56)
    End If

    ' 3. BOLT Sheet Dump
    If Not wsB Is Nothing And lastRowB >= 5 Then
        Call FlushBufferBlock(wsB, bufBolt, 5, lastRowB - 1, 2, 5)
        If colB > 6 Then Call FlushBufferBlock(wsB, bufBolt, 5, lastRowB - 1, 6, colB - 1)
        Call FlushBufferBlock(wsB, bufBolt, 5, lastRowB - 1, 54, 54)
    End If

    ' 4. SUM Sheet Dump
    If Not wsS Is Nothing And lastRowS >= 5 Then
        Call FlushBufferBlock(wsS, bufSum, 5, lastRowS - 1, 1, 2)
        Call FlushBufferBlock(wsS, bufSum, 5, lastRowS - 1, 8, 9)
    End If
    On Error GoTo 0
End Sub

Public Sub FormatGreenRow(ws As Worksheet, rw As Long)
    On Error Resume Next
    With ws.Range(ws.Cells(rw, 2), ws.Cells(rw, 4))
        .Interior.Color = RGB(0, 153, 76)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = False
    End With
    On Error GoTo 0
End Sub

Public Sub CleanTemplateRanges(wsA As Worksheet, wsP As Worksheet, wsB As Worksheet, wsS As Worksheet)
    On Error Resume Next
    Dim lastRow As Long
    
    If Not wsA Is Nothing Then
        lastRow = wsA.Cells(wsA.Rows.Count, "I").End(xlUp).Row
        If lastRow < 1503 Then lastRow = 1503
        wsA.Range("A5:C" & lastRow).SpecialCells(xlCellTypeConstants).ClearContents
        wsA.Range("E5:F" & lastRow).SpecialCells(xlCellTypeConstants).ClearContents
        wsA.Range("I4:BA" & lastRow).SpecialCells(xlCellTypeConstants).ClearContents
        ' D, G, H, BC, BD (turuncu, formüllü) sütunlarýna dokunulmaz; not sütunu BE
        wsA.Range("BE5:BG" & lastRow).SpecialCells(xlCellTypeConstants).ClearContents
End If
    
    If Not wsP Is Nothing Then
        lastRow = wsP.Cells(wsP.Rows.Count, "H").End(xlUp).Row
        If lastRow < 1503 Then lastRow = 1503
        wsP.Range("A5:F" & lastRow).SpecialCells(xlCellTypeConstants).ClearContents
        wsP.Range("H4:BA" & lastRow).SpecialCells(xlCellTypeConstants).ClearContents
        ' G, BB, BC (turuncu, formüllü) sütunlarýna dokunulmaz; not sütunu BD
        wsP.Range("BD5:BF" & lastRow).SpecialCells(xlCellTypeConstants).ClearContents
End If
    
    If Not wsB Is Nothing Then
        lastRow = wsB.Cells(wsB.Rows.Count, "F").End(xlUp).Row
        If lastRow < 502 Then lastRow = 502
        wsB.Range("B5:E" & lastRow).SpecialCells(xlCellTypeConstants).ClearContents
        wsB.Range("F4:AY" & lastRow).SpecialCells(xlCellTypeConstants).ClearContents
        wsB.Range("BB5:BG" & lastRow).SpecialCells(xlCellTypeConstants).ClearContents
        
        ' Sadece makronun boyadýðý B:E (yeþil somun/pul satýrý, turuncu tahmini aðýrlýk) sýfýrlanýr;
        ' þablonun kendi renkli (formüllü) sütunlarýna dokunulmaz
        wsB.Range("B5:E" & lastRow).Interior.ColorIndex = xlNone
        wsB.Range("B5:E" & lastRow).Font.ColorIndex = xlAutomatic
        wsB.Range("B5:E" & lastRow).Font.Bold = False
End If
    
    If Not wsS Is Nothing Then
        lastRow = wsS.Cells(wsS.Rows.Count, "A").End(xlUp).Row
        If lastRow < 5000 Then lastRow = 5000
        wsS.Range("A5:B" & lastRow).SpecialCells(xlCellTypeConstants).ClearContents
        wsS.Range("H5:H" & lastRow).SpecialCells(xlCellTypeConstants).ClearContents
        wsS.Range("B5:B" & lastRow).Interior.ColorIndex = xlNone
        wsS.Cells(4, 8).Value = "Total Angle+Plate weight"
        wsS.Cells(4, 8).Font.Bold = True
        ' Eskiden kalan Mutabakat, Aktarilan Poz ve Parca sutunlarini/renklerini tamamen temizle:
        Call SafeClear(wsS.Range("I4:I" & lastRow))
        Call SafeClear(wsS.Range("J4:J" & lastRow))
        Call SafeClear(wsS.Range("K4:K" & lastRow))
        wsS.Range("I4:K" & lastRow).Interior.ColorIndex = xlNone
        wsS.Range("I4:K" & lastRow).Font.Bold = False
End If
    On Error GoTo 0
End Sub

Public Function CountExistingOutputFiles(ByVal wsS As Worksheet) As Long
    Dim lastRow As Long, r As Long, n As Long
    On Error GoTo SafeExit
    If wsS Is Nothing Then Exit Function
    lastRow = wsS.Cells(wsS.Rows.Count, 1).End(xlUp).Row
    If lastRow < 5 Then Exit Function
    For r = 5 To lastRow
        If Trim(CStr(wsS.Cells(r, 1).Value)) <> "" Then n = n + 1
    Next r
    CountExistingOutputFiles = n
SafeExit:
End Function

Public Function GetNextOutputRow(ByVal ws As Worksheet, ByVal keyCol As Long, Optional ByVal firstRow As Long = 5) As Long
    Dim lastRow As Long
    On Error GoTo SafeExit
    If ws Is Nothing Then GoTo SafeExit
    lastRow = ws.Cells(ws.Rows.Count, keyCol).End(xlUp).Row
    If lastRow < firstRow Then
        GetNextOutputRow = firstRow
Else
        GetNextOutputRow = lastRow + 1
End If
    Exit Function
SafeExit:
    GetNextOutputRow = firstRow
End Function

' Poz anahtarý: ayný poz no FARKLI kesit / ölçü / boy ile gelirse ayrý satýr olur (toplanmaz)
Public Function MakePosKeyAngle(ByVal posNo As String, ByVal code As Variant, ByVal lengthVal As Double) As String
    MakePosKeyAngle = UCase$(Trim$(posNo)) & "|" & NormCodeText(code) & "|" & NumToText(lengthVal)
End Function

Public Function MakePosKeyPlate(ByVal posNo As String, ByVal thick As Double, ByVal width As Double, ByVal lengthVal As Double) As String
    MakePosKeyPlate = UCase$(Trim$(posNo)) & "|" & NumToText(thick) & "x" & NumToText(width) & "|" & NumToText(lengthVal)
End Function

' TYPE sütunu (A): dosya adý (uzantýsýz). fileName zaten uzantýsýzdýr; adýn içindeki noktalar korunur,
' yalnýzca bilinen bir liste uzantýsýyla bitiyorsa o kýsým atýlýr.
Public Function FileTypeLabel(ByVal fn As String) As String
    Dim p As Long, e As String
    fn = Trim$(fn)
    p = InStrRev(fn, ".")
    If p > 1 Then
        e = LCase$(Mid$(fn, p + 1))
        Select Case e
            Case "xsr", "txt", "xls", "xlsx", "xlsm", "xlsb", "csv", "pdf"
                fn = Left$(fn, p - 1)
        End Select
    End If
    FileTypeLabel = fn
End Function

' Birleþtirilen satýra baþka dosyadan adet gelirse TYPE'a o dosyanýn adý da eklenir
Public Sub AddTypeLabel(ByVal ws As Worksheet, ByVal rw As Long, ByVal isAngle As Boolean, ByVal fn As String)
    Dim cur As String, lbl As String
    On Error Resume Next
    lbl = FileTypeLabel(fn)
    If lbl = "" Then Exit Sub
    If isAngle Then cur = CStr(bufAngle(rw, 1)) Else cur = CStr(bufPlate(rw, 1))
    cur = Trim$(cur)
    If cur = "" Then
        Call SafeWrite(ws, rw, 1, lbl)
    ElseIf InStr(" / " & cur & " / ", " / " & lbl & " / ") = 0 Then
        Call SafeWrite(ws, rw, 1, cur & " / " & lbl)
    End If
End Sub

Public Sub LoadExistingPositionDictionaries(ByVal wsA As Worksheet, ByVal wsP As Worksheet, ByVal dictA As Object, ByVal dictP As Object)
    Dim lastRow As Long, r As Long, key As String, pos As String
    On Error Resume Next
    If Not wsA Is Nothing Then
        lastRow = wsA.Cells(wsA.Rows.Count, 2).End(xlUp).Row
        For r = 5 To lastRow
            pos = Trim(CStr(wsA.Cells(r, 2).Value))
            If pos <> "" Then
                key = MakePosKeyAngle(pos, wsA.Cells(r, 3).Value, BufNum(wsA.Cells(r, 6).Value))
                If Not dictA.Exists(key) Then dictA.Add key, r
End If
        Next r
End If
    If Not wsP Is Nothing Then
        lastRow = wsP.Cells(wsP.Rows.Count, 2).End(xlUp).Row
        For r = 5 To lastRow
            pos = Trim(CStr(wsP.Cells(r, 2).Value))
            If pos <> "" Then
                key = MakePosKeyPlate(pos, BufNum(wsP.Cells(r, 3).Value), BufNum(wsP.Cells(r, 4).Value), BufNum(wsP.Cells(r, 6).Value))
                If Not dictP.Exists(key) Then dictP.Add key, r
End If
        Next r
End If
    On Error GoTo 0
End Sub

Public Sub ShowProgress(ByVal pct As Double, ByVal msg As String)
    Dim ws As Worksheet, shpBG As Object, shpBar As Object, shpTxt As Object
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("SUM")
    
    Set shpBG = ws.Shapes("ProgBG")
    If shpBG Is Nothing Then
        Set shpBG = ws.Shapes.AddShape(1, 150, 150, 400, 40)
        shpBG.Name = "ProgBG"
        shpBG.Fill.ForeColor.RGB = RGB(220, 220, 220)
        shpBG.Line.ForeColor.RGB = RGB(0, 0, 0)
End If
    
    Set shpBar = ws.Shapes("ProgBar")
    If shpBar Is Nothing Then
        Set shpBar = ws.Shapes.AddShape(1, 150, 150, 1, 40)
        shpBar.Name = "ProgBar"
        shpBar.Fill.ForeColor.RGB = RGB(0, 153, 76)
        shpBar.Line.Visible = 0
End If
    
    Set shpTxt = ws.Shapes("ProgTxt")
    If shpTxt Is Nothing Then
        Set shpTxt = ws.Shapes.AddShape(1, 150, 150, 400, 40)
        shpTxt.Name = "ProgTxt"
        shpTxt.Fill.Visible = 0
        shpTxt.Line.Visible = 0
        shpTxt.TextFrame.Characters.Font.Color = RGB(0, 0, 0)
        shpTxt.TextFrame.Characters.Font.Bold = True
        shpTxt.TextFrame.Characters.Font.Size = 12
        shpTxt.TextFrame.HorizontalAlignment = -4108
        shpTxt.TextFrame.VerticalAlignment = -4108
End If
    
    shpBar.width = IIf(400 * pct < 1, 1, 400 * pct)
    shpTxt.TextFrame.Characters.Text = msg & " (% " & Format(pct * 100, "0") & ")"
    DoEvents
    On Error GoTo 0
End Sub

Public Sub RemoveProgress()
    On Error Resume Next
    ThisWorkbook.Sheets("SUM").Shapes("ProgBG").Delete
    ThisWorkbook.Sheets("SUM").Shapes("ProgBar").Delete
    ThisWorkbook.Sheets("SUM").Shapes("ProgTxt").Delete
    On Error GoTo 0
End Sub

Public Sub ResetExcelState()
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayAlerts = True
    Application.EnableEvents = True
End Sub

' =========================================================================
' SÜTUN GENÝÞLÝÐÝ: yazý uzunluðuna göre (sadece veri satýrlarý ölçülür,
' baþlýklar ve formüller etkilenmez). Min/Maks sýnýrlar þablonu korur.
' =========================================================================
Public Sub AutoFitOutputColumns()
    Dim ws As Worksheet, nm As String
    On Error Resume Next
    For Each ws In ThisWorkbook.Worksheets
        nm = UCase$(ws.Name)
        If nm = "SUM" Or nm Like "SUM_P*" Then
            Call FitDataColumn(ws, 1, 5, 30, 80, True)      ' A: dosya adlarý
        ElseIf nm = "ANGLE" Or nm Like "ANGLE_P*" Then
            Call FitDataColumn(ws, 57, 5, 10, 60, False)    ' BE: NOT
        ElseIf nm = "PLATE" Or nm Like "PLATE_P*" Then
            Call FitDataColumn(ws, 56, 5, 10, 60, False)    ' BD: NOT
        ElseIf nm = "BOLTS&WASHER" Or nm Like "BOLTS_P*" Then
            Call FitDataColumn(ws, 3, 5, 18, 45, False)     ' C: malzeme adý
            Call FitDataColumn(ws, 54, 5, 10, 60, False)    ' BB: NOT
        End If
    Next ws
    On Error GoTo 0
End Sub

Public Sub FitDataColumn(ByVal ws As Worksheet, ByVal col As Long, ByVal firstRow As Long, _
                          ByVal minW As Double, ByVal maxW As Double, ByVal keepShapes As Boolean)
    Dim lastRow As Long, rng As Range, shp As Shape, w As Double
    On Error Resume Next
    lastRow = ws.Cells(ws.Rows.Count, col).End(xlUp).Row
    If lastRow < firstRow Then Exit Sub

    ' Logo gibi resimler sütun geniþleyince esnemesin
    If keepShapes Then
        For Each shp In ws.Shapes
            If shp.TopLeftCell.Column <= col And shp.BottomRightCell.Column >= col Then shp.Placement = xlMove
        Next shp
    End If

    Set rng = ws.Range(ws.Cells(firstRow, col), ws.Cells(lastRow, col))
    rng.WrapText = False
    rng.Columns.AutoFit                     ' sadece bu hücrelere göre ölçer
    w = ws.Columns(col).ColumnWidth + 1
    If w < minW Then w = minW
    If w > maxW Then w = maxW
    ws.Columns(col).ColumnWidth = w
    If w >= maxW Then rng.ShrinkToFit = True   ' çok uzun yazý: sütun sýnýrda kalýr, yazý küçülür
    On Error GoTo 0
End Sub

' =========================================================================
' v2.3 - AÐIRLIK MUTABAKATI (SUM I:K) ve KÜTÜPHANE SAÐLIÐI
' =========================================================================
Public Sub WriteSumReconciliation(ByVal ws As Worksheet, ByVal lastRow As Long)
    ' H = müþterinin aðýrlýðý (makro yazar), C+D = þablonun hesapladýðý aðýrlýk
    ' I = Aktarýlan poz, J = Fark % = (C+D-H)/H  (ayrý "hesaplanan" sütunu yok, C ve D zaten gösteriyor)
    Dim r As Long
    On Error Resume Next
    ws.Range("I3").Value = "Fark toleransý:"
    ws.Range("I3").HorizontalAlignment = xlRight
    ws.Range("J3").Value = prmWeightTolPct / 100#
    ws.Range("J3").NumberFormat = "0.0%"
    ws.Range("K3").ClearContents
    ws.Range("I4").Value = "Aktarýlan Poz"
    ws.Range("J4").Value = "Fark % (C+D vs H)"
    ws.Range("K4").ClearContents
    ws.Range("I4:J4").Font.Bold = True
    ws.Range("I4:J4").HorizontalAlignment = xlCenter
    ' Dosya olmayan satýrlarda eski çalýþmalardan kalan I/J/K deðerlerini temizle
    Call SafeClear(ws.Range(ws.Cells(IIf(lastRow < 5, 5, lastRow + 1), 9), ws.Cells(5000, 11)))
    If lastRow < 5 Then Exit Sub
    For r = 5 To lastRow
        If Trim$(ws.Cells(r, 1).Text) <> "" Then
            ws.Cells(r, 10).Formula = "=IF(OR(H" & r & "="""",H" & r & "=0),"""",(C" & r & "+D" & r & "-H" & r & ")/H" & r & ")"
        End If
    Next r
    ws.Range("I5:I" & lastRow).NumberFormat = "0"
    ws.Range("J5:J" & lastRow).NumberFormat = "+0.0%;-0.0%;0.0%"
    ws.Range("I5:J" & lastRow).HorizontalAlignment = xlCenter
End Sub

' AÐIRLIK MUTABAKATI: þablonun hesapladýðý aðýrlýk (C+D) ile listedeki aðýrlýk (H) tolerans (AYARLAR) dýþýnda
' farklý olan dosyalar. Her biri AUDIT'e yazýlýr; dönen metin son mesajda gösterilir (her dosya bir satýr).
Public Function CollectReconciliation(ByVal ws As Worksheet, ByVal lastRow As Long, ByRef nOut As Long, _
                                      ByRef nNoRef As Long, ByRef firstBadRow As Long) As String
    Dim r As Long, h As Double, cd As Double, diff As Double, tol As Double, s As String, nm As String
    On Error Resume Next
    If lastRow < 5 Then Exit Function
    Application.Calculate
    tol = prmWeightTolPct / 100#
    For r = 5 To lastRow
        If Trim$(ws.Cells(r, 1).Text) <> "" Then
            nm = Trim$(ws.Cells(r, 11).Text)                 ' K = kýsa ad
            If nm = "" Then nm = Trim$(ws.Cells(r, 1).Text)
            h = CellNumber(ws.Cells(r, 8))
            cd = CellNumber(ws.Cells(r, 3)) + CellNumber(ws.Cells(r, 4))
            If h <= 0 Then
                nNoRef = nNoRef + 1
            Else
                diff = (cd - h) / h
                If Abs(diff) > tol Then
                    nOut = nOut + 1
                    If firstBadRow = 0 Then firstBadRow = r
                    s = s & "  - " & nm & ": " & Format$(diff * 100, "+0.0;-0.0") & "%  (þablon " & NumToText(Round(cd, 1)) & _
                        " kg / liste " & NumToText(Round(h, 1)) & " kg)" & vbLf
                    AuditRecord Trim$(ws.Cells(r, 1).Text), "SUM", r, ws.Name, "Aðýrlýk mutabakatý", "FILE", "WARNING", 50, _
                                "Fark " & Format$(diff * 100, "+0.0;-0.0") & "%", _
                                "Þablon aðýrlýðý (C+D) listedekinden (H) tolerans dýþýnda farklý: tanýnmayan/atlanan parça, " & _
                                "kütüphanede eksik veya yanlýþ kg/m ya da boy birimi olabilir."
                End If
            End If
        End If
    Next r
    CollectReconciliation = s
End Function

' Metnin ilk n satýrý (fazlasý "... ve N dosya daha")
Public Function FirstLines(ByVal s As String, ByVal n As Long) As String
    Dim a As Variant, i As Long, cnt As Long, out As String
    a = Split(s, vbLf)
    For i = 0 To UBound(a)
        If Len(a(i)) > 0 Then
            cnt = cnt + 1
            If cnt <= n Then out = out & a(i) & vbCrLf
        End If
    Next i
    If cnt > n Then out = out & "  ... ve " & (cnt - n) & " dosya daha (AUDIT)" & vbCrLf
    FirstLines = out
End Function

' ANGLE'da kodu kütüphanede olmayan veya kg/m'si boþ satýrlarý AUDIT'e yazar
Public Sub CheckLibraryHealth(ByVal wsA As Worksheet, ByVal lastRow As Long)
    Dim r As Long, dv As Variant, gv As Variant, issue As String
    On Error Resume Next
    If lastRow < 5 Then Exit Sub
    wsA.Calculate
    For r = 5 To lastRow
        If Trim$(wsA.Cells(r, 3).Text) <> "" Then
            dv = wsA.Cells(r, 4).Value
            gv = wsA.Cells(r, 7).Value
            issue = ""
            If IsError(dv) Or IsEmpty(dv) Then
                issue = "Malzeme kodu kütüphanede yok."
            ElseIf CStr(dv) = "0" Then
                issue = "Malzeme kodu kütüphanede yok."
            ElseIf IsError(gv) Or IsEmpty(gv) Then
                issue = "Kütüphanede kg/m (D sütunu) boþ."
            ElseIf IsNumeric(gv) Then
                If CDbl(gv) = 0 Then issue = "Kütüphanede kg/m (D sütunu) boþ."
            End If
            If issue <> "" Then
                AuditRecord "", "LIBRARY", r, wsA.Name, "Poz " & wsA.Cells(r, 2).Text & " | Kod " & wsA.Cells(r, 3).Text, _
                            "ANGLE", "WARNING", 50, "", issue & " Aðýrlýk 0 hesaplanýyor."
            End If
        End If
    Next r
End Sub

' Kütüphanede olmayan cývata/somun/pul: aðýrlýk hücresi turuncu + AUDIT
Public Sub FlagEstimatedWeight(ByVal ws As Worksheet, ByVal rw As Long, ByVal code As String, ByVal nm As String, ByVal wt As Double)
    On Error Resume Next
    ws.Cells(rw, 5).Interior.Color = RGB(255, 204, 153)
    AuditRecord "", "LIBRARY", rw, ws.Name, code & " | " & nm, "BOLT", "WARNING", 70, _
                "Tahmini aðýrlýk = " & NumToText(wt) & " kg", "BOLT-LIBRARY'de yok; aðýrlýk tahmin edildi. Kütüphaneye ekleyin."
End Sub

' Koþullu biçimlendirme: fonksiyon adý kullanýlmaz (Türkçe Excel'de de çalýþsýn diye)
Public Sub ApplyHealthFormatting()
    Dim ws As Worksheet, nm As String, lastR As Long
    Dim prevSheet As Object
    On Error Resume Next
    Set prevSheet = ActiveSheet
    For Each ws In ThisWorkbook.Worksheets
        If ws.Visible = xlSheetVisible Then
            nm = UCase$(ws.Name)
            If nm = "ANGLE" Or nm Like "ANGLE_P*" Then
                ' Uyarý rengi turuncu (formüllü) D, G, H sütunlarýna uygulanmaz
                Call RemoveCFByAddress(ws, "B5:H3003")        ' eski sürümün kuralý
                Call AddExprCF(ws, "B5:C3003,E5:F3003", "=($C5<>"""")*(($D5=0)+($G5=0))", RGB(255, 199, 206))
            ElseIf nm = "PLATE" Or nm Like "PLATE_P*" Then
                ' Uyarý rengi turuncu (formüllü) G sütununa uygulanmaz
                Call RemoveCFByAddress(ws, "B5:G1503")        ' eski sürümün kuralý
                Call AddExprCF(ws, "B5:F1503", "=($B5<>"""")*(($C5=0)+($D5=0)+($F5=0))", RGB(255, 199, 206))
            ElseIf nm = "SUM" Or nm Like "SUM_P*" Then
                Call RemoveCFByAddress(ws, "K5:K5000")        ' v2.3'ün eski kuralý
                Call AddExprCF(ws, "J5:J5000", "=($J5<>"""")*(($J5>$J$3)+($J5<-$J$3))", RGB(255, 199, 206))
                Call AddExprCF(ws, "J5:J5000", "=($J5<>"""")*($J5<=$J$3)*($J5>=-$J$3)", RGB(198, 239, 206), False)
            End If
        End If
    Next ws
    If Not prevSheet Is Nothing Then prevSheet.Activate
End Sub

Public Sub AddExprCF(ByVal ws As Worksheet, ByVal addr As String, ByVal f As String, ByVal fillColor As Long, _
                      Optional ByVal removeOld As Boolean = True)
    Dim fc As Object
    On Error Resume Next
    If removeOld Then Call RemoveCFByAddress(ws, addr)
    ' Göreli adresler aktif hücreye göre yorumlanýr -> aralýðýn ilk hücresini seç
    ws.Activate
    ws.Range(addr).Cells(1, 1).Select
    Set fc = ws.Range(addr).FormatConditions.Add(Type:=xlExpression, Formula1:=f)
    If Not fc Is Nothing Then
        fc.Interior.Color = fillColor
        fc.StopIfTrue = False
    End If
End Sub

' Sadece bu makronun eklediði (ayný aralýktaki) kurallarý siler; þablondaki diðer kurallar kalýr
Public Sub RemoveCFByAddress(ByVal ws As Worksheet, ByVal addr As String)
    Dim i As Long, target As String, fc As Object
    On Error Resume Next
    target = ws.Range(addr).Address
    For i = ws.Cells.FormatConditions.Count To 1 Step -1
        Set fc = ws.Cells.FormatConditions(i)
        If fc.AppliesTo.Address = target Then fc.Delete
    Next i
End Sub


' =========================================================================
' POZ ÇAKIÞMA KONTROLÜ
' Ayný poz tekrar geldiðinde (birleþtirme modunda farklý dosyalardan ya da
' ayný dosyada iki kez) profil / ölçü / boy farklýysa uyarýr.
' Adetler eskisi gibi toplanýr; NOT sütununa "ÇAKIÞMA!" yazýlýr, AUDIT'e detay düþer.
' =========================================================================
Public Sub CheckPosConflict(ByVal kind As String, ByVal existRow As Long, ByVal posNo As String, _
                            ByVal newCode As String, ByVal newT As Double, ByVal newW As Double, ByVal newL As Double, _
                            ByVal srcFile As String, ByVal srcType As String, ByVal srcRow As Long)
    Dim diff As String, oldCode As String, oldT As Double, oldW As Double, oldL As Double, curNote As String
    On Error GoTo Done
    If existRow <= 0 Then Exit Sub

    If kind = "ANGLE" Then
        oldCode = NormCodeText(bufAngle(existRow, 3))
        oldL = BufNum(bufAngle(existRow, 6))
        If oldCode <> NormCodeText(newCode) Then diff = "Kod: " & oldCode & " <> " & NormCodeText(newCode)
        If Abs(oldL - newL) > 0.5 Then diff = AddIssue(diff, "Boy: " & NumToText(oldL) & " <> " & NumToText(newL))
    Else
        oldT = BufNum(bufPlate(existRow, 3))
        oldW = BufNum(bufPlate(existRow, 4))
        oldL = BufNum(bufPlate(existRow, 6))
        If Abs(oldT - newT) > 0.05 Or Abs(oldW - newW) > 0.5 Then
            diff = "Ölçü: " & NumToText(oldT) & "x" & NumToText(oldW) & " <> " & NumToText(newT) & "x" & NumToText(newW)
        End If
        If Abs(oldL - newL) > 0.5 Then diff = AddIssue(diff, "Boy: " & NumToText(oldL) & " <> " & NumToText(newL))
    End If
    If diff = "" Then Exit Sub

    conflictCount = conflictCount + 1
    If kind = "ANGLE" Then
        curNote = CStr(bufAngle(existRow, 57))
        If InStr(curNote, "ÇAKIÞMA") = 0 Then bufAngle(existRow, 57) = Trim$(curNote & " ÇAKIÞMA!")
    Else
        curNote = CStr(bufPlate(existRow, 56))
        If InStr(curNote, "ÇAKIÞMA") = 0 Then bufPlate(existRow, 56) = Trim$(curNote & " ÇAKIÞMA!")
    End If
    AuditRecord srcFile, srcType, srcRow, "CONFLICT", "Poz " & posNo & " (þablon satýrý " & existRow & ")", kind, "WARNING", 40, _
                diff, "Ayný poz farklý bilgiyle tekrar geldi. Adetler toplandý; hangisinin doðru olduðunu kontrol edin."
Done:
End Sub

Public Function NormCodeText(ByVal v As Variant) As String
    On Error Resume Next
    If IsError(v) Or IsEmpty(v) Then Exit Function
    Select Case VarType(v)
        Case vbDouble, vbSingle, vbLong, vbInteger, vbCurrency, vbDecimal
            NormCodeText = NumToText(CDbl(v))
        Case Else
            NormCodeText = UCase$(Trim$(CStr(v)))
            If NormCodeText Like "#*" And Not NormCodeText Like "*[!0-9.]*" Then NormCodeText = NumToText(ParseBOMNumber(NormCodeText))
    End Select
End Function

Public Function BufNum(ByVal v As Variant) As Double
    On Error Resume Next
    If IsError(v) Or IsEmpty(v) Then Exit Function
    Select Case VarType(v)
        Case vbDouble, vbSingle, vbLong, vbInteger, vbCurrency, vbDecimal
            BufNum = CDbl(v)
        Case Else
            BufNum = ParseBOMNumber(CStr(v))
    End Select
End Function

' =========================================================================
' DOSYA BAÞLIKLARI (ANGLE / PLATE / BOLTS 2. satýr)
' Dar dikey sütunlara uzun dosya adý sýðmýyor. Tüm dosyalarda ORTAK olan parçalar
' atýlýr, sadece farklý kýsým gösterilir:
'   "P-0641482-A-23-Z175 --- D-2-E-2017.2_PC_WA160WAZ-42,0-30,0 ESH"  ->  "Z175 ESH"
' Kýsa adlar SUM!K sütununa yazýlýr; tam ad SUM!A'da kalýr.
' =========================================================================
Public Sub UpdateFileHeaders(ByVal wsS As Worksheet, ByVal wsA As Worksheet, ByVal wsP As Worksheet, ByVal wsB As Worksheet)
    Dim lastR As Long, n As Long, i As Long, c As Long, k As Long
    Dim names() As String, toks() As Variant, cnt As Object, seen As Object, used As Object, t As Variant
    Dim shortNm As String, sumRef As String
    On Error Resume Next
    lastR = wsS.Cells(wsS.Rows.Count, 1).End(xlUp).Row
    If lastR < 5 Then Exit Sub
    n = lastR - 4
    ReDim names(1 To n)
    ReDim toks(1 To n)
    Set cnt = CreateObject("Scripting.Dictionary")
    cnt.CompareMode = 1

    For i = 1 To n
        names(i) = Trim$(wsS.Cells(i + 4, 1).Text)
        toks(i) = FileNameTokens(names(i))
        Set seen = CreateObject("Scripting.Dictionary")
        seen.CompareMode = 1
        For Each t In toks(i)
            If Not seen.Exists(t) Then
                seen.Add t, 1
                cnt(t) = cnt(t) + 1
            End If
        Next t
    Next i

    Set used = CreateObject("Scripting.Dictionary")
    used.CompareMode = 1
    wsS.Range("K4").Value = "Kýsa Ad (baþlýk)"
    wsS.Range("K4").Font.Bold = True
    wsS.Range("K4").HorizontalAlignment = xlCenter
    For i = 1 To n
        shortNm = ""
        If names(i) <> "" Then
            If n >= 2 Then
                For Each t In toks(i)
                    If cnt(t) < n Then
                        If shortNm <> "" Then shortNm = shortNm & " "
                        shortNm = shortNm & t
                    End If
                Next t
            End If
            If shortNm = "" Then shortNm = names(i)
            If Len(shortNm) > 30 Then shortNm = Left$(shortNm, 30)
            If used.Exists(shortNm) Then shortNm = shortNm & " #" & i
            used(shortNm) = 1
        End If
        wsS.Cells(i + 4, 11).Value = shortNm
    Next i
    wsS.Range("K5:K" & lastR).HorizontalAlignment = xlCenter

    ' ANGLE 2. satýr: kýsa ad varsa onu göster (PLATE ve BOLTS zaten ANGLE'a baðlý)
    sumRef = "'" & Replace(wsS.Name, "'", "''") & "'!"
    For c = 9 To 54
        k = c - 4
        wsA.Cells(2, c).Formula = "=IF(" & sumRef & "K" & k & "=""""," & sumRef & "A" & k & "," & sumRef & "K" & k & ")"
    Next c
    Call FitHeaderCells(wsA.Range(wsA.Cells(2, 9), wsA.Cells(2, 54)))
    If Not wsP Is Nothing Then Call FitHeaderCells(wsP.Range(wsP.Cells(2, 8), wsP.Cells(2, 53)))
    If Not wsB Is Nothing Then Call FitHeaderCells(wsB.Range(wsB.Cells(2, 6), wsB.Cells(2, 51)))
End Sub

Private Sub FitHeaderCells(ByVal rng As Range)
    Dim cel As Range
    On Error Resume Next
    For Each cel In rng.Cells
        With cel.MergeArea
            .WrapText = False
            .ShrinkToFit = True
        End With
    Next cel
End Sub

Public Function FileNameTokens(ByVal s As String) As Variant
    Dim t As String
    t = s
    t = Replace(t, "_", " ")
    t = Replace(t, "-", " ")
    t = Replace(t, "(", " ")
    t = Replace(t, ")", " ")
    t = Application.WorksheetFunction.Trim(t)
    FileNameTokens = Split(t, " ")
End Function

' Birleþtirilmiþ hücre yüzünden toplu temizleme hata verirse hücre hücre temizler
Public Sub SafeClear(ByVal rng As Range)
    Dim cel As Range
    On Error Resume Next
    Err.Clear
    rng.ClearContents
    If Err.Number = 0 Then Exit Sub
    Err.Clear
    For Each cel In rng.Cells
        If cel.MergeCells Then
            If cel.Address = cel.MergeArea.Cells(1, 1).Address Then cel.MergeArea.ClearContents
        ElseIf Not IsEmpty(cel.Value) Then
            cel.ClearContents
        End If
    Next cel
End Sub

