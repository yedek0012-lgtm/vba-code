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
        If limitRowA > 5 And rw >= limitRowA Then
            Call NoteOverflow(ws, rw)
        ElseIf rw <= UBound(bufAngle, 1) And col <= UBound(bufAngle, 2) Then
            bufAngle(rw, col) = finalVal
        Else
            Call NoteOverflow(ws, rw)
        End If
    ElseIf ws Is curWsPlateRef Then
        If limitRowP > 5 And rw >= limitRowP Then
            Call NoteOverflow(ws, rw)
        ElseIf rw <= UBound(bufPlate, 1) And col <= UBound(bufPlate, 2) Then
            bufPlate(rw, col) = finalVal
        Else
            Call NoteOverflow(ws, rw)
        End If
    ElseIf ws Is curWsBoltRef Then
        If limitRowB > 5 And rw >= limitRowB Then
            Call NoteOverflow(ws, rw)
        ElseIf rw <= UBound(bufBolt, 1) And col <= UBound(bufBolt, 2) Then
            bufBolt(rw, col) = finalVal
        Else
            Call NoteOverflow(ws, rw)
        End If
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
    ' veri sýnýrýnýn altýna (þablonun formül / toplam satýrlarý) ASLA yazýlmaz
    If limitRowA > 5 And lastRowA > limitRowA Then lastRowA = limitRowA
    If limitRowP > 5 And lastRowP > limitRowP Then lastRowP = limitRowP
    If limitRowB > 5 And lastRowB > limitRowB Then lastRowB = limitRowB
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
        ' C (MATERIAL) ve E (UNIT WEIGHT) turuncu formüllü sütunlar: formüller korunur
        Call FlushBoltBlock(wsB, 5, lastRowB - 1)
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

' BOLTS: B (kod) ve D (kalite) deðer olarak yazýlýr. C ve E'de þablon formülü varsa
' dokunulmaz (BOLT-LIBRARY'den kendisi çeker); formül boþ/0 dönerse (kod kütüphanede yok
' ya da formülün aralýðý dýþýnda) makronun deðeri yazýlýr ve kýrmýzý kalýn yazýyla iþaretlenir.
' Formül, bir sonraki SIFIRDAN temizliðinde (CleanTemplateRanges) geri konur.
Private Sub FlushBoltBlock(ByVal ws As Worksheet, ByVal rStart As Long, ByVal rEnd As Long)
    If ws Is Nothing Or rEnd < rStart Then Exit Sub
    Call FlushBufferBlock(ws, bufBolt, rStart, rEnd, 2, 2)
    Call FlushBufferBlock(ws, bufBolt, rStart, rEnd, 4, 4)
    Call FlushKeepFormula(ws, rStart, rEnd, 3)
    Call FlushKeepFormula(ws, rStart, rEnd, 5)
End Sub

Private Sub FlushKeepFormula(ByVal ws As Worksheet, ByVal rStart As Long, ByVal rEnd As Long, ByVal col As Long)
    Dim r As Long, bv As Variant, cel As Range, cv As Variant, useMacro As Boolean
    On Error Resume Next
    ws.Range(ws.Cells(rStart, col), ws.Cells(rEnd, col)).Calculate
    For r = rStart To rEnd
        bv = bufBolt(r, col)
        Set cel = ws.Cells(r, col)
        If Not cel.HasFormula Then
            If Not IsEmpty(bv) Then cel.Value = bv
        ElseIf Not IsEmpty(bv) And Not IsEmpty(bufBolt(r, 2)) Then
            cv = cel.Value
            useMacro = False
            If IsError(cv) Then
                useMacro = True
            ElseIf Trim$(CStr(cv)) = "" Then
                useMacro = True
            ElseIf col = 5 Then
                If CellNumber(cel) = 0 And BufNum(bv) > 0 Then useMacro = True
            End If
            If useMacro Then
                If col = 5 Then pnlBoltFallback = pnlBoltFallback + 1
                cel.Value = bv
                cel.Font.Color = RGB(192, 0, 0)
                cel.Font.Bold = True
                AuditRecord "", "LIBRARY", r, ws.Name, CStr(bufBolt(r, 2)) & " | " & CStr(bufBolt(r, 3)), "BOLT", "WARNING", 70, _
                            IIf(col = 5, "Aðýrlýk", "Malzeme adý") & " = " & CStr(bv), _
                            "Þablon formülü boþ döndü (kod BOLT-LIBRARY'de yok ya da formülün aralýðý dýþýnda); makro deðeri yazýldý (kýrmýzý)."
            End If
        End If
    Next r
    On Error GoTo 0
End Sub

' Temizlikte, makronun deðer yazdýðý (formülü silinmiþ) C/E hücrelerine þablon formülünü geri koy
Private Sub RestoreColumnFormula(ByVal ws As Worksheet, ByVal col As Long, ByVal lastRow As Long)
    Dim r As Long, firstF As Long, lastF As Long, fR1C1 As String
    On Error Resume Next
    For r = 5 To lastRow
        If ws.Cells(r, col).HasFormula Then
            If firstF = 0 Then firstF = r: fR1C1 = ws.Cells(r, col).FormulaR1C1
            lastF = r
        End If
    Next r
    If firstF = 0 Or fR1C1 = "" Then Exit Sub
    For r = firstF To lastF
        If Not ws.Cells(r, col).HasFormula Then
            If IsEmpty(ws.Cells(r, col).Value) Then ws.Cells(r, col).FormulaR1C1 = fR1C1
        End If
    Next r
End Sub

Public Sub FormatGreenRow(ws As Worksheet, rw As Long)
    On Error Resume Next
    ' Sadece B (kod) ve D (kalite); turuncu formüllü C ve E'nin rengine dokunulmaz
    With ws.Range(ws.Cells(rw, 2).Address & "," & ws.Cells(rw, 4).Address)
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
        Call ClearConstants(wsA.Range("A5:C" & lastRow))
        Call ClearConstants(wsA.Range("E5:F" & lastRow))
        Call ClearConstants(wsA.Range("I4:BA" & lastRow))
        ' D, G, H, BC, BD (turuncu, formüllü) sütunlarýna dokunulmaz; not sütunu BE
        Call ClearConstants(wsA.Range("BE5:BG" & lastRow))
End If
    
    If Not wsP Is Nothing Then
        lastRow = wsP.Cells(wsP.Rows.Count, "H").End(xlUp).Row
        If lastRow < 1503 Then lastRow = 1503
        Call ClearConstants(wsP.Range("A5:F" & lastRow))
        Call ClearConstants(wsP.Range("H4:BA" & lastRow))
        ' G, BB, BC (turuncu, formüllü) sütunlarýna dokunulmaz; not sütunu BD
        Call ClearConstants(wsP.Range("BD5:BF" & lastRow))
End If
    
    If Not wsB Is Nothing Then Call CleanBoltSheet(wsB)
    
    If Not wsS Is Nothing Then
        lastRow = wsS.Cells(wsS.Rows.Count, "A").End(xlUp).Row
        If lastRow < 5000 Then lastRow = 5000
        Call ClearConstants(wsS.Range("A5:B" & lastRow))
        Call ClearConstants(wsS.Range("H5:H" & lastRow))
        wsS.Range("B5:B" & lastRow).Interior.ColorIndex = xlNone
        wsS.Cells(4, 8).Value = "Total Angle+Plate weight"
        wsS.Cells(4, 8).Font.Bold = True
        ' Eskiden kalan Mutabakat, Aktarilan Poz ve Parca sutunlarini/renklerini tamamen temizle:
        Call SafeClear(wsS.Range("I4:I" & lastRow))
        Call SafeClear(wsS.Range("J4:J" & lastRow))
        Call SafeClear(wsS.Range("K4:K" & lastRow))
        Call SafeClear(wsS.Range("L4:L" & lastRow))
        wsS.Range("I4:L" & lastRow).Interior.ColorIndex = xlNone
        wsS.Range("I4:L" & lastRow).Font.Bold = False
End If
    On Error GoTo 0
End Sub

' BOLTS&WASHER veri alanýný temizler (formüller korunur / geri konur); SIFIRDAN ve ekleme modunda liste baþtan dizilirken
Public Sub CleanBoltSheet(ByVal wsB As Worksheet)
    Dim lastRow As Long
    On Error Resume Next
        lastRow = wsB.Cells(wsB.Rows.Count, "F").End(xlUp).Row
        If wsB.Cells(wsB.Rows.Count, "B").End(xlUp).Row > lastRow Then lastRow = wsB.Cells(wsB.Rows.Count, "B").End(xlUp).Row
        If lastRow < 502 Then lastRow = 502
        Call ClearConstants(wsB.Range("B5:E" & lastRow))
        ' makronun deðer yazdýðý C/E hücrelerine þablon formülünü geri koy
        Call RestoreColumnFormula(wsB, 3, lastRow)
        Call RestoreColumnFormula(wsB, 5, lastRow)
        Call ClearConstants(wsB.Range("F4:AY" & lastRow))
        Call ClearConstants(wsB.Range("BB5:BG" & lastRow))
        
        ' Sadece makronun boyadýðý B ve D (yeþil somun/pul satýrý) sýfýrlanýr;
        ' turuncu formüllü C ve E'nin dolgu rengine dokunulmaz (yazý rengi/kalýnlýk sýfýrlanýr)
        wsB.Range("B5:B" & lastRow).Interior.ColorIndex = xlNone
        wsB.Range("D5:D" & lastRow).Interior.ColorIndex = xlNone
        wsB.Range("B5:E" & lastRow).Font.ColorIndex = xlAutomatic
        wsB.Range("B5:E" & lastRow).Font.Bold = False
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
    Call SafeClear(ws.Range(ws.Cells(IIf(lastRow < 5, 5, lastRow + 1), 9), ws.Cells(5000, 12)))
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
                Call PanelSetWeight(ws.Name, r, False, 0)
            Else
                diff = (cd - h) / h
                Call PanelSetWeight(ws.Name, r, True, diff)
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

' ADET MUTABAKATI (dosya baþýna; Excel / CSV / PDF motoru sayar, XSR sayýlmaz)
' Kaynaktaki parça adedi = ANGLE + PLATE'e yazýlan (dosyanýn sütunu, tampondan okunur) + cývata
'                          + bilinçli atlanan (somun/pul seti, çift sayým) + tanýnmayan (OGRENME)
' Fark ya da adedi okunamayan satýr varsa sonuç saklanýr; CollectQtyCheck SUM L'ye yazar ve raporlar.
Public Sub RecordQtyCheck(ByVal sumR As Long, ByVal colA As Long, ByVal colP As Long)
    Dim r As Long, placed As Double
    On Error Resume Next
    If qtySrc <= 0 And qtyNoRows = 0 Then Exit Sub
    If colA >= 1 And colA <= UBound(bufAngle, 2) Then
        For r = 5 To UBound(bufAngle, 1)
            If Not IsEmpty(bufAngle(r, colA)) Then placed = placed + BufNum(bufAngle(r, colA))
        Next r
    End If
    If colP >= 1 And colP <= UBound(bufPlate, 2) Then
        For r = 5 To UBound(bufPlate, 1)
            If Not IsEmpty(bufPlate(r, colP)) Then placed = placed + BufNum(bufPlate(r, colP))
        Next r
    End If
    If dictQtyCheck Is Nothing Then Set dictQtyCheck = CreateObject("Scripting.Dictionary")
    dictQtyCheck(sumR) = Array(fileName, qtySrc, placed + qtyBolt, qtySkip, qtyUnk, _
                               qtySrc - (placed + qtyBolt + qtySkip + qtyUnk), qtyNoRows)
End Sub

' SUM L = adet farký (kaynak - aktarýlan - ayrýlan). 0 = tamam. Sorunlu dosyalar AUDIT'e ve dönen metne.
Public Function CollectQtyCheck(ByVal ws As Worksheet, ByRef nBad As Long) As String
    Dim k As Variant, a As Variant, r As Long, s As String, ln As String
    On Error Resume Next
    If dictQtyCheck Is Nothing Then Exit Function
    ws.Range("L4").Value = "Adet Farký"
    ws.Range("L4").Font.Bold = True
    ws.Range("L4").HorizontalAlignment = xlCenter
    For Each k In dictQtyCheck.Keys
        a = dictQtyCheck(k)
        r = CLng(k)
        ws.Cells(r, 12).Value = a(5)
        ws.Cells(r, 12).NumberFormat = "+0;-0;0"
        ws.Cells(r, 12).HorizontalAlignment = xlCenter
        If Abs(a(5)) > 0.001 Or a(6) > 0 Then
            nBad = nBad + 1
            ws.Cells(r, 12).Interior.Color = RGB(255, 199, 206)
            ln = "kaynak " & NumToText(a(1)) & " adet, aktarýlan " & NumToText(a(2))
            If a(3) > 0 Then ln = ln & ", ayrýlan (set/çift) " & NumToText(a(3))
            If a(4) > 0 Then ln = ln & ", tanýnmayan " & NumToText(a(4))
            If a(5) > 0.001 Then ln = ln & " -> " & NumToText(a(5)) & " adet EKSÝK"
            If a(5) < -0.001 Then ln = ln & " -> " & NumToText(-a(5)) & " adet FAZLA"
            If a(6) > 0 Then ln = ln & "; " & a(6) & " satýrda adet okunamadý"
            s = s & "  - " & CStr(a(0)) & ": " & ln & vbLf
            AuditRecord CStr(a(0)), "SUM", r, ws.Name, "Adet mutabakatý", "FILE", "WARNING", 50, ln, _
                        "Kaynaktaki parça adedi þablona yazýlan + ayrýlan adetle tutmuyor ya da bazý satýrlarda adet okunamadý."
        Else
            ws.Cells(r, 12).Interior.ColorIndex = xlColorIndexNone
        End If
    Next k
    Set dictQtyCheck = Nothing
    CollectQtyCheck = s
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
                pnlAngleZero = pnlAngleZero + 1
                AuditRecord "", "LIBRARY", r, wsA.Name, "Poz " & wsA.Cells(r, 2).Text & " | Kod " & wsA.Cells(r, 3).Text, _
                            "ANGLE", "WARNING", 50, "", issue & " Aðýrlýk 0 hesaplanýyor."
            End If
        End If
    Next r
End Sub

' Kütüphanede olmayan cývata/somun/pul: AUDIT kaydý (aðýrlýk hücresi yazýlýrken kýrmýzý iþaretlenir;
' turuncu formüllü E sütununun dolgusu deðiþtirilmez)
Public Sub FlagEstimatedWeight(ByVal ws As Worksheet, ByVal rw As Long, ByVal code As String, ByVal nm As String, ByVal wt As Double)
    On Error Resume Next
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
    If pnlActive Then pnlConf = pnlConf + 1
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
' Baþlýk = SUM!K doluysa K (kullanýcýnýn yazdýðý ad), boþsa SUM!A (tam dosya adý).
' K'ya elle yazýlan adlar KISA_ADLAR'da saklanýr ve sonraki çalýþtýrmalarda geri yazýlýr.
' =========================================================================
Public Sub UpdateFileHeaders(ByVal wsS As Worksheet, ByVal wsA As Worksheet, ByVal wsP As Worksheet, ByVal wsB As Worksheet)
    ' Baþlýk adý: SUM K boþsa SUM A'daki tam dosya adý; K'ya elle yazýlan ad korunur (KISA_ADLAR).
    ' (Eskiden makro K'ya otomatik kýsaltma yazýyordu; kýsaltma adlarý bozduðu için kaldýrýldý.)
    Dim lastR As Long, r As Long, c As Long, k As Long, nm As String, sumRef As String
    On Error Resume Next
    lastR = wsS.Cells(wsS.Rows.Count, 1).End(xlUp).Row
    wsS.Range("K4").Value = "Baþlýk adý (boþ = dosya adý)"
    wsS.Range("K4").Font.Bold = True
    wsS.Range("K4").HorizontalAlignment = xlCenter
    If lastR >= 5 Then
        For r = 5 To lastR
            nm = Trim$(wsS.Cells(r, 1).Text)
            wsS.Cells(r, 11).Value = ""
            If nm <> "" And Not dictShortNames Is Nothing Then
                If dictShortNames.Exists(nm) Then wsS.Cells(r, 11).Value = dictShortNames(nm)
            End If
        Next r
        wsS.Range("K5:K" & lastR).HorizontalAlignment = xlCenter
    End If

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

' Veri satýrlarýnýn yazý tipini sütun sütun tek tipe getirir (þablonda satýrlar farklý biçimlenmiþ olabilir).
' Her sütunda ilk 100 satýrda en çok kullanýlan yazý tipi / boy / kalýnlýk bütün veri satýrlarýna uygulanýr;
' adet sütunlarý (qFirst..qLast) ilk adet sütununun biçimini alýr. Renklere ve formüllü sütunlara dokunulmaz.
Public Sub NormalizeDataFonts(ByVal ws As Worksheet, ByVal lastRow As Long, ByVal cols As Variant, _
                              ByVal qFirst As Long, ByVal qLast As Long)
    Dim i As Long
    On Error Resume Next
    If ws Is Nothing Or lastRow < 5 Then Exit Sub
    For i = LBound(cols) To UBound(cols)
        Call ApplyColumnFontMode(ws, CLng(cols(i)), CLng(cols(i)), CLng(cols(i)), lastRow)
    Next i
    If qLast >= qFirst Then Call ApplyColumnFontMode(ws, qFirst, qFirst, qLast, lastRow)
End Sub

Private Sub ApplyColumnFontMode(ByVal ws As Worksheet, ByVal sampleCol As Long, ByVal c1 As Long, ByVal c2 As Long, ByVal lastRow As Long)
    Dim r As Long, rEnd As Long, key As String, best As String, cnt As Object, k As Variant, pr As Variant
    On Error Resume Next
    Set cnt = CreateObject("Scripting.Dictionary")
    rEnd = lastRow
    If rEnd > 104 Then rEnd = 104
    For r = 5 To rEnd
        With ws.Cells(r, sampleCol).Font
            key = .Name & "|" & .Size & "|" & CStr(.Bold)
        End With
        cnt(key) = cnt(key) + 1
    Next r
    For Each k In cnt.Keys
        If best = "" Then best = CStr(k)
        If cnt(k) > cnt(best) Then best = CStr(k)
    Next k
    pr = Split(best, "|")
    If UBound(pr) <> 2 Then Exit Sub
    With ws.Range(ws.Cells(5, c1), ws.Cells(lastRow, c2)).Font
        .Name = pr(0)
        .Size = CDbl(pr(1))
        .Bold = (pr(2) = "True")
    End With
End Sub

' Dosya adý baþlýklarý (dikey yazý): yazý KÜÇÜLTÜLMEZ (þablonun yazý boyu kalýr, SUM'daki gibi);
' en uzun ad sýðmýyorsa baþlýk satýrý (2) uzatýlýr. Önceki sürümün "sýðdýrmak için küçült" ayarý kaldýrýlýr.
Private Sub FitHeaderCells(ByVal rng As Range)
    Dim cel As Range, maxLen As Long, fs As Double, need As Double, have As Double
    On Error Resume Next
    Dim key As String, best As String, cnt As Object, k As Variant, pr As Variant
    ' Baþlýklar tek tip: bu satýrda en çok kullanýlan yazý tipi / boyu / kalýnlýk hepsine uygulanýr
    Set cnt = CreateObject("Scripting.Dictionary")
    For Each cel In rng.Cells
        key = cel.Font.Name & "|" & cel.Font.Size & "|" & CStr(cel.Font.Bold)
        cnt(key) = cnt(key) + 1
    Next cel
    For Each k In cnt.Keys
        If best = "" Then best = CStr(k)
        If cnt(k) > cnt(best) Then best = CStr(k)
    Next k
    pr = Split(best, "|")
    For Each cel In rng.Cells
        With cel.MergeArea
            .ShrinkToFit = False
            .WrapText = False
            If UBound(pr) = 2 Then
                .Font.Name = pr(0)
                .Font.Size = CDbl(pr(1))
                .Font.Bold = (pr(2) = "True")
            End If
        End With
        If Len(cel.Text) > maxLen Then maxLen = Len(cel.Text)
    Next cel
    If maxLen = 0 Then Exit Sub
    fs = rng.Cells(1, 1).Font.Size
    If fs <= 0 Then fs = 8
    need = maxLen * fs * 0.62 + 10                  ' dikey yazý boyu (pt), kabaca
    have = rng.Cells(1, 1).MergeArea.Height
    If need > have Then
        With rng.Worksheet.Rows(rng.Row)
            If .RowHeight + (need - have) <= 409 Then
                .RowHeight = .RowHeight + (need - have)
            Else
                .RowHeight = 409
            End If
        End With
    End If
End Sub


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

' SUM K'ya elle yazýlan baþlýk adlarýný (dosya adý -> ad) kalýcý saklar: çok gizli KISA_ADLAR sayfasý.
' Her çalýþtýrmanýn BAÞINDA (temizlikten önce) çaðrýlýr. K boþaltýlan dosyanýn kaydý silinir.
' Sayfa ilk kez oluþturuluyorsa K'daki deðerler eski sürümün otomatik kýsaltmalarýdýr: alýnmaz.
Public Sub CaptureShortNames()
    Dim ws As Worksheet, st As Worksheet, r As Long, lastR As Long, nm As String, kv As String
    Dim firstTime As Boolean, k As Variant
    On Error Resume Next
    Set dictShortNames = CreateObject("Scripting.Dictionary")
    dictShortNames.CompareMode = 1
    Set st = ThisWorkbook.Sheets("KISA_ADLAR")
    If st Is Nothing Then
        firstTime = True
        Set st = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        If st Is Nothing Then Exit Sub
        st.Name = "KISA_ADLAR"
        st.Visible = xlSheetVeryHidden
    End If
    lastR = st.Cells(st.Rows.Count, 1).End(xlUp).Row
    For r = 1 To lastR
        nm = Trim$(st.Cells(r, 1).Text)
        If nm <> "" Then dictShortNames(nm) = CStr(st.Cells(r, 2).Value)
    Next r
    If Not firstTime Then
        For Each ws In ThisWorkbook.Worksheets
            If UCase$(ws.Name) = "SUM" Or UCase$(ws.Name) Like "SUM_P#*" Then
                lastR = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
                For r = 5 To lastR
                    nm = Trim$(ws.Cells(r, 1).Text)
                    kv = Trim$(ws.Cells(r, 11).Text)
                    If nm <> "" Then
                        If kv <> "" And kv <> nm Then
                            dictShortNames(nm) = kv
                        ElseIf dictShortNames.Exists(nm) Then
                            dictShortNames.Remove nm
                        End If
                    End If
                Next r
            End If
        Next ws
    End If
    st.Cells.ClearContents
    r = 0
    For Each k In dictShortNames.Keys
        r = r + 1
        st.Cells(r, 1).Value = "'" & CStr(k)
        st.Cells(r, 2).Value = "'" & dictShortNames(k)
    Next k
End Sub

' =========================================================================
' ÞABLON TOPLAM SATIRI ve KAPASÝTE
' ANGLE / PLATE / BOLTS'ta veri satýrlarýnýn altýnda her dosyanýn toplam aðýrlýk formülleri olan bir satýr
' vardýr (ör. 1504); SUM C / D / E bu satýrý gösterir (=ANGLE!I1504). Veri bu satýra yazýlýrsa formüller
' silinir ve SUM aðýrlýklarý 0 olur. Toplam satýrý SUM'daki formülden bulunur; bulunamazsa ilk adet
' sütununda formül olan ilk satýr alýnýr. 0 = bulunamadý (sýnýr yok).
' =========================================================================
Public Function TemplateTotalsRow(ByVal ws As Worksheet, ByVal wsSum As Worksheet, ByVal sumCol As Long, ByVal qtyCol As Long) As Long
    Dim f As String, rx As Object, mc As Object, r As Long, lastR As Long
    On Error Resume Next
    If ws Is Nothing Then Exit Function
    If Not wsSum Is Nothing Then
        f = CStr(wsSum.Cells(5, sumCol).Formula)
        Set rx = CreateObject("VBScript.RegExp")
        rx.IgnoreCase = True
        rx.Pattern = "'?" & RxEscapeText(ws.Name) & "'?!\$?[A-Z]{1,3}\$?(\d+)"
        Set mc = rx.Execute(f)
        If Not mc Is Nothing Then
            If mc.Count > 0 Then
                r = CLng(mc(0).SubMatches(0))
                If r > 5 Then
                    TemplateTotalsRow = r
                    Exit Function
                End If
            End If
        End If
    End If
    lastR = ws.Cells(ws.Rows.Count, qtyCol).End(xlUp).Row
    If lastR > 6000 Then lastR = 6000
    For r = 5 To lastR
        If ws.Cells(r, qtyCol).HasFormula Then
            TemplateTotalsRow = r
            Exit Function
        End If
    Next r
End Function

Private Function RxEscapeText(ByVal s As String) As String
    Dim i As Long, ch As String, out As String
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If InStr(".^$*+?()[]{}|\", ch) > 0 Then out = out & "\"
        out = out & ch
    Next i
    RxEscapeText = out
End Function

' Veri alanýndaki son dolu satýr (toplam satýrý ve altý hariç); veri yoksa 4
Public Function LastDataRow(ByVal ws As Worksheet, ByVal keyCol As Long, ByVal totRow As Long) As Long
    Dim r As Long
    On Error Resume Next
    If totRow > 5 Then
        If Trim$(ws.Cells(totRow - 1, keyCol).Text) <> "" Then
            r = totRow - 1
        Else
            r = ws.Cells(totRow - 1, keyCol).End(xlUp).Row
        End If
    Else
        r = ws.Cells(ws.Rows.Count, keyCol).End(xlUp).Row
    End If
    If r < 5 Then r = 4
    LastDataRow = r
End Function

' Toplam satýrý onarýmý: eski sürüm (Liste Üstüne Ekle) toplam satýrýný veri sanýp üzerine deðer yazýyordu.
' Ayný satýrda formülü saðlam kalan bir adet sütunu (henüz dosya yazýlmamýþ saðdaki sütunlar) varsa onun formülü
' (R1C1, sütuna göre göreli) formülü silinmiþ hücrelere yazýlýr. Onarýlamazsa uyarý metni döner.
Public Function TotalsRowDamage(ByVal ws As Worksheet, ByVal totRow As Long, ByVal qFirst As Long, ByVal qLast As Long) As String
    Dim c As Long, pat As String, n As Long
    On Error Resume Next
    If ws Is Nothing Or totRow <= 5 Then Exit Function
    For c = qLast To qFirst Step -1
        If ws.Cells(totRow, c).HasFormula Then
            pat = ws.Cells(totRow, c).FormulaR1C1
            Exit For
        End If
    Next c
    If pat = "" Then
        TotalsRowDamage = "[" & ws.Name & "] " & totRow & ". satýrdaki TOPLAM formülleri silinmiþ ve onarýlamadý (SUM aðýrlýklarý 0 çýkar). " & _
                          "Bu sayfayý temiz þablondan (BOS BOM) yeniden alýn." & vbCrLf
        Exit Function
    End If
    For c = qFirst To qLast
        If Not ws.Cells(totRow, c).HasFormula Then
            ws.Cells(totRow, c).FormulaR1C1 = pat
            n = n + 1
        End If
    Next c
    If n > 0 Then
        AuditRecord "", "SYSTEM", totRow, ws.Name, "Toplam satýrý onarýldý", "SYSTEM", "WARNING", 90, _
                    n & " hücreye toplam formülü geri yazýldý (" & pat & ")", _
                    "Önceki bir çalýþtýrma toplam satýrýnýn üzerine yazmýþtý; ayný satýrdaki saðlam formül kopyalandý. Aðýrlýklarý kontrol edin."
    End If
End Function


' Aralýktaki SABÝT deðerleri siler, formüllere dokunmaz. Excel'in SpecialCells'i 8192'den fazla ayrýk alan
' olunca HATA verir (daðýnýk adet sütunlarýnda olur); eski sürümde bu hata yutulduðu için temizlik HÝÇ
' yapýlmýyordu ve önceki çalýþtýrmanýn adetleri kalýp yeni adetlere ekleniyordu. Hata olursa sütun sütun silinir.
Public Sub ClearConstants(ByVal rng As Range)
    Dim i As Long, c As Range
    If rng Is Nothing Then Exit Sub
    On Error Resume Next
    Err.Clear
    Set c = rng.SpecialCells(xlCellTypeConstants)
    If Err.Number = 0 Then
        If Not c Is Nothing Then c.ClearContents
        If Err.Number = 0 Then Exit Sub
    End If
    ' ("hücre bulunamadý" da buraya düþer; sütun sütun denemek zararsýz)
    For i = 1 To rng.Columns.Count
        Err.Clear
        Set c = Nothing
        Set c = rng.Columns(i).SpecialCells(xlCellTypeConstants)
        If Err.Number = 0 Then
            If Not c Is Nothing Then c.ClearContents
        End If
    Next i
    Err.Clear
End Sub

' Veri alanýndaki formül sütunlarýný (ilk veri satýrlarýnda formül olanlar: kg/m, birim aðýrlýk, toplam...)
' toplam satýrýnýn bir üstüne kadar doldurur. Þablonda formül erken bitmiþse (ör. PLATE G 979'da) o satýrlarýn
' aðýrlýðý hesaplanmýyordu. Sadece BOÞ hücrelere yazýlýr.
' skipCols: makronun kendi yazdýðý sütunlar ("|1|2|3|") ve adet sütunlarý (qFirst..qLast) ASLA doldurulmaz
Public Sub FillFormulaColumns(ByVal ws As Worksheet, ByVal totRow As Long, ByVal skipCols As String, _
                              ByVal qFirst As Long, ByVal qLast As Long)
    Dim c As Long, lastC As Long, r0 As Long, pat As String, blanks As Range, n As Long
    On Error Resume Next
    If ws Is Nothing Or totRow <= 9 Then Exit Sub
    lastC = ws.Cells(4, ws.Columns.Count).End(xlToLeft).Column
    If lastC < 60 Then lastC = 60
    For c = 1 To lastC
        If (c >= qFirst And c <= qLast) Or InStr(skipCols, "|" & c & "|") > 0 Then GoTo NextCol
        pat = ""
        For r0 = 5 To 8
            If ws.Cells(r0, c).HasFormula Then
                pat = ws.Cells(r0, c).FormulaR1C1
                Exit For
            End If
        Next r0
        If pat <> "" Then
            Err.Clear
            Set blanks = Nothing
            Set blanks = ws.Range(ws.Cells(r0 + 1, c), ws.Cells(totRow - 1, c)).SpecialCells(xlCellTypeBlanks)
            If Err.Number = 0 Then
                If Not blanks Is Nothing Then
                    n = n + blanks.Count
                    blanks.FormulaR1C1 = pat
                End If
            End If
            Err.Clear
        End If
NextCol:
    Next c
    If n > 0 Then
        AuditRecord "", "SYSTEM", 0, ws.Name, "Þablon formülleri tamamlandý", "SYSTEM", "WARNING", 90, _
                    n & " boþ hücreye formül yazýldý (toplam satýrý " & totRow & ")", _
                    "Þablonda formüller veri alanýnýn sonuna kadar gitmiyordu; o satýrlarýn aðýrlýðý hesaplanmýyordu."
    End If
End Sub

' VERÝ SINIRI: adet sütununda (veri satýrlarýndan sonra) ilk formül olan satýr. Ýki bloklu þablonda (ANGLE:
' 5-1503 adet, 1505-3003 satýr aðýrlýklarý, 3004 toplam) 1504/1505'tir; tek bloklu þablonda toplam satýrý.
' Formül bulunamazsa (blok silinmiþ olabilir) ve toplam satýrý 1504'ten aþaðýdaysa þablonun özgün 1503 satýr
' tasarýmý korunur. 0 = sýnýr yok.
Public Function TemplateDataLimit(ByVal ws As Worksheet, ByVal totRow As Long, ByVal qtyCol As Long) As Long
    Dim scanEnd As Long, f As Range, ar As Range, lim As Long
    On Error Resume Next
    If ws Is Nothing Then Exit Function
    lim = totRow
    scanEnd = totRow - 1
    If totRow <= 5 Then scanEnd = 6000
    If scanEnd >= 5 Then
        Set f = ws.Range(ws.Cells(5, qtyCol), ws.Cells(scanEnd, qtyCol)).SpecialCells(xlCellTypeFormulas)
        If Not f Is Nothing Then
            For Each ar In f.Areas
                If lim = 0 Or ar.Row < lim Then lim = ar.Row
            Next ar
        End If
    End If
    Err.Clear
    If lim > 1504 And (lim = totRow Or totRow <= 5) Then lim = 1504
    If lim <= 5 Then lim = 0
    TemplateDataLimit = lim
End Function

' Ýki bloklu þablonda satýr aðýrlýðý bloðu (veri sýnýrýnýn altý ile toplam satýrý arasý, adet sütunlarý) saðlam mý?
' Eski sürümler bu bloða veri / deðer yazýp formülleri silmiþti. Bloðun formülü her satýrda aynýdýr (R1C1):
' saðlam bir sütundan alýnýr, formülü eksik sütunlara yazýlýr. Hiç saðlam sütun yoksa uyarý metni döner.
Public Function RepairWeightBlock(ByVal ws As Worksheet, ByVal limRow As Long, ByVal totRow As Long, _
                                  ByVal qFirst As Long, ByVal qLast As Long) As String
    Dim c As Long, r1 As Long, r2 As Long, pat As String, f As Range, nFix As Long, nRows As Long
    On Error Resume Next
    If ws Is Nothing Or limRow <= 5 Or totRow <= 5 Then Exit Function
    r1 = limRow + 1                                   ' veri sýnýrý satýrý (1504) ara toplam olabilir; blok altý
    r2 = totRow - 1
    If r2 - r1 < 2 Then Exit Function                 ' tek bloklu þablon: aðýrlýk bloðu yok
    nRows = r2 - r1 + 1
    For c = qLast To qFirst Step -1
        If ws.Cells(r1, c).HasFormula And ws.Cells(r2, c).HasFormula Then
            If ws.Cells(r1, c).FormulaR1C1 = ws.Cells(r2, c).FormulaR1C1 Then
                pat = ws.Cells(r1, c).FormulaR1C1
                Exit For
            End If
        End If
    Next c
    If pat = "" Then
        RepairWeightBlock = "[" & ws.Name & "] " & r1 & "-" & r2 & ". satýrlardaki aðýrlýk formülleri silinmiþ (aðýrlýklar 0 çýkar). " & _
                            "Alt+F8 > Yedekten_Geri_Yukle ile iþlem öncesi yedeðe dönün ya da temiz þablon kullanýn." & vbCrLf
        Exit Function
    End If
    For c = qFirst To qLast
        Err.Clear
        Set f = Nothing
        Set f = ws.Range(ws.Cells(r1, c), ws.Cells(r2, c)).SpecialCells(xlCellTypeFormulas)
        If f Is Nothing Or Err.Number <> 0 Then
            ws.Range(ws.Cells(r1, c), ws.Cells(r2, c)).FormulaR1C1 = pat
            nFix = nFix + 1
        ElseIf f.Count <> nRows Then
            ws.Range(ws.Cells(r1, c), ws.Cells(r2, c)).FormulaR1C1 = pat
            nFix = nFix + 1
        End If
    Next c
    Err.Clear
    If nFix > 0 Then
        AuditRecord "", "SYSTEM", r1, ws.Name, "Aðýrlýk bloðu onarýldý", "SYSTEM", "WARNING", 90, _
                    nFix & " sütunda " & r1 & "-" & r2 & ". satýrlara formül geri yazýldý (" & pat & ")", _
                    "Eski bir sürüm bu satýrlara veri yazýp formülleri silmiþti. Aðýrlýklarý kontrol edin."
    End If
End Function
