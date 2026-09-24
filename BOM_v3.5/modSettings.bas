Attribute VB_Name = "modSettings"
Option Explicit
Option Private Module

' =========================================================================
' AYARLAR: anahtar kelimeler ve parametreler
' =========================================================================

Public Sub CheckAndInitSettingsSheet()
    Dim wsSet As Worksheet, hwDefs As Variant, gbDefs As Variant, plDefs As Variant, anDefs As Variant, i As Long
    On Error Resume Next
    Set wsSet = ThisWorkbook.Sheets("AYARLAR")
    On Error GoTo 0
    If wsSet Is Nothing Then
        Set wsSet = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsSet.Name = "AYARLAR"
        wsSet.Cells(1, 1).Value = "Hýrdavat Kelimeleri"
        wsSet.Cells(1, 2).Value = "Çöp Kelimeler"
        wsSet.Cells(1, 3).Value = "Plaka Kelimeleri"
        wsSet.Cells(1, 4).Value = "Profil Kelimeleri"
        wsSet.Range("A1:D1").Font.Bold = True
        wsSet.Range("A1:D1").Interior.Color = RGB(0, 51, 102)
        wsSet.Range("A1:D1").Font.Color = RGB(255, 255, 255)
        hwDefs = Array("SCHRAUBE", "SCHEIBE", "FEDERRING", "MUTTER", "BOLT", "WASHER", "NUT", "FUTTERRING", "FUTTER", "STEIGBOLZEN", "FU-RING", "RING", "STB", "TB.M", "-STB")
        gbDefs = Array("EINZEL-", "GESAMT-", "FERTIGUNG", "WERKSTATTZEICHNUNG", "GEWICHT", "OBERFLACHE", "OBERFLÄCHE", "NAME", "DATUM", "SUMME", "STAHLGEWICHT", "ZINKAUFLAGE", "SCHRAUBEN", "FEUERVERZINKT", "BENENNUNG", "MONTAGESCHRAUBEN", "GEGENSTAND", "BEZEICHNUNG", "VERBINDUNGSMATERIAL")
        plDefs = Array("FL", "BL", "BLE", "PL", "FLAT", "KF", "PLT")
        anDefs = Array("UNP", "UPE", "HEA", "HEB", "HEM", "IPE", "IPN", "PFC", "UB", "UC", "INP", "UPN", "HD", "HL", "HP")
        For i = LBound(hwDefs) To UBound(hwDefs)
            wsSet.Cells(i + 2, 1).Value = hwDefs(i)
        Next i
        For i = LBound(gbDefs) To UBound(gbDefs)
            wsSet.Cells(i + 2, 2).Value = gbDefs(i)
        Next i
        For i = LBound(plDefs) To UBound(plDefs)
            wsSet.Cells(i + 2, 3).Value = plDefs(i)
        Next i
        For i = LBound(anDefs) To UBound(anDefs)
            wsSet.Cells(i + 2, 4).Value = anDefs(i)
        Next i
        wsSet.Columns("A:D").AutoFit
End If
End Sub

Public Sub LoadDictionaryArrays()
    Dim wsSet As Worksheet, r As Long, i As Long
    On Error Resume Next
    Set wsSet = ThisWorkbook.Sheets("AYARLAR")
    On Error GoTo 0
    If wsSet Is Nothing Then Exit Sub
    
    r = wsSet.Cells(wsSet.Rows.Count, 1).End(xlUp).Row
    If r >= 2 Then
        ReDim hwKeywords(1 To r - 1)
        For i = 2 To r
            hwKeywords(i - 1) = NormalizeBOMText(wsSet.Cells(i, 1).Text)
        Next i
Else
        ReDim hwKeywords(1 To 1)
        hwKeywords(1) = "SCHRAUBE"
    End If
    
    r = wsSet.Cells(wsSet.Rows.Count, 2).End(xlUp).Row
    If r >= 2 Then
        ReDim garbageKeywords(1 To r - 1)
        For i = 2 To r
            garbageKeywords(i - 1) = NormalizeBOMText(wsSet.Cells(i, 2).Text)
        Next i
Else
        ReDim garbageKeywords(1 To 1)
        garbageKeywords(1) = "GEWICHT"
    End If
    
    r = wsSet.Cells(wsSet.Rows.Count, 3).End(xlUp).Row
    If r >= 2 Then
        ReDim plateKeywords(1 To r - 1)
        For i = 2 To r
            plateKeywords(i - 1) = NormalizeBOMText(wsSet.Cells(i, 3).Text)
        Next i
Else
        ReDim plateKeywords(1 To 1)
        plateKeywords(1) = "PL"
    End If
    
    r = wsSet.Cells(wsSet.Rows.Count, 4).End(xlUp).Row
    If r >= 2 Then
        ReDim angleKeywords(1 To r - 1)
        For i = 2 To r
            angleKeywords(i - 1) = NormalizeBOMText(wsSet.Cells(i, 4).Text)
        Next i
Else
        ReDim angleKeywords(1 To 1)
        angleKeywords(1) = "UNP"
    End If
End Sub

' =========================================================================
' v2.3 - PARAMETRELER (AYARLAR!F:H)
' F sütunundaki etiketleri deðiþtirmeyin; G sütunundaki deðerleri deðiþtirin.
' =========================================================================
Public Sub LoadParameters()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("AYARLAR")
    On Error GoTo 0

    ' Varsayýlanlar (sayfa yoksa / deðer bozuksa bunlar kullanýlýr)
    prmStepBoltAutoSet = False
    prmWeightTolPct = 2
    prmQualH = "S355"
    prmQualDefault = "S275"
    prmBoltGrade = "8.8"
    prmBoltLen = 40
    prmBackup = True
    prmBackupFolder = ""
    prmBackupKeep = 20
    prmRevAuto = True
    If ws Is Nothing Then Exit Sub

    Call EnsureParameterBlock(ws)

    prmStepBoltAutoSet = ParseYesNo(GetParam(ws, "Basamak cývatasý için otomatik set", "HAYIR"), False)
    prmWeightTolPct = ParseBOMNumber(GetParam(ws, "Aðýrlýk fark toleransý (%)", "2"))
    If prmWeightTolPct <= 0 Then prmWeightTolPct = 2
    prmQualH = UCase$(Trim$(GetParam(ws, "Kalite yoksa (H iþaretli parça)", "S355")))
    If prmQualH = "" Then prmQualH = "S355"
    prmQualDefault = UCase$(Trim$(GetParam(ws, "Kalite yoksa (diðer)", "S275")))
    If prmQualDefault = "" Then prmQualDefault = "S275"
    prmBoltGrade = Replace(Trim$(GetParam(ws, "Cývata kalitesi yazmýyorsa", "8.8")), ",", ".")
    If Not IsValidBoltGrade(prmBoltGrade) Then prmBoltGrade = "8.8"
    prmBoltLen = CLng(ParseBOMNumber(GetParam(ws, "Cývata boyu yazmýyorsa (mm)", "40")))
    If prmBoltLen <= 0 Then prmBoltLen = 40
    prmBackup = ParseYesNo(GetParam(ws, "Otomatik yedek al", "EVET"), True)
    prmBackupFolder = Trim$(GetParam(ws, "Yedek klasörü (boþ = otomatik)", ""))
    prmBackupKeep = CLng(ParseBOMNumber(GetParam(ws, "Saklanacak yedek sayýsý", "20")))
    If prmBackupKeep <= 0 Then prmBackupKeep = 20
    prmRevAuto = ParseYesNo(GetParam(ws, "Üstüne eklemede REV otomatik artsýn", "EVET"), True)
End Sub

Public Sub EnsureParameterBlock(ByVal ws As Worksheet)
    Dim labels As Variant, defs As Variant, notes As Variant, i As Long
    labels = Array("Basamak cývatasý için otomatik set", "Aðýrlýk fark toleransý (%)", _
                   "Kalite yoksa (H iþaretli parça)", "Kalite yoksa (diðer)", _
                   "Cývata kalitesi yazmýyorsa", "Cývata boyu yazmýyorsa (mm)", _
                   "Otomatik yedek al", "Yedek klasörü (boþ = otomatik)", _
                   "Saklanacak yedek sayýsý", "Üstüne eklemede REV otomatik artsýn")
    defs = Array("HAYIR", "2", "S355", "S275", "8.8", "40", "EVET", "", "20", "EVET")
    notes = Array("EVET: STGB/STB için somun+pul+yaylý rondela üretilir. HAYIR: özel bölüme yazýlýr.", _
                  "SUM!J: müþteri aðýrlýðý (H) ile þablon aðýrlýðý (C+D) farký bu yüzdeyi geçerse kýrmýzý.", _
                  "Malzeme kalitesi boþ ve isimde/pozda H varsa.", _
                  "Malzeme kalitesi boþsa.", _
                  "4.6 / 5.6 / 8.8 / 10.9", _
                  "Sadece çap yazýlmýþ cývatalar (örn: M16) için.", _
                  "EVET: her çalýþtýrmadan önce dosyanýn tarihli kopyasý alýnýr.", _
                  "Boþsa: dosyanýn yanýnda YEDEK klasörü (OneDrive ise Belgeler\\BOM_YEDEK).", _
                  "Bu sayýdan eski yedekler silinir.", _
                  "EVET: 'Liste üstüne ekle' her çalýþtýðýnda REV +1. Temizle düðmesi REV'i 0 yapar.")

    If UCase$(Trim$(ws.Range("F1").Text)) = "PARAMETRE" Then
        ' Blok var: eksik satýr varsa sona ekle (yeni sürümle gelen parametreler)
        For i = LBound(labels) To UBound(labels)
            If FindParamRow(ws, CStr(labels(i))) = 0 Then
                Dim nr As Long
                nr = ws.Cells(ws.Rows.Count, 6).End(xlUp).Row + 1
                ws.Cells(nr, 6).Value = labels(i)
                ws.Cells(nr, 7).NumberFormat = "@"
                ws.Cells(nr, 7).Value = defs(i)
                ws.Cells(nr, 8).Value = notes(i)
                ws.Cells(nr, 7).Interior.Color = RGB(255, 242, 204)
            End If
        Next i
        Exit Sub
    End If

    ws.Range("F1").Value = "PARAMETRE"
    ws.Range("G1").Value = "DEÐER"
    ws.Range("H1").Value = "AÇIKLAMA"
    With ws.Range("F1:H1")
        .Font.Bold = True
        .Interior.Color = RGB(0, 51, 102)
        .Font.Color = RGB(255, 255, 255)
    End With
    ws.Range("G2").Resize(UBound(labels) + 1, 1).NumberFormat = "@"   ' "8.8" tarihe dönmesin
    For i = LBound(labels) To UBound(labels)
        ws.Cells(i + 2, 6).Value = labels(i)
        ws.Cells(i + 2, 7).Value = defs(i)
        ws.Cells(i + 2, 8).Value = notes(i)
    Next i
    ws.Range("G2").Resize(UBound(labels) + 1, 1).Interior.Color = RGB(255, 242, 204)
    ws.Range("F1").Resize(UBound(labels) + 2, 3).Borders.LineStyle = 1
    ws.Columns("F").ColumnWidth = 38
    ws.Columns("G").ColumnWidth = 14
    ws.Columns("H").ColumnWidth = 80
End Sub

Public Function FindParamRow(ByVal ws As Worksheet, ByVal label As String) As Long
    Dim r As Long, lastR As Long
    lastR = ws.Cells(ws.Rows.Count, 6).End(xlUp).Row
    For r = 2 To lastR
        If StrComp(Trim$(ws.Cells(r, 6).Text), label, vbTextCompare) = 0 Then
            FindParamRow = r
            Exit Function
        End If
    Next r
End Function

Public Function GetParam(ByVal ws As Worksheet, ByVal label As String, ByVal defVal As String) As String
    Dim r As Long
    r = FindParamRow(ws, label)
    If r = 0 Then
        GetParam = defVal
    Else
        GetParam = Trim$(ws.Cells(r, 7).Text)
    End If
End Function

Public Function ParseYesNo(ByVal txt As String, ByVal defVal As Boolean) As Boolean
    Select Case FoldText(txt)
        Case "EVET", "E", "YES", "Y", "TRUE", "DOGRU", "1", "VAR", "ACIK"
            ParseYesNo = True
        Case "HAYIR", "H", "NO", "N", "FALSE", "YANLIS", "0", "YOK", "KAPALI"
            ParseYesNo = False
        Case Else
            ParseYesNo = defVal
    End Select
End Function

' Sayýsal kodlar sayý olarak yazýlýr (þablondaki VLOOKUP eþleþsin), diðerleri metin
Public Function LibraryCodeValue(ByVal code As String) As Variant
    If code Like "*[!0-9]*" Or Left$(code, 1) = "0" Then
        LibraryCodeValue = code
    Else
        LibraryCodeValue = CDbl(code)
    End If
End Function
