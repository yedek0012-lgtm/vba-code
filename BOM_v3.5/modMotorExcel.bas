Attribute VB_Name = "modMotorExcel"
Option Explicit
Option Private Module

' =========================================================================
' EXCEL / ERP motoru: baþlýk algýlama, sütun eþleme, satýr sýnýflandýrma
' =========================================================================

Public Sub ProcessExcelFile()
    ' Excel / ERP malzeme listesi. Sütun rolleri (mbCols): 1=POS 2=ADET 3=ÝSÝM 4=BOY 5=MALZEME
    ' 6=NOT 7=DIN 8=TOPLAM AÐIRLIK 9=KALINLIK 10=GENÝÞLÝK 11=AÐIRLIK(genel) 12=BÝRÝM AÐIRLIK 13=ÖLÇÜ
    Dim aQtyE As Double, boltKeyE As String, boltNoteE As String, boltQualE As String
    Dim cleanNoSpace As String, dinVal As String, dValE As Long, existRow As Long
    Dim finalBoltQty As Long, finalPlateNote As String, hRow As Long, hwCodeE As String
    Dim hwNameE As String, isFuRingE As Boolean, lastColExt As Long, lastRowExt As Long
    Dim lengthVal As Double, lValE As Long, matCodeStrE As String, maxC As Long, maxR As Long
    Dim notStr As String, parts() As String, pNoteE As String, posKeyAngleE As String
    Dim posKeyPlateE As String, posNo As String, pQtyE As Double, pThickE As Double
    Dim pWidthE As Double, qtyKeyE As String, quality As String, quantVal As Long, r As Long
    Dim rawName As String, rawNameUpper As String, remVal As String, scanC As Long
    Dim srcData As Variant, wbIn As Workbook, wbOpenAlready As Boolean, wsIn As Worksheet

    Dim mbCols(1 To 16) As Long
    Dim mbMyMark As String
    Dim mbDisc As Boolean, mbDia2 As Double
    Dim mbQtyD As Double
    Dim mbKind As String, mbDinRaw As String, mbTmp As String
    Dim mbRowWt As Double, mbAngPlateWt As Double, mbFooterWt As Double
    Dim mbHwCust As Object, mbBoltDia As Object, mbHwKey As Variant
    Dim mbDia As Long, mbCustQty As Long, mbOurQty As Long
    Dim mbIssue As String, mbStatus As String, mbConf As Long
    Dim mbSheetsUsed As Long
    Dim mbLenMul As Double
    Dim mbLrn As Variant, mbLrnKind As String, mbLrnCode As String, mbLrnName As String, mbLrnT As Double, mbLrnW As Double

    Set wbIn = Nothing
    wbOpenAlready = False
    mbMyMark = FileMark(fileName)

    If Not pdfTempWb Is Nothing Then
        Set wbIn = pdfTempWb                 ' PDF motorunun oluþturduðu geçici çalýþma kitabý
    Else
        On Error Resume Next
        Set wbIn = Application.Workbooks(fso.GetFileName(vItem))
        On Error GoTo 0
    End If

    If Not wbIn Is Nothing Then
        wbOpenAlready = True
    Else
        On Error Resume Next
        If ext = "csv" Then
            Set wbIn = Application.Workbooks.Open(fileName:=vItem, ReadOnly:=True, Local:=True)
        Else
            Set wbIn = Application.Workbooks.Open(fileName:=vItem, ReadOnly:=True, UpdateLinks:=False, AddToMru:=False)
        End If
        On Error GoTo 0
    End If

    If wbIn Is Nothing Then
        AuditRecord fileName, "EXCEL", 0, "FILE", CStr(vItem), "FILE", "ERROR", 0, "", "Dosya açýlamadý (þifreli / bozuk / kilitli olabilir)."
    Else
        Set mbHwCust = CreateObject("Scripting.Dictionary")
        Set mbBoltDia = CreateObject("Scripting.Dictionary")
        mbHwCust.CompareMode = 1
        mbBoltDia.CompareMode = 1
        mbAngPlateWt = 0
        mbFooterWt = 0
        mbSheetsUsed = 0

        For Each wsIn In wbIn.Worksheets
            If pdfTempWb Is Nothing Then auditCurrentSheet = wsIn.Name Else auditCurrentSheet = ""
            mbTmp = FoldText(wsIn.Name)
            If InStr(mbTmp, "ZUSAMMEN") > 0 Or InStr(mbTmp, "SUMMARY") > 0 Then GoTo NextSheetExcel

            lastRowExt = 0
            lastColExt = 0
            On Error Resume Next
            lastRowExt = wsIn.Cells.Find(What:="*", After:=wsIn.Cells(1, 1), LookIn:=xlFormulas, SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
            lastColExt = wsIn.Cells.Find(What:="*", After:=wsIn.Cells(1, 1), LookIn:=xlFormulas, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
            On Error GoTo 0
            If lastRowExt < 2 Or lastColExt < 2 Then GoTo NextSheetExcel

            srcData = wsIn.Range(wsIn.Cells(1, 1), wsIn.Cells(lastRowExt, lastColExt)).Value2
            If Not IsArray(srcData) Then GoTo NextSheetExcel
            maxR = UBound(srcData, 1)
            maxC = UBound(srcData, 2)

            ' 1) Baþlýk satýrýný HÜCRE BAZLI bul (antet bloðundaki "Benennung:" gibi etiketler sayýlmaz)
            hRow = 0
            For r = 1 To MinLong(maxR, 80)
                If IsBOMHeaderRow(srcData, r, maxC) Then
                    hRow = r
                    Exit For
                End If
            Next r

            If hRow = 0 Then
                AuditRecord fileName, "EXCEL", 0, "HEADER", "Sayfa: " & wsIn.Name, "SHEET", "WARNING", 0, "", _
                            "Baþlýk satýrý bulunamadý (Poz / Adet / Benennung). Sayfa atlandý."
                GoTo NextSheetExcel
            End If

            Call MapBOMHeader(srcData, hRow, maxR, maxC, mbCols)
            mbLenMul = LengthUnitFactor(srcData, hRow, mbCols(4), maxR, maxC)
            If mbLenMul <> 1 Then AuditRecord fileName, "EXCEL", hRow, "HEADER", "Sayfa: " & wsIn.Name, "SHEET", "WARNING", 90, _
                        "Boy birimi mm deðil", "Boylar x" & mbLenMul & " ile mm'ye çevrildi (þablon F sütunu mm)."
            mbSheetsUsed = mbSheetsUsed + 1
            AuditRecord fileName, "EXCEL", hRow, "HEADER", "Sayfa: " & wsIn.Name, "SHEET", "EXACT", 100, _
                        DescribeBOMColumns(mbCols), ""

            ' 2) Veri satýrlarý
            For r = hRow + 1 To maxR
                ' Sayfa içinde tekrar eden baþlýk (yeni sayfa / yeni blok) -> sütunlarý yeniden eþle
                If HeaderRoleOf(BomCellText(srcData, r, mbCols(2), maxC)) = 2 Or _
                   HeaderRoleOf(BomCellText(srcData, r, mbCols(1), maxC)) = 1 Or _
                   HeaderRoleOf(BomCellText(srcData, r, mbCols(3), maxC)) = 3 Then
                    If IsBOMHeaderRow(srcData, r, maxC) Then
                        Call MapBOMHeader(srcData, r, maxR, maxC, mbCols)
                        mbLenMul = LengthUnitFactor(srcData, r, mbCols(4), maxR, maxC)
                        GoTo NextRowExcel
                    End If
                End If

                rawName = BomCellText(srcData, r, mbCols(3), maxC)

                ' Parçalanmýþ profil adý (ör: "L" | 200 | "x" | 20) - sadece rolü OLMAYAN komþu sütunlardan
                For scanC = mbCols(3) + 1 To mbCols(3) + IIf(mbCols(3) > 0, 4, 0)
                    If scanC > maxC Then Exit For
                    If IsMappedBOMCol(mbCols, scanC) Then Exit For
                    mbTmp = BomCellText(srcData, r, scanC, maxC)
                    If mbTmp = "" Then Exit For
                    If RxTest(mbTmp, "^[\d\.,]+$") Or UCase$(mbTmp) = "X" Or Left$(mbTmp, 1) = "+" Then
                        rawName = rawName & " " & mbTmp
                    Else
                        Exit For
                    End If
                Next scanC

                ' Ayrý "Abmessung/Ölçü" sütunu varsa birleþtir; "Winkel 150x12" -> "L 150x12"
                rawName = BuildItemName(rawName, BomCellText(srcData, r, mbCols(13), maxC))
                rawName = StripRevisionMarks(rawName)
                If rawName = "" Then GoTo NextRowExcel

                posNo = BomCellText(srcData, r, mbCols(1), maxC)
                quality = BomCellText(srcData, r, mbCols(5), maxC)
                lengthVal = BomCellNum(srcData, r, mbCols(4), maxC) * mbLenMul
                mbRowWt = BomCellNum(srcData, r, mbCols(8), maxC)
                mbDinRaw = BomCellText(srcData, r, mbCols(7), maxC)
                remVal = BomCellText(srcData, r, mbCols(6), maxC)

                quantVal = 0
                mbQtyD = BomCellNum(srcData, r, mbCols(2), maxC)
                If mbQtyD > 0 And mbQtyD < 1000000 Then quantVal = CLng(mbQtyD)
                ' Toplam aðýrlýk hücresi boþsa: adet x parça aðýrlýðý (PDF listelerinde satýr kayabiliyor)
                If mbRowWt <= 0 And mbCols(15) > 0 And quantVal > 0 Then mbRowWt = quantVal * BomCellNum(srcData, r, mbCols(15), maxC)

                ' Notlar (eski kuralla ayný: DIN geçen not alaný yazýlmaz)
                dinVal = mbDinRaw
                If UCase$(dinVal) Like "*DIN*" Then dinVal = ""
                If UCase$(remVal) Like "*BEMERKUNG*" Then remVal = ""
                notStr = ""
                If dinVal <> "" And remVal <> "" And dinVal <> remVal Then
                    notStr = dinVal & ", " & remVal
                ElseIf remVal <> "" Then
                    notStr = remVal
                ElseIf dinVal <> "" Then
                    notStr = dinVal
                End If

                mbKind = ClassifyBOMItem(rawName, mbDinRaw, quality, posNo)
                ' Öðrenilen parça (OGRENME'de türü seçilip aktarýlan): tür ve deðerler OGRENILEN'den
                mbLrnKind = "": mbLrnCode = "": mbLrnName = "": mbLrnT = 0: mbLrnW = 0
                If Not dictLearned Is Nothing Then
                    mbTmp = LearnedKey(rawName)
                    If dictLearned.Exists(mbTmp) Then
                        mbLrn = dictLearned(mbTmp)
                        mbLrnKind = CStr(mbLrn(0))
                        mbLrnCode = CStr(mbLrn(1))
                        mbLrnName = CStr(mbLrn(2))
                        mbLrnT = CDbl(mbLrn(3))
                        mbLrnW = CDbl(mbLrn(4))
                        Select Case mbLrnKind
                            Case "ANGLE": mbKind = "ANGLE"
                            Case "PLATE": mbKind = "PLATE"
                            Case "BOLT": mbKind = "LBOLT"
                        End Select
                    End If
                End If
                ' Yuvarlak kesit (D460 / Ø20): boy < çap ise DAÝRE PLAKA, deðilse YUVARLAK ÇUBUK (kütüphane Y kodu)
                mbDisc = False
                If mbKind = "ROUND" Then
                    mbDia2 = ParseBOMNumber(Mid$(Replace(Replace(UCase$(rawName), " ", ""), "RD", "D"), 2))
                    If lengthVal > 0 And lengthVal < mbDia2 Then
                        mbKind = "PLATE"
                        mbDisc = True
                    Else
                        mbKind = "ANGLE"
                        rawName = "Y" & NumToText(mbDia2)
                    End If
                End If
                rawNameUpper = UCase$(rawName)

                ' Montaj çift sayým korumasý: bu poz baþka bir dosyanýn kendi listesi ise (ör. FAC51410A) atla
                If (mbKind = "ANGLE" Or mbKind = "PLATE") And posNo <> "" And Not dictFileMarks Is Nothing Then
                    If dictFileMarks.Exists(UCase$(Trim$(posNo))) And UCase$(Trim$(posNo)) <> mbMyMark Then
                        auditSkipped = auditSkipped + 1
                        AuditRecord fileName, "EXCEL", r, wsIn.Name, rawName & " | Poz: " & posNo, mbKind, "SKIPPED", 100, "", _
                                    "Bu montajýn kendi parça listesi de iþleniyor; çift sayýlmamasý için atlandý."
                        GoTo NextRowExcel
                    End If
                End If

                Select Case mbKind

                    Case "EMPTY"
                        ' boþ

                    Case "GARBAGE"
                        ' Toplam satýrý ise aðýrlýðý yedek olarak al
                        If RxTest(FoldText(rawName), "(STAHLGEWICHT|GESAMTGEWICHT|GRAND TOTAL|TOTAL WEIGHT)") Then
                            If mbRowWt > 0 Then mbFooterWt = mbRowWt
                        End If

                    ' ---------------- CIVATA ----------------
                    Case "BOLT", "FURING"
                        If quantVal <= 0 Then
                            AuditRecord fileName, "EXCEL", r, wsIn.Name, rawName, "BOLT", "WARNING", 40, "", "Adet okunamadý, satýr atlandý."
                            GoTo NextRowExcel
                        End If
                        finalBoltQty = IIf(extraPct > 0, quantVal + Int((quantVal * (extraPct / 100#)) + 0.5), quantVal)
                        Call ParseBoltSpecs(rawName, dValE, lValE, hwCodeE, hwNameE, isFuRingE, boltNoteE, quality, boltQualE)
                        If notStr <> "" Then
                            If boltNoteE <> "" Then
                                boltNoteE = boltNoteE & ", " & notStr
                            Else
                                boltNoteE = notStr
                            End If
                        End If

                        If dValE > 0 Then
                            boltKeyE = Format(dValE, "00") & "|" & Format(lValE, "0000") & "|" & hwCodeE & "|" & boltQualE
                            If Not dictBoltList.Exists(boltKeyE) Then
                                dictBoltList.Add boltKeyE, dValE & "^" & lValE & "^" & hwCodeE & "^" & hwNameE & "^" & boltQualE & "^" & CStr(isFuRingE) & "^" & boltNoteE
                            End If
                            qtyKeyE = boltKeyE & "|" & currentColBolt
                            dictBoltQty(qtyKeyE) = IIf(dictBoltQty.Exists(qtyKeyE), dictBoltQty(qtyKeyE) + finalBoltQty, finalBoltQty)
                            If Not dictDiameters.Exists(CStr(dValE)) Then dictDiameters.Add CStr(dValE), dValE
                            If Not isFuRingE Then mbBoltDia(CStr(dValE)) = mbBoltDia(CStr(dValE)) + quantVal
                            totalProcessed = totalProcessed + 1

                            mbStatus = "EXACT": mbConf = 100: mbIssue = ""
                            If InStr(boltNoteE, "varsay") > 0 Or InStr(boltNoteE, "Özel") > 0 Then
                                mbStatus = "WARNING": mbConf = 70: mbIssue = boltNoteE
                            End If
                            AuditRecord fileName, "EXCEL", r, wsIn.Name, rawName, IIf(isFuRingE, "BOLT-SPECIAL", "BOLT"), mbStatus, mbConf, _
                                        "KOD=" & hwCodeE & "; " & hwNameE & "; KALITE=" & boltQualE & "; ADET=" & quantVal, mbIssue
                        Else
                            AuditRecord fileName, "EXCEL", r, wsIn.Name, rawName, "BOLT", "ERROR", 20, "", "Cývata çapý/boyu çözülemedi."
                            If Not dictFeedback.Exists(rawNameUpper) Then dictFeedback.Add rawNameUpper, fileName & " (Satýr: " & r & ")"
                        End If

                    ' ---------- ÖÐRENÝLEN CIVATA / PUL / BAÐLANTI PARÇASI ----------
                    ' Kod, cins ve aðýrlýk BOLT-LIBRARY'den; özel bölüme yazýlýr (otomatik somun/pul seti üretilmez)
                    Case "LBOLT"
                        If quantVal <= 0 Then
                            AuditRecord fileName, "EXCEL", r, wsIn.Name, rawName, "BOLT", "WARNING", 40, "", "Adet okunamadý, satýr atlandý."
                            GoTo NextRowExcel
                        End If
                        finalBoltQty = IIf(extraPct > 0, quantVal + Int((quantVal * (extraPct / 100#)) + 0.5), quantVal)
                        hwCodeE = mbLrnCode
                        hwNameE = IIf(mbLrnName <> "", mbLrnName, rawName)
                        boltKeyE = "99|0000|" & hwCodeE & "|" & quality
                        If Not dictBoltList.Exists(boltKeyE) Then
                            dictBoltList.Add boltKeyE, "99^0^" & hwCodeE & "^" & hwNameE & "^" & quality & "^True^" & notStr
                        End If
                        qtyKeyE = boltKeyE & "|" & currentColBolt
                        dictBoltQty(qtyKeyE) = IIf(dictBoltQty.Exists(qtyKeyE), dictBoltQty(qtyKeyE) + finalBoltQty, finalBoltQty)
                        If Not dictDiameters.Exists("99") Then dictDiameters.Add "99", 99
                        totalProcessed = totalProcessed + 1
                        AuditRecord fileName, "EXCEL", r, wsIn.Name, rawName, "BOLT-SPECIAL", "EXACT", 100, _
                                    "KOD=" & hwCodeE & "; " & hwNameE & "; ADET=" & quantVal & " (öðrenilen parça)", ""

                    ' ---------- SOMUN / PUL / YAYLI RONDELA (müþteri satýrý) ----------
                    ' Þablon bunlarý her çap için cývatalardan otomatik üretir (yeþil 3'lü set).
                    ' Müþterinin yazdýðý adet sadece kontrol amaçlý toplanýr.
                    Case "NUT", "WASHER", "SPRING"
                        mbDia = ParseHardwareDia(rawName, mbKind)
                        If mbDia > 0 And quantVal > 0 Then
                            mbHwKey = mbKind & "|" & CStr(mbDia)
                            mbHwCust(mbHwKey) = mbHwCust(mbHwKey) + quantVal
                        End If
                        auditSkipped = auditSkipped + 1
                        AuditRecord fileName, "EXCEL", r, wsIn.Name, rawName, mbKind, "SKIPPED", 100, "M" & mbDia & "; ADET=" & quantVal, _
                                    "Otomatik set (somun/pul/yaylý rondela) ile üretilir; adet kontrol için kaydedildi."

                    ' ---------------- PLAKA ----------------
                    Case "PLATE"
                        If quantVal <= 0 Then
                            AuditRecord fileName, "EXCEL", r, wsIn.Name, rawName, "PLATE", "WARNING", 40, "", "Adet okunamadý, satýr atlandý."
                            GoTo NextRowExcel
                        End If
                        Call ApplyDefaultQuality(quality, rawName, posNo)
                        Call ParsePlateDimensions(rawName, pThickE, pWidthE, pNoteE)
                        ' Ýsimde ölçü yoksa Dicke / Breite sütunlarýndan al
                        If pThickE <= 0 And mbCols(9) > 0 Then pThickE = BomCellNum(srcData, r, mbCols(9), maxC)
                        If pWidthE <= 0 And mbCols(10) > 0 Then pWidthE = BomCellNum(srcData, r, mbCols(10), maxC)
                        ' Öðrenilen plaka: kalýnlýk / geniþlik OGRENILEN'den
                        If mbLrnKind = "PLATE" Then
                            If mbLrnT > 0 Then pThickE = mbLrnT
                            If mbLrnW > 0 Then pWidthE = mbLrnW
                            If pNoteE = "" And mbLrnName <> "" And UCase$(mbLrnName) <> UCase$(rawName) Then pNoteE = mbLrnName
                        End If
                        ' Daire plaka: Ø x Ø kare pafta; kalýnlýk Dicke sütunundan, yoksa LENGTH'ten (Tekla: D460, L=4)
                        If mbDisc Then
                            pWidthE = mbDia2
                            pThickE = 0
                            If mbCols(9) > 0 Then pThickE = BomCellNum(srcData, r, mbCols(9), maxC)
                            If pThickE <= 0 Then pThickE = lengthVal
                            lengthVal = mbDia2
                            pNoteE = ChrW(216) & NumToText(mbDia2) & " daire (kare paftadan)"
                        End If

                        finalPlateNote = notStr
                        If pNoteE <> "" Then
                            If finalPlateNote <> "" Then
                                finalPlateNote = finalPlateNote & ", " & pNoteE
                            Else
                                finalPlateNote = pNoteE
                            End If
                        End If

                        If posNo = "" Then
                            posKeyPlateE = "NOPOS_" & r & "|" & wsIn.Name & "|" & fileName
                        ElseIf isMerge Then
                            posKeyPlateE = MakePosKeyPlate(posNo, pThickE, pWidthE, lengthVal)
                        Else
                            posKeyPlateE = MakePosKeyPlate(posNo, pThickE, pWidthE, lengthVal) & "|" & fileName
                        End If

                        If Not isDryRun Then
                            If Not dictPosPlate.Exists(posKeyPlateE) Then
                                Call SafeWrite(curWsPlate, targetRowPlate, 1, FileTypeLabel(fileName))
                                Call SafeWrite(curWsPlate, targetRowPlate, 2, posNo)
                                Call SafeWrite(curWsPlate, targetRowPlate, 3, pThickE)
                                Call SafeWrite(curWsPlate, targetRowPlate, 4, pWidthE)
                                Call SafeWrite(curWsPlate, targetRowPlate, 5, quality)
                                Call SafeWrite(curWsPlate, targetRowPlate, 6, lengthVal)
                                Call SafeWrite(curWsPlate, targetRowPlate, currentColPlate, quantVal)
                                If finalPlateNote <> "" Then Call SafeWrite(curWsPlate, targetRowPlate, 56, finalPlateNote)
                                dictPosPlate.Add posKeyPlateE, targetRowPlate
                                targetRowPlate = targetRowPlate + 1
                            Else
                                existRow = dictPosPlate(posKeyPlateE)
                                Call AddTypeLabel(curWsPlate, existRow, False, fileName)
                                Call CheckPosConflict("PLATE", existRow, posNo, "", pThickE, pWidthE, lengthVal, fileName, "EXCEL", r)
                                pQtyE = Val(bufPlate(existRow, currentColPlate))
                                Call SafeWrite(curWsPlate, existRow, currentColPlate, pQtyE + quantVal)
                            End If
                        End If
                        totalProcessed = totalProcessed + 1
                        mbAngPlateWt = mbAngPlateWt + mbRowWt

                        mbStatus = "EXACT": mbConf = 100: mbIssue = ""
                        If pThickE <= 0 Or pWidthE <= 0 Then mbStatus = "ERROR": mbConf = 20: mbIssue = "Plaka kalýnlýk/geniþlik çözülemedi."
                        If posNo = "" Then mbStatus = IIf(mbStatus = "ERROR", "ERROR", "WARNING"): mbIssue = AddIssue(mbIssue, "Poz no boþ.")
                        If lengthVal <= 0 Then mbStatus = IIf(mbStatus = "ERROR", "ERROR", "WARNING"): mbIssue = AddIssue(mbIssue, "Boy boþ.")
                        AuditRecord fileName, "EXCEL", r, wsIn.Name, rawName & " | Poz: " & posNo, "PLATE", mbStatus, mbConf, _
                                    "T=" & pThickE & "; W=" & pWidthE & "; L=" & lengthVal & "; " & quality & "; ADET=" & quantVal, mbIssue

                    ' ---------------- PROFÝL / KÖÞEBENT ----------------
                    Case "ANGLE"
                        If quantVal <= 0 Then
                            AuditRecord fileName, "EXCEL", r, wsIn.Name, rawName, "ANGLE", "WARNING", 40, "", "Adet okunamadý, satýr atlandý."
                            GoTo NextRowExcel
                        End If
                        Call ApplyDefaultQuality(quality, rawName, posNo)
                        rawNameUpper = UCase$(rawName)
                        cleanNoSpace = Replace(rawNameUpper, " ", "")
                        mbIssue = ""
                        If mbLrnKind <> "ANGLE" And Not ProfileKnown(rawNameUpper, dictProfileLib) Then
                            mbIssue = "Profil kütüphanede yok."
                            If Not dictFeedback.Exists(rawNameUpper) Then
                                dictFeedback.Add rawNameUpper, fileName
                                ' (kütüphaneye ekleme artýk OGRENME sayfasýndan yapýlýr)
                            End If
                        End If

                        matCodeStrE = GetProfileMaterialCode(rawName, dictProfileLib)
                        If mbLrnKind = "ANGLE" And mbLrnCode <> "" Then matCodeStrE = mbLrnCode

                        If posNo = "" Then
                            posKeyAngleE = "NOPOS_" & r & "|" & wsIn.Name & "|" & fileName
                        ElseIf isMerge Then
                            posKeyAngleE = MakePosKeyAngle(posNo, matCodeStrE, lengthVal)
                        Else
                            posKeyAngleE = MakePosKeyAngle(posNo, matCodeStrE, lengthVal) & "|" & fileName
                        End If

                        If Not isDryRun Then
                            If Not dictPosAngle.Exists(posKeyAngleE) Then
                                Call SafeWrite(curWsAngle, targetRowAngle, 1, FileTypeLabel(fileName))
                                Call SafeWrite(curWsAngle, targetRowAngle, 2, posNo)
                                Call SafeWrite(curWsAngle, targetRowAngle, 3, matCodeStrE)
                                Call SafeWrite(curWsAngle, targetRowAngle, 5, quality)
                                Call SafeWrite(curWsAngle, targetRowAngle, 6, lengthVal)
                                Call SafeWrite(curWsAngle, targetRowAngle, currentColAngle, quantVal)
                                If notStr <> "" Then Call SafeWrite(curWsAngle, targetRowAngle, 57, notStr)
                                dictPosAngle.Add posKeyAngleE, targetRowAngle
                                targetRowAngle = targetRowAngle + 1
                            Else
                                existRow = dictPosAngle(posKeyAngleE)
                                Call AddTypeLabel(curWsAngle, existRow, True, fileName)
                                Call CheckPosConflict("ANGLE", existRow, posNo, matCodeStrE, 0, 0, lengthVal, fileName, "EXCEL", r)
                                aQtyE = Val(bufAngle(existRow, currentColAngle))
                                Call SafeWrite(curWsAngle, existRow, currentColAngle, aQtyE + quantVal)
                            End If
                        End If
                        totalProcessed = totalProcessed + 1
                        mbAngPlateWt = mbAngPlateWt + mbRowWt

                        mbStatus = IIf(mbIssue = "", "EXACT", "WARNING")
                        If posNo = "" Then mbStatus = "WARNING": mbIssue = AddIssue(mbIssue, "Poz no boþ.")
                        If lengthVal <= 0 Then mbStatus = "WARNING": mbIssue = AddIssue(mbIssue, "Boy boþ.")
                        AuditRecord fileName, "EXCEL", r, wsIn.Name, rawName & " | Poz: " & posNo, "ANGLE", mbStatus, IIf(mbStatus = "EXACT", 100, 70), _
                                    "KOD=" & matCodeStrE & "; L=" & lengthVal & "; " & quality & "; ADET=" & quantVal, mbIssue

                    ' ---------------- TANINMAYAN ----------------
                    Case Else
                        If quantVal > 0 Then
                            If Not dictFeedback.Exists(rawNameUpper) Then
                                dictFeedback.Add rawNameUpper, fileName & " (Poz: " & posNo & ")"
                                ' (kütüphaneye ekleme artýk OGRENME sayfasýndan yapýlýr)
                            End If
                            Call AuditRecord(fileName, "EXCEL", r, wsIn.Name, rawName & " | Poz: " & posNo, "UNMATCHED", "UNMATCHED", 0, "", _
                                             "Köþebent/Profil, Plaka veya Cývata olarak sýnýflandýrýlamadý.")
                        End If
                End Select
NextRowExcel:
            Next r
NextSheetExcel:
        Next wsIn
        auditCurrentSheet = ""

        If mbSheetsUsed = 0 Then
            AuditRecord fileName, "EXCEL", 0, "FILE", CStr(vItem), "FILE", "ERROR", 0, "", _
                        "Hiçbir sayfada malzeme listesi baþlýðý bulunamadý."
        End If

        ' Müþterinin yazdýðý somun/pul/rondela adetlerini çevrilen cývata adetleriyle karþýlaþtýr
        For Each mbHwKey In mbHwCust.keys
            parts = Split(CStr(mbHwKey), "|")
            mbDia = CLng(Val(parts(1)))
            mbCustQty = CLng(mbHwCust(mbHwKey))
            mbOurQty = 0
            If mbBoltDia.Exists(CStr(mbDia)) Then mbOurQty = CLng(mbBoltDia(CStr(mbDia)))
            If mbCustQty = mbOurQty Then
                AuditRecord fileName, "EXCEL-CHECK", 0, "HARDWARE", parts(0) & " M" & mbDia, parts(0), "EXACT", 100, _
                            "Müþteri=" & mbCustQty & "; Cývata(set üretilen)=" & mbOurQty, ""
            Else
                AuditRecord fileName, "EXCEL-CHECK", 0, "HARDWARE", parts(0) & " M" & mbDia, parts(0), "WARNING", 60, _
                            "Müþteri=" & mbCustQty & "; Cývata(set üretilen)=" & mbOurQty, _
                            "Müþteri listesindeki adet, otomatik üretilecek set adedinden farklý. Kontrol edin (ör: basamak cývatalarý, özel parçalar)."
            End If
        Next mbHwKey

        ' Dosya aðýrlýðý: müþterinin profil+plaka satýr toplamý; yoksa "Stahlgewicht" satýrý
        If mbAngPlateWt > 0 Then
            fileWeight = mbAngPlateWt
        Else
            fileWeight = mbFooterWt
        End If

        Set mbHwCust = Nothing
        Set mbBoltDia = Nothing
        If Not wbOpenAlready Then wbIn.Close SaveChanges:=False
    End If
End Sub

Public Function BomCellText(ByRef data As Variant, ByVal r As Long, ByVal c As Long, ByVal maxC As Long) As String
    Dim v As Variant
    On Error GoTo Fail
    If c <= 0 Or c > maxC Then Exit Function
    If r < LBound(data, 1) Or r > UBound(data, 1) Then Exit Function
    v = data(r, c)
    If IsError(v) Or IsEmpty(v) Then Exit Function
    Select Case VarType(v)
        Case vbDouble, vbSingle, vbCurrency, vbInteger, vbLong, vbDecimal, vbByte
            BomCellText = NumToText(CDbl(v))
        Case vbBoolean
            BomCellText = ""
        Case Else
            BomCellText = CStr(v)
            BomCellText = Replace(BomCellText, ChrW(160), " ")
            BomCellText = Replace(BomCellText, vbCr, " ")
            BomCellText = Replace(BomCellText, vbLf, " ")
            BomCellText = Replace(BomCellText, """", "")
            BomCellText = Trim$(BomCellText)
    End Select
    Exit Function
Fail:
    BomCellText = ""
End Function

' Boy sütununun birimi: þablon F sütunu mm bekler.
' Baþlýkta birim yazýyorsa ona göre ("Länge [m]", "L (m)", "Length in m", "cm"); yazmýyorsa
' boylar metre gibi görünüyorsa (en büyük boy <= 30 ve ondalýklý deðer var) x1000.
Public Function LengthUnitFactor(ByRef data As Variant, ByVal hRow As Long, ByVal c As Long, _
                                 ByVal maxR As Long, ByVal maxC As Long) As Double
    Dim t As String, k As Long, r As Long, v As Double, n As Long, mx As Double, hasFrac As Boolean
    LengthUnitFactor = 1
    If c <= 0 Or c > maxC Then Exit Function
    ' baþlýk ve (sayý deðilse) altýndaki birim satýrý
    t = FoldText(BomCellText(data, hRow, c, maxC))
    If hRow + 1 <= maxR Then
        If BomCellNum(data, hRow + 1, c, maxC) = 0 Then t = t & " " & FoldText(BomCellText(data, hRow + 1, c, maxC))
    End If
    t = Trim$(t)
    If RxTest(t, "\bMM\b|MILLI") Then Exit Function
    If RxTest(t, "\bCM\b") Then LengthUnitFactor = 10: Exit Function
    If RxTest(t, "[\(\[]\s*M\s*[\)\]]|\bIN\s+M\b|\bM$|METER|METRE") Then LengthUnitFactor = 1000: Exit Function
    ' baþlýkta birim yok: deðerlere bak
    For r = hRow + 1 To maxR
        If r > hRow + 3000 Then Exit For
        v = BomCellNum(data, r, c, maxC)
        If v > 0 Then
            n = n + 1
            If v > mx Then mx = v
            If v <> Int(v) Then hasFrac = True
        End If
    Next r
    If n >= 3 And mx <= 30 And hasFrac Then LengthUnitFactor = 1000
End Function

Public Function BomCellNum(ByRef data As Variant, ByVal r As Long, ByVal c As Long, ByVal maxC As Long) As Double
    Dim v As Variant
    On Error GoTo Fail
    If c <= 0 Or c > maxC Then Exit Function
    If r < LBound(data, 1) Or r > UBound(data, 1) Then Exit Function
    v = data(r, c)
    If IsError(v) Or IsEmpty(v) Then Exit Function
    Select Case VarType(v)
        Case vbDouble, vbSingle, vbCurrency, vbInteger, vbLong, vbDecimal, vbByte
            BomCellNum = CDbl(v)
        Case vbString
            BomCellNum = ParseBOMNumber(CStr(v))
    End Select
    Exit Function
Fail:
    BomCellNum = 0
End Function

' Baþlýk hücresinin rolü:
' 1=POS 2=ADET 3=ÝSÝM 4=BOY 5=MALZEME 6=NOT 7=DIN 8=TOPLAM AÐIRLIK 9=KALINLIK 10=GENÝÞLÝK
' 11=AÐIRLIK(genel) 12=BÝRÝM AÐIRLIK 13=ÖLÇÜ (Abmessung)
Public Function HeaderRoleOf(ByVal cellTxt As String) As Long
    Dim t As String, c As String
    t = FoldText(cellTxt)
    If t = "" Then Exit Function
    If Right$(t, 1) = ":" Then Exit Function      ' antet etiketi: "Benennung:", "Zeichnungs-Nr.:"
    If Len(t) > 40 Then Exit Function             ' uzun metin baþlýk deðildir
    c = t
    c = Replace(c, " ", "")
    c = Replace(c, ".", "")
    c = Replace(c, "-", "")
    c = Replace(c, "_", "")

    Select Case c
        Case "POS", "POZ", "POSNR", "POSNO", "POSITION", "POSITIONSNR", "ITEM", "ITEMNO", "MARK", "MARKNO", _
             "PART", "PARTNO", "PARTMARK", "PIECEMARK", "TEILENR", "TEILNR", "PARCANO", "POZNO"
            HeaderRoleOf = 1: Exit Function
        Case "STCK", "STK", "STUCK", "STUECK", "ANZ", "ANZAHL", "MENGE", "QTY", "QTE", "QUANTITY", _
             "PCS", "PIECE", "PIECES", "ADET", "MIKTAR"
            HeaderRoleOf = 2: Exit Function
    End Select
    If c Like "QTY*" Or c Like "STCK*" Or c Like "ANZAHL*" Or c Like "MENGE*" Or c Like "ADET*" Or c Like "QUANTITY*" _
       Or c = "ANTALL" Or c = "ANT" Then
        HeaderRoleOf = 2: Exit Function
    End If
    ' "No." / "Nr." belirsiz: yanýnda "Item No." varsa ADET, yoksa POZ (MapBOMHeader karar verir)
    If c = "NO" Or c = "NR" Or c = "NOS" Or c = "NUMBER" Or c = "NUM" Then
        HeaderRoleOf = 16: Exit Function
    End If

    If InStr(c, "LANGE") > 0 Or InStr(c, "LAENGE") > 0 Or InStr(c, "LENGTH") > 0 Or InStr(c, "UZUNLUK") > 0 Or InStr(c, "LENGDE") > 0 _
       Or c = "BOY" Or c = "L" Or c Like "L(MM*" Or c Like "L[[]MM*" Then
        HeaderRoleOf = 4: Exit Function
    End If

    If InStr(c, "ABMESSUNG") > 0 Or c = "DIMENSION" Or c = "DIMENSIONS" Or c = "SIZE" Or c = "OLCU" Or c = "OLCULER" Then
        HeaderRoleOf = 13: Exit Function
    End If

    ' Profil/kesit sütunu ayrý rol (14): varsa isim olarak bu kullanýlýr (DESCRIPTION sadece "PLATE" gibi tür yazar)
    If InStr(c, "PROFIL") > 0 Or InStr(c, "SECTION") > 0 Or c = "KESIT" Then
        HeaderRoleOf = 14: Exit Function
    End If

    If InStr(c, "BENENNUNG") > 0 Or InStr(c, "BEZEICHNUNG") > 0 Or InStr(c, "DESCRIPTION") > 0 Or InStr(c, "BESKRIVELSE") > 0 _
       Or c = "DESC" Or InStr(c, "GEGENSTAND") > 0 Or InStr(c, "VERBINDUNGSMATERIAL") > 0 _
       Or c = "TANIM" Or c = "MALZEMECINSI" Then
        HeaderRoleOf = 3: Exit Function
    End If

    If InStr(c, "DICKE") > 0 Or InStr(c, "THICK") > 0 Or InStr(c, "KALINLIK") > 0 Then HeaderRoleOf = 9: Exit Function
    If InStr(c, "BREITE") > 0 Or InStr(c, "WIDTH") > 0 Or InStr(c, "GENISLIK") > 0 Then HeaderRoleOf = 10: Exit Function

    If InStr(c, "GESAMTGEWICHT") > 0 Or InStr(c, "TOTALWEIGHT") > 0 Or InStr(c, "T/WEIGHT") > 0 Or InStr(c, "TOPLAMAGIRLIK") > 0 _
       Or c Like "TOT*WEIGHT*" Or c = "GESAMT" Or c = "TOTAL" Or c = "TOTAL(KG)" Or c = "GESAMT(KG)" Or c = "GESAMTKG" Then
        HeaderRoleOf = 8: Exit Function
    End If
    ' Parça baþý aðýrlýk (15): adet x bu = satýr aðýrlýðý
    If InStr(c, "/PCS") > 0 Or InStr(c, "/PC") > 0 Or InStr(c, "PERPC") > 0 Or InStr(c, "/STK") > 0 Or c = "KG/STK" Or c = "KG/ST" _
       Or c Like "EINZELGEW*" Or InStr(c, "STUCKGEW") > 0 Or InStr(c, "UNITWEIGHT") > 0 Or InStr(c, "BIRIMAGIRLIK") > 0 Then
        HeaderRoleOf = 15: Exit Function
    End If
    ' Metre aðýrlýðý (12)
    If c = "KG/M" Or c = "(KG/M)" Or c Like "EINZEL*" Then
        HeaderRoleOf = 12: Exit Function
    End If
    If InStr(c, "GEWICHT") > 0 Or InStr(c, "WEIGHT") > 0 Or InStr(c, "AGIRLIK") > 0 Or InStr(c, "VEKT") > 0 Or c = "MASS" Or c = "MASSE" Or c = "KG" Or c = "(KG)" Then
        HeaderRoleOf = 11: Exit Function
    End If

    If Left$(c, 3) = "DIN" Or InStr(c, "ARTIKEL") > 0 Or InStr(c, "NORM") > 0 Or InStr(c, "STANDARD") > 0 Then HeaderRoleOf = 7: Exit Function

    If InStr(c, "BEMERKUNG") > 0 Or InStr(c, "REMARK") > 0 Or InStr(c, "NOTE") > 0 Or InStr(c, "COMMENT") > 0 _
       Or InStr(c, "ACIKLAMA") > 0 Or c = "NOT" Or c = "NOTLAR" Then
        HeaderRoleOf = 6: Exit Function
    End If

    If (InStr(c, "MATERIAL") > 0 Or InStr(c, "WERKSTOFF") > 0 Or InStr(c, "GUTE") > 0 Or InStr(c, "GUETE") > 0 _
        Or InStr(c, "MALZEME") > 0 Or InStr(c, "GRADE") > 0 Or InStr(c, "KALITE") > 0 Or InStr(c, "QUALITY") > 0 _
        Or InStr(c, "STEEL") > 0 Or c = "MAT" Or c = "STAHLSORTE") And InStr(c, "AUSZUG") = 0 Then
        HeaderRoleOf = 5: Exit Function
    End If
End Function

' Bir satýr malzeme listesi baþlýðý mý? (hücre bazlý; en az 3 farklý rol + Poz/Adet/Ýsim'den ikisi)
Public Function IsBOMHeaderRow(ByRef data As Variant, ByVal r As Long, ByVal maxC As Long) As Boolean
    Dim c As Long, role As Long, nRoles As Long
    Dim seen(1 To 16) As Boolean
    Dim hasPos As Boolean, hasQty As Boolean, hasName As Boolean
    On Error GoTo Fail
    For c = 1 To maxC
        role = HeaderRoleOf(BomCellText(data, r, c, maxC))
        If role > 0 Then
            If Not seen(role) Then
                seen(role) = True
                nRoles = nRoles + 1
            End If
        End If
    Next c
    hasPos = seen(1)
    hasQty = seen(2)
    If seen(16) Then
        If hasPos Then hasQty = True Else hasPos = True
    End If
    hasName = seen(3) Or seen(13) Or seen(14)
    IsBOMHeaderRow = ((hasQty And hasName) Or (hasPos And hasName) Or (hasPos And hasQty)) And nRoles >= 3
    Exit Function
Fail:
    IsBOMHeaderRow = False
End Function

' Baþlýk + alt baþlýk (birim satýrý) okuyarak sütun rollerini doldurur
Public Sub MapBOMHeader(ByRef data As Variant, ByVal hr As Long, ByVal maxR As Long, ByVal maxC As Long, ByRef cols() As Long)
    Dim c As Long, role As Long, t As String
    For c = LBound(cols) To UBound(cols)
        cols(c) = 0
    Next c

    For c = 1 To maxC
        role = HeaderRoleOf(BomCellText(data, hr, c, maxC))
        If role > 0 And role <= UBound(cols) Then
            If cols(role) = 0 Then cols(role) = c
        End If
    Next c

    ' Alt baþlýk satýrý: "kg/m" | "Gesamt" | "mm" gibi (birleþtirilmiþ "Gewicht" baþlýðýnýn altý)
    If hr + 1 <= maxR Then
        For c = 1 To maxC
            t = BomCellText(data, hr + 1, c, maxC)
            If t <> "" And Not RxTest(t, "^[\d\.,\s]+$") Then
                role = HeaderRoleOf(t)
                Select Case role
                    Case 0
                    Case 8
                        cols(8) = c
                    Case 12
                        cols(12) = c
                    Case Else
                        If role <= UBound(cols) Then
                            If cols(role) = 0 Then cols(role) = c
                        End If
                End Select
            End If
        Next c
    End If

    ' Toplam aðýrlýk yoksa genel "Gewicht" sütununu kullan (birim aðýrlýk sütunu deðilse)
    If cols(8) = 0 And cols(11) > 0 And cols(11) <> cols(12) Then cols(8) = cols(11)
    ' "No." sütunu: "Item No." varsa adet, yoksa poz
    If cols(16) > 0 Then
        If cols(1) > 0 And cols(2) = 0 Then
            cols(2) = cols(16)
        ElseIf cols(1) = 0 Then
            cols(1) = cols(16)
        End If
    End If
    ' Profil sütunu varsa isim olarak o kullanýlýr
    If cols(14) > 0 Then cols(3) = cols(14)
    ' Ýsim yoksa ölçü sütunu isim olur
    If cols(3) = 0 And cols(13) > 0 Then
        cols(3) = cols(13)
        cols(13) = 0
    End If
End Sub

Public Function IsMappedBOMCol(ByRef cols() As Long, ByVal c As Long) As Boolean
    Dim i As Long
    For i = LBound(cols) To UBound(cols)
        If cols(i) = c Then IsMappedBOMCol = True: Exit Function
    Next i
End Function

Public Function ColLetter(ByVal c As Long) As String
    If c <= 0 Then ColLetter = "-": Exit Function
    ColLetter = Split(ThisWorkbook.Sheets(1).Cells(1, c).Address(True, False), "$")(0)
End Function

Public Function DescribeBOMColumns(ByRef cols() As Long) As String
    DescribeBOMColumns = "POS=" & ColLetter(cols(1)) & "; ADET=" & ColLetter(cols(2)) & "; ISIM=" & ColLetter(cols(3)) & _
                         "; BOY=" & ColLetter(cols(4)) & "; MALZEME=" & ColLetter(cols(5)) & "; DIN=" & ColLetter(cols(7)) & _
                         "; NOT=" & ColLetter(cols(6)) & "; AGIRLIK=" & ColLetter(cols(8)) & _
                         "; KALINLIK=" & ColLetter(cols(9)) & "; GENISLIK=" & ColLetter(cols(10)) & "; OLCU=" & ColLetter(cols(13))
End Function

' Ýsim + ölçü sütunlarýný tek parça adýna çevirir.
' "Winkel" + "150x12" -> "L 150x12" ; "Blech" + "10x200" -> "BL 10x200"
Public Function BuildItemName(ByVal nameTxt As String, ByVal dimTxt As String) As String
    Dim n As String, d As String, k As String, nf As String
    n = Trim$(nameTxt)
    d = Trim$(dimTxt)
    If d = "" Then
        BuildItemName = NormalizeItemName(n)
        Exit Function
    End If
    If n = "" Then
        BuildItemName = NormalizeItemName(d)
        Exit Function
    End If
    k = ClassifyBOMItem(d, "", "")
    Select Case k
        Case "BOLT", "FURING", "PLATE", "ANGLE", "NUT", "WASHER", "SPRING", "ROUND"
            BuildItemName = d
            Exit Function
    End Select
    nf = FoldText(n)
    If RxTest(nf, "^(WINKEL|ANGLE|KOSEBENT|L-PROFIL|L PROFIL)") Then
        BuildItemName = "L " & d
    ElseIf RxTest(nf, "^(BLECH|PLATE|PLAKA|FLACHSTAHL|FLACH|FLAT|LAMA|SAC)") Then
        BuildItemName = "BL " & d
    Else
        BuildItemName = NormalizeItemName(n & " " & d)
    End If
End Function

' "Winkel 150x12" -> "L 150x12" ; "Blech 10x200" -> "BL 10x200" (sadece ardýndan rakam gelirse)
Public Function NormalizeItemName(ByVal txt As String) As String
    Dim f As String, m As String
    NormalizeItemName = Trim$(txt)
    f = FoldText(txt)
    m = RxFirst(f, "^(WINKELSTAHL|WINKEL|ANGLE|KOSEBENT)\s*")
    If m <> "" Then
        If RxTest(Mid$(f, Len(m) + 1), "^\d") Then NormalizeItemName = "L " & Trim$(Mid$(f, Len(m) + 1))
        Exit Function
    End If
    m = RxFirst(f, "^(BLECH|PLATE|PLAKA|FLACHSTAHL|LAMA|SAC)\s*")
    If m <> "" Then
        If RxTest(Mid$(f, Len(m) + 1), "^\d") Then NormalizeItemName = "BL " & Trim$(Mid$(f, Len(m) + 1))
    End If
End Function

' Anahtar kelime ile baþlýyor mu? requireDigit=True ise ardýndan (.,-,_ atlanarak) rakam gelmeli
Public Function StartsWithKeyword(ByVal c As String, ByVal arr As Variant, ByVal requireDigit As Boolean) As Boolean
    Dim i As Long, k As String, nxt As String
    On Error GoTo Fail
    If Not IsArray(arr) Then Exit Function
    For i = LBound(arr) To UBound(arr)
        k = Replace(CStr(arr(i)), " ", "")
        If k <> "" And Len(c) >= Len(k) Then
            If Left$(c, Len(k)) = k Then
                If Not requireDigit Then
                    StartsWithKeyword = True
                    Exit Function
                End If
                nxt = Mid$(c, Len(k) + 1)
                Do While Left$(nxt, 1) = "." Or Left$(nxt, 1) = "-" Or Left$(nxt, 1) = "_"
                    nxt = Mid$(nxt, 2)
                Loop
                If nxt Like "#*" Then
                    StartsWithKeyword = True
                    Exit Function
                End If
            End If
        End If
    Next i
    Exit Function
Fail:
    StartsWithKeyword = False
End Function

' Satýr sýnýflandýrýcý: BOLT / FURING / NUT / WASHER / SPRING / PLATE / ANGLE / GARBAGE / UNKNOWN / EMPTY
' Sýra önemlidir: önce KESÝN kalýplar (önek+rakam, M16x45...), sonra çöp, en son zayýf anahtar kelimeler.
Public Function ClassifyBOMItem(ByVal rawName As String, ByVal dinTxt As String, ByVal qualTxt As String, Optional ByVal hintTxt As String = "") As String
    Dim u As String, c As String, d As String, q As String, sepX As String
    u = FoldText(rawName)
    c = Replace(u, " ", "")
    d = Replace(FoldText(dinTxt), " ", "")
    q = Replace(Trim$(qualTxt), ",", ".")
    sepX = "[X\*/\-]"

    If c = "" Then ClassifyBOMItem = "EMPTY": Exit Function

    ' 1) Futterring (özel bölüm)
    ' FUBL = Futterblech (küçük dolgu parçasý, ör: "FUBL 18X10" = delik 18 (M16), 10 mm)
    If InStr(c, "FUTTER") > 0 Or InStr(c, "FU-RING") > 0 Or Left$(c, 4) = "FUBL" Then ClassifyBOMItem = "FURING": Exit Function

    ' Yuvarlak kesit: D460 / RD20 / Ø20 (motor boy/çap oranýna göre daire plaka ya da çubuk yapar)
    If RxTest(c, "^(RD|D|" & ChrW(216) & ")\d{1,4}([\.,]\d+)?$") Then ClassifyBOMItem = "ROUND": Exit Function

    ' 2) KESÝN plaka / profil: önek + rakam
    If RxTest(c, "^(BLECH|BLE|BL|PLATE|PLT|PL|FLACH|FLAT|FLA|FL|KF)[\.\-]?\d") Then ClassifyBOMItem = "PLATE": Exit Function
    If StartsWithKeyword(c, plateKeywords, True) Then ClassifyBOMItem = "PLATE": Exit Function
    If RxTest(c, "^(L|<|\?|" & ChrW(8736) & ")[\.\-]?\d") Then ClassifyBOMItem = "ANGLE": Exit Function
    If RxTest(c, "^(UNP|UPN|UPE|UAP|HEA|HEB|HEM|IPE|IPN|INP|PFC|UB|UC|HD|HL|HP|CFCHS|CFSHS|CFRHS|CHS|RHS|SHS|U|I|C|T|W|F)[\.\-]?\d") Then ClassifyBOMItem = "ANGLE": Exit Function
    If StartsWithKeyword(c, angleKeywords, True) Then ClassifyBOMItem = "ANGLE": Exit Function

    ' 3) KESÝN cývata kalýplarý
    '    "M7990 16 45+3" (DIN önce), "STGB 27 310" (basamak), "M16x45", "SSM24x345-R"
    If RxTest(u, "^M\s*(7990|7968|7969|931|933|601|558|6914|4014|4017)\s+\d{1,2}(?!\d)") Then ClassifyBOMItem = "BOLT": Exit Function
    If RxTest(u, "^(STGB|STEIGB|STB\b|TB\.?M)") Then ClassifyBOMItem = "BOLT": Exit Function
    If RxTest(u, "M\s*\d{1,2}(?!\d)\s*" & sepX & "\s*\d{2,3}") Then ClassifyBOMItem = "BOLT": Exit Function

    ' 4) Müþterinin yazdýðý somun / pul / yaylý rondela
    '    a) Ýpucu sütunu (poz sütununda "Nut M16", "Wash M16") ve isimdeki ISO/DIN standardý
    Dim hnt As String
    hnt = FoldText(hintTxt)
    If RxTest(hnt, "^(SPRING|FEDER|FDRG)") Or RxTest(u, "DIN\s*(127|128|6796)\b") Then ClassifyBOMItem = "SPRING": Exit Function
    If RxTest(hnt, "^(WASH|WASHER|SCHEIBE|U-SCHEIBE|PUL|SCHB)") Or RxTest(u, "(ISO\s*(7089|7090|7091|7093)|DIN\s*(125|126|433|434|435|436|6916|7349|7989))\b") Then
        ClassifyBOMItem = "WASHER": Exit Function
    End If
    If RxTest(hnt, "^(NUT|MUTTER|MU|SOMUN)\b") Or RxTest(u, "(ISO\s*(4032|4033|4034|7040)|DIN\s*(934|555|985|6915))\b") Then
        ClassifyBOMItem = "NUT": Exit Function
    End If
    If RxTest(hnt, "^(BOLT|SCHRAUBE|CIVATA|SCREW)") And RxTest(u, "M\s*\d{1,2}(?!\d)") Then ClassifyBOMItem = "BOLT": Exit Function
    '    b) Ýsimdeki anahtar kelimeler
    If RxTest(u, "^(FDRG|FEDERRING|FEDERSCHEIBE|SPRING|SW)\b") Or InStr(c, "FEDERRING") > 0 Or InStr(c, "SPRINGWASHER") > 0 _
       Or RxTest(d, "DIN(127|128|6796)") Then ClassifyBOMItem = "SPRING": Exit Function
    If RxTest(u, "^(U-SCHB|U-SCHEIBE|SCHB|SCH|SCHEIBE|PW|PUL|WASHER)\b") Or InStr(c, "SCHEIBE") > 0 Or InStr(c, "WASHER") > 0 _
       Or RxTest(d, "(DIN(125|126|433|434|435|436|6916|7349|7989)|ISO(7089|7090|7091|7093))") Then ClassifyBOMItem = "WASHER": Exit Function
    If RxTest(u, "^(MU|MUTTER|NUT|SOMUN)\b") Or InStr(c, "MUTTER") > 0 Or RxTest(u, "\bNUT\b") _
       Or RxTest(d, "(DIN(934|555|6915|985)|ISO(4032|4033|4034|7040))") Then ClassifyBOMItem = "NUT": Exit Function

    ' 5) Çöp / toplam satýrlarý
    If RxTest(c, "^[\d\.,\+\-%]+$") Then ClassifyBOMItem = "GARBAGE": Exit Function
    If CheckArrayMatch(u, garbageKeywords) Then ClassifyBOMItem = "GARBAGE": Exit Function

    ' 6) Zayýf cývata iþaretleri
    If RxTest(u, "(SCHRAUBE|BOLT|SCREW|CIVATA|STEIGBOLZEN|ANKER|ANCHOR)") Then ClassifyBOMItem = "BOLT": Exit Function
    If RxTest(u, "(^|[^A-Z])M\s*\d{1,2}(?!\d)") Then
        If IsValidBoltGrade(q) Or RxTest(d, "(7990|7968|7969|931|933|601|4014|4016|4017)") Then ClassifyBOMItem = "BOLT": Exit Function
    End If
    If RxTest(d, "(DIN(7990|7968|7969|931|933|601)|ISO(4014|4016|4017))") Then ClassifyBOMItem = "BOLT": Exit Function
    If CheckArrayMatch(u, hwKeywords) Then ClassifyBOMItem = "BOLT": Exit Function

    ' 7) Zayýf plaka iþaretleri (boru/çubuk/Ø, CHK, sadece "BL" yazýp ölçüyü Dicke/Breite'ye koyanlar)
    If RxTest(u, "(PIPE|ROHR|\bROD\b|ROUND|RUNDSTAHL)") Or InStr(u, ChrW(216)) > 0 Or InStr(c, "CHK") > 0 Then ClassifyBOMItem = "PLATE": Exit Function
    If StartsWithKeyword(c, plateKeywords, False) Then ClassifyBOMItem = "PLATE": Exit Function

    ClassifyBOMItem = "UNKNOWN"
End Function

' Somun/pul/rondela satýrýndan cývata çapý. "SCH 18 5" = delik 18 -> M16 ; "FDRG 16" -> M16
Public Function ParseHardwareDia(ByVal rawName As String, ByVal kind As String) As Long
    Dim u As String, t As String, n As Long
    On Error GoTo Fail
    u = FoldText(rawName)
    t = RxFirst(u, "M\s*\d{1,2}(?!\d)")
    If t <> "" Then
        ParseHardwareDia = CLng(Val(Mid$(Replace(t, " ", ""), 2)))
        Exit Function
    End If
    t = RxFirst(u, "\d{1,2}(?!\d)")
    If t = "" Then Exit Function
    n = CLng(Val(t))
    If kind = "WASHER" And RxTest(u, "^(U-SCHB|U-SCHEIBE|SCHB|SCH)\b") Then n = WasherHoleToDia(n)
    ParseHardwareDia = n
    Exit Function
Fail:
    ParseHardwareDia = 0
End Function

Public Function WasherHoleToDia(ByVal hole As Long) As Long
    Dim ds As Variant, i As Long
    ds = Array(42, 39, 36, 33, 30, 27, 24, 22, 20, 18, 16, 14, 12, 10, 8, 6)
    For i = LBound(ds) To UBound(ds)
        If ds(i) < hole And hole - ds(i) <= 4 Then
            WasherHoleToDia = ds(i)
            Exit Function
        End If
    Next i
    WasherHoleToDia = hole
End Function

' Eski kural korunur: kalite boþsa isimde/pozda "H" -> S355 (H ekini temizle), deðilse S275.
' Fark: HEA/HEB/HD/HL/HP profil adýndaki H artýk S355 sebebi sayýlmaz.
Public Sub ApplyDefaultQuality(ByRef quality As String, ByRef rawName As String, ByVal posNo As String)
    Dim u As String, cleanProf As String
    If Trim$(quality) <> "" Then Exit Sub
    u = UCase$(rawName)
    If (InStr(u, "H") > 0 And Not RxTest(Replace(u, " ", ""), "^(HE|HD|HL|HP)")) Or InStr(UCase$(posNo), "H") > 0 Then
        quality = prmQualH
        cleanProf = u
        cleanProf = Replace(cleanProf, "H X", "X")
        cleanProf = Replace(cleanProf, "HX", "X")
        cleanProf = Replace(cleanProf, "H *", "*")
        cleanProf = Replace(cleanProf, "H*", "*")
        If Right$(cleanProf, 1) = "H" Then cleanProf = Left$(cleanProf, Len(cleanProf) - 1)
        rawName = Trim$(cleanProf)
    Else
        quality = prmQualDefault
    End If
End Sub

' Tekla çizimlerindeki revizyon iþaretlerini atar: "PL40*518 02 03" -> "PL40*518"
' Sadece ilk parça ölçülü bir profilse (rakam + * veya X) ve sonrakiler 2 haneli sayýysa.
Public Function StripRevisionMarks(ByVal txt As String) As String
    Dim p() As String, i As Long, lastKeep As Long, out As String
    StripRevisionMarks = Trim$(txt)
    If Trim$(txt) = "" Then Exit Function
    p = Split(Application.WorksheetFunction.Trim(txt), " ")
    If UBound(p) < 1 Then Exit Function
    If Not RxTest(p(0), "^[A-Z]*\d[\d\.,]*([\*X/][\d\.,]+)+$") Then Exit Function
    lastKeep = UBound(p)
    Do While lastKeep >= 1
        If p(lastKeep) Like "##" Then
            lastKeep = lastKeep - 1
        Else
            Exit Do
        End If
    Loop
    If lastKeep = UBound(p) Then Exit Function
    out = p(0)
    For i = 1 To lastKeep
        out = out & " " & p(i)
    Next i
    StripRevisionMarks = out
End Function

