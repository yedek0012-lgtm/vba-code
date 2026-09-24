Attribute VB_Name = "modMotorXSR"
Option Explicit
Option Private Module

' =========================================================================
' XSR / TXT (Tekla) motoru ve assembly ayrýþtýrýcýlarý
' =========================================================================

Public Sub ProcessXsrFile()
    ' XSR / TXT dosyasý (Tekla raporu). Paylaþýlan durum: modState
    Dim aQty As Double, assemblyApproxMatch As Boolean, assemblyArea As Double
    Dim assemblyConfidence As Long, assemblyIsAngle As Boolean, assemblyIsPlate As Boolean
    Dim assemblyIssue As String, assemblyLength As Double, assemblyMatCode As String
    Dim assemblyName As String, assemblyPos As String, assemblyProfile As String
    Dim assemblyQty As Long, assemblyStatus As String, assemblyThickness As Double
    Dim assemblyType As String, assemblyValid As Boolean, assemblyWeight As Double
    Dim assemblyWidth As Double, boltKeyX As String, boltNoteX As String, boltQualX As String
    Dim bQtyX As Long, cleanNoSpace As String, currentSection As String, dValX As Long
    Dim existRow As Long, fileNo As Integer, finalBoltQty As Long, finalXsrPlateNote As String
    Dim hwCodeX As String, hwNameX As String, i As Long, isBoltItem As Boolean, isFuRingX As Boolean
    Dim isInsideParens As Boolean, lengthVal As Double, lineStr As String, loopGuard As Long
    Dim lValX As Long, matCodeStr As String, partApprox As Boolean, partAssembly As String
    Dim partConf As Long, partGrade As String, partIsPlate As Boolean, partIsProfile As Boolean
    Dim partIssue As String, partLength As Double, partMatCode As String, partParsed As String
    Dim partPos As String, partProfile As String, partQty As Long, parts() As String
    Dim partStatus As String, partValid As Boolean, partWeight As Double, pNoteX As String
    Dim posKeyAngle As String, posKeyPlate As String, posNo As String, pQty As Double
    Dim pThick As Double, pWidth As Double, qIdx As Long, qtyKeyX As String, quality As String
    Dim quantIdx As Long, quantVal As Long, rawLine As String, rawLineUpper As String
    Dim rawTypeUpper As String, sourceLineNo As Long, tmpWeightStr As String, xsrNote As String

    fileNo = FreeFile
    Open vItem For Input As #fileNo
    currentSection = ""
    sourceLineNo = 0
    
    Do While Not EOF(fileNo)
        Line Input #fileNo, rawLine
        sourceLineNo = sourceLineNo + 1
        lineStr = Trim(rawLine)

        If IsAssemblyPartListHeader(lineStr) Then
            currentSection = "ASSEMBLY_PART"
        ElseIf IsAssemblyListHeader(lineStr) Then
            currentSection = "ASSEMBLY"
        End If

        If InStr(1, UCase(lineStr), "TOTAL FOR", vbTextCompare) > 0 Then
            If InStr(1, UCase(lineStr), "ASSEMBL", vbTextCompare) > 0 Then
                assemblyWeight = GetLastNumericValue(lineStr)
                If assemblyWeight > 0 Then
                    fileWeight = fileWeight + assemblyWeight
End If
            End If
End If

        If currentSection = "ASSEMBLY_PART" Then
            partValid = ParseAssemblyPartListRow(lineStr, partAssembly, partPos, partQty, _
                                                 partProfile, partGrade, partLength, partWeight)

            If partValid Then
                partIsPlate = False
                partIsProfile = False
                partMatCode = ""
                partApprox = False
                partStatus = "EXACT"
                partConf = 100
                partIssue = ""
                partParsed = ""

                DetermineAssemblyPartType partProfile, partIsPlate, partIsProfile

                If partIsPlate Then
                    ParseAssemblyPartPlateDimensions partProfile, assemblyThickness, assemblyWidth, partIssue
                    If assemblyThickness <= 0 Or assemblyWidth <= 0 Then
                        partStatus = "ERROR"
                        partConf = 20
                        partIssue = AddIssue(partIssue, "PLATE ölçüsü çözülemedi.")
End If

                    partParsed = "ASSEMBLY=" & partAssembly & "; PART=" & partPos & "; QTY=" & CStr(partQty) & _
                                 "; PROFILE=" & partProfile & "; GRADE=" & partGrade & _
                                 "; THK_MM=" & Format$(assemblyThickness, "0.0") & _
                                 "; WIDTH_MM=" & Format$(assemblyWidth, "0.0") & _
                                 "; LENGTH_MM=" & Format$(partLength, "0.0")

                    If isMerge Then
                        posKeyPlate = partPos
Else
                        posKeyPlate = partPos & "|" & fileName
End If
                    
                    If Not isDryRun Then
                        If Not dictPosPlate.Exists(posKeyPlate) Then
                            SafeWrite curWsPlate, targetRowPlate, 2, partPos
                            SafeWrite curWsPlate, targetRowPlate, 3, assemblyThickness
                            SafeWrite curWsPlate, targetRowPlate, 4, assemblyWidth
                            SafeWrite curWsPlate, targetRowPlate, 5, partGrade
                            SafeWrite curWsPlate, targetRowPlate, 6, partLength
                            SafeWrite curWsPlate, targetRowPlate, currentColPlate, partQty
                            If partIssue <> "" Then
                                SafeWrite curWsPlate, targetRowPlate, 56, partIssue
End If
                            dictPosPlate.Add posKeyPlate, targetRowPlate
                            targetRowPlate = targetRowPlate + 1
Else
                            existRow = dictPosPlate(posKeyPlate)
                            Call CheckPosConflict("PLATE", existRow, partPos, "", assemblyThickness, assemblyWidth, partLength, fileName, "XSR", sourceLineNo)
                            pQty = val(bufPlate(existRow, currentColPlate))
                            SafeWrite curWsPlate, existRow, currentColPlate, pQty + partQty
End If
                    End If
                    AuditRecord fileName, "XSR-ASSEMBLY-PART", sourceLineNo, "ASSEMBLY_PART", lineStr, "PLATE", partStatus, partConf, partParsed, partIssue
                    totalProcessed = totalProcessed + 1
                                ElseIf partIsProfile Then
                    partMatCode = GetAssemblyPartMaterialCode(partProfile, dictProfileLib, partApprox)
                    
                    If partMatCode = "" Then
                        partStatus = "WARNING"
                        partConf = 60
                        partIssue = AddIssue(partIssue, "Profil kütüphanesinde eþleþme bulunamadý.")
                    ElseIf partApprox Then
                        partStatus = "WARNING"
                        partConf = 85
                        partIssue = AddIssue(partIssue, "Profil kütüphanesinde yaklaþýk eþleþme kullanýldý.")
End If

                    partParsed = "ASSEMBLY=" & partAssembly & "; PART=" & partPos & "; QTY=" & CStr(partQty) & _
                                 "; PROFILE=" & partProfile & "; GRADE=" & partGrade & _
                                 "; CODE=" & partMatCode & "; LENGTH_MM=" & Format$(partLength, "0.0")

                    If isMerge Then
                        posKeyAngle = partPos
Else
                        posKeyAngle = partPos & "|" & fileName
End If
                    
                    If Not isDryRun Then
                        If Not dictPosAngle.Exists(posKeyAngle) Then
                            SafeWrite curWsAngle, targetRowAngle, 2, partPos
                            SafeWrite curWsAngle, targetRowAngle, 3, partMatCode
                            SafeWrite curWsAngle, targetRowAngle, 5, partGrade
                            SafeWrite curWsAngle, targetRowAngle, 6, partLength
                            SafeWrite curWsAngle, targetRowAngle, currentColAngle, partQty
                            If partIssue <> "" Then
                                SafeWrite curWsAngle, targetRowAngle, 57, partIssue
End If
                            dictPosAngle.Add posKeyAngle, targetRowAngle
                            targetRowAngle = targetRowAngle + 1
Else
                            existRow = dictPosAngle(posKeyAngle)
                            Call CheckPosConflict("ANGLE", existRow, partPos, partMatCode, 0, 0, partLength, fileName, "XSR", sourceLineNo)
                            aQty = val(bufAngle(existRow, currentColAngle))
                            SafeWrite curWsAngle, existRow, currentColAngle, aQty + partQty
End If
                    End If
                    AuditRecord fileName, "XSR-ASSEMBLY-PART", sourceLineNo, "ASSEMBLY_PART", lineStr, "PROFILE", partStatus, partConf, partParsed, partIssue
                    totalProcessed = totalProcessed + 1
Else
                    AuditRecord fileName, "XSR-ASSEMBLY-PART", sourceLineNo, "ASSEMBLY_PART", lineStr, "UNMATCHED", "UNMATCHED", 0, "PROFILE=" & partProfile, "Tanýnmadý"
                End If
End If
        End If

        If currentSection = "ASSEMBLY" Then
            assemblyValid = ParseAssemblyListRow(lineStr, assemblyPos, assemblyQty, assemblyName, assemblyProfile, assemblyArea, assemblyWeight)
            If assemblyValid Then
                assemblyType = ""
                assemblyStatus = "EXACT"
                assemblyIssue = ""
                assemblyConfidence = 100
                DetermineAssemblyItemType assemblyProfile, assemblyIsPlate, assemblyIsAngle

                If assemblyIsAngle Then
                    assemblyMatCode = GetAssemblyProfileMaterialCode(assemblyProfile, dictProfileLib, assemblyApproxMatch)
                    assemblyLength = EstimateAssemblyLength(assemblyProfile, assemblyWeight, assemblyQty, assemblyIssue)
                    If assemblyMatCode <> "" Then
                        assemblyStatus = "EXACT"
                    Else
                        assemblyStatus = "WARNING"
                        assemblyIssue = AddIssue(assemblyIssue, "Profil kütüphanesi eþleþmesi bulunamadý.")
End If
                    If assemblyLength <= 0 Then
                        assemblyStatus = "WARNING"
                    End If
                    
                    If isMerge Then
                        posKeyAngle = assemblyPos
Else
                        posKeyAngle = assemblyPos & "|" & fileName
End If

                    If Not isDryRun Then
                        If Not dictPosAngle.Exists(posKeyAngle) Then
                            SafeWrite curWsAngle, targetRowAngle, 2, assemblyPos
                            SafeWrite curWsAngle, targetRowAngle, 3, assemblyMatCode
                            SafeWrite curWsAngle, targetRowAngle, 6, assemblyLength
                            SafeWrite curWsAngle, targetRowAngle, currentColAngle, assemblyQty
                            dictPosAngle.Add posKeyAngle, targetRowAngle
                            targetRowAngle = targetRowAngle + 1
Else
                            existRow = dictPosAngle(posKeyAngle)
                            Call CheckPosConflict("ANGLE", existRow, assemblyPos, assemblyMatCode, 0, 0, assemblyLength, fileName, "XSR", sourceLineNo)
                            aQty = val(bufAngle(existRow, currentColAngle))
                            SafeWrite curWsAngle, existRow, currentColAngle, aQty + assemblyQty
End If
                    End If
                    AuditRecord fileName, "XSR-ASSEMBLY", sourceLineNo, "ASSEMBLY", lineStr, "ANGLE", assemblyStatus, assemblyConfidence, assemblyProfile, assemblyIssue
                    totalProcessed = totalProcessed + 1
                                ElseIf assemblyIsPlate Then
                    ParseAssemblyPlateDimensions assemblyProfile, assemblyThickness, assemblyWidth, assemblyIssue
                    assemblyLength = EstimateAssemblyLength(assemblyProfile, assemblyWeight, assemblyQty, assemblyIssue)
                    If assemblyThickness <= 0 Or assemblyWidth <= 0 Then
                        assemblyStatus = "ERROR"
                    Else
                        assemblyStatus = "EXACT"
                    End If

                    If isMerge Then
                        posKeyPlate = assemblyPos
Else
                        posKeyPlate = assemblyPos & "|" & fileName
End If

                    If Not isDryRun Then
                        If Not dictPosPlate.Exists(posKeyPlate) Then
                            SafeWrite curWsPlate, targetRowPlate, 2, assemblyPos
                            SafeWrite curWsPlate, targetRowPlate, 3, assemblyThickness
                            SafeWrite curWsPlate, targetRowPlate, 4, assemblyWidth
                            SafeWrite curWsPlate, targetRowPlate, 6, assemblyLength
                            SafeWrite curWsPlate, targetRowPlate, currentColPlate, assemblyQty
                            dictPosPlate.Add posKeyPlate, targetRowPlate
                            targetRowPlate = targetRowPlate + 1
Else
                            existRow = dictPosPlate(posKeyPlate)
                            Call CheckPosConflict("PLATE", existRow, assemblyPos, "", assemblyThickness, assemblyWidth, assemblyLength, fileName, "XSR", sourceLineNo)
                            pQty = val(bufPlate(existRow, currentColPlate))
                            SafeWrite curWsPlate, existRow, currentColPlate, pQty + assemblyQty
End If
                    End If
                    AuditRecord fileName, "XSR-ASSEMBLY", sourceLineNo, "ASSEMBLY", lineStr, "PLATE", assemblyStatus, 100, assemblyProfile, assemblyIssue
                    totalProcessed = totalProcessed + 1
End If
            End If
End If
        
        If InStr(1, UCase(lineStr), "TOTAL WEIGHT PROFILES:") > 0 Then
            tmpWeightStr = Mid(lineStr, InStr(1, UCase(lineStr), "TOTAL WEIGHT PROFILES:") + 22)
            tmpWeightStr = Replace(UCase(tmpWeightStr), "KG", "")
            tmpWeightStr = Replace(tmpWeightStr, "=", "")
            fileWeight = fileWeight + val(Trim(Replace(tmpWeightStr, ",", ".")))
End If
        If InStr(1, UCase(lineStr), "TOTAL WEIGHT PLATES:") > 0 Then
            tmpWeightStr = Mid(lineStr, InStr(1, UCase(lineStr), "TOTAL WEIGHT PLATES:") + 20)
            tmpWeightStr = Replace(UCase(tmpWeightStr), "KG", "")
            tmpWeightStr = Replace(tmpWeightStr, "=", "")
            fileWeight = fileWeight + val(Trim(Replace(tmpWeightStr, ",", ".")))
End If
        
        If InStr(1, UCase(lineStr), "PRODUCTION SUMMARY") > 0 Or InStr(1, UCase(lineStr), "SUMMARY QUANTITIES") > 0 Or InStr(1, UCase(lineStr), "WEIGHT TOWER") > 0 Then
            Exit Do
End If
        If LCase(lineStr) Like "*profile*pos*" Then currentSection = "PROFILE"
        If LCase(lineStr) Like "*plate*pos*" Then currentSection = "PLATE"
        If LCase(lineStr) Like "*bolts*included*" Or LCase(lineStr) Like "*nuts*included*" Then currentSection = "HARDWARE"

        If (currentSection = "PROFILE" Or currentSection = "PLATE") And qualRegexGlobal.Test(lineStr) And Not LCase(lineStr) Like "* m *" And Not LCase(lineStr) Like "*total*" Then
            loopGuard = 0
            Do While InStr(lineStr, "  ") > 0
                lineStr = Replace(lineStr, "  ", " ")
                loopGuard = loopGuard + 1
                If loopGuard > 50 Then Exit Do
            Loop
            parts = Split(lineStr, " ")
            qIdx = -1
            For i = 0 To UBound(parts)
                If qualRegexGlobal.Test(parts(i)) Then
                    qIdx = i
                    Exit For
End If
            Next i

            If qIdx >= 2 And UBound(parts) >= 2 Then
                rawTypeUpper = UCase(parts(0))
                posNo = parts(1)
                quality = UCase(parts(qIdx))
                xsrNote = ""
                If (qIdx + 1) <= UBound(parts) Then
                    If UCase(parts(qIdx + 1)) = "Y" Then xsrNote = "M.D"
                End If
                
                quantIdx = qIdx + 1
                If quantIdx <= UBound(parts) Then
                    If UCase(parts(quantIdx)) = "Y" Then quantIdx = quantIdx + 1
                    If quantIdx <= UBound(parts) Then
                        If Not IsNumeric(Replace(parts(quantIdx), ",", ".")) Then quantIdx = quantIdx + 1
End If
                End If

                quantVal = 0
                lengthVal = 0
                If quantIdx <= UBound(parts) Then quantVal = CLng(val(parts(quantIdx)))
                If (quantIdx + 1) <= UBound(parts) Then lengthVal = CDbl(Replace(parts(quantIdx + 1), ",", "."))
                
                If currentSection = "PROFILE" And Not rawTypeUpper Like "KF*" Then
                    If Not (rawTypeUpper Like "R*" Or rawTypeUpper Like "M*") Then
                        cleanNoSpace = Replace(rawTypeUpper, " ", "")
                        If Not ProfileKnown(rawTypeUpper, dictProfileLib) Then
                            If Not dictFeedback.Exists(rawTypeUpper) Then
                                dictFeedback.Add rawTypeUpper, fileName
                                ' (kütüphaneye ekleme artýk OGRENME sayfasýndan yapýlýr)
End If
                        End If
                        
                        matCodeStr = GetProfileMaterialCode(parts(0), dictProfileLib)
                        If isMerge Then
                            posKeyAngle = MakePosKeyAngle(posNo, matCodeStr, lengthVal)
Else
                            posKeyAngle = MakePosKeyAngle(posNo, matCodeStr, lengthVal) & "|" & fileName
End If
                        
                        If Not isDryRun Then
                            If Not dictPosAngle.Exists(posKeyAngle) Then
                                Call SafeWrite(curWsAngle, targetRowAngle, 1, FileTypeLabel(fileName))
                                Call SafeWrite(curWsAngle, targetRowAngle, 2, posNo)
                                Call SafeWrite(curWsAngle, targetRowAngle, 3, matCodeStr)
                                Call SafeWrite(curWsAngle, targetRowAngle, 5, quality)
                                Call SafeWrite(curWsAngle, targetRowAngle, 6, lengthVal)
                                Call SafeWrite(curWsAngle, targetRowAngle, currentColAngle, quantVal)
                                If xsrNote <> "" Then Call SafeWrite(curWsAngle, targetRowAngle, 57, xsrNote)
                                dictPosAngle.Add posKeyAngle, targetRowAngle
                                targetRowAngle = targetRowAngle + 1
Else
                                existRow = dictPosAngle(posKeyAngle)
                                Call AddTypeLabel(curWsAngle, existRow, True, fileName)
                                Call CheckPosConflict("ANGLE", existRow, posNo, matCodeStr, 0, 0, lengthVal, fileName, "XSR", sourceLineNo)
                                aQty = val(bufAngle(existRow, currentColAngle))
                                Call SafeWrite(curWsAngle, existRow, currentColAngle, aQty + quantVal)
End If
                        End If
                        totalProcessed = totalProcessed + 1
End If
                ElseIf currentSection = "PLATE" Or rawTypeUpper Like "KF*" Then
                    If Not (rawTypeUpper Like "L*" And Not rawTypeUpper Like "KF*") Then
                        Call ParsePlateDimensions(parts(0), pThick, pWidth, pNoteX)
                        finalXsrPlateNote = pNoteX
                        If xsrNote <> "" Then
                            If finalXsrPlateNote <> "" Then
                                finalXsrPlateNote = finalXsrPlateNote & ", " & xsrNote
Else
                                finalXsrPlateNote = xsrNote
End If
                        End If
                        
                        If isMerge Then
                            posKeyPlate = MakePosKeyPlate(posNo, pThick, pWidth, lengthVal)
Else
                            posKeyPlate = MakePosKeyPlate(posNo, pThick, pWidth, lengthVal) & "|" & fileName
End If
                        
                        If Not isDryRun Then
                            If Not dictPosPlate.Exists(posKeyPlate) Then
                                Call SafeWrite(curWsPlate, targetRowPlate, 1, FileTypeLabel(fileName))
                                Call SafeWrite(curWsPlate, targetRowPlate, 2, posNo)
                                Call SafeWrite(curWsPlate, targetRowPlate, 3, pThick)
                                Call SafeWrite(curWsPlate, targetRowPlate, 4, pWidth)
                                Call SafeWrite(curWsPlate, targetRowPlate, 5, quality)
                                Call SafeWrite(curWsPlate, targetRowPlate, 6, lengthVal)
                                Call SafeWrite(curWsPlate, targetRowPlate, currentColPlate, quantVal)
                                If finalXsrPlateNote <> "" Then Call SafeWrite(curWsPlate, targetRowPlate, 56, finalXsrPlateNote)
                                dictPosPlate.Add posKeyPlate, targetRowPlate
                                targetRowPlate = targetRowPlate + 1
Else
                                existRow = dictPosPlate(posKeyPlate)
                                Call AddTypeLabel(curWsPlate, existRow, False, fileName)
                                Call CheckPosConflict("PLATE", existRow, posNo, "", pThick, pWidth, lengthVal, fileName, "XSR", sourceLineNo)
                                pQty = val(bufPlate(existRow, currentColPlate))
                                Call SafeWrite(curWsPlate, existRow, currentColPlate, pQty + quantVal)
End If
                        End If
                        totalProcessed = totalProcessed + 1
End If
                End If
End If
            
        ElseIf currentSection = "HARDWARE" Then
            If lineStr <> "" And Not LCase(lineStr) Like "*included*" And Not LCase(lineStr) Like "*---*" Then
                loopGuard = 0
                Do While InStr(lineStr, "  ") > 0
                    lineStr = Replace(lineStr, "  ", " ")
                    loopGuard = loopGuard + 1
                    If loopGuard > 50 Then Exit Do
                Loop
                parts = Split(lineStr, " ")
                If UBound(parts) >= 1 Then
                    rawLineUpper = UCase(lineStr)
                    
                    If Not (rawLineUpper Like "*NUT*" Or rawLineUpper Like "*MUTTER*" Or rawLineUpper Like "*WASHER*" Or rawLineUpper Like "*SCHEIBE*" Or rawLineUpper Like "*FEDERRING*") Or (rawLineUpper Like "*FUTTER*" Or rawLineUpper Like "*FU-RING*") Then
                        isBoltItem = False
                        If CheckArrayMatch(rawLineUpper, hwKeywords) Then
                            isBoltItem = True
                        ElseIf InStr(1, rawLineUpper, "M") > 0 Or InStr(1, rawLineUpper, "X") > 0 Or InStr(1, rawLineUpper, "*") > 0 Then
                            isBoltItem = True
End If
                        
                        If isBoltItem Then
                            bQtyX = 0
                            For i = UBound(parts) To 0 Step -1
                                If IsNumeric(parts(i)) Then
                                    If val(parts(i)) > 0 And InStr(parts(i), ".") = 0 And InStr(parts(i), ",") = 0 Then
                                        isInsideParens = False
                                        If i > 0 Then
                                            If parts(i - 1) = "(" Then isInsideParens = True
End If
                                        If i < UBound(parts) Then
                                            If parts(i + 1) = ")" Then isInsideParens = True
End If
                                        If Not isInsideParens Then
                                            bQtyX = CLng(val(parts(i)))
                                            qIdx = i
                                            Exit For
End If
                                    End If
End If
                            Next i
                            
                            If bQtyX > 0 Then
                                finalBoltQty = IIf(extraPct > 0, bQtyX + Int((bQtyX * (extraPct / 100#)) + 0.5), bQtyX)
                                boltQualX = "8.8"
                                For i = 0 To UBound(parts)
                                    If i <> qIdx Then
                                        If parts(i) Like "*8.8*" Or parts(i) Like "*10.9*" Or parts(i) Like "*5.6*" Then
                                            boltQualX = parts(i)
                                            Exit For
End If
                                    End If
                                Next i
                                
                                Call ParseBoltSpecs(lineStr, dValX, lValX, hwCodeX, hwNameX, isFuRingX, boltNoteX, boltQualX, boltQualX)
                                If dValX > 0 Then
                                    boltKeyX = Format(dValX, "00") & "|" & Format(lValX, "0000") & "|" & hwCodeX & "|" & boltQualX
                                    If Not dictBoltList.Exists(boltKeyX) Then
                                        dictBoltList.Add boltKeyX, dValX & "^" & lValX & "^" & hwCodeX & "^" & hwNameX & "^" & boltQualX & "^" & CStr(isFuRingX) & "^" & boltNoteX
End If
                                    qtyKeyX = boltKeyX & "|" & currentColBolt
                                    dictBoltQty(qtyKeyX) = IIf(dictBoltQty.Exists(qtyKeyX), dictBoltQty(qtyKeyX) + finalBoltQty, finalBoltQty)
                                        If Not dictDiameters.Exists(CStr(dValX)) Then dictDiameters.Add CStr(dValX), dValX
                                    totalProcessed = totalProcessed + 1
End If
                            End If
End If
                    End If
End If
            End If
End If
    Loop
    Close #fileNo
End Sub

Public Function IsAssemblyPartListHeader(ByVal txt As String) As Boolean
    Dim u As String
    u = UCase$(Trim$(txt))
    IsAssemblyPartListHeader = (InStr(1, u, "ASSEMBLY PART LIST", vbTextCompare) > 0) Or _
                               (InStr(1, u, "ASSEMBLY   PART", vbTextCompare) > 0 And InStr(1, u, "LENGTH(MM)", vbTextCompare) > 0) Or _
                               (InStr(1, u, "ASSEMBLY", vbTextCompare) > 0 And InStr(1, u, "PART", vbTextCompare) > 0 And InStr(1, u, "GRADE", vbTextCompare) > 0 And InStr(1, u, "LENGTH(MM)", vbTextCompare) > 0)
End Function

Public Function ParseAssemblyPartListRow(ByVal txt As String, ByRef assemblyPos As String, ByRef partPos As String, ByRef qty As Long, ByRef profile As String, ByRef grade As String, ByRef lengthMm As Double, ByRef weightKg As Double) As Boolean
    Dim t As String, rx As Object, m As Object
    On Error GoTo Fail
    assemblyPos = ""
    partPos = ""
    qty = 0
    profile = ""
    grade = ""
    lengthMm = 0
    weightKg = 0
    
    t = Trim$(txt)
    If t = "" Then Exit Function
    If Left$(t, 3) = "---" Then Exit Function
    If InStr(1, UCase$(t), "TEKLA STRUCTURES", vbTextCompare) > 0 Then Exit Function
    If InStr(1, UCase$(t), "TOTAL:", vbTextCompare) > 0 Then Exit Function
    If InStr(1, UCase$(t), "ASSEMBLY", vbTextCompare) > 0 And InStr(1, UCase$(t), "PART", vbTextCompare) > 0 Then Exit Function
    
    Set rx = CreateObject("VBScript.RegExp")
    rx.Global = False
    rx.IgnoreCase = True
    rx.Pattern = "^\s*(\S+)\s+(\S+)\s+([0-9]+)\s+(.+?)\s+([A-Za-z][A-Za-z0-9\+\-]*)\s+([0-9]+(?:[\.,][0-9]+)?)\s+([0-9]+(?:[\.,][0-9]+)?)\s*$"
    If rx.Test(t) Then
        Set m = rx.Execute(t)(0)
        assemblyPos = Trim$(m.SubMatches(0))
        partPos = Trim$(m.SubMatches(1))
        qty = CLng(val(m.SubMatches(2)))
        profile = Trim$(m.SubMatches(3))
        grade = Trim$(m.SubMatches(4))
        lengthMm = val(Replace(m.SubMatches(5), ",", "."))
        weightKg = val(Replace(m.SubMatches(6), ",", "."))
        If partPos = "" Or qty <= 0 Or profile = "" Or lengthMm <= 0 Then Exit Function
        ParseAssemblyPartListRow = True
        Exit Function
End If
    
    rx.Pattern = "^\s*(\S+)\s+([0-9]+)\s+(.+?)\s+([A-Za-z][A-Za-z0-9\+\-]*)\s+([0-9]+(?:[\.,][0-9]+)?)\s+([0-9]+(?:[\.,][0-9]+)?)\s*$"
    If rx.Test(t) Then
        Set m = rx.Execute(t)(0)
        assemblyPos = ""
        partPos = Trim$(m.SubMatches(0))
        qty = CLng(val(m.SubMatches(1)))
        profile = Trim$(m.SubMatches(2))
        grade = Trim$(m.SubMatches(3))
        lengthMm = val(Replace(m.SubMatches(4), ",", "."))
        weightKg = val(Replace(m.SubMatches(5), ",", "."))
        If partPos = "" Or qty <= 0 Or profile = "" Or lengthMm <= 0 Then Exit Function
        ParseAssemblyPartListRow = True
End If
    Exit Function
Fail:
    ParseAssemblyPartListRow = False
End Function

Public Sub ParseAssemblyPartPlateDimensions(ByVal rawProfile As String, ByRef thicknessMm As Double, ByRef widthMm As Double, ByRef issueText As String)
    Dim t As String, body As String, sepPos As Long, a As String, b As String
    thicknessMm = 0
    widthMm = 0
    issueText = ""
    On Error GoTo Fail
    
    t = UCase$(Trim$(rawProfile))
    t = Replace(t, Chr(34), "")
    t = Replace(t, "×", "*")
    t = Replace(t, "X", "*")
    
    If Left$(t, 2) = "PL" Then
        body = Trim$(Mid$(t, 3))
    ElseIf Left$(t, 4) = "FLAT" Then
        body = Trim$(Mid$(t, 5))
    ElseIf Left$(t, 2) = "FL" Then
        body = Trim$(Mid$(t, 3))
Else
        Exit Sub
End If
    
    sepPos = InStr(1, body, "*", vbTextCompare)
    If sepPos <= 1 Then Exit Sub
    
    a = Trim$(Left$(body, sepPos - 1))
    b = Trim$(Mid$(body, sepPos + 1))
    
    If IsNumeric(Replace(a, ",", ".")) And IsNumeric(Replace(b, ",", ".")) Then
        thicknessMm = CDbl(Replace(a, ",", "."))
        widthMm = CDbl(Replace(b, ",", "."))
Else
        issueText = AddIssue(issueText, "Metric PLATE ölçüsü çözülemedi.")
End If
    Exit Sub
Fail:
    thicknessMm = 0
    widthMm = 0
    issueText = AddIssue(issueText, "Metric PLATE ölçüsü çözülemedi.")
End Sub

Public Sub DetermineAssemblyPartType(ByVal rawProfile As String, ByRef isPlate As Boolean, ByRef isProfile As Boolean)
    Dim t As String
    t = UCase$(Trim$(rawProfile))
    t = Replace(t, " ", "")
    isPlate = False
    isProfile = False
    If Left$(t, 2) = "PL" Or Left$(t, 4) = "FLAT" Or Left$(t, 2) = "FL" Or Left$(t, 2) = "BL" Then
        isPlate = True
        Exit Sub
End If
    If Left$(t, 1) = "L" Or Left$(t, 2) = "UN" Or Left$(t, 2) = "HE" Or Left$(t, 3) = "IPE" Or Left$(t, 3) = "PFC" Or Left$(t, 2) = "UB" Or Left$(t, 2) = "UC" Or Left$(t, 1) = "D" Then
        isProfile = True
End If
End Sub

Public Function GetAssemblyPartMaterialCode(ByVal rawProfile As String, ByRef dictLib As Object, Optional ByRef approximateMatch As Boolean = False) As String
    Dim t As String, key As String
    Dim leg1Mm As Double, leg2Mm As Double, thickMm As Double, issue As String
    Dim bestCode As String, bestScore As Double
    
    approximateMatch = False
    On Error GoTo Fallback
    t = UCase$(Trim$(rawProfile))
    
    GetAssemblyPartMaterialCode = GetProfileMaterialCode(t, dictLib)
    If GetAssemblyPartMaterialCode <> "" And GetAssemblyPartMaterialCode <> "0" Then Exit Function
    
    If Left$(t, 1) = "L" Then
        If ParseAssemblyAngleDimensions(t, leg1Mm, leg2Mm, thickMm, issue) Then
            If FindClosestProfileLibraryCode(dictLib, leg1Mm, leg2Mm, thickMm, True, bestCode, bestScore) Then
                GetAssemblyPartMaterialCode = bestCode
                approximateMatch = True
                Exit Function
End If
        End If
End If
    
    key = Replace(t, " ", "")
    If dictLib.Exists(key) Then
        GetAssemblyPartMaterialCode = dictLib(key)
        Exit Function
End If
Fallback:
    If GetAssemblyPartMaterialCode = "" Then GetAssemblyPartMaterialCode = ""
End Function

Public Function IsAssemblyListHeader(ByVal txt As String) As Boolean
    Dim u As String
    u = UCase$(Trim$(txt))
    IsAssemblyListHeader = ((InStr(1, u, "ASSEMBLY", vbTextCompare) > 0 Or InStr(1, u, "ASMBLY", vbTextCompare) > 0 Or InStr(1, u, "ASMLY", vbTextCompare) > 0) And _
                            InStr(1, u, "POS.", vbTextCompare) > 0 And InStr(1, u, "PROFILE", vbTextCompare) > 0 And InStr(1, u, "WEIGHT", vbTextCompare) > 0)
End Function

Public Function GetLastNumericValue(ByVal txt As String) As Double
    Dim rx As Object, m As Object
    On Error GoTo Fail
    Set rx = CreateObject("VBScript.RegExp")
    rx.Global = True
    rx.IgnoreCase = True
    rx.Pattern = "([0-9]+(?:[\.,][0-9]+)?)\s*$"
    If rx.Test(txt) Then
        Set m = rx.Execute(txt)(0)
        GetLastNumericValue = val(Replace(m.SubMatches(0), ",", "."))
End If
    Exit Function
Fail:
    GetLastNumericValue = 0
End Function

Public Function ParseAssemblyListRow(ByVal txt As String, ByRef posNo As String, ByRef qty As Long, ByRef itemName As String, ByRef profile As String, ByRef area As Double, ByRef weight As Double) As Boolean
    Dim t As String, rest As String, p As Long
    Dim rx As Object, matches As Object, firstToken As String, qtyToken As String
    On Error GoTo Fail
    
    t = Trim$(txt)
    If t = "" Then Exit Function
    If Left$(t, 1) = "-" Then Exit Function
    If InStr(1, UCase$(t), "ASMBLY POS.", vbTextCompare) > 0 Then Exit Function
    If InStr(1, UCase$(t), "TOTAL FOR", vbTextCompare) > 0 Then Exit Function
    If InStr(1, UCase$(t), "TEKLA STRUCTURES", vbTextCompare) > 0 Then Exit Function
    
    p = InStr(1, t, " ")
    If p <= 1 Then Exit Function
    firstToken = Trim$(Left$(t, p - 1))
    rest = Trim$(Mid$(t, p + 1))
    
    p = InStr(1, rest, " ")
    If p <= 1 Then Exit Function
    qtyToken = Trim$(Left$(rest, p - 1))
    rest = Trim$(Mid$(rest, p + 1))
    
    If Not IsNumeric(qtyToken) Then Exit Function
    qty = CLng(val(qtyToken))
    If qty <= 0 Then Exit Function
    posNo = firstToken
    
    Set rx = CreateObject("VBScript.RegExp")
    rx.Global = False
    rx.IgnoreCase = True
    rx.Pattern = "^(.*)\s+([0-9]+(?:[\.,][0-9]+)?)\s+([0-9]+(?:[\.,][0-9]+)?)\s*$"
    If Not rx.Test(rest) Then Exit Function
    
    Set matches = rx.Execute(rest)(0)
    profile = Trim$(matches.SubMatches(0))
    area = val(Replace(matches.SubMatches(1), ",", "."))
    weight = val(Replace(matches.SubMatches(2), ",", "."))
    
    itemName = ""
    t = NormalizeBOMText(profile)
    If Left$(t, 5) = "BEAM " Then
        itemName = "BEAM"
        profile = Trim$(Mid$(t, 6))
    ElseIf Left$(t, 6) = "PLATE " Then
        itemName = "PLATE"
        profile = Trim$(Mid$(t, 7))
Else
        profile = Trim$(t)
End If
    
    If profile = "" Then Exit Function
    ParseAssemblyListRow = True
    Exit Function
Fail:
    ParseAssemblyListRow = False
End Function

Public Sub DetermineAssemblyItemType(ByVal rawProfile As String, ByRef isPlate As Boolean, ByRef isAngle As Boolean)
    Dim t As String
    t = NormalizeBOMText(rawProfile)
    isPlate = False
    isAngle = False
    If Left$(t, 2) = "PL" Or Left$(t, 4) = "FLAT" Or Left$(t, 2) = "FL" Then
        isPlate = True
    ElseIf Left$(t, 1) = "L" And Len(t) > 1 Then
        If IsNumeric(Mid$(t, 2, 1)) Then isAngle = True
End If
End Sub

Public Function MixedFractionToInches(ByVal txt As String, ByRef ok As Boolean) As Double
    Dim t As String, parts() As String, wholeVal As Double, fracParts() As String, n As Double, d As Double
    On Error GoTo Fail
    ok = False
    t = Trim$(txt)
    t = Replace(t, Chr(34), "")
    t = Replace(t, "IN", "", 1, -1, vbTextCompare)
    If t = "" Then Exit Function
    If InStr(1, t, "***", vbTextCompare) > 0 Then Exit Function
    
    parts = Split(t, " ")
    If UBound(parts) >= 1 Then
        If IsNumeric(parts(0)) Then wholeVal = CDbl(parts(0))
        fracParts = Split(parts(1), "/")
        If UBound(fracParts) = 1 Then
            If IsNumeric(fracParts(0)) And IsNumeric(fracParts(1)) Then
                n = CDbl(fracParts(0))
                d = CDbl(fracParts(1))
                If d <> 0 Then
                    MixedFractionToInches = wholeVal + n / d
                    ok = True
                    Exit Function
End If
            End If
End If
    End If
    
    If InStr(1, t, "/", vbTextCompare) > 0 Then
        fracParts = Split(t, "/")
        If UBound(fracParts) = 1 Then
            If IsNumeric(fracParts(0)) And IsNumeric(fracParts(1)) Then
                d = CDbl(fracParts(1))
                If d <> 0 Then
                    MixedFractionToInches = CDbl(fracParts(0)) / d
                    ok = True
                    Exit Function
End If
            End If
End If
    End If
    
    If IsNumeric(t) Then
        MixedFractionToInches = CDbl(t)
        ok = True
End If
    Exit Function
Fail:
    MixedFractionToInches = 0
    ok = False
End Function

Public Function ParseAssemblyDimensionPair(ByVal rawProfile As String, ByRef firstMm As Double, ByRef secondMm As Double, ByRef issueText As String) As Boolean
    Dim t As String, body As String, leftPart As String, rightPart As String
    Dim ok1 As Boolean, ok2 As Boolean, p As Long, prefixLen As Long
    Dim loopGuard As Long
    On Error GoTo Fail
    
    firstMm = 0
    secondMm = 0
    issueText = ""
    t = UCase$(Trim$(rawProfile))
    t = Replace(t, Chr(34), "")
    t = Replace(t, "×", "X")
    loopGuard = 0
    Do While InStr(t, "  ") > 0
        t = Replace(t, "  ", " ")
        loopGuard = loopGuard + 1
        If loopGuard > 50 Then Exit Do
    Loop
    
    If Left$(t, 2) = "PL" Then
        prefixLen = 2
    ElseIf Left$(t, 4) = "FLAT" Then
        prefixLen = 4
    ElseIf Left$(t, 1) = "L" Then
        prefixLen = 1
Else
        Exit Function
End If
    
    body = Trim$(Mid$(t, prefixLen + 1))
    p = InStr(1, body, "X", vbTextCompare)
    If p = 0 Then Exit Function
    
    leftPart = Trim$(Left$(body, p - 1))
    rightPart = Trim$(Mid$(body, p + 1))
    
    firstMm = MixedFractionToInches(leftPart, ok1) * 25.4
    secondMm = MixedFractionToInches(rightPart, ok2) * 25.4
    
    If Not ok1 Or Not ok2 Then
        issueText = AddIssue(issueText, "Inch/kesir ölçüsü tam çözülemedi.")
        ParseAssemblyDimensionPair = False
        Exit Function
End If
    
    ParseAssemblyDimensionPair = (firstMm > 0 And secondMm > 0)
    Exit Function
Fail:
    ParseAssemblyDimensionPair = False
End Function

Public Function ParseAssemblyAngleDimensions(ByVal rawProfile As String, ByRef leg1Mm As Double, ByRef leg2Mm As Double, ByRef thickMm As Double, ByRef issueText As String) As Boolean
    Dim t As String, body As String, parts() As String
    Dim ok1 As Boolean, ok2 As Boolean, ok3 As Boolean
    Dim loopGuard As Long
    On Error GoTo Fail
    
    leg1Mm = 0
    leg2Mm = 0
    thickMm = 0
    issueText = ""
    t = UCase$(Trim$(rawProfile))
    t = Replace(t, Chr(34), "")
    t = Replace(t, "×", "X")
    If Left$(t, 1) <> "L" Then Exit Function
    
    body = Trim$(Mid$(t, 2))
    loopGuard = 0
    Do While InStr(body, "  ") > 0
        body = Replace(body, "  ", " ")
        loopGuard = loopGuard + 1
        If loopGuard > 50 Then Exit Do
    Loop
    
    parts = Split(body, "X")
    If UBound(parts) = 1 Then
        leg1Mm = MixedFractionToInches(Trim$(parts(0)), ok1) * 25.4
        leg2Mm = leg1Mm
        thickMm = MixedFractionToInches(Trim$(parts(1)), ok2) * 25.4
        ok3 = ok2
    ElseIf UBound(parts) = 2 Then
        leg1Mm = MixedFractionToInches(Trim$(parts(0)), ok1) * 25.4
        leg2Mm = MixedFractionToInches(Trim$(parts(1)), ok2) * 25.4
        thickMm = MixedFractionToInches(Trim$(parts(2)), ok3) * 25.4
Else
        Exit Function
End If
    
    If Not ok1 Or Not ok2 Or Not ok3 Then
        issueText = AddIssue(issueText, "Angle inch/kesir ölçüsü tam çözülemedi.")
        Exit Function
End If
    
    ParseAssemblyAngleDimensions = (leg1Mm > 0 And leg2Mm > 0 And thickMm > 0)
    Exit Function
Fail:
    ParseAssemblyAngleDimensions = False
End Function

Public Sub ParseAssemblyPlateDimensions(ByVal rawProfile As String, ByRef thicknessMm As Double, ByRef widthMm As Double, ByRef issueText As String)
    Dim t As String
    thicknessMm = 0
    widthMm = 0
    issueText = ""
    
    t = UCase$(Trim$(rawProfile))
    If Left$(t, 2) = "PL" Then
        If ParseAssemblyDimensionPair(t, thicknessMm, widthMm, issueText) Then Exit Sub
        If InStr(1, t, "***", vbTextCompare) > 0 Then
            issueText = AddIssue(issueText, "Kaynakta *** karakteri bulundu; ölçü eksik.")
End If
    ElseIf Left$(t, 4) = "FLAT" Then
        t = Mid$(t, 5)
        t = Replace(t, "X", "*")
        If InStr(t, "*") > 0 Then
            thicknessMm = val(Left$(t, InStr(t, "*") - 1))
            widthMm = val(Mid$(t, InStr(t, "*") + 1))
            If thicknessMm <= 0 Or widthMm <= 0 Then issueText = AddIssue(issueText, "FLAT ölçüsü çözülemedi.")
End If
    End If
End Sub

Public Function SplitBeforeX(ByVal txt As String) As String
    Dim p As Long
    p = InStr(1, txt, "X", vbTextCompare)
    If p > 0 Then SplitBeforeX = Trim$(Left$(txt, p - 1))
End Function

Public Function SplitAfterX(ByVal txt As String) As String
    Dim p As Long
    p = InStr(1, txt, "X", vbTextCompare)
    If p > 0 Then SplitAfterX = Trim$(Mid$(txt, p + 1))
End Function

Public Function GetAssemblyProfileMaterialCode(ByVal rawProfile As String, ByRef dictLib As Object, Optional ByRef approximateMatch As Boolean = False) As String
    Dim t As String, leg1Mm As Double, leg2Mm As Double, thickMm As Double
    Dim bMm As Double, issue As String, k As String, candidates As Variant, i As Long
    Dim bestCode As String, bestScore As Double

    approximateMatch = False
    On Error GoTo Fallback
    t = UCase$(Trim$(rawProfile))

    If Left$(t, 1) = "L" And Len(t) > 1 Then
        If ParseAssemblyAngleDimensions(t, leg1Mm, leg2Mm, thickMm, issue) Then
            candidates = Array( _
                "L" & FormatNumberForKey(leg1Mm) & "*" & FormatNumberForKey(leg2Mm) & "*" & FormatNumberForKey(thickMm), _
                "L" & Format$(Round(leg1Mm, 1), "0.0") & "*" & Format$(Round(leg2Mm, 1), "0.0") & "*" & Format$(Round(thickMm, 1), "0.0"), _
                "L" & CStr(CLng(Round(leg1Mm, 0))) & "*" & CStr(CLng(Round(leg2Mm, 0))) & "*" & Format$(Round(thickMm, 1), "0.0"), _
                FormatNumberForKey(leg1Mm) & "*" & FormatNumberForKey(leg2Mm) & "*" & FormatNumberForKey(thickMm), _
                CStr(CLng(Round(leg1Mm, 0))) & "*" & CStr(CLng(Round(leg2Mm, 0))) & "*" & Format$(Round(thickMm, 1), "0.0") _
            )
            For i = LBound(candidates) To UBound(candidates)
                k = StandardizeProfileName(CStr(candidates(i)))
                If dictLib.Exists(k) Then
                    GetAssemblyProfileMaterialCode = dictLib(k)
                    Exit Function
End If
            Next i

            If FindClosestProfileLibraryCode(dictLib, leg1Mm, leg2Mm, thickMm, True, bestCode, bestScore) Then
                GetAssemblyProfileMaterialCode = bestCode
                approximateMatch = True
                Exit Function
End If
        End If

    ElseIf Left$(t, 2) = "PL" Then
        If ParseAssemblyDimensionPair(t, thickMm, bMm, issue) Then
            candidates = Array( _
                "PL" & FormatNumberForKey(thickMm) & "*" & FormatNumberForKey(bMm), _
                FormatNumberForKey(thickMm) & "*" & FormatNumberForKey(bMm) _
            )
            For i = LBound(candidates) To UBound(candidates)
                k = StandardizeProfileName(CStr(candidates(i)))
                If dictLib.Exists(k) Then
                    GetAssemblyProfileMaterialCode = dictLib(k)
                    Exit Function
End If
            Next i
End If

    ElseIf Left$(t, 4) = "FLAT" Then
        k = StandardizeProfileName(t)
        If dictLib.Exists(k) Then GetAssemblyProfileMaterialCode = dictLib(k)
End If

    Exit Function
Fallback:
    GetAssemblyProfileMaterialCode = ""
End Function

Public Function FindClosestProfileLibraryCode(ByRef dictLib As Object, ByVal d1 As Double, ByVal d2 As Double, _
                                               ByVal d3 As Double, ByVal isAngle As Boolean, _
                                               ByRef outCode As String, ByRef outScore As Double) As Boolean
    Dim keys As Variant, i As Long, k As String, parts() As String
    Dim vals(0 To 5) As Double, n As Long, j As Long
    Dim score As Double, a As Double, b As Double, c As Double, best As Double
    best = 1E+30
    Dim keyNorm As String

    outCode = ""
    outScore = 1E+30
    On Error GoTo Fail
    keys = dictLib.keys
    For i = LBound(keys) To UBound(keys)
        k = CStr(keys(i))
        keyNorm = Replace(k, ",", ".")
        parts = Split(keyNorm, "*")
        n = 0
        For j = LBound(parts) To UBound(parts)
            If IsNumeric(Trim$(parts(j))) Then
                If n <= 5 Then
                    vals(n) = CDbl(Trim$(parts(j)))
                    n = n + 1
End If
            End If
        Next j

        If isAngle And n >= 3 Then
            a = vals(0)
            b = vals(1)
            c = vals(2)
            score = Abs(a - d1) + Abs(b - d2) + (Abs(c - d3) * 2#)
            If score < best Then
                best = score
                If score < outScore Then
                    outScore = score
                    outCode = dictLib(keys(i))
End If
            End If
End If
    Next i

    If outCode <> "" And outScore <= 3# Then
        FindClosestProfileLibraryCode = True
End If
    Exit Function
Fail:
    FindClosestProfileLibraryCode = False
End Function

Public Function FormatNumberForKey(ByVal n As Double) As String
    Dim s As String
    s = Format$(n, "0.0####")
    Do While InStr(s, ",") > 0 And Right$(s, 1) = "0"
        s = Left$(s, Len(s) - 1)
    Loop
    If Right$(s, 1) = "," Then s = Left$(s, Len(s) - 1)
    s = Replace(s, ",", ".")
    FormatNumberForKey = s
End Function

Public Function EstimateAssemblyLength(ByVal rawProfile As String, ByVal totalWeight As Double, ByVal qty As Long, ByRef issueText As String) As Double
    Const STEEL_DENSITY_KG_MM3 As Double = 0.00000785
    Dim t As String, aIn As Double, wIn As Double
    Dim tMm As Double, areaMm2 As Double, pieceWeight As Double

    issueText = ""
    If qty <= 0 Or totalWeight <= 0 Then Exit Function
    pieceWeight = totalWeight / qty

    t = UCase$(Trim$(rawProfile))

    If Left$(t, 1) = "L" Then
        Dim leg1Mm As Double, leg2Mm As Double
        If ParseAssemblyAngleDimensions(t, leg1Mm, leg2Mm, tMm, issueText) Then
            areaMm2 = (leg1Mm * tMm) + (leg2Mm * tMm) - (tMm * tMm)
            If areaMm2 > 0 Then
                EstimateAssemblyLength = pieceWeight / (areaMm2 * STEEL_DENSITY_KG_MM3)
End If
        End If
    ElseIf Left$(t, 2) = "PL" Then
        If ParseAssemblyDimensionPair(t, tMm, wIn, issueText) Then
            areaMm2 = tMm * wIn
            If areaMm2 > 0 Then
                EstimateAssemblyLength = pieceWeight / (areaMm2 * STEEL_DENSITY_KG_MM3)
End If
        End If
    ElseIf Left$(t, 4) = "FLAT" Then
        t = Mid$(t, 5)
        If InStr(t, "X") > 0 Then
            aIn = val(Left$(t, InStr(t, "X") - 1))
            wIn = val(Mid$(t, InStr(t, "X") + 1))
            If aIn > 0 And wIn > 0 Then
                areaMm2 = aIn * wIn
                EstimateAssemblyLength = pieceWeight / (areaMm2 * STEEL_DENSITY_KG_MM3)
End If
        End If
End If

    If EstimateAssemblyLength > 0 Then EstimateAssemblyLength = Round(EstimateAssemblyLength, 1)
End Function
