Attribute VB_Name = "modBolts"
Option Explicit
Option Private Module

' =========================================================================
' CIVATA: ayrýþtýrma, somun/pul/rondela kodlarý, aðýrlýklar
' =========================================================================

Public Sub ParseBoltSpecs(ByVal rawName As String, ByRef dVal As Long, ByRef lVal As Long, ByRef hwCode As String, ByRef hwName As String, ByRef isFuRing As Boolean, ByRef boltNote As String, Optional ByVal rawQual As String = "", Optional ByRef boltQualOut As String = "")
    Dim temp As String, xPos As Long, strD As String, strL As String, cleanTempPlus3 As String
    Dim aVal As Double, bVal As Long, aMod As Double, codePrefix As String, nameStr As String, matches As Object
    Dim i As Long, sepX As String, isStepBolt As Boolean, lenAssumed As Boolean
    On Error Resume Next
    temp = UCase(Trim(rawName))
    temp = Replace(temp, ChrW(215), "X")
    temp = Replace(temp, ChrW(160), " ")
    isFuRing = False
    boltNote = ""
    boltQualOut = "5.6"
    dVal = 0
    lVal = 0
    
    If InStr(temp, "FUTTERRING") > 0 Or InStr(temp, "FU-RING") > 0 Or InStr(temp, "FUTTER") > 0 Or Left$(Replace(temp, " ", ""), 4) = "FUBL" Then
        isFuRing = True
        xPos = InStr(temp, "X")
        If xPos = 0 Then xPos = InStr(temp, Chr(42))
        If xPos > 1 Then
            strD = ""
            For i = xPos - 1 To 1 Step -1
                If IsNumeric(Mid(temp, i, 1)) Or Mid(temp, i, 1) = "." Or Mid(temp, i, 1) = "," Then
                    strD = Mid(temp, i, 1) & strD
                Else
                    If strD <> "" Then Exit For
                End If
            Next i
            aVal = Val(Replace(strD, ",", "."))
            strL = ""
            For i = xPos + 1 To Len(temp)
                If IsNumeric(Mid(temp, i, 1)) Then
                    strL = strL & Mid(temp, i, 1)
                Else
                    If strL <> "" Then Exit For
                End If
            Next i
            bVal = Val(strL)
            If aVal > 0 And bVal > 0 Then
                If aVal < 30 Then
                    aMod = aVal - 0.5
                    If aMod = Fix(aMod) Then
                        codePrefix = CStr(CLng(aMod))
                    Else
                        codePrefix = Replace(Replace(CStr(aMod), ".", ""), ",", "")
                    End If
                    nameStr = Replace(CStr(aMod), ".", ",")
                Else
                    aMod = aVal - 3
                    codePrefix = CStr(CLng(aMod))
                    nameStr = CStr(CLng(aMod))
                End If
                hwCode = codePrefix & CStr(bVal)
                hwName = "#" & CStr(bVal) & "/" & nameStr
                boltQualOut = "G26"
                dVal = 99
                lVal = bVal
                If Left$(Replace(temp, " ", ""), 4) = "FUBL" Or InStr(temp, "FUTTERBL") > 0 Then
                    boltNote = "Futterblech: delik " & Replace(CStr(aVal), ",", ".") & " mm, kalýnlýk " & CStr(bVal) & " mm"
                End If
                Exit Sub
            End If
        End If
    End If
    
    cleanTempPlus3 = Replace(temp, " ", "")
    If InStr(cleanTempPlus3, "+3/MU") > 0 Or InStr(temp, "+3 MU") > 0 Or InStr(temp, "+3 / MU") > 0 Or InStr(temp, "+3") > 0 Then
        boltNote = "3mm uzun imal edilecek"
    End If
    
    If globalRegBolt Is Nothing Then
        Set globalRegBolt = CreateObject("VBScript.RegExp")
        globalRegBolt.IgnoreCase = True
        globalRegBolt.Global = False
    End If
    
    sepX = "[\*xX\s\-\/]"
    isStepBolt = False
    lenAssumed = False
    
    ' A) DIN önce yazým (Alman direk listeleri): "M7990 16 45+3" = M16x45 DIN 7990
    globalRegBolt.Pattern = "M\s*(?:7990|7968|7969|931|933|601|558|6914|4014|4017)\s+(\d{1,2})(?!\d)\s*" & sepX & "\s*(\d{2,3})"
    If globalRegBolt.Test(temp) Then
        Set matches = globalRegBolt.Execute(temp)
        dVal = CLng(matches(0).SubMatches(0))
        lVal = CLng(matches(0).SubMatches(1))
    Else
        ' B) Basamak cývatasý: "STGB 27 310", "STB M16x160", "TB.M16x160", "Steigbolzen M16x160"
        globalRegBolt.Pattern = "^(?:STGB|STEIGB[A-Z]*|STB|TB\.?)\s*[\.\-]?\s*M?\s*(\d{1,2})(?!\d)\s*" & sepX & "?\s*(\d{2,3})"
        If globalRegBolt.Test(temp) Then
            Set matches = globalRegBolt.Execute(temp)
            dVal = CLng(matches(0).SubMatches(0))
            lVal = CLng(matches(0).SubMatches(1))
            isStepBolt = True
        Else
            ' C) Standart: M16x45, M16 x 45, M16*45, M16-45, M8x20
            globalRegBolt.Pattern = "M\s*(\d{1,2})(?!\d)\s*" & sepX & "\s*(\d{2,3})"
            If globalRegBolt.Test(temp) Then
                Set matches = globalRegBolt.Execute(temp)
                dVal = CLng(matches(0).SubMatches(0))
                lVal = CLng(matches(0).SubMatches(1))
            Else
                ' D) Sadece çap (boy yok): M16
                globalRegBolt.Pattern = "M\s*(\d{1,2})(?!\d)"
                If globalRegBolt.Test(temp) Then
                    Set matches = globalRegBolt.Execute(temp)
                    dVal = CLng(matches(0).SubMatches(0))
                    lVal = prmBoltLen
                    lenAssumed = True
                End If
            End If
        End If
    End If
    
    ' M5 - M64 dýþý = hatalý eþleþme (ör: "M1" gibi montaj etiketleri)
    If dVal < 5 Or dVal > 64 Then dVal = 0
    
    If dVal <= 0 Then
        dVal = 0
        lVal = 0
        hwCode = ""
        hwName = ""
        boltQualOut = ""
        Exit Sub
    End If
    
    If lVal <= 0 Then lVal = prmBoltLen: lenAssumed = True
    If InStr(temp, "10.9") > 0 Then
        boltQualOut = "10.9"
    ElseIf InStr(temp, "8.8") > 0 Then
        boltQualOut = "8.8"
    ElseIf InStr(temp, "5.6") > 0 Then
        boltQualOut = "5.6"
    ElseIf InStr(temp, "4.6") > 0 Then
        boltQualOut = "4.6"
    ElseIf IsValidBoltGrade(rawQual) Then
        boltQualOut = Replace(UCase$(Trim$(rawQual)), ",", ".")
    Else
        boltQualOut = prmBoltGrade
    End If
    
    If Not isStepBolt Then isStepBolt = RxTest(temp, "(^|[^A-Z])(STGB|STEIGB|STB)")
    
    If isStepBolt Then
        ' Kütüphanedeki basamak cývatasý kodu: 161607990SB / M16x160-DIN7990 S.BOLT
        hwCode = CStr(dVal) & CStr(lVal) & "7990SB"
        hwName = "M" & CStr(dVal) & "x" & CStr(lVal) & "-DIN7990 S.BOLT"
        isFuRing = Not prmStepBoltAutoSet
        If boltNote <> "" Then boltNote = boltNote & ", "
        boltNote = boltNote & "Basamak cývatasý"
    Else
        hwCode = CStr(dVal) & CStr(lVal) & "7990"
        hwName = "M" & CStr(dVal) & "x" & CStr(lVal) & "-DIN7990"
        If CheckArrayMatch(temp, Split("FU-RING,FUTTERRING,FUTTER,RING", ",")) Then isFuRing = True
    End If
    
    ' Baþýnda M olmayan özel parça: "SSM24x345-R" (ör: Grabmayr) -> kontrol notu
    If Not isStepBolt And RxTest(temp, "^[A-Z]{2,}M\s*\d") Then
        If boltNote <> "" Then boltNote = boltNote & ", "
        boltNote = boltNote & "Özel parça (kontrol): " & Trim(rawName)
    End If
    If lenAssumed Then
        If boltNote <> "" Then boltNote = boltNote & ", "
        boltNote = boltNote & "Boy bulunamadý, " & prmBoltLen & " varsayýldý"
    End If
    On Error GoTo 0
End Sub

Public Function GetNutCode(ByVal d As Long) As String
    GetNutCode = CStr(d) & "N4032"
End Function

Public Function GetNutName(ByVal d As Long) As String
    GetNutName = "M" & d & "NUT-ISO 4032"
End Function

Public Function GetNutWeight(ByVal d As Long) As Double
    Select Case d
        Case 12: GetNutWeight = 0.0153
        Case 16: GetNutWeight = 0.0365
        Case 20: GetNutWeight = 0.0693
        Case 24: GetNutWeight = 0.119
        Case 27: GetNutWeight = 0.174
        Case 30: GetNutWeight = 0.242
        Case Else: GetNutWeight = 0.02
    End Select
End Function

Public Function GetPWCode(ByVal d As Long) As String
    GetPWCode = CStr(d) & "PW7989 #6"
End Function

Public Function GetPWName(ByVal d As Long) As String
    GetPWName = "M" & d & "PW-DIN7989 #6"
End Function

Public Function GetPWWeight(ByVal d As Long) As Double
    Select Case d
        Case 12: GetPWWeight = 0.0141
        Case 16: GetPWWeight = 0.0177
        Case 20: GetPWWeight = 0.0273
        Case 24: GetPWWeight = 0.0384
        Case 27: GetPWWeight = 0.054
        Case 30: GetPWWeight = 0.075
        Case Else: GetPWWeight = 0.02
    End Select
End Function

Public Function GetSWCode(ByVal d As Long) As String
    GetSWCode = CStr(d) & "SW128"
End Function

Public Function GetSWName(ByVal d As Long) As String
    GetSWName = "M" & d & "SW-DIN128"
End Function

Public Function GetSWWeight(ByVal d As Long) As Double
    Select Case d
        Case 12: GetSWWeight = 0.0032
        Case 16: GetSWWeight = 0.007
        Case 20: GetSWWeight = 0.0122
        Case 24: GetSWWeight = 0.0215
        Case 27: GetSWWeight = 0.031
        Case 30: GetSWWeight = 0.044
        Case Else: GetSWWeight = 0.005
    End Select
End Function

Public Function GetBoltWeight(ByVal d As Long, ByVal l As Long) As Double
    Select Case d
        Case 12
            Select Case l
                Case 35: GetBoltWeight = 0.0473
                Case 40: GetBoltWeight = 0.0517
                Case 45: GetBoltWeight = 0.0561
                Case 50: GetBoltWeight = 0.0605
                Case Else: GetBoltWeight = EstimateBoltWeight(d, l)
            End Select
        Case 16
            Select Case l
                Case 40: GetBoltWeight = 0.0962
                Case 45: GetBoltWeight = 0.1042
                Case 50: GetBoltWeight = 0.1122
                Case 55: GetBoltWeight = 0.1202
                Case 60: GetBoltWeight = 0.1282
                Case 65: GetBoltWeight = 0.1362
                Case Else: GetBoltWeight = EstimateBoltWeight(d, l)
            End Select
        Case 20
            Select Case l
                Case 45: GetBoltWeight = 0.1917
                Case 50: GetBoltWeight = 0.2037
                Case 55: GetBoltWeight = 0.2157
                Case 60: GetBoltWeight = 0.2277
                Case 65: GetBoltWeight = 0.2397
                Case Else: GetBoltWeight = EstimateBoltWeight(d, l)
            End Select
        Case 24
            Select Case l
                Case 50: GetBoltWeight = 0.292
                Case 55: GetBoltWeight = 0.31
                Case 60: GetBoltWeight = 0.328
                Case 65: GetBoltWeight = 0.346
                Case Else: GetBoltWeight = EstimateBoltWeight(d, l)
            End Select
        Case 27
            Select Case l
                Case 50: GetBoltWeight = 0.5785
                Case 60: GetBoltWeight = 0.6225
                Case 70: GetBoltWeight = 0.6665
                Case 75: GetBoltWeight = 0.6885
                Case 90: GetBoltWeight = 0.7545
                Case Else: GetBoltWeight = EstimateBoltWeight(d, l)
            End Select
        Case Else
            GetBoltWeight = EstimateBoltWeight(d, l)
    End Select
End Function

' Kütüphanede/tabloda olmayan cývata için yaklaþýk aðýrlýk (gövde + altýgen baþ), kg
Public Function EstimateBoltWeight(ByVal d As Long, ByVal l As Long) As Double
    Dim shank As Double, head As Double
    If d <= 0 Or l <= 0 Then EstimateBoltWeight = 0.1: Exit Function
    shank = 3.14159265 / 4# * d * d * l * 0.00000785
    head = 0.866 * (1.5 * d) ^ 2 * (0.7 * d) * 0.00000785
    EstimateBoltWeight = Round(shank + head, 4)
End Function

Public Function BoltInputHasExplicitSize(ByVal rawText As String) As Boolean
    Dim rx As Object
    On Error GoTo Fail
    Set rx = CreateObject("VBScript.RegExp")
    rx.IgnoreCase = True
    rx.Global = False
    rx.Pattern = "M\s*\d{2,3}\s*[\*xX\-\/]\s*\d{2,3}"
    BoltInputHasExplicitSize = rx.Test(UCase$(rawText))
    Exit Function
Fail:
    BoltInputHasExplicitSize = False
End Function

Public Function IsValidBoltGrade(ByVal gradeText As String) As Boolean
    Dim g As String
    g = Replace(UCase$(Trim$(gradeText)), ",", ".")
    Select Case g
        Case "4.6", "5.6", "8.8", "10.9", "G26"
            IsValidBoltGrade = True
        Case Else
            IsValidBoltGrade = False
    End Select
End Function

Public Function LibraryContainsBoltCode(ByVal code As String, ByRef dictLib As Object) As Boolean
    On Error GoTo Fail
    If Trim$(code) = "" Then Exit Function
    LibraryContainsBoltCode = dictLib.Exists(Trim$(code))
    Exit Function
Fail:
    LibraryContainsBoltCode = False
End Function
