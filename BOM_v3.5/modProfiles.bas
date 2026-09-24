Attribute VB_Name = "modProfiles"
Option Explicit
Option Private Module

' =========================================================================
' PROFÝL / PLAKA: kütüphane kodu, ölçü ayrýþtýrma
' =========================================================================

Public Function GetProfileMaterialCode(ByVal typeStr As String, ByRef dictLib As Object) As String
    Dim temp As String, stdKey As String, cleanNoSpace As String
    On Error Resume Next
    Dim lkFound As Boolean, lkCode As String, ckFb As String
    lkCode = LookupProfileCode(typeStr, dictLib, lkFound)
    If lkFound Then
        GetProfileMaterialCode = lkCode
        Exit Function
    End If
    temp = UCase(Trim(typeStr))
    cleanNoSpace = Replace(temp, " ", "")
    If dictLib.Exists(cleanNoSpace) Then
        GetProfileMaterialCode = dictLib(cleanNoSpace)
        Exit Function
End If
    stdKey = StandardizeProfileName(typeStr)
    If dictLib.Exists(stdKey) Then
        GetProfileMaterialCode = dictLib(stdKey)
        Exit Function
End If
    ' U280 -> UNP280 / UPE280 Akýllý Eþleme
    If Left(cleanNoSpace, 1) = "U" And Len(cleanNoSpace) > 1 Then
        If IsNumeric(Mid(cleanNoSpace, 2, 1)) Then
            If dictLib.Exists("UNP" & Mid(cleanNoSpace, 2)) Then
                GetProfileMaterialCode = dictLib("UNP" & Mid(cleanNoSpace, 2))
                Exit Function
End If
            If dictLib.Exists("UPE" & Mid(cleanNoSpace, 2)) Then
                GetProfileMaterialCode = dictLib("UPE" & Mid(cleanNoSpace, 2))
                Exit Function
End If
        End If
    ' I200 -> IPE200 / IPN200 Akýllý Eþleme
    ElseIf Left(cleanNoSpace, 1) = "I" And Len(cleanNoSpace) > 1 Then
        If IsNumeric(Mid(cleanNoSpace, 2, 1)) Then
            If dictLib.Exists("IPE" & Mid(cleanNoSpace, 2)) Then
                GetProfileMaterialCode = dictLib("IPE" & Mid(cleanNoSpace, 2))
                Exit Function
End If
            If dictLib.Exists("IPN" & Mid(cleanNoSpace, 2)) Then
                GetProfileMaterialCode = dictLib("IPN" & Mid(cleanNoSpace, 2))
                Exit Function
End If
        End If
End If
    ' Boru / kutu profil kütüphanede yoksa kütüphane biçiminde kod üret (CHS48,3x3) -> kýrmýzý + OGRENME
    ckFb = CanonicalSectionKey(typeStr)
    If RxTest(ckFb, "^(CFCHS|CFSHS|CFRHS|CHS|RHS|SHS)\d") Then
        If Left$(ckFb, 3) = "RHS" And UBound(Split(Mid$(ckFb, 4), "X")) = 1 Then ckFb = "CHS" & Mid$(ckFb, 4)
        GetProfileMaterialCode = RxFirst(ckFb, "^[A-Z]+") & Replace(Mid$(ckFb, Len(RxFirst(ckFb, "^[A-Z]+")) + 1), "X", "x")
        Exit Function
    End If
    GetProfileMaterialCode = FormatMaterialCode(typeStr)
    On Error GoTo 0
End Function

Public Function StandardizeProfileName(ByVal typeStr As String) As String
    Dim temp As String, parts() As String
    Dim loopGuard As Long
    temp = UCase(Trim(typeStr))
    
    If Left(temp, 2) = "L " Or Left(temp, 2) = "L." Then
        temp = Trim(Mid(temp, 3))
    ElseIf Left(temp, 1) = "L" And Len(temp) > 1 Then
        If IsNumeric(Mid(temp, 2, 1)) Then temp = Trim(Mid(temp, 2))
End If
    
    temp = Replace(temp, "?", "")
    temp = Replace(temp, " ", "")
    temp = Replace(temp, ".", ",")
    temp = Replace(temp, "X", Chr(42))
    temp = Replace(temp, "-", Chr(42))
    temp = Replace(temp, "/", Chr(42))
    
    loopGuard = 0
    Do While InStr(temp, Chr(42) & Chr(42)) > 0
        temp = Replace(temp, Chr(42) & Chr(42), Chr(42))
        loopGuard = loopGuard + 1
        If loopGuard > 50 Then Exit Do
    Loop
    
    parts = Split(temp, Chr(42))
    If UBound(parts) = 1 Then
        If IsNumToken(parts(0)) And IsNumToken(parts(1)) Then
            temp = parts(0) & Chr(42) & parts(0) & Chr(42) & parts(1)
End If
    End If
    
    StandardizeProfileName = temp
End Function

Public Function FormatMaterialCode(typeStr As String) As String
    Dim temp As String, combined As String, i As Long, parts() As String
    Dim loopGuard As Long
    temp = UCase(Trim(typeStr))
    If Left(temp, 1) = "L" Then temp = Mid(temp, 2)
    temp = Replace(temp, "?", "")
    temp = Replace(temp, ",", ".")
    temp = Replace(temp, "X", Chr(42))
    temp = Replace(temp, "-", Chr(42))
    temp = Replace(temp, "/", Chr(42))
    loopGuard = 0
    Do While InStr(temp, Chr(42) & Chr(42)) > 0
        temp = Replace(temp, Chr(42) & Chr(42), Chr(42))
        loopGuard = loopGuard + 1
        If loopGuard > 50 Then Exit Do
    Loop
    parts = Split(temp, Chr(42))
    If UBound(parts) = 2 Then
        If Trim(parts(0)) = Trim(parts(1)) Then
            combined = Trim(parts(0)) & Trim(parts(2))
Else
            combined = Trim(parts(0)) & Trim(parts(1)) & Trim(parts(2))
End If
    ElseIf UBound(parts) = 1 Then
        combined = Trim(parts(0)) & Trim(parts(1))
Else
        combined = ""
        For i = 1 To Len(temp)
            If IsNumeric(Mid(temp, i, 1)) Then combined = combined & Mid(temp, i, 1)
        Next i
End If
    If combined = "" Then combined = "0"
    FormatMaterialCode = CStr(Int(val(combined)))
End Function

Public Sub ParsePlateDimensions(ByVal typeStr As String, ByRef thickness As Double, ByRef width As Double, ByRef pNote As String)
    Dim temp As String, parts() As String, subParts() As String, i As Long, dimVals() As Double, valCount As Long, v1 As Double, v2 As Double
    Dim loopGuard As Long
    pNote = ""
    temp = UCase(Trim(typeStr))
    If InStr(temp, "PIPE") > 0 Or InStr(temp, "ROD") > 0 Or InStr(temp, "ROUND") > 0 Or InStr(temp, "Ø") > 0 Then
        pNote = typeStr
        thickness = 0
        width = 0
        Exit Sub
End If
    If InStr(temp, "CHK") > 0 Then
        pNote = "CHK'D PL"
        If InStr(temp, "PL") > 0 Then
            temp = Mid(temp, InStr(temp, "PL") + 2)
Else
            temp = Replace(temp, "CHK'D", "")
            temp = Replace(temp, "CHK", "")
End If
    ElseIf Left(temp, 3) = "BLE" Or Left(temp, 3) = "FLA" Or Left(temp, 3) = "PLT" Then
        temp = Mid(temp, 4)
    ElseIf Left(temp, 2) = "PL" Or Left(temp, 2) = "FL" Or Left(temp, 2) = "BL" Or Left(temp, 2) = "KF" Then
        temp = Mid(temp, 3)
    ElseIf Left(temp, 1) = "P" Then
        temp = Mid(temp, 2)
End If
    temp = Replace(temp, " ", "")
    temp = Replace(temp, ",", ".")
    temp = Replace(temp, "X", Chr(42))
    temp = Replace(temp, "-", Chr(42))
    loopGuard = 0
    Do While InStr(temp, Chr(42) & Chr(42)) > 0
        temp = Replace(temp, Chr(42) & Chr(42), Chr(42))
        loopGuard = loopGuard + 1
        If loopGuard > 50 Then Exit Do
    Loop
    parts = Split(temp, Chr(42))
    valCount = 0
    ReDim dimVals(0 To 30)
    For i = LBound(parts) To UBound(parts)
        If InStr(parts(i), "/") > 0 Then
            pNote = parts(i)
            subParts = Split(parts(i), "/")
            If UBound(subParts) >= 0 Then
                v1 = val(subParts(0))
                If UBound(subParts) >= 1 Then
                    v2 = val(subParts(1))
                    If v2 > v1 Then v1 = v2
End If
                dimVals(valCount) = v1
                valCount = valCount + 1
End If
Else
            If val(parts(i)) > 0 Then
                dimVals(valCount) = val(parts(i))
                valCount = valCount + 1
End If
        End If
    Next i
    thickness = 0
    width = 0
    If valCount >= 2 Then
        If dimVals(0) < dimVals(1) Then
            thickness = dimVals(0)
            width = dimVals(1)
Else
            thickness = dimVals(1)
            width = dimVals(0)
End If
    ElseIf valCount = 1 Then
        thickness = dimVals(0)
End If
End Sub

Public Sub AutoAppendToLibrary(wsLib As Worksheet, ByVal newProfile As String, Optional ByVal mappedCode As String = "YENI_EKLE")
    On Error Resume Next
    If Not wsLib Is Nothing Then
        Dim nextRow As Long
        nextRow = wsLib.Cells(wsLib.Rows.Count, 1).End(xlUp).Row + 1
        wsLib.Cells(nextRow, 1).Value = mappedCode
        wsLib.Cells(nextRow, 2).Value = newProfile
        wsLib.Cells(nextRow, 1).Interior.Color = RGB(255, 200, 200)
End If
    On Error GoTo 0
End Sub

' Kütüphane anahtarý: "CHS Ø 508x20" / "CHS508*20" / "CHS508X20,0" -> "CHS508X20"
Public Function CanonicalSectionKey(ByVal s As String) As String
    Dim t As String
    t = UCase$(Trim$(s))
    t = Replace(t, ChrW(216), "")
    t = Replace(t, " ", "")
    t = Replace(t, "*", "X")
    t = Replace(t, ChrW(215), "X")
    t = Replace(t, ".", ",")
    t = RxReplaceAll(t, ",0+(?=X|$)", "")
    CanonicalSectionKey = t
End Function

' Kütüphanede var mý? (boþluksuz ad, standart ad, kanonik ad, RHS->CHS)
Public Function LookupProfileCode(ByVal typeStr As String, ByRef dictLib As Object, ByRef found As Boolean) As String
    Dim temp As String, cleanNoSpace As String, stdKey As String, ck As String
    On Error Resume Next
    found = True
    temp = UCase$(Trim$(typeStr))
    cleanNoSpace = Replace(temp, " ", "")
    If dictLib.Exists(cleanNoSpace) Then LookupProfileCode = dictLib(cleanNoSpace): Exit Function
    stdKey = StandardizeProfileName(typeStr)
    If stdKey <> "" Then
        If dictLib.Exists(stdKey) Then LookupProfileCode = dictLib(stdKey): Exit Function
    End If
    ck = CanonicalSectionKey(typeStr)
    If ck <> "" Then
        If dictLib.Exists(ck) Then LookupProfileCode = dictLib(ck): Exit Function
        ' Müþteri boruyu "RHS48.3X3.0" diye yazabiliyor: iki ölçülü RHS = boru (CHS)
        If Left$(ck, 3) = "RHS" And UBound(Split(Mid$(ck, 4), "X")) = 1 Then
            If dictLib.Exists("CHS" & Mid$(ck, 4)) Then LookupProfileCode = dictLib("CHS" & Mid$(ck, 4)): Exit Function
        End If
    End If
    found = False
    LookupProfileCode = ""
End Function

Public Function ProfileKnown(ByVal typeStr As String, ByRef dictLib As Object) As Boolean
    Dim f As Boolean
    Call LookupProfileCode(typeStr, dictLib, f)
    ProfileKnown = f
End Function

