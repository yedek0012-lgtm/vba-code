Attribute VB_Name = "modText"
Option Explicit
Option Private Module

' =========================================================================
' METÝN / SAYI yardýmcýlarý (regex, sadeleþtirme, bölgeden baðýmsýz sayý)
' =========================================================================

' =========================================================================
' MOTOR B v2 - YARDIMCI FONKSÝYONLAR (Genel Excel okuyucu)
' =========================================================================

Public Sub InitRxHelper()
    If rxHelper Is Nothing Then
        Set rxHelper = CreateObject("VBScript.RegExp")
        rxHelper.IgnoreCase = True
        rxHelper.Global = False
    End If
End Sub

Public Function RxTest(ByVal txt As String, ByVal pattern As String) As Boolean
    On Error GoTo Fail
    InitRxHelper
    rxHelper.pattern = pattern
    RxTest = rxHelper.Test(txt)
    Exit Function
Fail:
    RxTest = False
End Function

Public Function RxFirst(ByVal txt As String, ByVal pattern As String) As String
    Dim mm As Object
    On Error GoTo Fail
    InitRxHelper
    rxHelper.pattern = pattern
    If rxHelper.Test(txt) Then
        Set mm = rxHelper.Execute(txt)
        RxFirst = mm(0).Value
    End If
    Exit Function
Fail:
    RxFirst = ""
End Function

' Büyük harf + Almanca/Türkçe karakter sadeleþtirme + tek boþluk
Public Function FoldText(ByVal txt As String) As String
    Dim t As String, loopGuard As Long
    t = UCase$(Trim$(txt))
    t = Replace(t, ChrW(160), " ")
    t = Replace(t, vbTab, " ")
    t = Replace(t, vbCr, " ")
    t = Replace(t, vbLf, " ")
    t = Replace(t, ChrW(196), "A")   ' Ä
    t = Replace(t, ChrW(214), "O")   ' Ö
    t = Replace(t, ChrW(220), "U")   ' Ü
    t = Replace(t, ChrW(223), "SS")  ' ß
    t = Replace(t, ChrW(7838), "SS") ' buyuk eszett
    t = Replace(t, ChrW(304), "I")   ' Ý
    t = Replace(t, ChrW(305), "I")   ' ý
    t = Replace(t, ChrW(350), "S")   ' Þ
    t = Replace(t, ChrW(199), "C")   ' Ç
    t = Replace(t, ChrW(286), "G")   ' Ð
    t = Replace(t, ChrW(215), "X")   ' ×
    loopGuard = 0
    Do While InStr(t, "  ") > 0
        t = Replace(t, "  ", " ")
        loopGuard = loopGuard + 1
        If loopGuard > 50 Then Exit Do
    Loop
    FoldText = Trim$(t)
End Function

' Sayýyý bölge ayarýndan baðýmsýz metne çevirir (5.6 -> "5.6", 403 -> "403")
Public Function NumToText(ByVal d As Double) As String
    Dim s As String
    If d = Fix(d) And Abs(d) < 1E+15 Then
        NumToText = Format$(d, "0")
    Else
        s = Trim$(Str$(d))
        If Left$(s, 1) = "." Then s = "0" & s
        If Left$(s, 2) = "-." Then s = "-0" & Mid$(s, 2)
        NumToText = s
    End If
End Function

' "1.234,5" / "1,234.5" / "12,5" / "4 Stk" -> sayý
' TEK SAYI OKUYUCU (v3.5): hücre sayýysa olduðu gibi, metinse ParseBOMNumber (bölge ayarýndan baðýmsýz).
' Tüm modüller hücreden sayý okurken bunu kullanýr (CDbl/Val metinde Türkçe/Ýngilizce ayara göre farklý sonuç verir).
Public Function CellNumber(ByVal c As Range) As Double
    Dim v As Variant
    On Error GoTo Fail
    v = c.Value
    If IsError(v) Or IsEmpty(v) Then Exit Function
    Select Case VarType(v)
        Case vbDouble, vbSingle, vbCurrency, vbLong, vbInteger, vbByte, vbDecimal
            CellNumber = CDbl(v)
        Case vbString
            CellNumber = ParseBOMNumber(CStr(v))
    End Select
    Exit Function
Fail:
    CellNumber = 0
End Function

' Metin tek baþýna bir sayý mý? ("12", "12.5", "12,5"; bölge ayarýndan baðýmsýz)
Public Function IsNumToken(ByVal s As String) As Boolean
    IsNumToken = RxTest(Trim$(s), "^[+-]?\d+([.,]\d+)?$")
End Function

Public Function ParseBOMNumber(ByVal s As String) As Double
    Dim t As String, lastDot As Long, lastCom As Long
    On Error GoTo Fail
    ' Ýlk sayýyý al; "850 101.5" gibi iki sayý birleþmiþse boþluðu SÝLME (850101,5 olurdu)
    s = Replace(s, ChrW(160), " ")
    t = RxFirst(s, "-?\d[\d\.,]*")
    If t = "" Then Exit Function
    ' Üst üste basýlmýþ iki sayý ("112115..76") okunamaz -> 0
    If InStr(t, "..") > 0 Or InStr(t, ",,") > 0 Then Exit Function
    Do While Right$(t, 1) = "." Or Right$(t, 1) = ","
        t = Left$(t, Len(t) - 1)
    Loop
    lastDot = InStrRev(t, ".")
    lastCom = InStrRev(t, ",")
    If lastDot > 0 And lastCom > 0 Then
        If lastCom > lastDot Then
            t = Replace(t, ".", "")
            t = Replace(t, ",", ".")
        Else
            t = Replace(t, ",", "")
        End If
    ElseIf lastCom > 0 Then
        If RxTest(t, "^-?\d{1,3}(,\d{3})+$") Then t = Replace(t, ",", "") Else t = Replace(t, ",", ".")
    ElseIf lastDot > 0 Then
        If Len(t) - Len(Replace(t, ".", "")) > 1 Then t = Replace(t, ".", "")
    End If
    ParseBOMNumber = Val(t)
    Exit Function
Fail:
    ParseBOMNumber = 0
End Function

Public Function NormalizeBOMText(ByVal txt As String) As String
    Dim t As String
    t = UCase$(Trim$(txt))
    t = Replace(t, ChrW(160), " ")
    t = Replace(t, "Ä", "A")
    t = Replace(t, "Ö", "O")
    t = Replace(t, "Ü", "U")
    t = Replace(t, "?", "SS")
    t = Replace(t, "ß", "SS")
    Dim loopGuard As Long
    loopGuard = 0
    Do While InStr(t, "  ") > 0
        t = Replace(t, "  ", " ")
        loopGuard = loopGuard + 1
        If loopGuard > 50 Then Exit Do
    Loop
    NormalizeBOMText = t
End Function

Public Function CheckArrayMatch(ByVal txt As String, ByVal arr As Variant) As Boolean
    Dim i As Long, normalizedTxt As String
    On Error GoTo ErrHand
    If Not IsArray(arr) Then Exit Function
    
    normalizedTxt = NormalizeBOMText(txt)
    For i = LBound(arr) To UBound(arr)
        If arr(i) <> "" Then
            ' Diziler LoadDictionaryArrays adýmýnda zaten normalize edildiði için doðrudan aranýr (Ciddi hýz artýþý)
            If InStr(1, normalizedTxt, CStr(arr(i)), vbTextCompare) > 0 Then
                CheckArrayMatch = True
                Exit Function
End If
        End If
    Next i
ErrHand:
    CheckArrayMatch = False
End Function

Public Function AddIssue(ByVal currentIssue As String, ByVal newIssue As String) As String
    If Trim$(newIssue) = "" Then
        AddIssue = currentIssue
    ElseIf Trim$(currentIssue) = "" Then
        AddIssue = newIssue
Else
        AddIssue = currentIssue & " | " & newIssue
End If
End Function

Public Function MinLong(ByVal a As Long, ByVal b As Long) As Long
    If a < b Then MinLong = a Else MinLong = b
End Function

Public Sub SortArrayInPlace(ByRef arr As Variant)
    Dim i As Long, j As Long, temp As Variant
    On Error Resume Next
    If Not IsArray(arr) Then Exit Sub
    
    For i = LBound(arr) To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If val(arr(i)) > val(arr(j)) Then
                temp = arr(i)
                arr(i) = arr(j)
                arr(j) = temp
            ElseIf val(arr(i)) = val(arr(j)) And CStr(arr(i)) > CStr(arr(j)) Then
                temp = arr(i)
                arr(i) = arr(j)
                arr(j) = temp
End If
        Next j
    Next i
    On Error GoTo 0
End Sub

Public Function RxReplaceAll(ByVal txt As String, ByVal pattern As String, ByVal repl As String) As String
    Dim rx As Object
    On Error GoTo Fail
    Set rx = CreateObject("VBScript.RegExp")
    rx.Global = True
    rx.IgnoreCase = True
    rx.pattern = pattern
    RxReplaceAll = rx.Replace(txt, repl)
    Exit Function
Fail:
    RxReplaceAll = txt
End Function

' Dosya adýndan çizim/montaj iþareti: "FAC51410A__Rev 03" -> "FAC51410A", "FA-F-A2C2P2 - Rev 05" -> "FA-F-A2C2P2"
Public Function FileMark(ByVal nm As String) As String
    Dim t As String, p As Long
    t = Trim$(nm)
    p = InStrRev(t, ".")
    If p > 0 And Len(t) - p <= 4 Then t = Left$(t, p - 1)
    t = RxReplaceAll(t, "[\s_\-]*rev[\s_\.\-]*\d+.*$", "")
    t = RxReplaceAll(t, "[\s_\-]+$", "")
    FileMark = UCase$(Trim$(t))
End Function

