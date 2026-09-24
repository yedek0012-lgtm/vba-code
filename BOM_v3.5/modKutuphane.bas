Attribute VB_Name = "modKutuphane"
Option Explicit
Option Private Module

' =========================================================================
' KÜTÜPHANE SAÐLIÐI (v3.5)
' L-U-I-O-Y-LIBRARY ve BOLT-LIBRARY'yi tarar, sorunlarý KUTUPHANE_KONTROL sayfasýna yazar.
' Kütüphanelere DOKUNMAZ; yalnýzca rapor verir (satýr numarasý týklanýnca ilgili satýra gider).
'   - Kod / cins / kg/m / yüzey alan / aðýrlýk boþ
'   - Eski YENI_EKLE / TANIMSIZ yer tutucularý
'   - Ayný cins farklý kod, ayný kod farklý kg/m veya aðýrlýk
'   - kg/m ve yüzey alan ölçüden hesaplanan deðerden çok farklý (L, CHS, RHS/SHS, yuvarlak, lama)
'   - Cývata: birim aðýrlýk x 1000 ile 1000 adet aðýrlýðý uyumsuz
' =========================================================================

Private Const TOL_KGM As Double = 0.15      ' kg/m için izin verilen sapma (%15)
Private Const TOL_AREA As Double = 0.2      ' yüzey alan için izin verilen sapma (%20)
Private Const STEEL As Double = 0.00785     ' kg / (mm² * m)

Private repWs As Worksheet
Private repRow As Long
Private nErr As Long
Private nWarn As Long

' Profil adýndan kg/m ve yüzey alan (m²/m) tahmini. Hesaplanamayan kesitlerde (HEA, IPE, U...) False.
Public Function EstimateSection(ByVal nm As String, ByRef kgm As Double, ByRef area As Double) As Boolean
    Dim u As String, pre As String, rest As String, parts As Variant, v() As Double, n As Long, i As Long
    Dim a As Double, b As Double, t As Double, d As Double
    Const PI As Double = 3.14159265358979
    kgm = 0: area = 0
    u = UCase$(Trim$(nm))
    u = Replace(u, " ", "")
    u = Replace(u, ChrW(215), "X")
    u = Replace(u, ChrW(216), "D")
    If u = "" Then Exit Function
    i = 1
    Do While i <= Len(u)
        If Mid$(u, i, 1) Like "[A-Z]" Then i = i + 1 Else Exit Do
    Loop
    pre = Left$(u, i - 1)
    rest = Replace(Replace(Mid$(u, i), "*", "X"), "/", "X")
    If rest = "" Then Exit Function
    parts = Split(rest, "X")
    ReDim v(0 To UBound(parts))
    For i = 0 To UBound(parts)
        If Not parts(i) Like "#*" Then Exit Function
        v(i) = ParseBOMNumber(CStr(parts(i)))
        If v(i) <= 0 Then Exit Function
    Next i
    n = UBound(parts) + 1

    Select Case pre
        Case "", "L", "LS"                       ' köþebent: a x b x t veya a x t (eþit kollu)
            If n = 3 Then
                a = v(0): b = v(1): t = v(2)
            ElseIf n = 2 Then
                a = v(0): b = v(0): t = v(1)
            Else
                Exit Function
            End If
            If t <= 0 Or t >= a Or t >= b Then Exit Function
            kgm = t * (a + b - t) * STEEL
            area = 2 * (a + b) / 1000#
        Case "CHS", "RO", "ROR"                  ' boru: d x t
            If n <> 2 Then Exit Function
            d = v(0): t = v(1)
            If t <= 0 Or 2 * t >= d Then Exit Function
            kgm = PI * t * (d - t) * STEEL
            area = PI * d / 1000#
        Case "RHS", "SHS", "KHP", "QHP", "QR", "RR", "HSS"   ' kutu: a x b x t veya a x t
            If n = 3 Then
                a = v(0): b = v(1): t = v(2)
            ElseIf n = 2 Then
                a = v(0): b = v(0): t = v(1)
            Else
                Exit Function
            End If
            If t <= 0 Or 2 * t >= a Or 2 * t >= b Then Exit Function
            kgm = 2 * t * (a + b - 2 * t) * STEEL
            area = 2 * (a + b) / 1000#
        Case "Y", "RD", "D"                      ' yuvarlak çubuk: d
            If n <> 1 Then Exit Function
            d = v(0)
            kgm = PI * d * d / 4 * STEEL
            area = PI * d / 1000#
        Case "FL", "FLA", "FLACH"                ' lama: b x t
            If n <> 2 Then Exit Function
            a = v(0): b = v(1)
            kgm = a * b * STEEL
            area = 2 * (a + b) / 1000#
        Case Else
            Exit Function
    End Select
    EstimateSection = (kgm > 0)
End Function

' Kütüphane / OGRENME deðeri tahminden çok farklýysa açýklama döner ("" = sorun yok ya da hesaplanamadý)
Public Function SectionValueWarning(ByVal nm As String, ByVal kgm As Double, ByVal area As Double) As String
    Dim eK As Double, eA As Double, s As String
    If Not EstimateSection(nm, eK, eA) Then Exit Function
    If kgm > 0 Then
        If Abs(kgm - eK) / eK > TOL_KGM Then s = "kg/m " & NumToText(kgm) & ", ölçüden tahmin " & NumToText(Round(eK, 2)) & " (%" & Format$(Abs(kgm - eK) / eK * 100, "0") & " fark)"
    End If
    If area > 0 Then
        If Abs(area - eA) / eA > TOL_AREA Then
            If s <> "" Then s = s & "; "
            s = s & "yüzey alan " & NumToText(area) & ", ölçüden tahmin " & NumToText(Round(eA, 3))
        End If
    End If
    SectionValueWarning = s
End Function

Private Sub Report(ByVal libName As String, ByVal r As Long, ByVal code As String, ByVal nm As String, _
                   ByVal isErr As Boolean, ByVal issue As String, ByVal detail As String)
    repRow = repRow + 1
    repWs.Cells(repRow, 1).Value = libName
    repWs.Cells(repRow, 2).Value = r
    repWs.Hyperlinks.Add Anchor:=repWs.Cells(repRow, 2), Address:="", SubAddress:="'" & libName & "'!A" & r, TextToDisplay:=CStr(r)
    repWs.Cells(repRow, 3).NumberFormat = "@"
    repWs.Cells(repRow, 3).Value = code
    repWs.Cells(repRow, 4).Value = nm
    repWs.Cells(repRow, 5).Value = IIf(isErr, "HATA", "UYARI")
    repWs.Cells(repRow, 5).Interior.Color = IIf(isErr, RGB(255, 199, 206), RGB(255, 235, 156))
    repWs.Cells(repRow, 6).Value = issue
    repWs.Cells(repRow, 7).Value = detail
    If isErr Then nErr = nErr + 1 Else nWarn = nWarn + 1
End Sub

Public Sub RunLibraryCheck()
    Dim ws As Worksheet, wsB As Worksheet, r As Long, lastR As Long
    Dim code As String, nm As String, kgm As Double, area As Double, w1 As Double, w1000 As Double
    Dim dName As Object, dCode As Object, key As String, prev As Variant, s As String

    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("L-U-I-O-Y-LIBRARY")
    Set wsB = ThisWorkbook.Sheets("BOLT-LIBRARY")
    Set repWs = ThisWorkbook.Sheets("KUTUPHANE_KONTROL")
    On Error GoTo 0
    If ws Is Nothing And wsB Is Nothing Then
        MsgBox "L-U-I-O-Y-LIBRARY ve BOLT-LIBRARY bulunamadý.", vbCritical
        Exit Sub
    End If
    Application.ScreenUpdating = False
    If repWs Is Nothing Then
        Call UnprotectForMacro
        Set repWs = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        repWs.Name = "KUTUPHANE_KONTROL"
        Call ProtectAfterMacro
    End If
    repWs.Cells.Clear
    repWs.Range("A1:G1").Value = Array("KÜTÜPHANE", "SATIR", "KOD", "CÝNS", "DÜZEY", "SORUN", "AYRINTI")
    With repWs.Range("A1:G1")
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(0, 51, 102)
    End With
    repRow = 1: nErr = 0: nWarn = 0

    ' ---------------- PROFÝL KÜTÜPHANESÝ ----------------
    If Not ws Is Nothing Then
        Set dName = CreateObject("Scripting.Dictionary"): dName.CompareMode = 1
        Set dCode = CreateObject("Scripting.Dictionary"): dCode.CompareMode = 1
        lastR = Application.WorksheetFunction.Max(ws.Cells(ws.Rows.Count, 1).End(xlUp).Row, ws.Cells(ws.Rows.Count, 2).End(xlUp).Row)
        For r = 2 To lastR
            code = Trim$(ws.Cells(r, 1).Text)
            nm = Trim$(ws.Cells(r, 2).Text)
            If code <> "" Or nm <> "" Then
                kgm = CellNumber(ws.Cells(r, 4))
                area = CellNumber(ws.Cells(r, 5))
                If UCase$(code) = "YENI_EKLE" Or UCase$(code) = "TANIMSIZ" Then
                    Call Report(ws.Name, r, code, nm, True, "Eski yer tutucu", "Kodu girin ya da Kutuphane_Temizle ile silin.")
                Else
                    If code = "" Then Call Report(ws.Name, r, code, nm, True, "Kod boþ", "Bu satýr eþleþmede kullanýlmaz.")
                    If nm = "" Then Call Report(ws.Name, r, code, nm, True, "Cins boþ", "")
                    If kgm <= 0 Then Call Report(ws.Name, r, code, nm, True, "kg/m boþ", "Aðýrlýk 0 hesaplanýr.")
                    If area <= 0 Then Call Report(ws.Name, r, code, nm, False, "Yüzey alan boþ", "")
                    s = SectionValueWarning(nm, kgm, area)
                    If s <> "" Then Call Report(ws.Name, r, code, nm, False, "Deðer þüpheli", s)

                    ' ayný cins farklý kod
                    If nm <> "" And code <> "" Then
                        key = CanonicalSectionKey(nm)
                        If dName.Exists(key) Then
                            prev = dName(key)
                            If UCase$(prev(1)) <> UCase$(code) Then
                                Call Report(ws.Name, r, code, nm, True, "Ayný cins farklý kod", "Satýr " & prev(0) & ": kod " & prev(1) & " (makro ilk bulduðunu kullanýr)")
                            End If
                        Else
                            dName.Add key, Array(r, code)
                        End If
                    End If
                    ' ayný kod farklý kg/m (ayný koda ait ham ad satýrlarý normaldir, deðerleri ayný olmalý)
                    If code <> "" And kgm > 0 Then
                        If dCode.Exists(code) Then
                            prev = dCode(code)
                            If Abs(prev(1) - kgm) > 0.01 * kgm Then
                                Call Report(ws.Name, r, code, nm, True, "Ayný kod farklý kg/m", "Satýr " & prev(0) & ": " & NumToText(prev(1)) & " kg/m, bu satýr: " & NumToText(kgm))
                            End If
                        Else
                            dCode.Add code, Array(r, kgm)
                        End If
                    End If
                End If
            End If
        Next r
    End If

    ' ---------------- CIVATA KÜTÜPHANESÝ ----------------
    If Not wsB Is Nothing Then
        Set dName = CreateObject("Scripting.Dictionary"): dName.CompareMode = 1
        Set dCode = CreateObject("Scripting.Dictionary"): dCode.CompareMode = 1
        lastR = Application.WorksheetFunction.Max(wsB.Cells(wsB.Rows.Count, 1).End(xlUp).Row, wsB.Cells(wsB.Rows.Count, 2).End(xlUp).Row)
        For r = 3 To lastR
            code = Trim$(wsB.Cells(r, 1).Text)
            nm = Trim$(wsB.Cells(r, 2).Text)
            If code <> "" Or nm <> "" Then
                w1 = CellNumber(wsB.Cells(r, 3))
                w1000 = CellNumber(wsB.Cells(r, 4))
                If code = "" Then Call Report(wsB.Name, r, code, nm, True, "Kod boþ", "")
                If w1 <= 0 And w1000 <= 0 Then
                    Call Report(wsB.Name, r, code, nm, True, "Aðýrlýk boþ", "Makro aðýrlýðý tahmin eder (turuncu).")
                ElseIf w1 > 0 And w1000 > 0 Then
                    If Abs(w1 * 1000 - w1000) > 0.02 * w1000 Then
                        Call Report(wsB.Name, r, code, nm, False, "Birim / 1000 adet uyumsuz", _
                                    "Birim " & NumToText(w1) & " x 1000 = " & NumToText(w1 * 1000) & ", 1000 adet: " & NumToText(w1000) & " (makro birim aðýrlýðý kullanýr)")
                    End If
                End If
                If w1 <= 0 Then w1 = w1000 / 1000#
                If code <> "" Then
                    If dCode.Exists(UCase$(code)) Then
                        prev = dCode(UCase$(code))
                        If w1 > 0 And Abs(prev(1) - w1) > 0.01 * w1 Then
                            Call Report(wsB.Name, r, code, nm, True, "Ayný kod farklý aðýrlýk", "Satýr " & prev(0) & ": " & NumToText(prev(1)) & " kg, bu satýr: " & NumToText(w1))
                        End If
                    Else
                        dCode.Add UCase$(code), Array(r, w1)
                    End If
                End If
                If nm <> "" And code <> "" Then
                    key = UCase$(Replace(nm, " ", ""))
                    If dName.Exists(key) Then
                        prev = dName(key)
                        If UCase$(prev(1)) <> UCase$(code) Then
                            Call Report(wsB.Name, r, code, nm, False, "Ayný cins farklý kod", "Satýr " & prev(0) & ": kod " & prev(1))
                        End If
                    Else
                        dName.Add key, Array(r, code)
                    End If
                End If
            End If
        Next r
    End If

    repWs.Columns("A").ColumnWidth = 20
    repWs.Columns("B").ColumnWidth = 8
    repWs.Columns("C").ColumnWidth = 14
    repWs.Columns("D").ColumnWidth = 32
    repWs.Columns("E").ColumnWidth = 8
    repWs.Columns("F").ColumnWidth = 26
    repWs.Columns("G").ColumnWidth = 80
    If repRow > 1 Then
        repWs.Range("A1").Resize(repRow, 7).AutoFilter
    Else
        repWs.Range("A2").Value = "Sorun bulunamadý."
    End If
    Application.ScreenUpdating = True
    repWs.Activate
    repWs.Range("A2").Select
    MsgBox nErr & " HATA, " & nWarn & " UYARI bulundu." & vbCrLf & _
           "Ayrýntý: KUTUPHANE_KONTROL sayfasý (SATIR numarasýna týklayýnca kütüphanedeki satýra gider)." & vbCrLf & _
           "Kütüphanelerde hiçbir þey deðiþtirilmedi.", IIf(nErr > 0, vbExclamation, vbInformation), "Kütüphane Kontrol"
End Sub
