Attribute VB_Name = "modPdfText"
Option Explicit
Option Compare Binary
Option Private Module

' =========================================================================
' PDF OKUYUCU (saf VBA) - v3.5
' Power Query KULLANMADAN metin tabanlý PDF'teki malzeme listelerini okur:
'  1) Dosya baytlarý okunur, "N 0 obj" nesneleri bulunur (sýkýþtýrýlmýþ nesne akýþlarý dahil).
'  2) Sayfa içeriði FlateDecode ile açýlýr (VBA inflate), yazýlar sayfadaki konumlarýyla toplanýr.
'  3) Tablolar PDF_BOM_Okuyucu.html ile ayný yöntemle kurulur (baþlýk satýrý + sütun hizasý).
' Binlerce çizgi içeren Tekla montaj çizimleri birkaç saniyede okunur.
' Taranmýþ (resim) PDF'te yazý yoktur -> 0 tablo döner.
' =========================================================================

' ---- dosya ve nesneler ----
Private fb() As Byte
Private fbLen As Long
Private oOff() As Long
Private oStm() As Long
Private oIdx() As Long
Private oMax As Long
Private stmLoaded As Boolean
Private stmBytes() As Variant

' ---- yazý parçalarý ----
Private tkS() As String
Private tkX0() As Double
Private tkX1() As Double
Private tkY() As Double
Private tkH() As Double
Private tkN As Long
Private pageTop As Double

' ---- fontlar ----
Private fntKey As Collection
Private fnCID() As Boolean
Private fnW() As Variant
Private fnDW() As Double
Private fnCW() As Variant
Private fnUni() As Variant
Private fnRng() As Variant
Private fnN As Long

' ---- grafik / metin durumu ----
Private gsSt() As Double
Private gsD As Long
Private ctm(5) As Double
Private tmx(5) As Double
Private tlm(5) As Double
Private curFont As Long
Private curSize As Double
Private tChar As Double
Private tWord As Double
Private tHs As Double
Private tLead As Double
Private tRise As Double
Private runOn As Boolean
Private runStr As String
Private runTx As Double
Private runM(5) As Double
Private runSize As Double
Private runHs As Double
Private runRise As Double

' ---- inflate ----
Private zi() As Byte
Private ziPos As Long
Private ziEnd As Long
Private zBit As Long
Private zCnt As Long
Private zEOF As Boolean
Private zo() As Byte
Private zoPos As Long
Private ltC(15) As Long
Private ltS(319) As Long
Private dtC(15) As Long
Private dtS(31) As Long
Private fastL(511) As Long       ' 9 bitlik hýzlý çözme tablosu: sembol * 16 + kod boyu (-1 = yok)
Private fastD(511) As Long
Private lenBase(28) As Long
Private lenExt(28) As Long
Private distBase(29) As Long
Private distExt(29) As Long
Private p2(30) As Long
Private winAnsi(255) As Long
Private tablesReady As Boolean

' ---- satýrlar ----
Private lnY() As Double
Private lnH() As Double
Private lnIt() As Variant
Private lnN As Long

' ---- sayfalar / çýktý satýrlarý ----
Private pgList() As Variant
Private pgN As Long
Private outRows() As Variant
Private outN As Long

' =========================================================================
' GÝRÝÞ
' Dönüþ: bulunan tablo sayýsý (0 = tablo yok, -1 = PDF okunamadý -> errMsg)
' outArr: (1..satýr, 1..sütun); her tablo baþlýk + satýrlar, tablolar arasýnda boþ satýr
' =========================================================================
Public Function PdfReadTables(ByVal pdfPath As String, ByRef outArr As Variant, ByRef errMsg As String) As Long
    Dim i As Long, nT As Long
    On Error GoTo Fail
    errMsg = ""
    outArr = Empty
    InitTables
    If Not LoadPdfFile(pdfPath) Then
        errMsg = "Dosya açýlamadý veya boþ."
        PdfReadTables = -1
        Exit Function
    End If
    ScanObjects
    If oMax = 0 Then
        errMsg = "PDF nesnesi bulunamadý (bozuk dosya)."
        PdfReadTables = -1
        Exit Function
    End If
    If FindLast(fb, fbLen, "/Encrypt") >= 0 Then
        errMsg = "Þifreli PDF."
        PdfReadTables = -1
        Exit Function
    End If
    Set fntKey = New Collection
    fnN = 0
    pgN = 0
    ReDim pgList(1 To 8)
    outN = 0
    ReDim outRows(1 To 64)
    CollectPages
    For i = 1 To pgN
        tkN = 0
        ReadPageText pgList(i)
        nT = nT + BuildTables()
    Next i
    If outN > 0 Then outArr = RowsToArray()
    PdfReadTables = nT
    Erase fb
    Erase zi
    Erase zo
    Exit Function
Fail:
    errMsg = "PDF okunamadý: " & Err.Description
    PdfReadTables = -1
End Function

' =========================================================================
' DOSYA / NESNE AYRIÞTIRICI
' =========================================================================
Private Function LoadPdfFile(ByVal pdfPath As String) As Boolean
    Dim f As Integer
    On Error GoTo Fail
    f = FreeFile
    Open pdfPath For Binary Access Read As #f
    fbLen = LOF(f)
    If fbLen < 32 Then
        Close #f
        Exit Function
    End If
    ReDim fb(0 To fbLen - 1)
    Get #f, 1, fb
    Close #f
    LoadPdfFile = True
    Exit Function
Fail:
    On Error Resume Next
    Close #f
End Function

' Dosyada "N G obj" konumlarýný bulur (sonraki taným öncekini ezer = artýmlý güncelleme)
Private Sub ScanObjects()
    Dim i As Long, j As Long, k As Long, num As Long, mul As Long
    oMax = 0
    ReDim oOff(0 To 1023)
    ReDim oStm(0 To 1023)
    ReDim oIdx(0 To 1023)
    stmLoaded = False
    ReDim stmBytes(0 To 0)
    For i = 1 To fbLen - 4
        If fb(i) = 111 Then
            If fb(i + 1) = 98 Then
                If fb(i + 2) = 106 Then
                    If IsWS(fb(i - 1)) And (IsWS(fb(i + 3)) Or IsDelim(fb(i + 3))) Then
                        j = i - 1
                        Do While j > 0
                            If Not IsWS(fb(j)) Then Exit Do
                            j = j - 1
                        Loop
                        If IsDigitB(fb(j)) Then
                            Do While j > 0
                                If Not IsDigitB(fb(j)) Then Exit Do
                                j = j - 1
                            Loop
                            If IsWS(fb(j)) Then
                                Do While j > 0
                                    If Not IsWS(fb(j)) Then Exit Do
                                    j = j - 1
                                Loop
                                If IsDigitB(fb(j)) Then
                                    num = 0: mul = 1: k = j
                                    Do While k >= 0
                                        If Not IsDigitB(fb(k)) Then Exit Do
                                        num = num + (fb(k) - 48) * mul
                                        mul = mul * 10
                                        k = k - 1
                                        If mul > 10000000 Then Exit Do
                                    Loop
                                    If num > 0 Then
                                        EnsureObj num
                                        oOff(num) = i + 3
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Next i
End Sub

Private Sub EnsureObj(ByVal num As Long)
    Dim nu As Long
    If num > UBound(oOff) Then
        nu = UBound(oOff) * 2
        If nu < num + 1 Then nu = num + 1024
        ReDim Preserve oOff(0 To nu)
        ReDim Preserve oStm(0 To nu)
        ReDim Preserve oIdx(0 To nu)
    End If
    If num > oMax Then oMax = num
End Sub

' Sýkýþtýrýlmýþ nesne akýþlarýndaki (ObjStm) nesneler
Private Sub LoadObjStms()
    Dim n As Long, p As Long, d As Variant, sb As Variant, bb() As Byte, cnt As Long, first As Long
    Dim i As Long, q As Long, onum As Long, objOfs As Long, bl As Long
    stmLoaded = True
    For n = 1 To oMax
        If oOff(n) > 0 Then
            p = oOff(n)
            SkipWS fb, p, fbLen
            If p + 1 < fbLen Then
                If fb(p) = 60 And fb(p + 1) = 60 Then
                    LetV d, ParseObj(fb, p, fbLen)
                    If NameOf(DGet(d, "Type")) = "/ObjStm" Then
                        sb = GetStreamBytes(n)
                        If IsArray(sb) Then
                            bb = sb
                            bl = UBound(bb) + 1
                            cnt = CLng(Nv(Resolve(DGet(d, "N"))))
                            first = CLng(Nv(Resolve(DGet(d, "First"))))
                            q = 0
                            For i = 1 To cnt
                                SkipWS bb, q, bl
                                If q >= bl Then Exit For
                                onum = CLng(ReadNumber(bb, q, bl))
                                SkipWS bb, q, bl
                                objOfs = CLng(ReadNumber(bb, q, bl))
                                If onum > 0 Then
                                    EnsureObj onum
                                    If oOff(onum) = 0 Then
                                        oStm(onum) = n
                                        oIdx(onum) = first + objOfs
                                    End If
                                End If
                            Next i
                            If n > UBound(stmBytes) Then
                                ReDim Preserve stmBytes(0 To oMax)
                            End If
                            stmBytes(n) = sb
                        End If
                    End If
                End If
            End If
        End If
    Next n
End Sub

Private Function GetObj(ByVal n As Long) As Variant
    Dim p As Long, r As Variant, b() As Byte
    r = Empty
    If n > 0 And n <= oMax Then
        If oOff(n) = 0 And oStm(n) = 0 And Not stmLoaded Then LoadObjStms
        If oOff(n) > 0 Then
            p = oOff(n)
            LetV r, ParseObj(fb, p, fbLen)
        ElseIf oStm(n) > 0 Then
            b = stmBytes(oStm(n))
            p = oIdx(n)
            If p <= UBound(b) Then LetV r, ParseObj(b, p, UBound(b) + 1)
        End If
    End If
    If IsObject(r) Then Set GetObj = r Else GetObj = r
End Function

' Akýþ nesnesinin (stream) çözülmüþ baytlarý; akýþ deðilse Empty
Private Function GetStreamBytes(ByVal n As Long) As Variant
    Dim p As Long, d As Variant, ln As Long, st As Long, raw() As Byte, i As Long, lv As Variant
    Dim flt As Variant, k As Long, v As Variant
    GetStreamBytes = Empty
    If n <= 0 Or n > oMax Then Exit Function
    If oOff(n) = 0 Then Exit Function
    p = oOff(n)
    LetV d, ParseObj(fb, p, fbLen)
    If Not IsDict(d) Then Exit Function
    SkipWS fb, p, fbLen
    If Not MatchAt(fb, p, fbLen, "stream") Then Exit Function
    p = p + 6
    If p < fbLen Then
        If fb(p) = 13 Then p = p + 1
    End If
    If p < fbLen Then
        If fb(p) = 10 Then p = p + 1
    End If
    st = p
    ln = -1
    LetV lv, Resolve(DGet(d, "Length"))
    If VarType(lv) = vbDouble Then ln = CLng(lv)
    If ln < 0 Or st + ln > fbLen Then
        ln = FindFrom(fb, fbLen, "endstream", st)
        If ln < 0 Then ln = fbLen
        ln = ln - st
    End If
    If ln <= 0 Then Exit Function
    ReDim raw(0 To ln - 1)
    For i = 0 To ln - 1
        raw(i) = fb(st + i)
    Next i
    v = raw
    LetV flt, Resolve(DGet(d, "Filter"))
    If VarType(flt) = vbString Then flt = Array(flt)
    If IsArray(flt) Then
        For k = 0 To UBound(flt)
            v = ApplyFilter(v, NameOf(Resolve(flt(k))))
            If Not IsArray(v) Then Exit Function
        Next k
    End If
    GetStreamBytes = v
End Function

Private Function ApplyFilter(ByVal v As Variant, ByVal fName As String) As Variant
    Dim b() As Byte
    ApplyFilter = Empty
    Select Case fName
        Case "/FlateDecode", "/Fl"
            b = v
            ApplyFilter = Inflate(b)
        Case "/ASCIIHexDecode", "/AHx"
            b = v
            ApplyFilter = HexDecode(b)
        Case Else
            ' resim / desteklenmeyen filtre: içerik okunmaz
    End Select
End Function

Private Function HexDecode(b() As Byte) As Variant
    Dim o() As Byte, k As Long, i As Long, hv As Long, hi As Long, c As Long
    ReDim o(0 To (UBound(b) + 1) \ 2 + 1)
    hi = -1
    For i = 0 To UBound(b)
        c = b(i)
        If c = 62 Then Exit For
        hv = HexNib(c)
        If hv >= 0 Then
            If hi < 0 Then
                hi = hv
            Else
                o(k) = hi * 16 + hv: k = k + 1: hi = -1
            End If
        End If
    Next i
    If hi >= 0 Then o(k) = hi * 16: k = k + 1
    If k = 0 Then
        HexDecode = Empty
    Else
        ReDim Preserve o(0 To k - 1)
        HexDecode = o
    End If
End Function

' ---- sözdizimi ----
Private Function IsWS(ByVal c As Long) As Boolean
    IsWS = (c = 32 Or c = 10 Or c = 13 Or c = 9 Or c = 12 Or c = 0)
End Function

Private Function IsDelim(ByVal c As Long) As Boolean
    IsDelim = (c = 40 Or c = 41 Or c = 60 Or c = 62 Or c = 91 Or c = 93 Or c = 123 Or c = 125 Or c = 47 Or c = 37)
End Function

Private Function IsDigitB(ByVal c As Long) As Boolean
    IsDigitB = (c >= 48 And c <= 57)
End Function

Private Function HexNib(ByVal c As Long) As Long
    If c >= 48 And c <= 57 Then
        HexNib = c - 48
    ElseIf c >= 65 And c <= 70 Then
        HexNib = c - 55
    ElseIf c >= 97 And c <= 102 Then
        HexNib = c - 87
    Else
        HexNib = -1
    End If
End Function

Private Sub SkipWS(b() As Byte, ByRef p As Long, ByVal n As Long)
    Dim c As Long
    Do While p < n
        c = b(p)
        If c = 32 Or c = 10 Or c = 13 Or c = 9 Or c = 12 Or c = 0 Then
            p = p + 1
        ElseIf c = 37 Then
            Do While p < n
                c = b(p)
                If c = 10 Or c = 13 Then Exit Do
                p = p + 1
            Loop
        Else
            Exit Do
        End If
    Loop
End Sub

Private Function MatchAt(b() As Byte, ByVal p As Long, ByVal n As Long, ByVal pat As String) As Boolean
    Dim j As Long
    If p < 0 Or p + Len(pat) > n Then Exit Function
    For j = 1 To Len(pat)
        If b(p + j - 1) <> AscW(Mid$(pat, j, 1)) Then Exit Function
    Next j
    MatchAt = True
End Function

Private Function FindFrom(b() As Byte, ByVal n As Long, ByVal pat As String, ByVal start As Long) As Long
    Dim i As Long, c0 As Long
    FindFrom = -1
    c0 = AscW(pat)
    For i = start To n - Len(pat)
        If b(i) = c0 Then
            If MatchAt(b, i, n, pat) Then
                FindFrom = i
                Exit Function
            End If
        End If
    Next i
End Function

Private Function FindLast(b() As Byte, ByVal n As Long, ByVal pat As String) As Long
    Dim i As Long, c0 As Long
    FindLast = -1
    c0 = AscW(pat)
    For i = n - Len(pat) To 0 Step -1
        If b(i) = c0 Then
            If MatchAt(b, i, n, pat) Then
                FindLast = i
                Exit Function
            End If
        End If
    Next i
End Function

Private Function ReadNumber(b() As Byte, ByRef p As Long, ByVal n As Long) As Double
    Dim neg As Boolean, v As Double, frac As Double, c As Long, seenDot As Boolean
    frac = 1
    If p < n Then
        If b(p) = 45 Then
            neg = True: p = p + 1
        ElseIf b(p) = 43 Then
            p = p + 1
        End If
    End If
    Do While p < n
        c = b(p)
        If c >= 48 And c <= 57 Then
            If seenDot Then
                frac = frac / 10
                v = v + (c - 48) * frac
            Else
                v = v * 10 + (c - 48)
            End If
        ElseIf c = 46 And Not seenDot Then
            seenDot = True
        ElseIf c = 45 Then
            ' "--5" gibi hatalý yazýmlar: yok say
        Else
            Exit Do
        End If
        p = p + 1
    Loop
    If neg Then v = -v
    ReadNumber = v
End Function

Private Function ReadName(b() As Byte, ByRef p As Long, ByVal n As Long) As String
    Dim s As String, c As Long
    p = p + 1
    Do While p < n
        c = b(p)
        If IsWS(c) Or IsDelim(c) Then Exit Do
        If c = 35 And p + 2 < n Then
            s = s & ChrW(MaxL(0, HexNib(b(p + 1))) * 16 + MaxL(0, HexNib(b(p + 2))))
            p = p + 3
        Else
            s = s & ChrW(c)
            p = p + 1
        End If
    Loop
    ReadName = s
End Function

Private Function ReadKeyword(b() As Byte, ByRef p As Long, ByVal n As Long) As String
    Dim s As String, c As Long, st As Long
    st = p
    Do While p < n
        c = b(p)
        If IsWS(c) Or IsDelim(c) Then Exit Do
        s = s & ChrW(c)
        p = p + 1
    Loop
    If p = st Then p = p + 1
    ReadKeyword = s
End Function

Private Sub AddB(o() As Byte, ByRef k As Long, ByVal v As Long)
    If k > UBound(o) Then
        ReDim Preserve o(0 To 2 * UBound(o) + 1)
    End If
    o(k) = v And 255
    k = k + 1
End Sub

' (...) dizgisi -> Byte() (boþsa Empty)
Private Function ReadLitBytes(b() As Byte, ByRef p As Long, ByVal n As Long) As Variant
    Dim o() As Byte, k As Long, depth As Long, c As Long, v As Long, j As Long
    ReDim o(0 To 63)
    p = p + 1
    depth = 1
    Do While p < n
        c = b(p)
        If c = 92 Then
            p = p + 1
            If p >= n Then Exit Do
            c = b(p)
            Select Case c
                Case 110: AddB o, k, 10: p = p + 1
                Case 114: AddB o, k, 13: p = p + 1
                Case 116: AddB o, k, 9: p = p + 1
                Case 98: AddB o, k, 8: p = p + 1
                Case 102: AddB o, k, 12: p = p + 1
                Case 48 To 55
                    v = 0: j = 0
                    Do While j < 3 And p < n
                        c = b(p)
                        If c < 48 Or c > 55 Then Exit Do
                        v = v * 8 + (c - 48)
                        p = p + 1
                        j = j + 1
                    Loop
                    AddB o, k, v
                Case 13
                    p = p + 1
                    If p < n Then
                        If b(p) = 10 Then p = p + 1
                    End If
                Case 10
                    p = p + 1
                Case Else
                    AddB o, k, c: p = p + 1
            End Select
        ElseIf c = 40 Then
            depth = depth + 1
            AddB o, k, c: p = p + 1
        ElseIf c = 41 Then
            depth = depth - 1
            p = p + 1
            If depth = 0 Then Exit Do
            AddB o, k, c
        Else
            AddB o, k, c: p = p + 1
        End If
    Loop
    If k = 0 Then
        ReadLitBytes = Empty
    Else
        ReDim Preserve o(0 To k - 1)
        ReadLitBytes = o
    End If
End Function

' <...> dizgisi -> Byte() (boþsa Empty)
Private Function ReadHexBytes(b() As Byte, ByRef p As Long, ByVal n As Long) As Variant
    Dim o() As Byte, k As Long, hv As Long, hi As Long, c As Long
    ReDim o(0 To 63)
    p = p + 1
    hi = -1
    Do While p < n
        c = b(p)
        p = p + 1
        If c = 62 Then Exit Do
        hv = HexNib(c)
        If hv >= 0 Then
            If hi < 0 Then
                hi = hv
            Else
                AddB o, k, hi * 16 + hv
                hi = -1
            End If
        End If
    Loop
    If hi >= 0 Then AddB o, k, hi * 16
    If k = 0 Then
        ReadHexBytes = Empty
    Else
        ReDim Preserve o(0 To k - 1)
        ReadHexBytes = o
    End If
End Function

Private Function BytesToText(ByVal v As Variant) As String
    Dim i As Long, s As String
    If Not IsArray(v) Then Exit Function
    For i = LBound(v) To UBound(v)
        s = s & ChrW(v(i))
    Next i
    BytesToText = s
End Function

' PDF nesnesi: sayý=Double, ad="/Ad", dizgi="(metin", baþvuru="@N", dizi=Variant(),
' sözlük=Array("<<", anahtarlar(), deðerler(), adet)
Private Function ParseObj(b() As Byte, ByRef p As Long, ByVal n As Long) As Variant
    Dim c As Long, arr() As Variant, cnt As Long, ks() As String, vs() As Variant, nk As Long, k As String, v As Variant
    Dim q As Long, num As Double, isDic As Boolean
    ParseObj = Empty
    SkipWS b, p, n
    If p >= n Then Exit Function
    c = b(p)
    Select Case c
        Case 47
            ParseObj = "/" & ReadName(b, p, n)
        Case 40
            ParseObj = "(" & BytesToText(ReadLitBytes(b, p, n))
        Case 60
            isDic = False
            If p + 1 < n Then isDic = (b(p + 1) = 60)
            If isDic Then
                p = p + 2
                nk = 0
                ReDim ks(0 To 7)
                ReDim vs(0 To 7)
                Do
                    SkipWS b, p, n
                    If p >= n Then Exit Do
                    If b(p) = 62 Then
                        p = p + 2
                        Exit Do
                    End If
                    If b(p) <> 47 Then
                        p = p + 1
                    Else
                        k = ReadName(b, p, n)
                        v = ParseObj(b, p, n)
                        If nk > UBound(ks) Then
                            ReDim Preserve ks(0 To 2 * nk)
                            ReDim Preserve vs(0 To 2 * nk)
                        End If
                        ks(nk) = k
                        vs(nk) = v
                        nk = nk + 1
                    End If
                Loop
                ParseObj = Array("<<", ks, vs, nk)
            Else
                ParseObj = "(" & BytesToText(ReadHexBytes(b, p, n))
            End If
        Case 91
            p = p + 1
            cnt = 0
            ReDim arr(0 To 7)
            Do
                SkipWS b, p, n
                If p >= n Then Exit Do
                If b(p) = 93 Then
                    p = p + 1
                    Exit Do
                End If
                v = Empty
                LetV v, ParseObj(b, p, n)
                If cnt > UBound(arr) Then
                    ReDim Preserve arr(0 To 2 * cnt)
                End If
                If IsObject(v) Then Set arr(cnt) = v Else arr(cnt) = v
                cnt = cnt + 1
            Loop
            If cnt = 0 Then
                ParseObj = Array()
            Else
                ReDim Preserve arr(0 To cnt - 1)
                ParseObj = arr
            End If
        Case 43, 45, 46, 48 To 57
            num = ReadNumber(b, p, n)
            ParseObj = num
            ' "N G R" baþvurusu mu?
            If num >= 0 And num = Int(num) Then
                q = p
                SkipWS b, p, n
                If p < n Then
                    If IsDigitB(b(p)) Then
                        ReadNumber b, p, n
                        SkipWS b, p, n
                        If p < n Then
                            If b(p) = 82 Then
                                If p + 1 >= n Then
                                    p = p + 1
                                    ParseObj = "@" & CStr(CLng(num))
                                    Exit Function
                                ElseIf IsWS(b(p + 1)) Or IsDelim(b(p + 1)) Then
                                    p = p + 1
                                    ParseObj = "@" & CStr(CLng(num))
                                    Exit Function
                                End If
                            End If
                        End If
                    End If
                End If
                p = q
            End If
        Case Else
            ParseObj = ReadKeyword(b, p, n)
    End Select
End Function

Private Sub LetV(ByRef dst As Variant, ByVal src As Variant)
    If IsObject(src) Then Set dst = src Else dst = src
End Sub

Private Function IsDict(ByVal v As Variant) As Boolean
    If Not IsArray(v) Then Exit Function
    If VarType(v) <> vbArray + vbVariant Then Exit Function
    If UBound(v) <> 3 Then Exit Function
    If VarType(v(0)) = vbString Then IsDict = (v(0) = "<<")
End Function

Private Function DGet(ByVal d As Variant, ByVal key As String) As Variant
    Dim ks As Variant, vs As Variant, i As Long
    DGet = Empty
    If Not IsDict(d) Then Exit Function
    ks = d(1)
    vs = d(2)
    For i = 0 To d(3) - 1
        If ks(i) = key Then
            DGet = vs(i)
            Exit Function
        End If
    Next i
End Function

Private Function RefNum(ByVal v As Variant) As Long
    If VarType(v) = vbString Then
        If Left$(v, 1) = "@" Then RefNum = CLng(Mid$(v, 2))
    End If
End Function

Private Function Resolve(ByVal v As Variant) As Variant
    Dim i As Long, r As Variant, rn As Long
    LetV r, v
    For i = 1 To 8
        If IsObject(r) Then Exit For
        rn = RefNum(r)
        If rn = 0 Then Exit For
        LetV r, GetObj(rn)
    Next i
    If IsObject(r) Then Set Resolve = r Else Resolve = r
End Function

Private Function NameOf(ByVal v As Variant) As String
    If VarType(v) = vbString Then NameOf = v
End Function

Private Function Nv(ByVal v As Variant) As Double
    Select Case VarType(v)
        Case vbDouble, vbLong, vbInteger, vbSingle, vbByte, vbCurrency
            Nv = v
    End Select
End Function

Private Function MaxD(ByVal a As Double, ByVal b As Double) As Double
    If a > b Then MaxD = a Else MaxD = b
End Function

Private Function MinD(ByVal a As Double, ByVal b As Double) As Double
    If a < b Then MinD = a Else MinD = b
End Function

Private Function MaxL(ByVal a As Long, ByVal b As Long) As Long
    If a > b Then MaxL = a Else MaxL = b
End Function

' =========================================================================
' SAYFALAR
' =========================================================================
Private Sub AddPage(ByVal node As Variant, ByVal res As Variant, ByVal mb As Variant)
    pgN = pgN + 1
    If pgN > UBound(pgList) Then
        ReDim Preserve pgList(1 To 2 * pgN)
    End If
    pgList(pgN) = Array(node, res, mb)
End Sub

Private Sub CollectPages()
    Dim p As Long, root As Variant, cat As Variant, pgs As Variant, n As Long, o As Variant, q As Long
    p = FindLast(fb, fbLen, "/Root")
    If p >= 0 Then
        p = p + 5
        LetV root, ParseObj(fb, p, fbLen)
        LetV cat, Resolve(root)
        LetV pgs, Resolve(DGet(cat, "Pages"))
        If IsDict(pgs) Then WalkPages pgs, Empty, Empty, 0
    End If
    If pgN = 0 Then
        ' yedek: tüm /Type /Page nesneleri
        For n = 1 To oMax
            If oOff(n) > 0 Then
                q = oOff(n)
                LetV o, ParseObj(fb, q, fbLen)
                If NameOf(DGet(o, "Type")) = "/Page" Then
                    AddPage o, Resolve(DGet(o, "Resources")), Resolve(DGet(o, "MediaBox"))
                End If
            End If
        Next n
    End If
End Sub

Private Sub WalkPages(ByVal node As Variant, ByVal res As Variant, ByVal mb As Variant, ByVal depth As Long)
    Dim kids As Variant, r As Variant, m As Variant, i As Long, k As Variant
    If depth > 32 Or Not IsDict(node) Then Exit Sub
    LetV r, Resolve(DGet(node, "Resources"))
    If Not IsDict(r) Then LetV r, res
    LetV m, Resolve(DGet(node, "MediaBox"))
    If Not IsArray(m) Then LetV m, mb
    LetV kids, Resolve(DGet(node, "Kids"))
    If IsArray(kids) Then
        For i = 0 To UBound(kids)
            LetV k, Resolve(kids(i))
            WalkPages k, r, m, depth + 1
        Next i
    Else
        AddPage node, r, m
    End If
End Sub

Private Sub ReadPageText(ByVal pg As Variant)
    Dim node As Variant, res As Variant, mb As Variant, cv As Variant, part As Variant
    Dim data() As Byte, dl As Long, i As Long, nParts As Long
    LetV node, pg(0)
    LetV res, pg(1)
    LetV mb, pg(2)
    pageTop = 842
    If IsArray(mb) Then
        If UBound(mb) >= 3 Then pageTop = MaxD(Nv(Resolve(mb(1))), Nv(Resolve(mb(3))))
    End If
    ReDim data(0 To 1023)
    dl = 0
    LetV cv, DGet(node, "Contents")
    If VarType(cv) = vbString Then
        part = GetStreamBytes(RefNum(cv))
        If IsArray(part) Then
            data = part
            dl = UBound(data) + 1
            nParts = 1
        Else
            LetV cv, Resolve(cv)
        End If
    End If
    If IsArray(cv) Then
        For i = 0 To UBound(cv)
            part = GetStreamBytes(RefNum(cv(i)))
            If IsArray(part) Then
                AppendBytes data, dl, part
                AppendBytes data, dl, Array(10)
                nParts = nParts + 1
            End If
        Next i
    End If
    If dl = 0 Then Exit Sub

    ' grafik durumu
    SetIdent ctm
    SetIdent tmx
    SetIdent tlm
    gsD = 0
    ReDim gsSt(0 To 12, 0 To 63)
    curFont = 0: curSize = 10: tChar = 0: tWord = 0: tHs = 1: tLead = 0: tRise = 0
    runOn = False
    RunContent data, dl, res, 0
End Sub

Private Sub AppendBytes(dst() As Byte, ByRef dl As Long, ByVal src As Variant)
    Dim i As Long, n As Long, need As Long, lb As Long
    lb = LBound(src)
    n = UBound(src) - lb + 1
    If n <= 0 Then Exit Sub
    need = dl + n
    If need > UBound(dst) + 1 Then
        ReDim Preserve dst(0 To MaxL(need, 2 * (UBound(dst) + 1)) - 1)
    End If
    For i = 0 To n - 1
        dst(dl + i) = src(lb + i)
    Next i
    dl = dl + n
End Sub

' =========================================================================
' ÝÇERÝK AKIÞI (metin operatörleri)
' =========================================================================
Private Sub SetIdent(m() As Double)
    m(0) = 1: m(1) = 0: m(2) = 0: m(3) = 1: m(4) = 0: m(5) = 0
End Sub

' R = A x B
Private Sub MatMul(a() As Double, b() As Double, r() As Double)
    Dim t0 As Double, t1 As Double, t2 As Double, t3 As Double, t4 As Double, t5 As Double
    t0 = a(0) * b(0) + a(1) * b(2)
    t1 = a(0) * b(1) + a(1) * b(3)
    t2 = a(2) * b(0) + a(3) * b(2)
    t3 = a(2) * b(1) + a(3) * b(3)
    t4 = a(4) * b(0) + a(5) * b(2) + b(4)
    t5 = a(4) * b(1) + a(5) * b(3) + b(5)
    r(0) = t0: r(1) = t1: r(2) = t2: r(3) = t3: r(4) = t4: r(5) = t5
End Sub

Private Sub PushGS()
    Dim i As Long
    If gsD > UBound(gsSt, 2) Then
        ReDim Preserve gsSt(0 To 12, 0 To 2 * UBound(gsSt, 2) + 1)
    End If
    For i = 0 To 5
        gsSt(i, gsD) = ctm(i)
    Next i
    gsSt(6, gsD) = curFont: gsSt(7, gsD) = curSize: gsSt(8, gsD) = tChar: gsSt(9, gsD) = tWord
    gsSt(10, gsD) = tHs: gsSt(11, gsD) = tLead: gsSt(12, gsD) = tRise
    gsD = gsD + 1
End Sub

Private Sub PopGS()
    Dim i As Long
    If gsD = 0 Then Exit Sub
    gsD = gsD - 1
    For i = 0 To 5
        ctm(i) = gsSt(i, gsD)
    Next i
    curFont = CLng(gsSt(6, gsD)): curSize = gsSt(7, gsD): tChar = gsSt(8, gsD): tWord = gsSt(9, gsD)
    tHs = gsSt(10, gsD): tLead = gsSt(11, gsD): tRise = gsSt(12, gsD)
End Sub

Private Sub TextMove(ByVal tx As Double, ByVal ty As Double)
    Dim i As Long
    FlushRun
    tlm(4) = tx * tlm(0) + ty * tlm(2) + tlm(4)
    tlm(5) = tx * tlm(1) + ty * tlm(3) + tlm(5)
    For i = 0 To 5
        tmx(i) = tlm(i)
    Next i
End Sub

' Ýçerikteki [...] dizisi (TJ): sayý ve Byte() öðeleri
Private Function ReadContentArray(b() As Byte, ByRef p As Long, ByVal n As Long) As Variant
    Dim arr() As Variant, cnt As Long, c As Long, v As Variant
    ReDim arr(0 To 15)
    p = p + 1
    Do While p < n
        SkipWS b, p, n
        If p >= n Then Exit Do
        c = b(p)
        v = Empty
        If c = 93 Then
            p = p + 1
            Exit Do
        ElseIf (c >= 48 And c <= 57) Or c = 45 Or c = 46 Or c = 43 Then
            v = ReadNumber(b, p, n)
        ElseIf c = 40 Then
            v = ReadLitBytes(b, p, n)
        ElseIf c = 60 Then
            v = ReadHexBytes(b, p, n)
        ElseIf c = 47 Then
            v = "/" & ReadName(b, p, n)
        Else
            ReadKeyword b, p, n
        End If
        If Not IsEmpty(v) Then
            If cnt > UBound(arr) Then
                ReDim Preserve arr(0 To 2 * cnt)
            End If
            arr(cnt) = v
            cnt = cnt + 1
        End If
    Loop
    If cnt = 0 Then
        ReadContentArray = Empty
    Else
        ReDim Preserve arr(0 To cnt - 1)
        ReadContentArray = arr
    End If
End Function

' << ... >> (iþaretli içerik özellikleri) atlanýr
Private Sub SkipDictBytes(b() As Byte, ByRef p As Long, ByVal n As Long)
    Dim depth As Long
    Do While p + 1 < n
        If b(p) = 60 And b(p + 1) = 60 Then
            depth = depth + 1: p = p + 2
        ElseIf b(p) = 62 And b(p + 1) = 62 Then
            depth = depth - 1: p = p + 2
            If depth <= 0 Then Exit Sub
        ElseIf b(p) = 40 Then
            ReadLitBytes b, p, n
        Else
            p = p + 1
        End If
    Loop
    p = n
End Sub

' BI ... ID <veri> EI (satýr içi resim) atlanýr
Private Sub SkipInlineImage(b() As Byte, ByRef p As Long, ByVal n As Long)
    Do While p + 2 < n
        If b(p) = 73 Then
            If b(p + 1) = 68 Then
                If IsWS(b(p + 2)) And IsWS(b(p - 1)) Then
                    p = p + 3
                    Exit Do
                End If
            End If
        End If
        p = p + 1
    Loop
    Do While p + 1 < n
        If b(p) = 69 Then
            If b(p + 1) = 73 And IsWS(b(p - 1)) Then
                If p + 2 >= n Then
                    p = n
                    Exit Sub
                ElseIf IsWS(b(p + 2)) Then
                    p = p + 2
                    Exit Sub
                End If
            End If
        End If
        p = p + 1
    Loop
    p = n
End Sub

Private Sub RunContent(b() As Byte, ByVal n As Long, ByVal res As Variant, ByVal depth As Long)
    Dim p As Long, c As Long, op As String, i As Long
    Dim nums(0 To 15) As Double, nN As Long, lastName As String, lastStr As Variant, lastArr As Variant
    Dim m(5) As Double, el As Variant, xo As Variant, xr As Variant, xn As Long, xd As Variant, xb As Variant
    Dim xbb() As Byte, isDic As Boolean
    p = 0
    Do While p < n
        c = b(p)
        If c = 32 Or c = 10 Or c = 13 Or c = 9 Or c = 12 Or c = 0 Then
            p = p + 1
        ElseIf (c >= 48 And c <= 57) Or c = 45 Or c = 46 Or c = 43 Then
            If nN > 15 Then
                For i = 0 To 14
                    nums(i) = nums(i + 1)
                Next i
                nN = 15
            End If
            nums(nN) = ReadNumber(b, p, n)
            nN = nN + 1
        ElseIf c = 47 Then
            lastName = "/" & ReadName(b, p, n)
        ElseIf c = 40 Then
            lastStr = ReadLitBytes(b, p, n)
        ElseIf c = 60 Then
            isDic = False
            If p + 1 < n Then isDic = (b(p + 1) = 60)
            If isDic Then
                SkipDictBytes b, p, n
            Else
                lastStr = ReadHexBytes(b, p, n)
            End If
        ElseIf c = 91 Then
            lastArr = ReadContentArray(b, p, n)
        ElseIf c = 37 Then
            SkipWS b, p, n
        ElseIf c = 93 Or c = 62 Or c = 41 Or c = 123 Or c = 125 Then
            p = p + 1
        Else
            op = ReadKeyword(b, p, n)
            Select Case op
                Case "q"
                    PushGS
                Case "Q"
                    PopGS
                Case "cm"
                    If nN >= 6 Then
                        For i = 0 To 5
                            m(i) = nums(nN - 6 + i)
                        Next i
                        MatMul m, ctm, ctm
                    End If
                Case "BT"
                    runOn = False
                    SetIdent tmx
                    SetIdent tlm
                Case "ET"
                    FlushRun
                Case "Tf"
                    If nN >= 1 Then
                        FlushRun
                        curFont = FontIndex(res, lastName)
                        curSize = nums(nN - 1)
                    End If
                Case "Td"
                    If nN >= 2 Then TextMove nums(nN - 2), nums(nN - 1)
                Case "TD"
                    If nN >= 2 Then
                        tLead = -nums(nN - 1)
                        TextMove nums(nN - 2), nums(nN - 1)
                    End If
                Case "Tm"
                    If nN >= 6 Then
                        FlushRun
                        For i = 0 To 5
                            tlm(i) = nums(nN - 6 + i)
                            tmx(i) = tlm(i)
                        Next i
                    End If
                Case "T*"
                    TextMove 0, -tLead
                Case "TL"
                    If nN >= 1 Then tLead = nums(nN - 1)
                Case "Tc"
                    If nN >= 1 Then tChar = nums(nN - 1)
                Case "Tw"
                    If nN >= 1 Then tWord = nums(nN - 1)
                Case "Tz"
                    If nN >= 1 Then tHs = nums(nN - 1) / 100
                Case "Ts"
                    If nN >= 1 Then tRise = nums(nN - 1)
                Case "Tj"
                    ShowBytes lastStr
                    FlushRun
                Case "'"
                    TextMove 0, -tLead
                    ShowBytes lastStr
                    FlushRun
                Case """"
                    If nN >= 2 Then
                        tWord = nums(nN - 2)
                        tChar = nums(nN - 1)
                    End If
                    TextMove 0, -tLead
                    ShowBytes lastStr
                    FlushRun
                Case "TJ"
                    If IsArray(lastArr) Then
                        el = lastArr
                        For i = 0 To UBound(el)
                            If VarType(el(i)) = vbDouble Then
                                TJAdjust el(i)
                            Else
                                ShowBytes el(i)
                            End If
                        Next i
                    End If
                    FlushRun
                Case "Do"
                    If Len(lastName) > 1 And depth < 8 Then
                        LetV xo, Resolve(DGet(res, "XObject"))
                        LetV xr, DGet(xo, Mid$(lastName, 2))
                        xn = RefNum(xr)
                        If xn > 0 Then
                            LetV xd, GetObj(xn)
                            If NameOf(DGet(xd, "Subtype")) = "/Form" Then
                                xb = GetStreamBytes(xn)
                                If IsArray(xb) Then
                                    xbb = xb
                                    PushGS
                                    LetV el, Resolve(DGet(xd, "Matrix"))
                                    If IsArray(el) Then
                                        If UBound(el) >= 5 Then
                                            For i = 0 To 5
                                                m(i) = Nv(Resolve(el(i)))
                                            Next i
                                            MatMul m, ctm, ctm
                                        End If
                                    End If
                                    LetV xr, Resolve(DGet(xd, "Resources"))
                                    If Not IsDict(xr) Then LetV xr, res
                                    RunContent xbb, UBound(xbb) + 1, xr, depth + 1
                                    PopGS
                                End If
                            End If
                        End If
                    End If
                Case "BI"
                    SkipInlineImage b, p, n
            End Select
            nN = 0
            If Len(lastName) > 0 Then lastName = ""
            If Not IsEmpty(lastStr) Then lastStr = Empty
            If Not IsEmpty(lastArr) Then lastArr = Empty
        End If
    Loop
    FlushRun
End Sub

' ---- yazý gösterimi ----
Private Sub StartRun()
    MatMul tmx, ctm, runM
    runStr = ""
    runTx = 0
    runSize = curSize
    runHs = tHs
    runRise = tRise
    runOn = True
End Sub

Private Sub ShowBytes(ByVal v As Variant)
    Dim i As Long, ub As Long, code As Long, w As Double, adv As Double, s As String, cid As Boolean
    If Not IsArray(v) Then Exit Sub
    If curFont > 0 Then cid = fnCID(curFont)
    If Not runOn Then StartRun
    i = LBound(v)
    ub = UBound(v)
    Do While i <= ub
        If cid Then
            code = CLng(v(i)) * 256
            If i + 1 <= ub Then code = code + v(i + 1)
            i = i + 2
        Else
            code = v(i)
            i = i + 1
        End If
        s = s & MapChar(code, cid)
        w = GlyphW(code, cid)
        adv = w / 1000# * curSize + tChar
        If code = 32 And Not cid Then adv = adv + tWord
        runTx = runTx + adv * tHs
    Loop
    runStr = runStr & s
End Sub

' TJ içindeki aralýk: küçükse ayný yazý, büyükse (0,3 em üstü) ayrý yazý
Private Sub TJAdjust(ByVal num As Double)
    Dim dx As Double
    dx = -num / 1000# * curSize * tHs
    If Abs(num) > 300 Or Not runOn Then
        FlushRun
        tmx(4) = tmx(4) + dx * tmx(0)
        tmx(5) = tmx(5) + dx * tmx(1)
    Else
        runTx = runTx + dx
    End If
End Sub

Private Sub FlushRun()
    If Not runOn Then Exit Sub
    EmitToken
    tmx(4) = tmx(4) + runTx * tmx(0)
    tmx(5) = tmx(5) + runTx * tmx(1)
    runOn = False
End Sub

Private Sub EmitToken()
    Dim a As Double, bb As Double, c As Double, d As Double, x0 As Double, x1 As Double, y As Double, h As Double
    Dim s As String, L As Long, i As Long, st As Long, cw As Double, seg As String
    If Len(Trim$(runStr)) = 0 Then Exit Sub
    a = runSize * runHs * runM(0): bb = runSize * runHs * runM(1)
    c = runSize * runM(2): d = runSize * runM(3)
    If a <= 0 Or d <= 0 Then Exit Sub                          ' ters / aynalý yazý
    If Abs(bb) > 0.05 * a Or Abs(c) > 0.05 * d Then Exit Sub     ' döndürülmüþ yazý (ölçüler)
    h = Sqr(c * c + d * d)
    x0 = runM(4)
    x1 = runTx * runM(0) + runM(4)
    If x1 < x0 Then
        y = x0: x0 = x1: x1 = y
    End If
    y = pageTop - (runM(5) + runRise * runM(3))
    ' 2+ boþlukla ayrýlmýþ parçalar ayrý yazý sayýlýr
    s = Replace(runStr, ChrW(160), " ")
    L = Len(s)
    cw = (x1 - x0) / L
    i = 1
    Do While i <= L
        Do While i <= L
            If Mid$(s, i, 1) <> " " Then Exit Do
            i = i + 1
        Loop
        If i > L Then Exit Do
        st = i
        Do While i <= L
            If Mid$(s, i, 2) = "  " Then Exit Do
            i = i + 1
        Loop
        seg = RTrim$(Mid$(s, st, i - st))
        If Len(seg) > 0 Then AddToken seg, x0 + (st - 1) * cw, x0 + (st - 1 + Len(seg)) * cw, y, h
    Loop
End Sub

Private Sub AddToken(ByVal s As String, ByVal x0 As Double, ByVal x1 As Double, ByVal y As Double, ByVal h As Double)
    Dim nu As Long
    tkN = tkN + 1
    If tkN = 1 Then
        ReDim tkS(1 To 256): ReDim tkX0(1 To 256): ReDim tkX1(1 To 256): ReDim tkY(1 To 256): ReDim tkH(1 To 256)
    ElseIf tkN > UBound(tkS) Then
        nu = UBound(tkS) * 2
        ReDim Preserve tkS(1 To nu): ReDim Preserve tkX0(1 To nu): ReDim Preserve tkX1(1 To nu)
        ReDim Preserve tkY(1 To nu): ReDim Preserve tkH(1 To nu)
    End If
    tkS(tkN) = s: tkX0(tkN) = x0: tkX1(tkN) = x1: tkY(tkN) = y: tkH(tkN) = h
End Sub

' =========================================================================
' FONTLAR
' =========================================================================
Private Function FontIndex(ByVal res As Variant, ByVal nameV As Variant) As Long
    Dim fd As Variant, ref As Variant, fi As Long, f As Variant
    If VarType(nameV) <> vbString Then Exit Function
    LetV fd, Resolve(DGet(res, "Font"))
    LetV ref, DGet(fd, Mid$(nameV, 2))
    If RefNum(ref) > 0 Then
        On Error Resume Next
        fi = 0
        fi = fntKey.Item(CStr(ref))
        On Error GoTo 0
        If fi = 0 Then
            LetV f, Resolve(ref)
            fi = LoadFont(f)
            If fi > 0 Then fntKey.Add fi, CStr(ref)
        End If
    ElseIf IsDict(ref) Then
        fi = LoadFont(ref)
    End If
    FontIndex = fi
End Function

Private Function LoadFont(ByVal fd As Variant) As Long
    Dim i As Long, desc As Variant, w As Variant, fc As Long, k As Long, j As Long, arr() As Double
    Dim tu As Variant, cmb As Variant, dw As Variant
    If Not IsDict(fd) Then Exit Function
    fnN = fnN + 1
    i = fnN
    ReDim Preserve fnCID(1 To i): ReDim Preserve fnW(1 To i): ReDim Preserve fnDW(1 To i)
    ReDim Preserve fnCW(1 To i): ReDim Preserve fnUni(1 To i): ReDim Preserve fnRng(1 To i)
    Set fnCW(i) = New Collection
    Set fnUni(i) = New Collection
    fnRng(i) = Empty
    fnDW(i) = 1000
    If NameOf(DGet(fd, "Subtype")) = "/Type0" Then
        fnCID(i) = True
        LetV desc, Resolve(DGet(fd, "DescendantFonts"))
        If IsArray(desc) Then
            If UBound(desc) >= 0 Then
                LetV desc, Resolve(desc(0))
                LetV dw, Resolve(DGet(desc, "DW"))
                If VarType(dw) = vbDouble Then fnDW(i) = dw
                LetV w, Resolve(DGet(desc, "W"))
                If IsArray(w) Then ParseCidW w, fnCW(i)
            End If
        End If
    Else
        fnCID(i) = False
        ReDim arr(0 To 255)
        For k = 0 To 255
            arr(k) = 500
        Next k
        arr(32) = 278
        fc = CLng(Nv(Resolve(DGet(fd, "FirstChar"))))
        LetV w, Resolve(DGet(fd, "Widths"))
        If IsArray(w) Then
            For j = 0 To UBound(w)
                If fc + j >= 0 And fc + j <= 255 Then arr(fc + j) = Nv(Resolve(w(j)))
            Next j
        End If
        fnW(i) = arr
    End If
    LetV tu, DGet(fd, "ToUnicode")
    If RefNum(tu) > 0 Then
        cmb = GetStreamBytes(RefNum(tu))
        If IsArray(cmb) Then ParseCMap cmb, i
    End If
    LoadFont = i
End Function

Private Sub ParseCidW(ByVal w As Variant, ByVal col As Collection)
    Dim j As Long, c1 As Long, c2 As Long, k As Long, ww As Double, sa As Variant
    On Error Resume Next
    j = 0
    Do While j < UBound(w)
        c1 = CLng(Nv(Resolve(w(j))))
        If IsArray(w(j + 1)) Then
            sa = w(j + 1)
            For k = 0 To UBound(sa)
                col.Add Nv(Resolve(sa(k))), CStr(c1 + k)
            Next k
            j = j + 2
        Else
            If j + 2 > UBound(w) Then Exit Do
            c2 = CLng(Nv(Resolve(w(j + 1))))
            ww = Nv(Resolve(w(j + 2)))
            If c2 - c1 < 20000 Then
                For k = c1 To c2
                    col.Add ww, CStr(k)
                Next k
            End If
            j = j + 3
        End If
    Loop
End Sub

Private Function GlyphW(ByVal code As Long, ByVal cid As Boolean) As Double
    Dim w As Double, col As Collection, wa As Variant
    If curFont <= 0 Then
        GlyphW = 500
        Exit Function
    End If
    If cid Then
        w = fnDW(curFont)
        Set col = fnCW(curFont)
        On Error Resume Next
        w = col.Item(CStr(code))
        On Error GoTo 0
        GlyphW = w
    ElseIf code >= 0 And code <= 255 Then
        wa = fnW(curFont)
        GlyphW = wa(code)
    Else
        GlyphW = 500
    End If
End Function

Private Function MapChar(ByVal code As Long, ByVal cid As Boolean) As String
    Dim col As Collection, s As String, found As Boolean, rg As Variant, k As Long, e As Variant, cp As Long
    If curFont > 0 Then
        Set col = fnUni(curFont)
        If col.Count > 0 Then
            On Error Resume Next
            Err.Clear
            s = col.Item(CStr(code))
            found = (Err.Number = 0)
            Err.Clear
            On Error GoTo 0
            If found Then
                MapChar = s
                Exit Function
            End If
        End If
        rg = fnRng(curFont)
        If IsArray(rg) Then
            For k = 0 To UBound(rg)
                e = rg(k)
                If code >= e(0) And code <= e(1) Then
                    cp = e(3) + code - e(0)
                    If cp >= 0 And cp <= 65535 Then MapChar = e(2) & ChrW(cp)
                    Exit Function
                End If
            Next k
        End If
    End If
    If cid Then
        If code >= 32 Then MapChar = ChrW(code)
    ElseIf code >= 32 And code <= 255 Then
        MapChar = ChrW(winAnsi(code))
    End If
End Function

' ToUnicode CMap: beginbfchar / beginbfrange
Private Sub ParseCMap(ByVal v As Variant, ByVal fi As Long)
    Dim s As String, a As Long, e As Long, toks As Variant, k As Long, lo As Long, hi As Long, pos As Long
    Dim dst As String, rgs() As Variant, nr As Long, col As Collection, sa As Variant, j As Long
    s = BytesToText(v)
    Set col = fnUni(fi)
    On Error Resume Next
    pos = 1
    Do
        a = InStr(pos, s, "beginbfchar")
        If a = 0 Then Exit Do
        e = InStr(a, s, "endbfchar")
        If e = 0 Then e = Len(s) + 1
        toks = CMapTokens(Mid$(s, a + 11, e - a - 11))
        If IsArray(toks) Then
            For k = 0 To UBound(toks) - 1 Step 2
                col.Add HexToUni(toks(k + 1)), CStr(HexToLong(toks(k)))
            Next k
        End If
        pos = e + 1
    Loop
    nr = 0
    ReDim rgs(0 To 7)
    pos = 1
    Do
        a = InStr(pos, s, "beginbfrange")
        If a = 0 Then Exit Do
        e = InStr(a, s, "endbfrange")
        If e = 0 Then e = Len(s) + 1
        toks = CMapTokens(Mid$(s, a + 12, e - a - 12))
        If IsArray(toks) Then
            k = 0
            Do While k + 2 <= UBound(toks)
                lo = HexToLong(toks(k)): hi = HexToLong(toks(k + 1))
                If toks(k + 2) = "[" Then
                    j = k + 3
                    Do While j <= UBound(toks)
                        If toks(j) = "]" Then Exit Do
                        col.Add HexToUni(toks(j)), CStr(lo + j - k - 3)
                        j = j + 1
                    Loop
                    k = j + 1
                Else
                    dst = HexToUni(toks(k + 2))
                    If Len(dst) > 0 Then
                        If nr > UBound(rgs) Then
                            ReDim Preserve rgs(0 To 2 * nr)
                        End If
                        rgs(nr) = Array(lo, hi, Left$(dst, Len(dst) - 1), AscW(Right$(dst, 1)) And 65535)
                        nr = nr + 1
                    End If
                    k = k + 3
                End If
            Loop
        End If
        pos = e + 1
    Loop
    On Error GoTo 0
    If nr > 0 Then
        ReDim Preserve rgs(0 To nr - 1)
        fnRng(fi) = rgs
    End If
End Sub

Private Function CMapTokens(ByVal s As String) As Variant
    Dim out() As String, n As Long, i As Long, ch As String, e As Long
    ReDim out(0 To 63)
    i = 1
    Do While i <= Len(s)
        ch = Mid$(s, i, 1)
        If ch = "<" Then
            e = InStr(i + 1, s, ">")
            If e = 0 Then Exit Do
            If n > UBound(out) Then
                ReDim Preserve out(0 To 2 * n)
            End If
            out(n) = Mid$(s, i + 1, e - i - 1): n = n + 1
            i = e + 1
        ElseIf ch = "[" Or ch = "]" Then
            If n > UBound(out) Then
                ReDim Preserve out(0 To 2 * n)
            End If
            out(n) = ch: n = n + 1
            i = i + 1
        Else
            i = i + 1
        End If
    Loop
    If n = 0 Then
        CMapTokens = Empty
    Else
        ReDim Preserve out(0 To n - 1)
        CMapTokens = out
    End If
End Function

Private Function HexToLong(ByVal h As String) As Long
    Dim i As Long, v As Long, d As Long
    For i = 1 To Len(h)
        d = HexNib(AscW(Mid$(h, i, 1)))
        If d >= 0 Then v = v * 16 + d
        If v > 16777215 Then Exit For
    Next i
    HexToLong = v
End Function

Private Function HexToUni(ByVal h As String) As String
    Dim i As Long, s As String, v As Long
    h = Replace(Replace(Replace(h, " ", ""), vbCr, ""), vbLf, "")
    If Len(h) <= 2 Then
        If Len(h) > 0 Then HexToUni = ChrW(HexToLong(h))
        Exit Function
    End If
    For i = 1 To Len(h) - 3 Step 4
        v = HexToLong(Mid$(h, i, 4))
        s = s & ChrW(v)
    Next i
    HexToUni = s
End Function

' =========================================================================
' INFLATE (zlib / RFC 1951)
' =========================================================================
Private Sub InitTables()
    Dim i As Long, b As Long
    If tablesReady Then Exit Sub
    p2(0) = 1
    For i = 1 To 30
        p2(i) = p2(i - 1) * 2
    Next i
    b = 3
    For i = 0 To 27
        If i < 8 Then lenExt(i) = 0 Else lenExt(i) = (i - 4) \ 4
        lenBase(i) = b
        b = b + p2(lenExt(i))
    Next i
    lenExt(28) = 0: lenBase(28) = 258
    b = 1
    For i = 0 To 29
        If i < 2 Then distExt(i) = 0 Else distExt(i) = (i - 2) \ 2
        distBase(i) = b
        b = b + p2(distExt(i))
    Next i
    For i = 0 To 255
        winAnsi(i) = i
    Next i
    winAnsi(128) = 8364: winAnsi(130) = 8218: winAnsi(131) = 402: winAnsi(132) = 8222
    winAnsi(133) = 8230: winAnsi(134) = 8224: winAnsi(135) = 8225: winAnsi(136) = 710
    winAnsi(137) = 8240: winAnsi(138) = 352: winAnsi(139) = 8249: winAnsi(140) = 338
    winAnsi(142) = 381: winAnsi(145) = 8216: winAnsi(146) = 8217: winAnsi(147) = 8220
    winAnsi(148) = 8221: winAnsi(149) = 8226: winAnsi(150) = 8211: winAnsi(151) = 8212
    winAnsi(152) = 732: winAnsi(153) = 8482: winAnsi(154) = 353: winAnsi(155) = 8250
    winAnsi(156) = 339: winAnsi(158) = 382: winAnsi(159) = 376
    tablesReady = True
End Sub

Private Function Inflate(src() As Byte) As Variant
    Dim bfinal As Long, btype As Long, ln As Long, i As Long, outB() As Byte
    Inflate = Empty
    zi = src
    ziPos = 0
    ziEnd = UBound(zi) + 1
    zBit = 0: zCnt = 0: zEOF = False
    ReDim zo(0 To MaxL(4096, ziEnd * 4))
    zoPos = 0
    If ziEnd >= 2 Then
        If (zi(0) And 15) = 8 And ((CLng(zi(0)) * 256 + zi(1)) Mod 31) = 0 Then ziPos = 2
    End If
    Do
        bfinal = GetBits(1)
        btype = GetBits(2)
        If zEOF Then Exit Do
        Select Case btype
            Case 0
                ' bayt sýnýrýna hizala: yarým baytý at, tamponda bekleyen tam baytlarý girdiye geri ver
                zBit = zBit \ p2(zCnt Mod 8)
                zCnt = zCnt - (zCnt Mod 8)
                ziPos = ziPos - zCnt \ 8
                zBit = 0: zCnt = 0
                If ziPos + 4 > ziEnd Then Exit Do
                ln = zi(ziPos) + 256& * zi(ziPos + 1)
                ziPos = ziPos + 4
                For i = 1 To ln
                    If ziPos >= ziEnd Then zEOF = True: Exit For
                    PutByte zi(ziPos)
                    ziPos = ziPos + 1
                Next i
            Case 1
                BuildFixed
                InflateBlock
            Case 2
                If Not DecodeTrees() Then Exit Do
                InflateBlock
            Case Else
                Exit Do
        End Select
        If zEOF Then Exit Do
    Loop Until bfinal = 1
    ' sonuç ayrý bir diziye kopyalanýr (bir sonraki çözme iþlemi bu tamponu yeniden kullanýr)
    If zoPos > 0 Then
        ReDim outB(0 To zoPos - 1)
        For i = 0 To zoPos - 1
            outB(i) = zo(i)
        Next i
        Inflate = outB
    End If
End Function

' Bit tamponu: zBit'te en fazla 24 bit bekler (LSB önce); ZFill en az 17 bite tamamlar
Private Sub ZFill()
    Do While zCnt <= 16
        If ziPos >= ziEnd Then Exit Do
        zBit = zBit + zi(ziPos) * p2(zCnt)
        ziPos = ziPos + 1
        zCnt = zCnt + 8
    Loop
End Sub

Private Function GetBits(ByVal n As Long) As Long
    If n <= 0 Then Exit Function
    If zCnt < n Then ZFill
    GetBits = zBit And (p2(n) - 1)
    zBit = zBit \ p2(n)
    zCnt = zCnt - n
    If zCnt < 0 Then
        zEOF = True
        zCnt = 0
    End If
End Function

Private Sub PutByte(ByVal v As Long)
    If zoPos > UBound(zo) Then
        ReDim Preserve zo(0 To 2 * UBound(zo) + 1)
    End If
    zo(zoPos) = v
    zoPos = zoPos + 1
End Sub

' which: 0 = uzunluk/harf tablosu, 1 = mesafe tablosu
' Kanonik Huffman tablosu (yavaþ yol) + 9 bite kadar kodlar için hýzlý arama tablosu
Private Sub BuildTree(lens() As Long, ByVal off As Long, ByVal num As Long, ByVal which As Long)
    Dim cnt(15) As Long, offs(15) As Long, nextc(15) As Long, i As Long, s As Long, L As Long
    Dim code As Long, rv As Long, b As Long, k As Long
    For i = 0 To num - 1
        cnt(lens(off + i)) = cnt(lens(off + i)) + 1
    Next i
    cnt(0) = 0
    s = 0
    For i = 0 To 15
        offs(i) = s
        s = s + cnt(i)
    Next i
    For i = 0 To 15
        If which = 0 Then ltC(i) = cnt(i) Else dtC(i) = cnt(i)
    Next i
    For i = 0 To 511
        If which = 0 Then fastL(i) = -1 Else fastD(i) = -1
    Next i
    code = 0
    For L = 1 To 15
        code = (code + cnt(L - 1)) * 2
        nextc(L) = code
    Next L
    For i = 0 To num - 1
        L = lens(off + i)
        If L > 0 Then
            If which = 0 Then ltS(offs(L)) = i Else dtS(offs(L)) = i
            offs(L) = offs(L) + 1
            code = nextc(L)
            nextc(L) = code + 1
            If L <= 9 Then
                ' deflate bitleri ters sýrada okur: kodu ters çevir, boþta kalan üst bitlerin tüm deðerlerini doldur
                rv = 0
                For b = 0 To L - 1
                    If (code And p2(b)) <> 0 Then rv = rv + p2(L - 1 - b)
                Next b
                k = rv
                Do While k < 512
                    If which = 0 Then fastL(k) = i * 16 + L Else fastD(k) = i * 16 + L
                    k = k + p2(L)
                Loop
            End If
        End If
    Next i
End Sub

Private Sub BuildFixed()
    Dim lens(319) As Long, i As Long
    For i = 0 To 143: lens(i) = 8: Next i
    For i = 144 To 255: lens(i) = 9: Next i
    For i = 256 To 279: lens(i) = 7: Next i
    For i = 280 To 287: lens(i) = 8: Next i
    BuildTree lens, 0, 288, 0
    For i = 0 To 29: lens(i) = 5: Next i
    BuildTree lens, 0, 30, 1
End Sub

Private Function DecodeSym(ByVal which As Long) As Long
    Dim cur As Long, s As Long, L As Long, cn As Long, e As Long
    ' hýzlý yol: 9 bite kadar kodlar tek tabloda
    If zCnt < 9 Then ZFill
    If which = 0 Then e = fastL(zBit And 511) Else e = fastD(zBit And 511)
    If e >= 0 Then
        L = e And 15
        If L <= zCnt Then
            zBit = zBit \ p2(L)
            zCnt = zCnt - L
            DecodeSym = e \ 16
            Exit Function
        End If
    End If
    ' yavaþ yol: uzun kodlar / dosya sonu (kanonik, bit bit)
    L = 0
    Do
        If zCnt = 0 Then
            ZFill
            If zCnt = 0 Then
                zEOF = True
                DecodeSym = 256
                Exit Function
            End If
        End If
        cur = 2 * cur + (zBit And 1)
        zBit = zBit \ 2
        zCnt = zCnt - 1
        L = L + 1
        If which = 0 Then cn = ltC(L) Else cn = dtC(L)
        s = s + cn
        cur = cur - cn
        If cur < 0 Then Exit Do
        If L >= 15 Then
            zEOF = True
            DecodeSym = 256
            Exit Function
        End If
    Loop
    If which = 0 Then DecodeSym = ltS(s + cur) Else DecodeSym = dtS(s + cur)
End Function

Private Function DecodeTrees() As Boolean
    Dim hlit As Long, hdist As Long, hclen As Long, i As Long, num As Long, sym As Long, ln As Long, prev As Long
    Dim cl(18) As Long, lens(319) As Long, order As Variant
    order = Array(16, 17, 18, 0, 8, 7, 9, 6, 10, 5, 11, 4, 12, 3, 13, 2, 14, 1, 15)
    hlit = GetBits(5) + 257
    hdist = GetBits(5) + 1
    hclen = GetBits(4) + 4
    For i = 0 To hclen - 1
        cl(order(i)) = GetBits(3)
    Next i
    BuildTree cl, 0, 19, 0
    num = 0
    Do While num < hlit + hdist
        sym = DecodeSym(0)
        If zEOF Then Exit Function
        Select Case sym
            Case 16
                If num = 0 Then Exit Function
                prev = lens(num - 1)
                ln = GetBits(2) + 3
            Case 17
                prev = 0
                ln = GetBits(3) + 3
            Case 18
                prev = 0
                ln = GetBits(7) + 11
            Case Else
                prev = sym
                ln = 1
        End Select
        If num + ln > hlit + hdist Then Exit Function
        For i = 1 To ln
            lens(num) = prev
            num = num + 1
        Next i
    Loop
    BuildTree lens, 0, hlit, 0
    BuildTree lens, hlit, hdist, 1
    DecodeTrees = True
End Function

Private Sub InflateBlock()
    Dim sym As Long, ln As Long, ds As Long, dist As Long, i As Long
    Do
        sym = DecodeSym(0)
        If zEOF Then Exit Do
        If sym < 256 Then
            If zoPos > UBound(zo) Then
                ReDim Preserve zo(0 To 2 * UBound(zo) + 1)
            End If
            zo(zoPos) = sym
            zoPos = zoPos + 1
        ElseIf sym = 256 Then
            Exit Do
        Else
            sym = sym - 257
            If sym > 28 Then zEOF = True: Exit Do
            ln = GetBits(lenExt(sym)) + lenBase(sym)
            ds = DecodeSym(1)
            If ds > 29 Or zEOF Then zEOF = True: Exit Do
            dist = GetBits(distExt(ds)) + distBase(ds)
            If dist > zoPos Then zEOF = True: Exit Do
            If zoPos + ln > UBound(zo) Then
                ReDim Preserve zo(0 To 2 * UBound(zo) + ln + 1)
            End If
            For i = 1 To ln
                zo(zoPos) = zo(zoPos - dist)
                zoPos = zoPos + 1
            Next i
        End If
    Loop
End Sub

' =========================================================================
' TABLO KURMA (PDF_BOM_Okuyucu.html ile ayný kurallar)
' =========================================================================
Private Function FoldPdf(ByVal s As String) As String
    Dim t As String
    t = UCase$(Trim$(Replace(s, ChrW(160), " ")))
    t = Replace(t, ChrW(196), "A"): t = Replace(t, ChrW(197), "A"): t = Replace(t, ChrW(198), "AE")
    t = Replace(t, ChrW(192), "A"): t = Replace(t, ChrW(193), "A"): t = Replace(t, ChrW(194), "A")
    t = Replace(t, ChrW(199), "C"): t = Replace(t, ChrW(200), "E"): t = Replace(t, ChrW(201), "E")
    t = Replace(t, ChrW(202), "E"): t = Replace(t, ChrW(214), "O"): t = Replace(t, ChrW(216), "O")
    t = Replace(t, ChrW(211), "O"): t = Replace(t, ChrW(220), "U"): t = Replace(t, ChrW(218), "U")
    t = Replace(t, ChrW(223), "SS"): t = Replace(t, ChrW(7838), "SS"): t = Replace(t, ChrW(304), "I")
    t = Replace(t, ChrW(305), "I"): t = Replace(t, ChrW(350), "S"): t = Replace(t, ChrW(286), "G")
    FoldPdf = t
End Function

Private Function StartsAny(ByVal c As String, ByVal list As String) As Boolean
    Dim a As Variant, i As Long
    a = Split(list, "|")
    For i = 0 To UBound(a)
        If Left$(c, Len(a(i))) = a(i) Then
            StartsAny = True
            Exit Function
        End If
    Next i
End Function

Private Function ContainsAny(ByVal c As String, ByVal list As String) As Boolean
    Dim a As Variant, i As Long
    a = Split(list, "|")
    For i = 0 To UBound(a)
        If InStr(c, a(i)) > 0 Then
            ContainsAny = True
            Exit Function
        End If
    Next i
End Function

Private Function InList(ByVal c As String, ByVal list As String) As Boolean
    InList = (InStr("|" & list & "|", "|" & c & "|") > 0)
End Function

' Baþlýk rolleri: 1 poz, 2 adet, 3 taným, 4 boy, 5 malzeme, 6 aðýrlýk, 7 profil, 8 doküman no, 16 "No." (belirsiz)
Private Function RoleOf(ByVal txt As String) As Long
    Dim t As String, c As String
    t = FoldPdf(txt)
    If t = "" Then Exit Function
    If Right$(t, 1) = ":" Or Len(t) > 30 Then Exit Function
    c = Replace(Replace(Replace(Replace(Replace(t, " ", ""), ".", ""), "-", ""), "_", ""), ":", "")
    If InList(c, "ITEMNO|ITEM|POS|POSNR|POSNO|POSITION|MARK|MARKNO|PARTNO|PARTMARK|POZ|POZNO|TEILENR") Then RoleOf = 1: Exit Function
    If StartsAny(c, "QTY|QUANTITY|STCK|STK|STUCK|ANZ|MENGE|ADET|MIKTAR|PCS|ANT") Then RoleOf = 2: Exit Function
    If InList(c, "NO|NR|NOS|NUMBER") Then RoleOf = 16: Exit Function
    If InList(c, "DOCNO|DOCNR|DOCUMENTNO|DOKNO|DOKUMENTNR|DRAWINGNO|DWGNO|ZEICHNUNGSNR|ZEICHNR") Then RoleOf = 8: Exit Function
    If ContainsAny(c, "LENGTH|LANGE|LAENGE|LENGDE|UZUNLUK") Or c = "BOY" Then RoleOf = 4: Exit Function
    If ContainsAny(c, "PROFIL|SECTION|KESIT") Then RoleOf = 7: Exit Function
    If ContainsAny(c, "DESCRIPTION|BENENNUNG|BEZEICHNUNG|BESKRIVELSE|TANIM") Then RoleOf = 3: Exit Function
    If ContainsAny(c, "WEIGHT|GEWICHT|VEKT|AGIRLIK") Then RoleOf = 6: Exit Function
    If ContainsAny(c, "MATERIAL|WERKSTOFF|GUTE|GRADE|QUALITY|KALITE|MALZEME") And InStr(c, "AUSZUG") = 0 Then RoleOf = 5: Exit Function
End Function

Private Function IsHeaderRoles(rl() As Long, ByVal a As Long, ByVal b As Long) As Boolean
    Dim i As Long, has(16) As Boolean, nDistinct As Long, pos As Boolean, qty As Boolean, nm As Boolean, core As Long
    For i = a To b
        If rl(i) > 0 Then
            If Not has(rl(i)) Then nDistinct = nDistinct + 1
            has(rl(i)) = True
        End If
    Next i
    pos = has(1): qty = has(2)
    If has(16) Then
        If pos Then qty = True Else pos = True
    End If
    nm = has(3) Or has(7)
    If pos Then core = core + 1
    If qty Then core = core + 1
    If nm Then core = core + 1
    IsHeaderRoles = (core >= 2 And nDistinct >= 3)
End Function

Private Sub QSortTok(ord() As Long, ByVal lo As Long, ByVal hi As Long)
    Dim i As Long, j As Long, pv As Long, t As Long
    If lo >= hi Then Exit Sub
    i = lo: j = hi
    pv = ord((lo + hi) \ 2)
    Do While i <= j
        Do While TokLess(ord(i), pv)
            i = i + 1
        Loop
        Do While TokLess(pv, ord(j))
            j = j - 1
        Loop
        If i <= j Then
            t = ord(i): ord(i) = ord(j): ord(j) = t
            i = i + 1: j = j - 1
        End If
    Loop
    If lo < j Then QSortTok ord, lo, j
    If i < hi Then QSortTok ord, i, hi
End Sub

Private Function TokLess(ByVal a As Long, ByVal b As Long) As Boolean
    If tkY(a) <> tkY(b) Then
        TokLess = (tkY(a) < tkY(b))
    Else
        TokLess = (tkX0(a) < tkX0(b))
    End If
End Function

Private Sub BuildLines()
    Dim ord() As Long, i As Long, k As Long, cur() As Long, cn As Long
    ReDim ord(1 To tkN)
    For i = 1 To tkN
        ord(i) = i
    Next i
    QSortTok ord, 1, tkN
    lnN = 0
    ReDim lnY(1 To 64): ReDim lnH(1 To 64): ReDim lnIt(1 To 64)
    ReDim cur(0 To 63)
    cn = 0
    For i = 1 To tkN
        k = ord(i)
        If lnN > 0 Then
            If Abs(tkY(k) - lnY(lnN)) <= MaxD(0.45 * MinD(tkH(k), lnH(lnN)), 0.8) Then
                If cn > UBound(cur) Then
                    ReDim Preserve cur(0 To 2 * cn)
                End If
                cur(cn) = k
                cn = cn + 1
                lnY(lnN) = (lnY(lnN) * (cn - 1) + tkY(k)) / cn
                lnH(lnN) = MaxD(lnH(lnN), tkH(k))
                GoTo NextTok
            End If
            FinishLine cur, cn
        End If
        lnN = lnN + 1
        If lnN > UBound(lnY) Then
            ReDim Preserve lnY(1 To 2 * lnN): ReDim Preserve lnH(1 To 2 * lnN): ReDim Preserve lnIt(1 To 2 * lnN)
        End If
        lnY(lnN) = tkY(k): lnH(lnN) = tkH(k)
        cur(0) = k
        cn = 1
NextTok:
    Next i
    If lnN > 0 Then FinishLine cur, cn
End Sub

' Satýrdaki yazýlarý x'e göre sýralayýp saklar
Private Sub FinishLine(cur() As Long, ByVal cn As Long)
    Dim a() As Long, i As Long, j As Long, t As Long
    ReDim a(0 To cn - 1)
    For i = 0 To cn - 1
        a(i) = cur(i)
    Next i
    For i = 1 To cn - 1
        t = a(i)
        j = i - 1
        Do While j >= 0
            If tkX0(a(j)) <= tkX0(t) Then Exit Do
            a(j + 1) = a(j)
            j = j - 1
        Loop
        a(j + 1) = t
    Next i
    lnIt(lnN) = a
End Sub

Private Function IsNoWord(ByVal s As String) As Boolean
    s = UCase$(Trim$(s))
    IsNoWord = (s = "NO" Or s = "NR" Or s = "NO." Or s = "NR.")
End Function

Private Function EndsItemWord(ByVal s As String) As Boolean
    Dim a As Variant, i As Long
    s = UCase$(Trim$(s))
    If Right$(s, 1) = "." Then s = Left$(s, Len(s) - 1)
    a = Array("ITEM", "POS", "PART", "MARK", "TEILE")
    For i = 0 To UBound(a)
        If Right$(s, Len(a(i))) = a(i) Then
            EndsItemWord = True
            Exit Function
        End If
    Next i
End Function

Private Function EndsNo(ByVal s As String) As Boolean
    s = UCase$(Trim$(s))
    EndsNo = (Right$(s, 2) = "NO" Or Right$(s, 3) = "NO.")
End Function

Private Function LastWord(ByVal s As String) As String
    Dim a As Variant
    a = Split(Trim$(s), " ")
    If UBound(a) >= 0 Then LastWord = a(UBound(a))
End Function

' Baþlýk kelimeleri: "Item"+"No." -> "Item No."; "NET"+"WEIGHT" -> "NET WEIGHT"
Private Sub MergeHeader(ByVal li As Long, mS() As String, mX0() As Double, mX1() As Double, mH() As Double, ByRef mN As Long)
    Dim it As Variant, k As Long, t As Long, q As Long, m As Long, joined As Boolean, near As Boolean, gap As Double
    it = lnIt(li)
    mN = 0
    ReDim mS(0 To UBound(it)): ReDim mX0(0 To UBound(it)): ReDim mX1(0 To UBound(it)): ReDim mH(0 To UBound(it))
    For k = 0 To UBound(it)
        t = it(k)
        joined = False
        near = False
        If mN > 0 Then
            gap = tkX0(t) - mX1(mN - 1)
            near = (gap < 1.6 * MinD(tkH(t), mH(mN - 1)) And Abs(tkH(t) - mH(mN - 1)) < 0.3 * mH(mN - 1))
        End If
        If IsNoWord(tkS(t)) Then
            q = -1
            For m = mN - 1 To MaxL(0, mN - 3) Step -1
                If EndsItemWord(mS(m)) Then
                    q = m
                    Exit For
                End If
            Next m
            If q >= 0 Then
                If Not EndsNo(mS(q)) And tkX0(t) - mX1(q) < 3 * MinD(tkH(t), mH(q)) Then
                    mS(q) = mS(q) & " " & tkS(t)
                    mX1(q) = MaxD(mX1(q), tkX1(t))
                    joined = True
                End If
            End If
        End If
        If Not joined And near Then
            If RoleOf(mS(mN - 1)) = 0 And RoleOf(tkS(t)) <> 0 Then
                If InList(UCase$(LastWord(mS(mN - 1))), "NET|GROSS|TOT|TOT.|TOTAL|UNIT|EINZEL|GESAMT|BRUTTO|NETTO") Then
                    mS(mN - 1) = mS(mN - 1) & " " & tkS(t)
                    mX1(mN - 1) = MaxD(mX1(mN - 1), tkX1(t))
                    joined = True
                End If
            End If
        End If
        If Not joined Then
            mS(mN) = tkS(t): mX0(mN) = tkX0(t): mX1(mN) = tkX1(t): mH(mN) = tkH(t)
            mN = mN + 1
        End If
    Next k
End Sub

' Satýr bir tablo baþlýðý mý? Evet ise sütun hücreleri (sol->sað) ve yazý boyu döner
Private Function HeaderCells(ByVal li As Long, cS() As String, cX0() As Double, cX1() As Double, cR() As Long, _
                             ByRef nC As Long, ByRef H As Double) As Boolean
    Dim mS() As String, mX0() As Double, mX1() As Double, mH() As Double, mN As Long
    Dim rl() As Long, hs() As Double, nIdx As Long, i As Long, j As Long, t As Double, maxGap As Double
    Dim lo As Long, hi As Long, gs As Long, bestS As Long, bestE As Long, bestLen As Long, lastI As Long
    MergeHeader li, mS, mX0, mX1, mH, mN
    If mN = 0 Then Exit Function
    ReDim rl(0 To mN - 1): ReDim hs(0 To mN - 1)
    lo = -1
    For i = 0 To mN - 1
        rl(i) = RoleOf(mS(i))
        If rl(i) > 0 Then
            If lo < 0 Then lo = i
            hi = i
            hs(nIdx) = mH(i)
            nIdx = nIdx + 1
        End If
    Next i
    If nIdx = 0 Then Exit Function
    ' ortanca yazý boyu
    For i = 1 To nIdx - 1
        t = hs(i): j = i - 1
        Do While j >= 0
            If hs(j) <= t Then Exit Do
            hs(j + 1) = hs(j)
            j = j - 1
        Loop
        hs(j + 1) = t
    Next i
    H = hs(nIdx \ 2)
    maxGap = 30 * H
    ' rollü kelimeler arasýnda büyük boþluk varsa ayrý gruplar; en uzun geçerli grup
    bestLen = 0
    gs = lo
    For i = lo To hi
        If i = hi Then
            If IsHeaderRoles(rl, gs, i) And i - gs + 1 > bestLen Then bestS = gs: bestE = i: bestLen = i - gs + 1
        ElseIf mX0(i + 1) - mX1(i) > maxGap Then
            If IsHeaderRoles(rl, gs, i) And i - gs + 1 > bestLen Then bestS = gs: bestE = i: bestLen = i - gs + 1
            gs = i + 1
        End If
    Next i
    If bestLen = 0 Then Exit Function
    ReDim cS(0 To mN - 1): ReDim cX0(0 To mN - 1): ReDim cX1(0 To mN - 1): ReDim cR(0 To mN - 1)
    nC = 0
    For i = bestS To bestE
        If rl(i) > 0 Or i = bestS Then
            cS(nC) = mS(i): cX0(nC) = mX0(i): cX1(nC) = mX1(i): cR(nC) = rl(i)
            nC = nC + 1
        End If
    Next i
    ' saða doðru ayný boydaki rolsüz baþlýklar ("Bundle")
    lastI = bestE
    Do While lastI + 1 < mN
        If mX0(lastI + 1) - mX1(lastI) < 7 * H And Abs(mH(lastI + 1) - H) < 0.25 * H And rl(lastI + 1) = 0 Then
            cS(nC) = mS(lastI + 1): cX0(nC) = mX0(lastI + 1): cX1(nC) = mX1(lastI + 1): cR(nC) = 0
            nC = nC + 1
            lastI = lastI + 1
        Else
            Exit Do
        End If
    Loop
    HeaderCells = True
End Function

Private Function ColFor(cX0() As Double, cX1() As Double, ByVal nC As Long, ByVal H As Double, ByVal t As Long) As Long
    Dim i As Long, bi As Long, bo As Double, ov As Double
    bi = -1
    For i = 0 To nC - 1
        ov = MinD(tkX1(t), cX1(i)) - MaxD(tkX0(t), cX0(i))
        If ov > bo Then bo = ov: bi = i
    Next i
    If bi >= 0 Then
        ColFor = bi
        Exit Function
    End If
    For i = nC - 1 To 0 Step -1
        If tkX0(t) >= cX0(i) - 0.6 * H Then
            ColFor = i
            Exit Function
        End If
    Next i
    ColFor = 0
End Function

Private Function BuildTables() As Long
    Dim i As Long, j As Long, k As Long, t As Long, c As Long, nT As Long, found As Boolean
    Dim cS() As String, cX0() As Double, cX1() As Double, cR() As Long, nC As Long, H As Double
    Dim dS() As String, dX0() As Double, dX1() As Double, dR() As Long, dN As Long, dH As Double
    Dim lft As Double, rgt As Double, lastY As Double, pitch As Double
    Dim rY() As Double, rT() As Variant, nR As Long, it As Variant, toks As Variant
    If tkN = 0 Then Exit Function
    BuildLines
    i = 1
    Do While i <= lnN
        If Not HeaderCells(i, cS, cX0, cX1, cR, nC, H) Then
            i = i + 1
        Else
            lft = cX0(0) - 3 * H
            rgt = cX1(nC - 1) + 1.5 * H
            nR = 0
            ReDim rY(1 To 16): ReDim rT(1 To 16)
            lastY = lnY(i)
            pitch = 1.9 * H
            For j = i + 1 To lnN
                If HeaderCells(j, dS, dX0, dX1, dR, dN, dH) Then Exit For
                it = lnIt(j)
                toks = NewArr(nC, "")
                found = False
                For k = 0 To UBound(it)
                    t = it(k)
                    If tkX1(t) > lft And tkX0(t) < rgt And tkH(t) > 0.55 * H And tkH(t) < 1.7 * H Then
                        c = ColFor(cX0, cX1, nC, H, t)
                        toks(c) = toks(c) & "," & CStr(t)
                        found = True
                    End If
                Next k
                If Not found Then
                    If lnY(j) - lastY > 3 * pitch Then Exit For
                Else
                    If lnY(j) - lastY > 3.2 * pitch And nR > 0 Then Exit For
                    If nR > 0 Then pitch = MaxD(0.9 * H, MinD(3 * H, (lnY(j) - lastY) * 0.5 + pitch * 0.5))
                    nR = nR + 1
                    If nR > UBound(rY) Then
                        ReDim Preserve rY(1 To 2 * nR): ReDim Preserve rT(1 To 2 * nR)
                    End If
                    rY(nR) = lnY(j)
                    rT(nR) = toks
                    lastY = lnY(j)
                End If
            Next j
            If nR > 0 Then
                If FinishTable(cS, cR, nC, rY, rT, nR, pitch) Then nT = nT + 1
            End If
            i = j
        End If
    Loop
    BuildTables = nT
End Function

Private Function IsNumTok(ByVal s As String) As Boolean
    Dim i As Long, ch As String, st As Long, nd1 As Long, nd2 As Long, sep As Boolean
    s = Trim$(s)
    If Len(s) = 0 Then Exit Function
    st = 1
    If Left$(s, 1) = "-" Then st = 2
    For i = st To Len(s)
        ch = Mid$(s, i, 1)
        If ch >= "0" And ch <= "9" Then
            If sep Then nd2 = nd2 + 1 Else nd1 = nd1 + 1
        ElseIf (ch = "." Or ch = ",") And Not sep Then
            sep = True
        Else
            Exit Function
        End If
    Next i
    IsNumTok = (nd1 > 0 And (Not sep Or nd2 > 0))
End Function

Private Function IsTwoDigits(ByVal s As String) As Boolean
    IsTwoDigits = (s Like "##")
End Function

' "PL40*518 02 03" -> "PL40*518"; "01 S355J2H" -> "S355J2H"
Private Function StripRev(ByVal txt As String) As String
    Dim a As Variant, p() As String, n As Long, i As Long, s As Long, e As Long, out As String
    a = Split(Trim$(txt), " ")
    ReDim p(0 To UBound(a) + 1)
    For i = 0 To UBound(a)
        If Len(a(i)) > 0 Then p(n) = a(i): n = n + 1
    Next i
    If n < 2 Then
        StripRev = txt
        Exit Function
    End If
    s = 0: e = n - 1
    Do While e > s
        If Not IsTwoDigits(p(e)) Then Exit Do
        e = e - 1
    Loop
    Do While e > s
        If Not IsTwoDigits(p(s)) Then Exit Do
        s = s + 1
    Loop
    For i = s To e
        If Len(out) > 0 Then out = out & " "
        out = out & p(i)
    Next i
    StripRev = out
End Function

' Her çaðrýda YENÝ bir dizi (saklanan satýrlar birbirini ezmesin)
Private Function NewArr(ByVal n As Long, ByVal fill As Variant) As Variant
    Dim a() As Variant, i As Long
    ReDim a(0 To n - 1)
    For i = 0 To n - 1
        a(i) = fill
    Next i
    NewArr = a
End Function

Private Function CountFilled(ByVal r As Variant) As Long
    Dim c As Long
    For c = 0 To UBound(r)
        If Len(r(c)) > 0 Then CountFilled = CountFilled + 1
    Next c
End Function

Private Function FinishTable(cS() As String, cR() As Long, ByVal nC As Long, _
                             rY() As Double, rT() As Variant, ByVal nR As Long, ByVal pitch As Double) As Boolean
    Dim hv() As Double, hc() As Long, nh As Long, r As Long, c As Long, parts As Variant, k As Long, t As Long, q As Long
    Dim dom As Double, found As Boolean, nNum As Long, nDom As Long, domOnly As Boolean, s As String
    Dim rows() As Variant, ys() As Double, cellTxt As Variant, mg() As Variant, mgY() As Double, nM As Long
    Dim cur As Variant, prv As Variant, filled As Long, ok As Boolean, m As Long, posCol As Long, hasName As Boolean
    Dim noName As Boolean, row As Variant, hdr As Variant, u As String, rv As Variant
    ' 1) tablonun baskýn yazý boyu (üst üste basýlmýþ revizyon yazýlarý farklý boyda)
    ReDim hv(1 To 8): ReDim hc(1 To 8)
    For r = 1 To nR
        rv = rT(r)
        For c = 0 To nC - 1
            If Len(rv(c)) > 0 Then
                parts = Split(Mid$(rv(c), 2), ",")
                For k = 0 To UBound(parts)
                    t = CLng(parts(k))
                    found = False
                    For q = 1 To nh
                        If Abs(hv(q) - tkH(t)) < 0.005 Then
                            hc(q) = hc(q) + 1
                            found = True
                            Exit For
                        End If
                    Next q
                    If Not found Then
                        nh = nh + 1
                        If nh > UBound(hv) Then
                            ReDim Preserve hv(1 To 2 * nh): ReDim Preserve hc(1 To 2 * nh)
                        End If
                        hv(nh) = tkH(t): hc(nh) = 1
                    End If
                Next k
            End If
        Next c
    Next r
    If nh = 0 Then Exit Function
    q = 1
    For k = 2 To nh
        If hc(k) > hc(q) Then q = k
    Next k
    dom = hv(q)

    ' 2) hücre metinleri
    ReDim rows(1 To nR): ReDim ys(1 To nR)
    For r = 1 To nR
        cellTxt = NewArr(nC, "")
        rv = rT(r)
        For c = 0 To nC - 1
            If Len(rv(c)) > 0 Then
                parts = Split(Mid$(rv(c), 2), ",")
                nNum = 0: nDom = 0
                For k = 0 To UBound(parts)
                    t = CLng(parts(k))
                    If IsNumTok(tkS(t)) Then
                        nNum = nNum + 1
                        If Abs(tkH(t) - dom) <= 0.03 * dom Then nDom = nDom + 1
                    End If
                Next k
                domOnly = (nNum > 1 And nNum = UBound(parts) + 1 And nDom > 0 And nDom < nNum)
                s = ""
                For k = 0 To UBound(parts)
                    t = CLng(parts(k))
                    If Not domOnly Or Abs(tkH(t) - dom) <= 0.03 * dom Then
                        If Len(s) > 0 Then s = s & " "
                        s = s & tkS(t)
                    End If
                Next k
                cellTxt(c) = s
            End If
        Next c
        rows(r) = cellTxt
        ys(r) = rY(r)
    Next r

    ' 3) yetim parçalarý önceki satýra birleþtir
    ReDim mg(1 To nR): ReDim mgY(1 To nR)
    nM = 0
    For r = 1 To nR
        cur = rows(r)
        filled = CountFilled(cur)
        found = False
        If nM > 0 And filled <= 2 And filled < nC / 2 Then
            prv = mg(nM)
            ok = True
            For c = 0 To nC - 1
                If Len(cur(c)) > 0 And Len(prv(c)) > 0 Then ok = False
            Next c
            If ok And Abs(ys(r) - mgY(nM)) < 1.2 * pitch Then
                For c = 0 To nC - 1
                    If Len(cur(c)) > 0 Then prv(c) = cur(c)
                Next c
                mg(nM) = prv
                found = True
            End If
        End If
        If Not found Then
            nM = nM + 1
            mg(nM) = cur
            mgY(nM) = ys(r)
        End If
    Next r

    ' 4) kendi satýrýndan biraz yukarý basýlmýþ parça: aþaðý çek
    m = 1
    Do While m < nM
        cur = mg(m): prv = mg(m + 1)
        ok = (CountFilled(cur) <= 2)
        If ok Then
            For c = 0 To nC - 1
                If Len(cur(c)) > 0 And Len(prv(c)) > 0 Then ok = False
            Next c
        End If
        If ok And Abs(mgY(m + 1) - mgY(m)) < 1.2 * pitch Then
            For c = 0 To nC - 1
                If Len(cur(c)) > 0 Then prv(c) = cur(c)
            Next c
            mg(m + 1) = prv
            For q = m To nM - 1
                mg(q) = mg(q + 1): mgY(q) = mgY(q + 1)
            Next q
            nM = nM - 1
        Else
            m = m + 1
        End If
    Loop

    ' 5) listenin altýndaki toplam satýrlarý ("Assemblies 19 3819.4", "Bolts 85.4") parça deðildir
    posCol = -1
    For c = 0 To nC - 1
        If posCol < 0 And (cR(c) = 1 Or cR(c) = 16) Then posCol = c
        If cR(c) = 3 Or cR(c) = 7 Then hasName = True
    Next c
    If posCol >= 0 And hasName Then
        For r = nM To 1 Step -1
            cur = mg(r)
            noName = True
            For c = 0 To nC - 1
                If cR(c) = 3 Or cR(c) = 7 Then
                    If Len(Trim$(cur(c))) > 0 Then noName = False
                End If
            Next c
            If noName Then
                u = FoldPdf(cur(posCol))
                If StartsAny(u, "ITEM|ASSEMBL|BOLT|TOTAL|SUM|TOPLAM|GESAMT") Then
                    For q = r To nM - 1
                        mg(q) = mg(q + 1): mgY(q) = mgY(q + 1)
                    Next q
                    nM = nM - 1
                End If
            End If
        Next r
    End If
    If nM = 0 Then Exit Function

    ' 6) çýktý: baþlýk + satýrlar (tablolar arasýnda boþ satýr)
    If outN > 0 Then AddOutRow Empty
    hdr = NewArr(nC, "")
    For c = 0 To nC - 1
        hdr(c) = cS(c)
    Next c
    AddOutRow hdr
    For r = 1 To nM
        cur = mg(r)
        row = NewArr(nC, "")
        For c = 0 To nC - 1
            s = Trim$(cur(c))
            If cR(c) = 3 Or cR(c) = 5 Or cR(c) = 7 Then s = StripRev(s)
            If IsNumTok(s) And InStr(s, ",") = 0 Then
                row(c) = Val(s)
            Else
                row(c) = s
            End If
        Next c
        AddOutRow row
    Next r
    FinishTable = True
End Function

Private Sub AddOutRow(ByVal v As Variant)
    outN = outN + 1
    If outN > UBound(outRows) Then
        ReDim Preserve outRows(1 To 2 * outN)
    End If
    outRows(outN) = v
End Sub

Private Function RowsToArray() As Variant
    Dim out() As Variant, r As Long, c As Long, mc As Long, it As Variant
    For r = 1 To outN
        it = outRows(r)
        If IsArray(it) Then
            If UBound(it) + 1 > mc Then mc = UBound(it) + 1
        End If
    Next r
    If mc = 0 Then mc = 1
    ReDim out(1 To outN, 1 To mc)
    For r = 1 To outN
        it = outRows(r)
        If IsArray(it) Then
            For c = 0 To UBound(it)
                out(r, c + 1) = it(c)
            Next c
        End If
    Next r
    RowsToArray = out
End Function
