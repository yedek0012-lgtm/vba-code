Attribute VB_Name = "modTest"
Option Explicit

' =========================================================================
' KONTROL LÝSTESÝ (regresyon testi) - v3.5
' Bir TEST klasöründeki her liste TEK BAÞINA, SIFIRDAN iþlenir; ANGLE / PLATE / BOLTS&WASHER sonucu
' "<klasör>\_beklenen\<dosya adý>.txt" ile karþýlaþtýrýlýr.
'   Test_Beklenen_Kaydet : sonuçlarý doðru olduðunu gördüðünüz halleriyle kaydeder (bir kez)
'   Test_Calistir        : her güncellemeden sonra çalýþtýrýn; farklarý TEST_SONUC sayfasýnda gösterir
' DÝKKAT: test, bu dosyadaki ANGLE / PLATE / BOLTS&WASHER / SUM içeriðini siler.
'         Testi þablonun bir KOPYASINDA çalýþtýrýn (ör. "BOS BOM-4 - TEST.xlsm").
' Karþýlaþtýrýlan: poz, kod / ölçü, kalite, boy, adet, cývata kodu / adý / kalitesi / adedi.
' Karþýlaþtýrýlmayan: aðýrlýklar (kütüphaneye baðlý), notlar, satýr sýrasý.
' =========================================================================

Sub Test_Beklenen_Kaydet()
    Call RunRegression(True)
End Sub

Sub Test_Calistir()
    Call RunRegression(False)
End Sub

Private Sub RunRegression(ByVal saveMode As Boolean)
    Dim folder As String, files As Collection, f As Variant, fso As Object, wsR As Worksheet
    Dim oldAppend As Boolean, oldPct As Double, oldMerge As Boolean
    Dim snap As String, expPath As String, expTxt As String, status As String, detail As String
    Dim nDiff As Long, r As Long, nPass As Long, nFail As Long, nNew As Long, nSaved As Long

    folder = GetTestFolder()
    If folder = "" Then Exit Sub
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set files = ListTestFiles(folder)
    If files.Count = 0 Then
        MsgBox "Test klasöründe liste dosyasý yok:" & vbCrLf & folder, vbExclamation, "Kontrol Listesi"
        Exit Sub
    End If
    If MsgBox(files.Count & " liste tek tek SIFIRDAN iþlenecek." & vbCrLf & vbCrLf & _
              "DÝKKAT: bu dosyadaki ANGLE / PLATE / BOLTS&WASHER / SUM içeriði silinecek." & vbCrLf & _
              "Testi þablonun bir KOPYASINDA çalýþtýrýn." & vbCrLf & vbCrLf & _
              IIf(saveMode, "Sonuçlar BEKLENEN olarak kaydedilecek (eskilerin üstüne yazýlýr).", "Sonuçlar beklenenlerle karþýlaþtýrýlacak.") & _
              vbCrLf & vbCrLf & "Devam edilsin mi?", vbYesNo + vbExclamation, "Kontrol Listesi") = vbNo Then Exit Sub

    If Not fso.FolderExists(folder & "\_beklenen") Then fso.CreateFolder folder & "\_beklenen"

    Set wsR = GetResultSheet()
    wsR.Range("A3:D" & wsR.Rows.Count).ClearContents
    wsR.Range("A3:D" & wsR.Rows.Count).Interior.ColorIndex = xlColorIndexNone
    wsR.Range("A3:D3").Value = Array("DOSYA", "SONUÇ", "FARK", "AYRINTI (EKSÝK = beklenip gelmeyen, FAZLA = beklenmeyen)")
    wsR.Range("A3:D3").Font.Bold = True
    r = 4

    oldAppend = usrAppendMode: oldPct = usrExtraPct: oldMerge = usrIsMerge
    usrAppendMode = False: usrExtraPct = 0: usrIsMerge = False
    runQuiet = True
    On Error GoTo Fail

    For Each f In files
        Application.StatusBar = "Kontrol listesi: " & fso.GetFileName(f)
        runPresetFiles = Array(CStr(f))
        Call Evrensel_BOM_Cevirici_Core
        snap = SnapshotOutputs()
        expPath = folder & "\_beklenen\" & fso.GetFileName(f) & ".txt"
        detail = ""
        nDiff = 0
        If saveMode Then
            Call WriteTextFile(expPath, snap)
            status = "KAYDEDÝLDÝ"
            nSaved = nSaved + 1
        ElseIf Not fso.FileExists(expPath) Then
            status = "BEKLENEN YOK"
            detail = "Önce Test_Beklenen_Kaydet çalýþtýrýn."
            nNew = nNew + 1
        Else
            expTxt = ReadTextFile(expPath)
            nDiff = CompareSnapshots(expTxt, snap, detail)
            If nDiff = 0 Then
                status = "GEÇTÝ"
                nPass = nPass + 1
            Else
                status = "FARK"
                nFail = nFail + 1
            End If
        End If
        wsR.Cells(r, 1).Value = fso.GetFileName(f)
        wsR.Cells(r, 2).Value = status
        wsR.Cells(r, 3).Value = nDiff
        wsR.Cells(r, 4).Value = detail
        Select Case status
            Case "GEÇTÝ", "KAYDEDÝLDÝ": wsR.Cells(r, 2).Interior.Color = RGB(198, 239, 206)
            Case "FARK": wsR.Cells(r, 2).Interior.Color = RGB(255, 199, 206)
            Case Else: wsR.Cells(r, 2).Interior.Color = RGB(255, 235, 156)
        End Select
        r = r + 1
    Next f

Done:
    runQuiet = False
    runPresetFiles = Empty
    usrAppendMode = oldAppend: usrExtraPct = oldPct: usrIsMerge = oldMerge
    Application.StatusBar = False
    wsR.Activate
    wsR.Range("D:D").WrapText = True
    If saveMode Then
        MsgBox nSaved & " listenin sonucu BEKLENEN olarak kaydedildi:" & vbCrLf & folder & "\_beklenen", vbInformation, "Kontrol Listesi"
    Else
        MsgBox nPass & " liste GEÇTÝ, " & nFail & " listede FARK var, " & nNew & " listenin beklenen sonucu yok." & vbCrLf & _
               "Ayrýntý: TEST_SONUC sayfasý.", IIf(nFail > 0, vbExclamation, vbInformation), "Kontrol Listesi"
    End If
    Exit Sub
Fail:
    wsR.Cells(r, 1).Value = "HATA"
    wsR.Cells(r, 4).Value = Err.Description
    Resume Done
End Sub

' ---- test klasörü: TEST_SONUC!B1'de saklanýr; yoksa / bulunamazsa sorulur ----
Private Function GetTestFolder() As String
    Dim ws As Worksheet, p As String, fd As FileDialog
    Set ws = GetResultSheet()
    p = Trim$(ws.Range("B1").Text)
    If p <> "" Then
        If CreateObject("Scripting.FileSystemObject").FolderExists(p) Then
            If MsgBox("Test klasörü:" & vbCrLf & p & vbCrLf & vbCrLf & "Bu klasör kullanýlsýn mý? (Hayýr = baþka klasör seç)", _
                      vbYesNo + vbQuestion, "Kontrol Listesi") = vbYes Then
                GetTestFolder = p
                Exit Function
            End If
        End If
    End If
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    fd.Title = "Test listelerinin bulunduðu klasörü seçin"
    If fd.Show <> -1 Then Exit Function
    p = fd.SelectedItems(1)
    ws.Range("B1").Value = p
    GetTestFolder = p
End Function

Private Function GetResultSheet() As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("TEST_SONUC")
    On Error GoTo 0
    If ws Is Nothing Then
        Call UnprotectForMacro
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = "TEST_SONUC"
        ws.Range("A1").Value = "TEST KLASÖRÜ"
        ws.Range("A1").Font.Bold = True
        ws.Columns("A").ColumnWidth = 45
        ws.Columns("B").ColumnWidth = 16
        ws.Columns("C").ColumnWidth = 8
        ws.Columns("D").ColumnWidth = 110
        Call ProtectAfterMacro
    End If
    Set GetResultSheet = ws
End Function

Private Function ListTestFiles(ByVal folder As String) As Collection
    Dim fso As Object, fl As Object, ext As String, c As New Collection
    Set fso = CreateObject("Scripting.FileSystemObject")
    For Each fl In fso.GetFolder(folder).Files
        ext = LCase$(fso.GetExtensionName(fl.Name))
        If Left$(fl.Name, 2) <> "~$" Then
            Select Case ext
                Case "xsr", "txt", "xls", "xlsx", "xlsm", "xlsb", "csv", "pdf"
                    c.Add fl.Path
            End Select
        End If
    Next fl
    Set ListTestFiles = c
End Function

' ---- çýktý özeti: her satýr bir kayýt, sýralý ----
Private Function CellKey(ByVal v As Variant) As String
    If IsError(v) Or IsEmpty(v) Then Exit Function
    Select Case VarType(v)
        Case vbDouble, vbSingle, vbLong, vbInteger, vbCurrency, vbDecimal
            CellKey = NumToText(CDbl(v))
        Case Else
            CellKey = Trim$(CStr(v))
    End Select
End Function

Private Function SnapshotOutputs() As String
    Dim lines() As String, n As Long, ws As Worksheet, r As Long, lastR As Long, ln As String
    ReDim lines(1 To 64)
    Set ws = ThisWorkbook.Sheets("ANGLE")
    lastR = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
    For r = 5 To lastR
        If CellKey(ws.Cells(r, 2).Value) <> "" Then
            ln = "ANGLE | POZ " & CellKey(ws.Cells(r, 2).Value) & " | KOD " & CellKey(ws.Cells(r, 3).Value) & _
                 " | " & CellKey(ws.Cells(r, 5).Value) & " | L " & CellKey(ws.Cells(r, 6).Value) & " | ADET " & CellKey(ws.Cells(r, 9).Value)
            GoSub AddLine
        End If
    Next r
    Set ws = ThisWorkbook.Sheets("PLATE")
    lastR = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
    For r = 5 To lastR
        If CellKey(ws.Cells(r, 2).Value) <> "" Then
            ln = "PLATE | POZ " & CellKey(ws.Cells(r, 2).Value) & " | " & CellKey(ws.Cells(r, 3).Value) & "x" & CellKey(ws.Cells(r, 4).Value) & _
                 " | " & CellKey(ws.Cells(r, 5).Value) & " | L " & CellKey(ws.Cells(r, 6).Value) & " | ADET " & CellKey(ws.Cells(r, 8).Value)
            GoSub AddLine
        End If
    Next r
    Set ws = ThisWorkbook.Sheets("BOLTS&WASHER")
    lastR = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
    For r = 5 To lastR
        If CellKey(ws.Cells(r, 2).Value) <> "" Then
            ln = "BOLT | " & CellKey(ws.Cells(r, 2).Value) & " | " & CellKey(ws.Cells(r, 3).Value) & " | " & CellKey(ws.Cells(r, 4).Value) & _
                 " | ADET " & CellKey(ws.Cells(r, 6).Value)
            GoSub AddLine
        End If
    Next r
    If n = 0 Then
        SnapshotOutputs = ""
        Exit Function
    End If
    ReDim Preserve lines(1 To n)
    Call SortStrings(lines, 1, n)
    SnapshotOutputs = Join(lines, vbLf)
    Exit Function
AddLine:
    n = n + 1
    If n > UBound(lines) Then ReDim Preserve lines(1 To 2 * n)
    lines(n) = ln
    Return
End Function

Private Sub SortStrings(a() As String, ByVal lo As Long, ByVal hi As Long)
    Dim i As Long, j As Long, pv As String, t As String
    If lo >= hi Then Exit Sub
    i = lo: j = hi
    pv = a((lo + hi) \ 2)
    Do While i <= j
        Do While StrComp(a(i), pv, vbBinaryCompare) < 0
            i = i + 1
        Loop
        Do While StrComp(a(j), pv, vbBinaryCompare) > 0
            j = j - 1
        Loop
        If i <= j Then
            t = a(i): a(i) = a(j): a(j) = t
            i = i + 1: j = j - 1
        End If
    Loop
    If lo < j Then SortStrings a, lo, j
    If i < hi Then SortStrings a, i, hi
End Sub

' Farklý satýr sayýsýný döner; ayrýntý en fazla 30 satýr
Private Function CompareSnapshots(ByVal expTxt As String, ByVal actTxt As String, ByRef detail As String) As Long
    Dim dE As Object, dA As Object, k As Variant, ln As Variant, n As Long, shown As Long
    Set dE = CreateObject("Scripting.Dictionary")
    Set dA = CreateObject("Scripting.Dictionary")
    expTxt = Replace(Replace(expTxt, vbCrLf, vbLf), vbCr, vbLf)
    For Each ln In Split(expTxt, vbLf)
        If Trim$(ln) <> "" Then dE(ln) = dE(ln) + 1
    Next ln
    For Each ln In Split(actTxt, vbLf)
        If Trim$(ln) <> "" Then dA(ln) = dA(ln) + 1
    Next ln
    For Each k In dE.keys
        If dE(k) > dA(k) Then
            n = n + 1
            If shown < 30 Then detail = detail & "EKSÝK: " & k & IIf(dE(k) - dA(k) > 1, "  (x" & (dE(k) - dA(k)) & ")", "") & vbLf: shown = shown + 1
        End If
    Next k
    For Each k In dA.keys
        If dA(k) > dE(k) Then
            n = n + 1
            If shown < 30 Then detail = detail & "FAZLA: " & k & IIf(dA(k) - dE(k) > 1, "  (x" & (dA(k) - dE(k)) & ")", "") & vbLf: shown = shown + 1
        End If
    Next k
    If n > shown Then detail = detail & "... ve " & (n - shown) & " fark daha"
    CompareSnapshots = n
End Function

' Türkçe karakterler bozulmasýn diye Unicode metin dosyasý
Private Sub WriteTextFile(ByVal path As String, ByVal txt As String)
    Dim ts As Object
    Set ts = CreateObject("Scripting.FileSystemObject").OpenTextFile(path, 2, True, -1)
    ts.Write txt
    ts.Close
End Sub

Private Function ReadTextFile(ByVal path As String) As String
    Dim ts As Object
    Set ts = CreateObject("Scripting.FileSystemObject").OpenTextFile(path, 1, False, -1)
    If Not ts.AtEndOfStream Then ReadTextFile = ts.ReadAll
    ts.Close
End Function
