Attribute VB_Name = "modAudit"
Option Explicit
Option Private Module

' =========================================================================
' RAPOR: AUDIT, OGRENME, GECMIS, yedek, REV/DATE
' =========================================================================

Public Sub InitAuditSheet()
    On Error GoTo AuditFail
    Set auditWs = Nothing
    Set auditFiles = CreateObject("Scripting.Dictionary")
    auditFiles.CompareMode = vbTextCompare

    On Error Resume Next
    Set auditWs = ThisWorkbook.Sheets("AUDIT")
    On Error GoTo AuditFail

    If auditWs Is Nothing Then
        Set auditWs = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        auditWs.Name = "AUDIT"
    Else
        auditWs.Cells.Clear
End If

    auditNextRow = 2
    auditProcessed = 0
    auditExact = 0
    auditWarning = 0
    auditError = 0
    auditUnmatched = 0
    auditSkipped = 0

    auditWs.Cells(1, 1).Value = "TARÝH"
    auditWs.Cells(1, 2).Value = "DOSYA"
    auditWs.Cells(1, 3).Value = "KAYNAK"
    auditWs.Cells(1, 4).Value = "SATIR"
    auditWs.Cells(1, 5).Value = "BÖLÜM"
    auditWs.Cells(1, 6).Value = "HAM VERÝ"
    auditWs.Cells(1, 7).Value = "TÝP"
    auditWs.Cells(1, 8).Value = "DURUM"
    auditWs.Cells(1, 9).Value = "GÜVEN"
    auditWs.Cells(1, 10).Value = "AYRIÞTIRILAN"
    auditWs.Cells(1, 11).Value = "AÇIKLAMA"
    auditWs.Range("A1:K1").Font.Bold = True
    auditWs.Range("A1:K1").Interior.Color = RGB(0, 51, 102)
    auditWs.Range("A1:K1").Font.Color = RGB(255, 255, 255)
    auditWs.Range("A1:K1").AutoFilter
    Exit Sub
AuditFail:
    Set auditWs = Nothing
End Sub

Public Sub AuditRecord(ByVal fileName As String, ByVal sourceType As String, ByVal sourceRow As Long, ByVal sectionName As String, ByVal rawText As String, ByVal itemType As String, ByVal status As String, ByVal confidence As Long, ByVal parsedText As String, ByVal issueText As String)
    On Error GoTo AuditExit
    auditProcessed = auditProcessed + 1
    Select Case UCase$(status)
        Case "EXACT": auditExact = auditExact + 1
        Case "WARNING": auditWarning = auditWarning + 1
        Case "ERROR": auditError = auditError + 1
        Case "UNMATCHED": auditUnmatched = auditUnmatched + 1
    End Select

    If auditWs Is Nothing Then Exit Sub
    auditWs.Cells(auditNextRow, 1).Value = Now
    auditWs.Cells(auditNextRow, 2).Value = fileName
    auditWs.Cells(auditNextRow, 3).Value = sourceType
    auditWs.Cells(auditNextRow, 4).Value = sourceRow
    auditWs.Cells(auditNextRow, 5).Value = sectionName
    auditWs.Cells(auditNextRow, 6).Value = Left$(rawText, 500)
    auditWs.Cells(auditNextRow, 7).Value = itemType
    auditWs.Cells(auditNextRow, 8).Value = status
    auditWs.Cells(auditNextRow, 9).Value = confidence
    auditWs.Cells(auditNextRow, 10).Value = Left$(parsedText, 500)
    auditWs.Cells(auditNextRow, 11).Value = Left$(issueText, 500)

    Select Case UCase$(status)
        Case "EXACT"
            auditWs.Cells(auditNextRow, 8).Interior.Color = RGB(0, 153, 76)
            auditWs.Cells(auditNextRow, 8).Font.Color = RGB(255, 255, 255)
        Case "WARNING"
            auditWs.Cells(auditNextRow, 8).Interior.Color = RGB(255, 192, 0)
        Case Else
            auditWs.Cells(auditNextRow, 8).Interior.Color = RGB(192, 0, 0)
            auditWs.Cells(auditNextRow, 8).Font.Color = RGB(255, 255, 255)
    End Select
    auditNextRow = auditNextRow + 1

    ' Köprü: SATIR hücresine týklayýnca kaynaða git
    On Error Resume Next
    Dim linkRow As Long, subAddr As String
    linkRow = auditNextRow - 1
    If sourceRow > 0 Then
        If UCase$(sourceType) = "LIBRARY" Then
            subAddr = "'" & Replace(sectionName, "'", "''") & "'!A" & sourceRow
            auditWs.Hyperlinks.Add Anchor:=auditWs.Cells(linkRow, 4), Address:="", SubAddress:=subAddr, TextToDisplay:=CStr(sourceRow)
        ElseIf auditCurrentPath <> "" Then
            subAddr = ""
            If auditCurrentSheet <> "" Then subAddr = "'" & Replace(auditCurrentSheet, "'", "''") & "'!A" & sourceRow
            auditWs.Hyperlinks.Add Anchor:=auditWs.Cells(linkRow, 4), Address:=auditCurrentPath, SubAddress:=subAddr, TextToDisplay:=CStr(sourceRow)
        End If
    End If
    On Error GoTo 0
AuditExit:
End Sub

Public Sub FinalizeAuditSheet()
    On Error GoTo AuditFail
    If auditWs Is Nothing Then Exit Sub
    With auditWs
        .Rows(1).AutoFilter
        .Columns("A:K").EntireColumn.AutoFit
        .Columns("F:F").ColumnWidth = 42
        .Columns("J:K").ColumnWidth = 42
        .Columns("A:A").NumberFormat = "dd.mm.yyyy hh:mm"

        .Cells(1, 13).Value = "ÖZET"
        .Cells(2, 13).Value = "Toplam kayýt"
        .Cells(2, 14).Value = auditProcessed
        .Cells(3, 13).Value = "Exact"
        .Cells(3, 14).Value = auditExact
        .Cells(4, 13).Value = "Warning"
        .Cells(4, 14).Value = auditWarning
        .Cells(5, 13).Value = "Error"
        .Cells(5, 14).Value = auditError
        .Cells(6, 13).Value = "Unmatched"
        .Cells(6, 14).Value = auditUnmatched
        .Cells(7, 13).Value = "Skipped"
        .Cells(7, 14).Value = auditSkipped
        .Range("M1:N7").Font.Bold = True

        .Cells(9, 13).Value = "ÇALIÞMA BÝLGÝSÝ"
        .Cells(9, 13).Font.Bold = True
        .Cells(10, 13).Value = "Kullanýcý": .Cells(10, 14).Value = Application.UserName & " (" & Environ$("USERNAME") & ")"
        .Cells(11, 13).Value = "Bilgisayar": .Cells(11, 14).Value = Environ$("COMPUTERNAME")
        .Cells(12, 13).Value = "Baþlangýç": .Cells(12, 14).Value = Format$(runStart, "dd.mm.yyyy hh:nn:ss")
        .Cells(13, 13).Value = "Bitiþ": .Cells(13, 14).Value = Format$(Now, "dd.mm.yyyy hh:nn:ss")
        .Cells(14, 13).Value = "Mod": .Cells(14, 14).Value = runMode
        .Cells(15, 13).Value = "REV": .Cells(15, 14).Value = runRev
        .Cells(16, 13).Value = "Yedek": .Cells(16, 14).Value = IIf(runBackupPath = "", "-", runBackupPath)
        .Cells(17, 13).Value = "Kod sürümü": .Cells(17, 14).Value = CODE_VERSION
        .Range("M10:M17").Font.Bold = True
        .Columns("M:N").AutoFit

        If auditError > 0 Or auditUnmatched > 0 Then
            .Tab.Color = RGB(192, 0, 0)
        ElseIf auditWarning > 0 Then
            .Tab.Color = RGB(255, 192, 0)
Else
            .Tab.Color = RGB(0, 153, 76)
End If
    End With
    Exit Sub
AuditFail:
End Sub

' =========================================================================
' ÖÐRENEN KÜTÜPHANE - TOPLU LÝSTE (v3.5: 4 veri giriþi)
' Tanýnmayan parçalar tek bir sayfada gösterilir. Her parça için kütüphanenin
' 4 sütunu girilir: MALZEME KODU, MALZEME CÝNSÝ, kg/m (birim aðýrlýk), yüzey alan (1000 adet aðýrlýðý).
' =========================================================================
Public Sub WriteUnknownsSheet(ByVal dictUnk As Object)
    Dim ws As Worksheet, k As Variant, r As Long, out() As Variant, btn As Object
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("OGRENME")
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = "OGRENME"
    End If

    On Error Resume Next
    ws.Buttons.Delete
    On Error GoTo 0
    ws.Cells.Clear
    ws.Cells.Validation.Delete

    ws.Range("A1:H1").Value = Array("TANINMAYAN PARÇA", "KAYNAK (Dosya / Poz)", "KÜTÜPHANE", _
                                    "MALZEME KODU", "MALZEME CÝNSÝ", _
                                    "kg/m" & vbLf & "(cývata: BÝRÝM AÐIRLIK)", _
                                    "YÜZEY ALAN" & vbLf & "(cývata: 1000 ADET AÐIRLIÐI)", "DURUM")
    With ws.Range("A1:H1")
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(0, 51, 102)
        .WrapText = True
        .VerticalAlignment = xlCenter
        .HorizontalAlignment = xlCenter
    End With
    ws.Rows(1).RowHeight = 32

    ReDim out(1 To dictUnk.Count, 1 To 5)
    r = 0
    For Each k In dictUnk.keys
        r = r + 1
        out(r, 1) = CStr(k)
        out(r, 2) = CStr(dictUnk(k))
        out(r, 3) = GuessUnknownLibrary(CStr(k))
        out(r, 4) = ""
        out(r, 5) = CStr(k)          ' MALZEME CÝNSÝ: tanýnmayan ad ile önceden doldurulur, deðiþtirilebilir
    Next k
    ws.Range("D2").Resize(r, 1).NumberFormat = "@"       ' kodun baþtaki sýfýrý kaybolmasýn
    ws.Range("A2").Resize(r, 5).Value = out
    ws.Range("F2").Resize(r, 2).NumberFormat = "0.000#"
    ws.Range("C2").Resize(r, 5).Interior.Color = RGB(255, 242, 204)
    ws.Range("A1").Resize(r + 1, 8).Borders.LineStyle = 1

    On Error Resume Next
    With ws.Range("C2").Resize(r, 1).Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="PROFÝL,CIVATA"
        .IgnoreBlank = True
    End With
    On Error GoTo 0

    ws.Columns("A").ColumnWidth = 30
    ws.Columns("B").ColumnWidth = 45
    ws.Columns("C").ColumnWidth = 12
    ws.Columns("D").ColumnWidth = 16
    ws.Columns("E").ColumnWidth = 24
    ws.Columns("F").ColumnWidth = 14
    ws.Columns("G").ColumnWidth = 16
    ws.Columns("H").ColumnWidth = 34

    ws.Range("J1").Value = "NASIL KULLANILIR"
    ws.Range("J1").Font.Bold = True
    ws.Range("J2").Value = "1) Sarý alandaki 4 veriyi girin: MALZEME KODU, MALZEME CÝNSÝ, kg/m, YÜZEY ALAN."
    ws.Range("J3").Value = "   PROFÝL -> L-U-I-O-Y-LIBRARY (A kod, B cins, D kg/m, E yüzey alan)."
    ws.Range("J4").Value = "   CIVATA -> BOLT-LIBRARY (A kod, B cins, C birim aðýrlýk, D 1000 adet aðýrlýðý;"
    ws.Range("J5").Value = "   ikisinden biri yeterli, diðeri hesaplanýr)."
    ws.Range("J6").Value = "2) KÜTÜPHANE sütununu kontrol edin (PROFÝL / CIVATA)."
    ws.Range("J7").Value = "3) Tanýmlamak istemediklerinizi boþ býrakýn (kütüphaneye eklenmez)."
    ws.Range("J8").Value = "4) Eksik verili satýrlar aktarýlmaz, DURUM sütununda eksik alan yazar."
    ws.Range("J9").Value = "5) 'KÜTÜPHANEYE AKTAR' düðmesine basýn, sonra listeyi yeniden iþleyin."

    On Error Resume Next
    Set btn = ws.Buttons.Add(ws.Range("J11").Left, ws.Range("J11").Top, 180, 30)
    btn.Caption = "KÜTÜPHANEYE AKTAR"
    btn.OnAction = "Kutuphaneye_Aktar"
    On Error GoTo 0
End Sub

' Tanýnmayan parçanýn hangi kütüphaneye gideceðini tahmin eder (kullanýcý deðiþtirebilir)
Private Function GuessUnknownLibrary(ByVal nm As String) As String
    nm = UCase$(nm)
    If RxTest(nm, "^M\d|BOLT|SCHRAUB|CIVATA|CÝVATA|SCREW|\bNUT\b|MUTTER|SOMUN|WASH|SCHEIBE|RONDELA|\bPUL\b|\bSCH\b|\bMU\b|\bFDRG\b") Then
        GuessUnknownLibrary = "CIVATA"
    Else
        GuessUnknownLibrary = "PROFÝL"
    End If
End Function

' =========================================================================
' v2.3 - TEMÝZLE düðmesi için rapor sayfalarý
' =========================================================================
Public Sub ClearReportSheets()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = Nothing
    Set ws = ThisWorkbook.Sheets("AUDIT")
    If Not ws Is Nothing Then
        ws.Cells.Clear
        ws.Tab.ColorIndex = xlColorIndexNone
    End If
    Set ws = Nothing
    Set ws = ThisWorkbook.Sheets("FEEDBACK")
    If Not ws Is Nothing Then
        ws.Cells.Clear
        ws.Tab.ColorIndex = xlColorIndexNone
    End If
    Set ws = Nothing
    Set ws = ThisWorkbook.Sheets("OGRENME")
    If Not ws Is Nothing Then
        ws.Buttons.Delete
        ws.Cells.Validation.Delete
        ws.Cells.Clear
    End If
    ' Önceki çalýþtýrmadan kalan SUM mutabakat sütunlarý
    Set ws = Nothing
    Set ws = ThisWorkbook.Sheets("SUM")
    If Not ws Is Nothing Then
        ws.Range("I3:K3").ClearContents
        ws.Range("K5:K5000").FormatConditions.Delete
    End If
    On Error GoTo 0
End Sub

' =========================================================================
' v2.3 - REV / DATE ve GEÇMÝÞ
' =========================================================================
Public Function UpdateRevDate(ByVal isAppend As Boolean) As String
    Dim ws As Worksheet, revNo As Long, t As String
    On Error GoTo Fail
    Set ws = ThisWorkbook.Sheets("SUM")
    t = RxFirst(ws.Range("G1").Text, "\d+")
    revNo = CLng(Val(t))
    If isAppend And prmRevAuto Then revNo = revNo + 1
    ws.Range("G1").Value = "REV:       " & revNo
    ws.Range("G2").Value = "DATE: " & Format$(Day(Date), "00") & "." & Format$(Month(Date), "00") & "." & Year(Date)
    UpdateRevDate = CStr(revNo)
    Exit Function
Fail:
    UpdateRevDate = "?"
End Function

' Kalýcý geçmiþ: her çalýþtýrma bir satýr (Temizle düðmesi bu sayfaya dokunmaz)
Public Sub AppendRunHistory(ByRef files() As String, ByVal fileCount As Long, ByVal posCount As Long)
    Dim ws As Worksheet, r As Long, i As Long, fl As String, fso As Object
    On Error GoTo Done
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("GECMIS")
    On Error GoTo Done
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = "GECMIS"
        ws.Range("A1:J1").Value = Array("TARÝH", "KULLANICI", "BÝLGÝSAYAR", "MOD", "REV", "DOSYA SAYISI", "POZ SAYISI", "KOD SÜRÜMÜ", "YEDEK", "DOSYALAR")
        With ws.Range("A1:J1")
            .Font.Bold = True
            .Interior.Color = RGB(0, 51, 102)
            .Font.Color = RGB(255, 255, 255)
        End With
        ws.Columns("A").ColumnWidth = 18
        ws.Columns("B:H").ColumnWidth = 14
        ws.Columns("I").ColumnWidth = 40
        ws.Columns("J").ColumnWidth = 100
    End If
    Set fso = CreateObject("Scripting.FileSystemObject")
    For i = 1 To fileCount
        If fl <> "" Then fl = fl & " ; "
        fl = fl & fso.GetFileName(files(i))
    Next i
    r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    ws.Cells(r, 1).Value = Now
    ws.Cells(r, 1).NumberFormat = "dd.mm.yyyy hh:mm"
    ws.Cells(r, 2).Value = Application.UserName
    ws.Cells(r, 3).Value = Environ$("COMPUTERNAME")
    ws.Cells(r, 4).Value = runMode
    ws.Cells(r, 5).Value = runRev
    ws.Cells(r, 6).Value = fileCount
    ws.Cells(r, 7).Value = posCount
    ws.Cells(r, 8).Value = CODE_VERSION
    ws.Cells(r, 9).Value = IIf(runBackupPath = "", "-", runBackupPath)
    ws.Cells(r, 10).Value = Left$(fl, 32000)
Done:
End Sub

' =========================================================================
' v2.3 - OTOMATÝK YEDEK
' =========================================================================
Public Function MakeBackup() As String
    Dim fso As Object, folder As String, baseName As String, ext As String, target As String
    On Error GoTo Fail
    Set fso = CreateObject("Scripting.FileSystemObject")

    folder = prmBackupFolder
    If folder = "" Then
        If ThisWorkbook.Path <> "" And LCase$(Left$(ThisWorkbook.Path, 4)) <> "http" Then
            folder = ThisWorkbook.Path & "\YEDEK"
        Else
            ' OneDrive/SharePoint (https://...) veya kaydedilmemiþ dosya
            folder = Environ$("USERPROFILE") & "\Documents\BOM_YEDEK"
        End If
    End If
    If Not EnsureFolder(fso, folder) Then GoTo Fail

    baseName = fso.GetBaseName(ThisWorkbook.Name)
    ext = fso.GetExtensionName(ThisWorkbook.Name)
    If ext = "" Then ext = "xlsm"
    target = folder & "\" & baseName & "_" & Format$(Now, "yyyymmdd_hhnnss") & "." & ext
    ThisWorkbook.SaveCopyAs target
    Call PruneBackups(fso, folder, baseName & "_", prmBackupKeep)
    MakeBackup = target
    Exit Function
Fail:
    MakeBackup = ""
    AuditRecord "", "SYSTEM", 0, "YEDEK", folder, "BACKUP", "WARNING", 50, "", "Yedek alýnamadý: " & Err.Description
End Function

Public Function EnsureFolder(ByVal fso As Object, ByVal folder As String) As Boolean
    On Error GoTo Fail
    If fso.FolderExists(folder) Then EnsureFolder = True: Exit Function
    If Not fso.FolderExists(fso.GetParentFolderName(folder)) Then
        If Not EnsureFolder(fso, fso.GetParentFolderName(folder)) Then Exit Function
    End If
    fso.CreateFolder folder
    EnsureFolder = True
    Exit Function
Fail:
    EnsureFolder = False
End Function

' En yeni 'keep' yedeði býrakýr, eskileri siler (sadece bu dosyanýn yedekleri)
Public Sub PruneBackups(ByVal fso As Object, ByVal folder As String, ByVal prefix As String, ByVal keep As Long)
    Dim f As Object, names() As String, dates() As Date, n As Long, i As Long, j As Long
    Dim tN As String, tD As Date
    On Error GoTo Done
    ReDim names(1 To 1000)
    ReDim dates(1 To 1000)
    For Each f In fso.GetFolder(folder).Files
        If LCase$(Left$(f.Name, Len(prefix))) = LCase$(prefix) Then
            n = n + 1
            If n > 1000 Then Exit For
            names(n) = f.Path
            dates(n) = f.DateLastModified
        End If
    Next f
    If n <= keep Then GoTo Done
    ' yeniden eskiye sýrala
    For i = 1 To n - 1
        For j = i + 1 To n
            If dates(j) > dates(i) Then
                tD = dates(i): dates(i) = dates(j): dates(j) = tD
                tN = names(i): names(i) = names(j): names(j) = tN
            End If
        Next j
    Next i
    For i = keep + 1 To n
        fso.DeleteFile names(i), True
    Next i
Done:
End Sub
