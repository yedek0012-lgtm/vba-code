Attribute VB_Name = "modProtect"
Option Explicit
Option Private Module

' =========================================================================
' KORUMA: sayfa ve çalýþma kitabý korumasý
' =========================================================================

' =========================================================================
' SAYFA KORUMASI
' - SUM / ANGLE / PLATE / BOLTS&WASHER (ve _P2.. sayfalarý): FORMÜL hücreleri kilitli,
'   deðer hücreleri (QTY, manuel adetler, notlar) düzenlenebilir. Filtre ve biçim serbest.
' - GECMIS: tamamen kilitli (kayýt deðiþtirilemez).
' - Çalýþma kitabý yapýsý kilitli: sayfa silme / yeniden adlandýrma / "Taþý veya Kopyala" kapalý.
' Makrolar çalýþýrken korumayý kendileri açýp kapatýr.
'
' PAROLA KODDA DEÐÝL (v3.5): çalýþma kitabýndaki gizli "BOM_KORUMA" isminde saklanýr.
' Ýlk kurulumda VBA penceresinden Yonetici_ParolaDegistir çalýþtýrýn (F5) ve VBA projesini kilitleyin.
' LEGACY_PWD yalnýzca geçiþ içindir: eski dosyalar bununla korunmuþtu; makro önce geçerli parolayý,
' olmazsa bunu dener ve ilk çalýþtýrmada sayfalarý geçerli parolayla yeniden korur.
' =========================================================================
Private Const LEGACY_PWD As String = "Sara#Bom2026!"
Private Const PWD_NAME As String = "BOM_KORUMA"
Private Const QT As String = """"

' Geçerli koruma parolasý: gizli isimden; tanýmlý deðilse eski parola
Private Function Pwd() As String
    Dim nm As Name, t As String
    On Error Resume Next
    Set nm = ThisWorkbook.Names(PWD_NAME)
    On Error GoTo 0
    If nm Is Nothing Then
        Pwd = LEGACY_PWD
        Exit Function
    End If
    t = nm.RefersTo                                   ' ="parola"
    If Left$(t, 2) = "=" & QT And Right$(t, 1) = QT Then t = Mid$(t, 3, Len(t) - 3)
    t = Replace(t, QT & QT, QT)
    If t = "" Then t = LEGACY_PWD
    Pwd = t
End Function

' Önce geçerli, olmazsa eski parola ile sayfa korumasýný açar
Private Sub UnprotectSheet(ByVal ws As Worksheet)
    On Error Resume Next
    ws.Unprotect Pwd()
    If ws.ProtectContents Then ws.Unprotect LEGACY_PWD
End Sub

Public Function IsOutputSheet(ByVal nm As String) As Boolean
    nm = UCase$(nm)
    IsOutputSheet = (nm = "SUM" Or nm = "ANGLE" Or nm = "PLATE" Or nm = "BOLTS&WASHER" Or _
                     nm Like "SUM_P#*" Or nm Like "ANGLE_P#*" Or nm Like "PLATE_P#*" Or nm Like "BOLTS_P#*")
End Function

Public Sub UnprotectForMacro()
    Dim ws As Worksheet
    On Error Resume Next
    ThisWorkbook.Unprotect Pwd()
    If ThisWorkbook.ProtectStructure Then ThisWorkbook.Unprotect LEGACY_PWD
    For Each ws In ThisWorkbook.Worksheets
        If IsOutputSheet(ws.Name) Or UCase$(ws.Name) = "GECMIS" Then Call UnprotectSheet(ws)
    Next ws
    On Error GoTo 0
End Sub

Public Sub ProtectAfterMacro()
    Dim ws As Worksheet, fr As Range, p As String
    On Error Resume Next
    p = Pwd()
    For Each ws In ThisWorkbook.Worksheets
        If IsOutputSheet(ws.Name) Then
            Call UnprotectSheet(ws)
            ws.Cells.Locked = False
            Set fr = Nothing
            Set fr = ws.UsedRange.SpecialCells(xlCellTypeFormulas)
            If Not fr Is Nothing Then fr.Locked = True
            ws.Protect Password:=p, DrawingObjects:=True, Contents:=True, Scenarios:=True, _
                       UserInterfaceOnly:=True, AllowFormattingCells:=True, AllowFormattingColumns:=True, _
                       AllowFormattingRows:=True, AllowFiltering:=True
            ws.EnableSelection = xlNoRestrictions
        ElseIf UCase$(ws.Name) = "GECMIS" Then
            Call UnprotectSheet(ws)
            ws.Cells.Locked = True
            ws.Protect Password:=p, UserInterfaceOnly:=True, AllowFiltering:=True
        End If
    Next ws
    If ThisWorkbook.ProtectStructure Then
        ThisWorkbook.Unprotect p
        If ThisWorkbook.ProtectStructure Then ThisWorkbook.Unprotect LEGACY_PWD
    End If
    ThisWorkbook.Protect Password:=p, Structure:=True
    On Error GoTo 0
End Sub

' Sadece yönetici (bakým için): VBA penceresinden çalýþtýrýn, parola sorar.
Public Sub Yonetici_KorumayiKaldir()
    If InputBox("Yönetici parolasý:", "Korumayý kaldýr") <> Pwd() Then
        MsgBox "Parola yanlýþ.", vbCritical
        Exit Sub
    End If
    Call UnprotectForMacro
    MsgBox "Koruma kaldýrýldý. Tekrar korumak için ProtectAfterMacro çalýþtýrýn.", vbInformation
End Sub

' Sadece yönetici: VBA penceresinden çalýþtýrýn (F5). Koruma parolasýný deðiþtirir.
' Parola kodda deðil, bu dosyanýn gizli "BOM_KORUMA" isminde saklanýr; çýktý sayfalarý yeni parolayla korunur.
Public Sub Yonetici_ParolaDegistir()
    Dim cur As String, p1 As String, p2 As String
    cur = InputBox("Mevcut koruma parolasý:", "Parola deðiþtir")
    If cur = "" Then Exit Sub
    If cur <> Pwd() And cur <> LEGACY_PWD Then
        MsgBox "Parola yanlýþ.", vbCritical
        Exit Sub
    End If
    p1 = InputBox("Yeni parola (en az 8 karakter):", "Parola deðiþtir")
    If Len(p1) < 8 Then
        MsgBox "Parola en az 8 karakter olmalý; deðiþtirilmedi.", vbExclamation
        Exit Sub
    End If
    p2 = InputBox("Yeni parolayý tekrar yazýn:", "Parola deðiþtir")
    If p1 <> p2 Then
        MsgBox "Parolalar ayný deðil; deðiþtirilmedi.", vbExclamation
        Exit Sub
    End If
    Call UnprotectForMacro                            ' eski parola ile aç
    On Error Resume Next
    ThisWorkbook.Names(PWD_NAME).Delete
    On Error GoTo 0
    ThisWorkbook.Names.Add Name:=PWD_NAME, RefersTo:="=" & QT & Replace(p1, QT, QT & QT) & QT, Visible:=False
    Call ProtectAfterMacro                            ' yeni parola ile koru
    MsgBox "Koruma parolasý deðiþtirildi ve sayfalar yeni parolayla korundu." & vbCrLf & _
           "Parolayý güvenli bir yerde saklayýn ve VBA projesini de kilitleyin" & vbCrLf & _
           "(Araçlar > VBAProject Özellikleri > Koruma).", vbInformation
End Sub
