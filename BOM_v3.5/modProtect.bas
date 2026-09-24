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
' PAROLAYI DEÐÝÞTÝRÝN ve VBA projesini de parolayla kilitleyin (kurulum notuna bakýn).
' =========================================================================
Public Const PROTECT_PWD As String = "Sara#Bom2026!"

Public Function IsOutputSheet(ByVal nm As String) As Boolean
    nm = UCase$(nm)
    IsOutputSheet = (nm = "SUM" Or nm = "ANGLE" Or nm = "PLATE" Or nm = "BOLTS&WASHER" Or _
                     nm Like "SUM_P#*" Or nm Like "ANGLE_P#*" Or nm Like "PLATE_P#*" Or nm Like "BOLTS_P#*")
End Function

Public Sub UnprotectForMacro()
    Dim ws As Worksheet
    On Error Resume Next
    ThisWorkbook.Unprotect PROTECT_PWD
    For Each ws In ThisWorkbook.Worksheets
        If IsOutputSheet(ws.Name) Or UCase$(ws.Name) = "GECMIS" Then ws.Unprotect PROTECT_PWD
    Next ws
    On Error GoTo 0
End Sub

Public Sub ProtectAfterMacro()
    Dim ws As Worksheet, fr As Range
    On Error Resume Next
    For Each ws In ThisWorkbook.Worksheets
        If IsOutputSheet(ws.Name) Then
            ws.Unprotect PROTECT_PWD
            ws.Cells.Locked = False
            Set fr = Nothing
            Set fr = ws.UsedRange.SpecialCells(xlCellTypeFormulas)
            If Not fr Is Nothing Then fr.Locked = True
            ws.Protect Password:=PROTECT_PWD, DrawingObjects:=True, Contents:=True, Scenarios:=True, _
                       UserInterfaceOnly:=True, AllowFormattingCells:=True, AllowFormattingColumns:=True, _
                       AllowFormattingRows:=True, AllowFiltering:=True
            ws.EnableSelection = xlNoRestrictions
        ElseIf UCase$(ws.Name) = "GECMIS" Then
            ws.Unprotect PROTECT_PWD
            ws.Cells.Locked = True
            ws.Protect Password:=PROTECT_PWD, UserInterfaceOnly:=True, AllowFiltering:=True
        End If
    Next ws
    ThisWorkbook.Protect Password:=PROTECT_PWD, Structure:=True
    On Error GoTo 0
End Sub

' Sadece siz (bakým için): VBA penceresinden çalýþtýrýn, parola sorar.
Public Sub Yonetici_KorumayiKaldir()
    If InputBox("Yönetici parolasý:", "Korumayý kaldýr") <> PROTECT_PWD Then
        MsgBox "Parola yanlýþ.", vbCritical
        Exit Sub
    End If
    Call UnprotectForMacro
    MsgBox "Koruma kaldýrýldý. Tekrar korumak için ProtectAfterMacro çalýþtýrýn.", vbInformation
End Sub
