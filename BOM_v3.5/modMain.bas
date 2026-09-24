Attribute VB_Name = "modMain"
Option Explicit

' =========================================================================
' GÝRÝÞ NOKTALARI: düðmeler, ana motor, temizleyici. (Makrolar listesinde görünen tek modül)
' =========================================================================

' =========================================================================
' BOM ÇEVÝRÝCÝ - v2 (Genel Excel motoru)
' DEÐÝÞÝKLÝKLER:
'  1) Baþlýk satýrý artýk HÜCRE bazlý bulunuyor. Eski kod satýrý birleþtirip
'     "TEILE" / "BENENNUNG" arýyordu; antet bloðundaki "Benennung: Teile Schuss..."
'     yanlýþlýkla baþlýk sanýlýyor ve sayfa atlanýyordu. ":" ile biten etiketler yok sayýlýr.
'  2) Ýlk 20 sütun / 20 satýr sýnýrý kaldýrýldý (80 satýr, tüm sütunlar taranýr).
'  3) Birleþtirilmiþ baþlýk + alt baþlýk ("Gewicht" -> "kg/m" | "Gesamt") okunur.
'  4) Sayfa içinde tekrar eden baþlýklar yeniden eþlenir.
'  5) Cývata: "M7990 16 45+3", "STGB 27 310", "SSM24x345-R", "M8x20" okunur.
'     Eski regex "M7990" -> M799 buluyordu.
'  6) Müþterinin yazdýðý SCH / FDRG / MU satýrlarý "tanýnmadý" olmaz; atlanýr ve
'     çevrilen cývata adediyle karþýlaþtýrýlýp AUDIT'e yazýlýr.
'  7) Sýnýflandýrma önek+rakam ile (ör: "UB" artýk "SCHRAUBE" içinde eþleþmez).
'  8) Sayýlar bölge ayarýndan baðýmsýz okunur (5.6 kalite artýk "5,6" -> 8.8 olmaz).
'  9) Ayrý Dicke/Breite ve Abmessung sütunlarý desteklenir; "Winkel 150x12" -> "L 150x12".
' 10) Dosya aðýrlýðý = müþterinin profil+plaka satýr toplamý (yoksa Stahlgewicht satýrý).
' 11) Kütüphanede olmayan cývata için aðýrlýk tahmini (sabit 0.11 / 0.65 yerine).
' 13) Dosya adý / NOT / malzeme adý sütunlarý yazý uzunluðuna göre geniþler.
' v2.3:
' 14) SUM I = Aktarýlan poz, J = Fark % (C+D vs H; tolerans J3, dýþý kýrmýzý).
' 15) Kütüphane saðlýðý: ANGLE/PLATE eksik kod-kg/m-ölçü satýrlarý kýrmýzý; tahmini
'     cývata aðýrlýklarý turuncu + AUDIT'e yazýlýr.
' 16) Parametreler AYARLAR!F:H bloðunda (kod açmadan deðiþtirilir).
' 17) Ýþlem öncesi otomatik yedek, REV/DATE otomatik, GECMIS sayfasý, AUDIT çalýþma bilgisi.
' 18) AUDIT satýr numarasý = köprü (müþteri dosyasýnda ilgili satýra gider).
' 19) Temizle düðmesi AUDIT, FEEDBACK, OGRENME sayfalarýný da temizler.
' 20) Üstüne ekle modunda SUM/BOLTS eski satýrlarý korunur.
' 12) FUBL (Futterblech, ör: "FUBL 18X10") Futterring gibi özel bölüme yazýlýr.
' =========================================================================
' --- GLOBAL KULLANICI AYARLARI (UserForm'dan gelir) ---
' Global Sözlük ve Regex Dizileri
' --- PARAMETRELER (AYARLAR sayfasý F:H bloðundan okunur, bkz. LoadParameters) ---
' --- ÇALIÞMA BÝLGÝSÝ (AUDIT özeti ve GECMIS sayfasý için) ---
' --- AUDIT köprüleri: iþlenen dosyanýn tam yolu ve sayfasý ---
' v3.0:
' 21) Kod 12 modüle ayrýldý (modMain dýþýndakiler Option Private Module).
' 22) Poz çakýþmasý kontrolü (ayný poz farklý profil/ölçü/boy -> uyarý + NOT: ÇAKIÞMA!).
' 23) Sayfa + çalýþma kitabý korumasý (formüller kilitli).
' 24) (v3.1) Lisans sistemi kaldýrýldý; lisanslama web sitesi üzerinden yapýlacak.
' 28) (v3.5) OGRENME sayfasýnda 4 veri giriþi: MALZEME KODU, MALZEME CÝNSÝ, kg/m, YÜZEY ALAN
'     (cývata için birim aðýrlýk / 1000 adet aðýrlýðý -> BOLT-LIBRARY). Eksik verili satýr aktarýlmaz.
' 27) (v3.4) 'No.' adet sütunu, Nut/Wash/Bolt ipucu, ISO standartlarý, montaj çift sayým korumasý,
'     iki sayýlý hücre düzeltmesi. PDF için PDF_BOM_Okuyucu.html aracý.
' 26) (v3.3) PDF malzeme listesi (Power Query, Excel 365/2021), PROFILE sütunu, D460 daire plaka,
'     CHS/RHS boru kodlarý, revizyon iþaretleri, kütüphane temizleme makrosu.
' 25) (v3.1) Sütun baþlýklarýnda dosya adýnýn sadece farklý kýsmý (Z175 ESH gibi); SUM!K = kýsa ad.

Sub Evrensel_BOM_Cevirici_Core()
    Application.EnableCancelKey = xlErrorHandler
    On Error GoTo ErrorHandler
    conflictCount = 0
    bufOverflowMsg = ""

    Dim fd As FileDialog
    Dim fileNo As Integer
    Dim lineStr As String, rawLine As String, rawLineUpper As String
    Dim curWsBolt As Worksheet, curWsSum As Worksheet, wsFeed As Worksheet
    Dim wsBoltLib As Worksheet, wsProfileLib As Worksheet
    Dim wbIn As Workbook, wsIn As Worksheet
        Dim srcData As Variant
    Dim maxR As Long, maxC As Long, lastRowExt As Long, lastColExt As Long
    Dim wbOpenAlready As Boolean
    Dim processedFolderPath As String
    
    Dim i As Long, r As Long, c As Long, scanC As Long, qIdx As Long, quantIdx As Long
    Dim targetRowBolt As Long, sumRow As Long
    Dim posNo As String, quality As String, rawName As String, rawNameUpper As String, rawTypeUpper As String
    Dim lengthVal As Double, quantVal As Long
    Dim currentSection As String
    Dim existRow As Long, duplicateCount As Long
    Dim parts() As String, logText As String
    
    Dim extraPctInput As String
    Dim finalBoltQty As Long, bQtyX As Long
    Dim dValX As Long, lValX As Long, isFuRingX As Boolean, boltNoteX As String, boltQualX As String
    Dim hwCodeX As String, hwNameX As String, boltKeyX As String, qtyKeyX As String
    Dim dValE As Long, lValE As Long, isFuRingE As Boolean, boltNoteE As String, boltQualE As String
    Dim hwCodeE As String, hwNameE As String, boltKeyE As String, qtyKeyE As String
    
    Dim currentSumRow As Long, weightCol As Long, tmpWeightStr As String
    Dim mergeAns As Integer
    Dim xsrNote As String, matCodeStr As String, posKeyAngle As String, aQty As Double
    Dim pThick As Double, pWidth As Double, pNoteX As String, finalXsrPlateNote As String, posKeyPlate As String, pQty As Double
    Dim pThickE As Double, pWidthE As Double, pNoteE As String, finalPlateNote As String, posKeyPlateE As String, pQtyE As Double
    Dim matCodeStrE As String, posKeyAngleE As String, aQtyE As Double
    Dim cleanNoSpace As String, isBoltItem As Boolean, isPlateItem As Boolean, isAngleItem As Boolean, isInsideParens As Boolean
    Dim notStr As String, dinVal As String, remVal As String, cleanProf As String
    
    Dim posCol As Long, qtyCol As Long, nameCol As Long, lenCol As Long, matCol As Long, remCol As Long, dinCol As Long
    Dim hRow As Long, lastR As Long, isGarbageRow As Boolean, cellVal As String, rText As String
    Dim hasPosKwd As Boolean, hasQtyKwd As Boolean, hasNameKwd As Boolean
    Dim loopGuard As Long
    
    Dim arrDiams As Variant, arrBoltKeys As Variant
    Dim dIdx As Long, dCurr As Long, bKey As Variant, colC As Long
    ReDim realBoltColTotals(6 To 500) As Long
    Dim bParts() As String, bWt As Double, isThisFuRing As Boolean, bNote As String
    Dim nutCode As String, nutWt As Double, pwCode As String, pwWt As Double, swCode As String, swWt As Double
    
    Dim dictBoltWeights As Object, dictSelectedFiles As Object
    Dim logFile As Object
    Dim rLib As Long, rPLib As Long, pCode As String, pName As String, normKey As String, stdKey As String
    
    Dim validFiles() As String, vfCount As Long
    Dim batchNum As Long, fIdx As Long, bStart As Long, bEnd As Long
    Dim logFilePath As String, finalMsg As String
    Dim feedKeys As Variant, fr As Long
    Dim appendMode As Boolean, existingFileCount As Long
    Dim sourceLineNo As Long
    Dim firstChar As String
    
    Dim assemblyPos As String, assemblyProfile As String, assemblyName As String
    Dim assemblyQty As Long, assemblyArea As Double, assemblyWeight As Double
    Dim assemblyType As String, assemblyStatus As String, assemblyIssue As String
    Dim assemblyConfidence As Long, assemblyParsed As String
    Dim assemblyThickness As Double, assemblyWidth As Double, assemblyLength As Double
    Dim assemblyMatCode As String, assemblyIsPlate As Boolean, assemblyIsAngle As Boolean
    Dim assemblyValid As Boolean, assemblyApproxMatch As Boolean

    Dim partValid As Boolean
    Dim partPos As String, partQty As Long, partProfile As String
    Dim partGrade As String, partLength As Double, partWeight As Double
    Dim partAssembly As String
    Dim partIsPlate As Boolean, partIsProfile As Boolean
    Dim partMatCode As String, partApprox As Boolean
    Dim partStatus As String, partConf As Long, partIssue As String
    Dim partParsed As String
    Dim rawQty As Long, fireliQty As Long
    Dim reconTxt As String, reconOut As Long, reconNoRef As Long, reconRow As Long, reconSheet As String, reconPrev As Long
    Dim capTxt As String

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    Call UnprotectForMacro

    Call CheckAndInitSettingsSheet
    Call LoadParameters
    Call InitAuditSheet
    Call LoadDictionaryArrays
    runStart = Now
    runBackupPath = ""
    runRev = ""
    auditCurrentPath = ""
    auditCurrentSheet = ""
    
    Set globalRegBolt = CreateObject("VBScript.RegExp")
    globalRegBolt.IgnoreCase = True
    globalRegBolt.Global = False

    Set qualRegexGlobal = CreateObject("VBScript.RegExp")
    qualRegexGlobal.IgnoreCase = True
    qualRegexGlobal.Global = False
    qualRegexGlobal.Pattern = "(S235|S275|S355|S460|ST37|ST44|ST52|A36|A572)"

    On Error Resume Next
    Set curWsAngle = ThisWorkbook.Sheets("ANGLE")
    Set curWsPlate = ThisWorkbook.Sheets("PLATE")
    Set curWsBolt = ThisWorkbook.Sheets("BOLTS&WASHER")
    Set curWsSum = ThisWorkbook.Sheets("SUM")
    Set wsBoltLib = ThisWorkbook.Sheets("BOLT-LIBRARY")
    Set wsProfileLib = ThisWorkbook.Sheets("L-U-I-O-Y-LIBRARY")
    On Error GoTo ErrorHandler
    
    If curWsAngle Is Nothing Or curWsPlate Is Nothing Or curWsBolt Is Nothing Or curWsSum Is Nothing Then
        Call ProtectAfterMacro
        Call ResetExcelState
        MsgBox "HATA: Gerekli ana sayfalar bulunamadý!", vbCritical
        Exit Sub
End If

    isDryRun = False

    ' Arayüzden (UserForm) gelen verileri al
    appendMode = usrAppendMode
    extraPct = usrExtraPct
    isMerge = usrIsMerge

    ' Dosyalar: hazýr liste (son iþlemi yenile / test) ya da dosya seçme penceresi
    Dim selFiles() As String, selCount As Long, sfI As Long
    If IsArray(runPresetFiles) Then
        selCount = UBound(runPresetFiles) - LBound(runPresetFiles) + 1
        If selCount > 0 Then
            ReDim selFiles(1 To selCount)
            For sfI = 1 To selCount
                selFiles(sfI) = CStr(runPresetFiles(LBound(runPresetFiles) + sfI - 1))
            Next sfI
        End If
        runPresetFiles = Empty
        If selCount <= 0 Then GoTo Cikis
    Else
        Set fd = Application.FileDialog(msoFileDialogFilePicker)
        fd.Title = "Ýþlenecek BOM Dosyalarýný Seçin"
        fd.Filters.Clear
        fd.Filters.Add "Tüm BOM Dosyalarý", "*.xsr; *.txt; *.xls; *.xlsx; *.xlsm; *.xlsb; *.csv; *.pdf"
        fd.AllowMultiSelect = True
        If fd.Show <> -1 Then
            GoTo Cikis
        End If
        selCount = fd.SelectedItems.Count
        If selCount = 0 Then
            GoTo Cikis
        End If
        ReDim selFiles(1 To selCount)
        For sfI = 1 To selCount
            selFiles(sfI) = fd.SelectedItems(sfI)
        Next sfI
    End If

    runMode = IIf(appendMode, "ÜSTÜNE EKLE", "SIFIRDAN")
    ' Ýþlemden önce þablonun tarihli yedeði
    If prmBackup And Not isDryRun And Not runQuiet Then runBackupPath = MakeBackup()

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set dictBoltWeights = CreateObject("Scripting.Dictionary")
    Set dictProfileLib = CreateObject("Scripting.Dictionary")
    Set dictSelectedFiles = CreateObject("Scripting.Dictionary")
    Set dictFeedback = CreateObject("Scripting.Dictionary")
    
    dictBoltWeights.CompareMode = 1
    dictProfileLib.CompareMode = 1
    dictSelectedFiles.CompareMode = 1
    dictFeedback.CompareMode = 1
    Call LoadLearnedItems
    
    If Not wsBoltLib Is Nothing Then
        For rLib = 3 To wsBoltLib.Cells(wsBoltLib.Rows.Count, 1).End(xlUp).Row
            If Trim(wsBoltLib.Cells(rLib, 1).Text) <> "" Then
                pCode = Trim(wsBoltLib.Cells(rLib, 1).Text)
                pName = Trim(wsBoltLib.Cells(rLib, 2).Text)
                bWt = CellNumber(wsBoltLib.Cells(rLib, 3))
                If bWt <= 0 Then bWt = CellNumber(wsBoltLib.Cells(rLib, 4)) / 1000#
                If pCode <> "" And bWt > 0 Then dictBoltWeights(pCode) = bWt
                If pName <> "" And bWt > 0 Then dictBoltWeights(pName) = bWt
            End If
        Next rLib
    End If

    If Not wsProfileLib Is Nothing Then
        For rPLib = 2 To wsProfileLib.Cells(wsProfileLib.Rows.Count, 1).End(xlUp).Row
            pCode = Trim(wsProfileLib.Cells(rPLib, 1).Text)
            pName = Trim(wsProfileLib.Cells(rPLib, 2).Text)
            ' Eski "YENI_EKLE" / "TANIMSIZ" yer tutucularý kod deðildir, kütüphaneye alýnmaz
            If UCase(pCode) = "YENI_EKLE" Or UCase(pCode) = "TANIMSIZ" Then pCode = ""
            If pName <> "" And pCode <> "" Then
                If Not dictProfileLib.Exists(CanonicalSectionKey(pName)) Then dictProfileLib.Add CanonicalSectionKey(pName), pCode
                If Not dictProfileLib.Exists(CanonicalSectionKey(pCode)) Then dictProfileLib.Add CanonicalSectionKey(pCode), pCode
                normKey = UCase(Replace(pName, " ", ""))
                If Not dictProfileLib.Exists(normKey) Then dictProfileLib.Add normKey, pCode
                stdKey = StandardizeProfileName(pName)
                If stdKey <> "" And Not dictProfileLib.Exists(stdKey) Then dictProfileLib.Add stdKey, pCode
            End If
        Next rPLib
    End If
    
    If appendMode And Not isDryRun Then
        existingFileCount = CountExistingOutputFiles(curWsSum)
        If (existingFileCount + selCount) > 45 Then
            MsgBox "DÝKKAT: Þablon kapasitesi (Maksimum 45 Dosya) aþýlýyor!" & vbCrLf & vbCrLf & _
                   "Mevcut Dosya: " & existingFileCount & vbCrLf & _
                   "Yeni Seçilen: " & selCount & vbCrLf & _
                   "AZ Formül sütununu korumak için iþlem durduruldu. Lütfen yeni liste oluþturun.", vbCritical, "Kapasite Ýhlali"
            GoTo Cikis
        End If
    End If

    logText = "=== MTL BOM OTOMASYON MOTORU RAPORU ===" & vbCrLf
    logText = logText & "Tarih: " & Now & vbCrLf & "-----------------------------------" & vbCrLf
    
    ReDim validFiles(1 To selCount)
    vfCount = 0
    duplicateCount = 0
    
    For sfI = 1 To selCount
        vItem = selFiles(sfI)
        fileName = fso.GetBaseName(vItem)
        If Not dictSelectedFiles.Exists(fileName) Then
            dictSelectedFiles.Add fileName, True
            vfCount = vfCount + 1
            validFiles(vfCount) = vItem
Else
            duplicateCount = duplicateCount + 1
            logText = logText & "- ATLANDI (Mükerrer): " & fileName & vbCrLf
End If
    Next sfI
    
    If vfCount = 0 Then
        MsgBox "Geçerli yeni bir dosya seçilmedi.", vbExclamation
        GoTo Cikis
End If

    curWsSum.Activate
    Call ShowProgress(0, "Motor Baþlatýlýyor...")

    batchNum = 1
    totalProcessed = 0

    ' Montaj çift sayým korumasý: iþlenen dosyalarýn ve SUM'daki eski dosyalarýn iþaretleri
    Set dictFileMarks = CreateObject("Scripting.Dictionary")
    dictFileMarks.CompareMode = 1
    Dim mkI As Long, mkS As String
    For mkI = 1 To vfCount
        mkS = FileMark(fso.GetFileName(validFiles(mkI)))
        If mkS <> "" And Not dictFileMarks.Exists(mkS) Then dictFileMarks.Add mkS, 1
    Next mkI
    If appendMode Then
        For mkI = 5 To ThisWorkbook.Sheets("SUM").Cells(ThisWorkbook.Sheets("SUM").Rows.Count, 1).End(xlUp).Row
            mkS = FileMark(ThisWorkbook.Sheets("SUM").Cells(mkI, 1).Text)
            If mkS <> "" And Not dictFileMarks.Exists(mkS) Then dictFileMarks.Add mkS, 1
        Next mkI
    End If
    
    For bStart = 1 To vfCount Step 45
        bEnd = bStart + 44
        If bEnd > vfCount Then bEnd = vfCount
        
        Set dictPosAngle = CreateObject("Scripting.Dictionary")
        Set dictPosPlate = CreateObject("Scripting.Dictionary")
        Set dictBoltQty = CreateObject("Scripting.Dictionary")
        Set dictBoltList = CreateObject("Scripting.Dictionary")
        Set dictDiameters = CreateObject("Scripting.Dictionary")
        dictPosAngle.CompareMode = 1
        dictPosPlate.CompareMode = 1
        dictBoltQty.CompareMode = 1
        dictBoltList.CompareMode = 1
        dictDiameters.CompareMode = 1
        
        If batchNum = 1 Then
            Set curWsAngle = ThisWorkbook.Sheets("ANGLE")
            Set curWsPlate = ThisWorkbook.Sheets("PLATE")
            Set curWsBolt = ThisWorkbook.Sheets("BOLTS&WASHER")
            Set curWsSum = ThisWorkbook.Sheets("SUM")

            If Not appendMode And Not isDryRun Then
                Call CleanTemplateRanges(curWsAngle, curWsPlate, curWsBolt, curWsSum)
End If
Else
            If Not isDryRun Then
                ThisWorkbook.Sheets("ANGLE").Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
                Set curWsAngle = ActiveSheet
                curWsAngle.Name = "ANGLE_P" & batchNum
                
                ThisWorkbook.Sheets("PLATE").Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
                Set curWsPlate = ActiveSheet
                curWsPlate.Name = "PLATE_P" & batchNum
                
                ThisWorkbook.Sheets("BOLTS&WASHER").Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
                Set curWsBolt = ActiveSheet
                curWsBolt.Name = "BOLTS_P" & batchNum
                
                ThisWorkbook.Sheets("SUM").Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
                Set curWsSum = ActiveSheet
                curWsSum.Name = "SUM_P" & batchNum
                
                Call CleanTemplateRanges(curWsAngle, curWsPlate, curWsBolt, curWsSum)
End If
        End If
        
        Set curWsAngleRef = curWsAngle
        Set curWsPlateRef = curWsPlate
        Set curWsBoltRef = curWsBolt
        Set curWsSumRef = curWsSum

        ReDim bufAngle(1 To 5000, 1 To 70)
        ReDim bufPlate(1 To 5000, 1 To 70)
        ReDim bufBolt(1 To 2000, 1 To 70)
        ReDim bufSum(1 To 200, 1 To 20)

        If appendMode And batchNum = 1 Then
            Dim rAInit As Long, rPInit As Long, vAInit As Variant, vPInit As Variant
            Dim rInit As Long, cInit As Long
            rAInit = curWsAngle.Cells(curWsAngle.Rows.Count, 2).End(xlUp).Row
            If rAInit >= 5 Then
                vAInit = curWsAngle.Range(curWsAngle.Cells(5, 1), curWsAngle.Cells(rAInit, 65)).Value2
                For rInit = 1 To UBound(vAInit, 1)
                    For cInit = 1 To UBound(vAInit, 2)
                        bufAngle(rInit + 4, cInit) = vAInit(rInit, cInit)
                    Next cInit
                Next rInit
            End If
            
            rPInit = curWsPlate.Cells(curWsPlate.Rows.Count, 2).End(xlUp).Row
            If rPInit >= 5 Then
                vPInit = curWsPlate.Range(curWsPlate.Cells(5, 1), curWsPlate.Cells(rPInit, 65)).Value2
                For rInit = 1 To UBound(vPInit, 1)
                    For cInit = 1 To UBound(vPInit, 2)
                        bufPlate(rInit + 4, cInit) = vPInit(rInit, cInit)
                    Next cInit
                Next rInit
            End If

            ' SUM ve BOLTS tamponlarýný da mevcut içerikle doldur (aksi halde flush eski satýrlarý siler)
            Dim vSInit As Variant, vBInit As Variant, rSInit As Long, rBInit As Long
            rSInit = curWsSum.Cells(curWsSum.Rows.Count, 1).End(xlUp).Row
            If rSInit > UBound(bufSum, 1) Then rSInit = UBound(bufSum, 1)
            If rSInit >= 5 Then
                vSInit = curWsSum.Range(curWsSum.Cells(5, 1), curWsSum.Cells(rSInit, UBound(bufSum, 2))).Formula
                For rInit = 1 To UBound(vSInit, 1)
                    For cInit = 1 To UBound(vSInit, 2)
                        bufSum(rInit + 4, cInit) = vSInit(rInit, cInit)
                    Next cInit
                Next rInit
            End If
            rBInit = curWsBolt.Cells(curWsBolt.Rows.Count, 2).End(xlUp).Row
            If rBInit > UBound(bufBolt, 1) Then rBInit = UBound(bufBolt, 1)
            If rBInit >= 5 Then
                vBInit = curWsBolt.Range(curWsBolt.Cells(5, 1), curWsBolt.Cells(rBInit, UBound(bufBolt, 2))).Formula
                For rInit = 1 To UBound(vBInit, 1)
                    For cInit = 1 To UBound(vBInit, 2)
                        bufBolt(rInit + 4, cInit) = vBInit(rInit, cInit)
                    Next cInit
                Next rInit
            End If

            existingFileCount = CountExistingOutputFiles(curWsSum)
            currentColAngle = 9 + existingFileCount
            currentColPlate = 8 + existingFileCount
            currentColBolt = 6 + existingFileCount
            targetRowAngle = GetNextOutputRow(curWsAngle, 2, 5)
            targetRowPlate = GetNextOutputRow(curWsPlate, 2, 5)
            targetRowBolt = GetNextOutputRow(curWsBolt, 2, 5)
            sumRow = GetNextOutputRow(curWsSum, 1, 5)
            If isMerge Then
                Call LoadExistingPositionDictionaries(curWsAngle, curWsPlate, dictPosAngle, dictPosPlate)
End If
Else
            currentColAngle = 9
            currentColPlate = 8
            currentColBolt = 6
            targetRowAngle = 5
            targetRowPlate = 5
            targetRowBolt = 5
            sumRow = 5
End If

        For fIdx = bStart To bEnd
            vItem = validFiles(fIdx)
            fileName = fso.GetBaseName(vItem)
            ext = LCase(fso.GetExtensionName(vItem))
            
            Call ShowProgress(fIdx / vfCount, "Ýþleniyor: " & fileName)
            
            fileWeight = 0
            currentSumRow = sumRow
            Dim fileStartCount As Long
            fileStartCount = totalProcessed
            auditCurrentPath = CStr(vItem)
            auditCurrentSheet = ""
            
            If Not isDryRun Then
                Call SafeWrite(curWsSum, currentSumRow, 1, fileName, 5000, 10)
                Call SafeWrite(curWsSum, currentSumRow, 2, 1, 5000, 10)
                curWsSum.Cells(currentSumRow, 2).Interior.Color = vbYellow
End If
            sumRow = sumRow + 1
            
            ' --- Dosya tipine göre motor (modMotorXSR / modMotorExcel) ---
            On Error GoTo ErrorHandler
            If ext = "xsr" Or ext = "txt" Then
                Call ProcessXsrFile
            ElseIf ext = "xls" Or ext = "xlsx" Or ext = "xlsm" Or ext = "xlsb" Or ext = "csv" Then
                Call ProcessExcelFile
            ElseIf ext = "pdf" Then
                Call ProcessPdfFile
            End If
            
            If Not isDryRun Then
                Call SafeWrite(curWsSum, currentSumRow, 8, fileWeight)
                Call SafeWrite(curWsSum, currentSumRow, 9, totalProcessed - fileStartCount)
                                                                currentColAngle = currentColAngle + 1
                currentColPlate = currentColPlate + 1
                currentColBolt = currentColBolt + 1
End If
        Next fIdx
        auditCurrentPath = ""
        auditCurrentSheet = ""
        
        ' -------------------------------------------------------------------------
        ' PAKET CIVATA VE YEÞÝL 3'LÜ SET DÝZÝMÝ
        ' -------------------------------------------------------------------------
        If dictDiameters.Count > 0 And Not isDryRun Then
            Call ShowProgress(1, "Cývatalar listeye diziliyor...")
            
            arrDiams = dictDiameters.keys
            Call SortArrayInPlace(arrDiams)
            arrBoltKeys = dictBoltList.keys
            Call SortArrayInPlace(arrBoltKeys)
            
            For dIdx = LBound(arrDiams) To UBound(arrDiams)
                dCurr = CLng(val(arrDiams(dIdx)))
                For colC = 6 To 500
                    realBoltColTotals(colC) = 0
                Next colC
                
                For Each bKey In arrBoltKeys
                    bParts = Split(dictBoltList(bKey), "^")
                    If CLng(val(bParts(0))) = dCurr Then
                        isThisFuRing = False
                        If UBound(bParts) >= 5 Then
                            If UCase(bParts(5)) = "TRUE" Then isThisFuRing = True
End If
                        
                        If Not isThisFuRing Then
                            Call SafeWrite(curWsBolt, targetRowBolt, 2, bParts(2))
                            Call SafeWrite(curWsBolt, targetRowBolt, 3, bParts(3))
                            Call SafeWrite(curWsBolt, targetRowBolt, 4, bParts(4))
                            
                            If dictBoltWeights.Exists(bParts(2)) Then
                                bWt = dictBoltWeights(bParts(2))
Else
                                bWt = GetBoltWeight(dCurr, CLng(val(bParts(1))))
                                Call FlagEstimatedWeight(curWsBolt, targetRowBolt, bParts(2), bParts(3), bWt)
End If
                            ' E formülü korunur; bu deðer yalnýzca formül boþ dönerse yazýlýr (FlushBoltBlock)
                            Call SafeWrite(curWsBolt, targetRowBolt, 5, bWt)
                            
                            If UBound(bParts) >= 6 Then
                                bNote = bParts(6)
                                If bNote <> "" Then Call SafeWrite(curWsBolt, targetRowBolt, 54, bNote)
End If
                            
                            For colC = 6 To (currentColBolt - 1)
                                rawQty = 0
                                fireliQty = 0
                                If dictBoltQty.Exists(bKey & "|" & colC) Then fireliQty = dictBoltQty(bKey & "|" & colC)
                                rawQty = Int(fireliQty / (1 + (extraPct / 100#)) + 0.1)
                                If fireliQty > 0 Then
                                    Call SafeWrite(curWsBolt, targetRowBolt, colC, fireliQty)
                                    realBoltColTotals(colC) = realBoltColTotals(colC) + rawQty
End If
                            Next colC
                            targetRowBolt = targetRowBolt + 1
End If
                    End If
                Next bKey
                
                If dCurr <> 99 Then
                    ' SOMUN
                    nutCode = GetNutCode(dCurr)
                    If dictBoltWeights.Exists(nutCode) Then
                        nutWt = dictBoltWeights(nutCode)
Else
                        nutWt = GetNutWeight(dCurr)
                        Call FlagEstimatedWeight(curWsBolt, targetRowBolt, nutCode, GetNutName(dCurr), nutWt)
End If
                    Call SafeWrite(curWsBolt, targetRowBolt, 2, nutCode)
                    Call SafeWrite(curWsBolt, targetRowBolt, 3, GetNutName(dCurr))
                    Call SafeWrite(curWsBolt, targetRowBolt, 4, "")
                    Call SafeWrite(curWsBolt, targetRowBolt, 5, nutWt)
                    For colC = 6 To (currentColBolt - 1)
                        If realBoltColTotals(colC) > 0 Then
                            Call SafeWrite(curWsBolt, targetRowBolt, colC, realBoltColTotals(colC) + Int((realBoltColTotals(colC) * (extraPct / 100#)) + 0.5))
End If
                    Next colC
                    Call FormatGreenRow(curWsBolt, targetRowBolt)
                    targetRowBolt = targetRowBolt + 1
                    
                    ' PUL
                    pwCode = GetPWCode(dCurr)
                    If dictBoltWeights.Exists(pwCode) Then
                        pwWt = dictBoltWeights(pwCode)
Else
                        pwWt = GetPWWeight(dCurr)
                        Call FlagEstimatedWeight(curWsBolt, targetRowBolt, pwCode, GetPWName(dCurr), pwWt)
End If
                    Call SafeWrite(curWsBolt, targetRowBolt, 2, pwCode)
                    Call SafeWrite(curWsBolt, targetRowBolt, 3, GetPWName(dCurr))
                    Call SafeWrite(curWsBolt, targetRowBolt, 4, "G26")
                    Call SafeWrite(curWsBolt, targetRowBolt, 5, pwWt)
                    For colC = 6 To (currentColBolt - 1)
                        If realBoltColTotals(colC) > 0 Then
                            Call SafeWrite(curWsBolt, targetRowBolt, colC, realBoltColTotals(colC) + Int((realBoltColTotals(colC) * (extraPct / 100#)) + 0.5))
End If
                    Next colC
                    Call FormatGreenRow(curWsBolt, targetRowBolt)
                    targetRowBolt = targetRowBolt + 1
                    
                    ' YAYLI RONDELA
                    swCode = GetSWCode(dCurr)
                    If dictBoltWeights.Exists(swCode) Then
                        swWt = dictBoltWeights(swCode)
Else
                        swWt = GetSWWeight(dCurr)
                        Call FlagEstimatedWeight(curWsBolt, targetRowBolt, swCode, GetSWName(dCurr), swWt)
End If
                    Call SafeWrite(curWsBolt, targetRowBolt, 2, swCode)
                    Call SafeWrite(curWsBolt, targetRowBolt, 3, GetSWName(dCurr))
                    Call SafeWrite(curWsBolt, targetRowBolt, 4, "")
                    Call SafeWrite(curWsBolt, targetRowBolt, 5, swWt)
                    For colC = 6 To (currentColBolt - 1)
                        If realBoltColTotals(colC) > 0 Then
                            Call SafeWrite(curWsBolt, targetRowBolt, colC, realBoltColTotals(colC) + Int((realBoltColTotals(colC) * (extraPct / 100#)) + 0.5))
End If
                    Next colC
                    Call FormatGreenRow(curWsBolt, targetRowBolt)
                    targetRowBolt = targetRowBolt + 1
End If
            Next dIdx
            
            ' ÖZEL FUTTERRING
            For Each bKey In arrBoltKeys
                bParts = Split(dictBoltList(bKey), "^")
                isThisFuRing = False
                If UBound(bParts) >= 5 Then
                    If UCase(bParts(5)) = "TRUE" Then isThisFuRing = True
End If
                
                If isThisFuRing Then
                    Call SafeWrite(curWsBolt, targetRowBolt, 2, bParts(2))
                    Call SafeWrite(curWsBolt, targetRowBolt, 3, bParts(3))
                    Call SafeWrite(curWsBolt, targetRowBolt, 4, bParts(4))
                    If dictBoltWeights.Exists(bParts(2)) Then
                        bWt = dictBoltWeights(bParts(2))
                    ElseIf CLng(Val(bParts(0))) <> 99 Then
                        bWt = GetBoltWeight(CLng(Val(bParts(0))), CLng(Val(bParts(1))))
                        Call FlagEstimatedWeight(curWsBolt, targetRowBolt, bParts(2), bParts(3), bWt)
                    Else
                        bWt = 0.05
                        Call FlagEstimatedWeight(curWsBolt, targetRowBolt, bParts(2), bParts(3), bWt)
                    End If
                    Call SafeWrite(curWsBolt, targetRowBolt, 5, bWt)
                    If UBound(bParts) >= 6 Then
                        If bParts(6) <> "" Then Call SafeWrite(curWsBolt, targetRowBolt, 54, bParts(6))
End If
                    For colC = 6 To (currentColBolt - 1)
                        quantVal = 0
                        If dictBoltQty.Exists(bKey & "|" & colC) Then quantVal = dictBoltQty(bKey & "|" & colC)
                        If quantVal > 0 Then Call SafeWrite(curWsBolt, targetRowBolt, colC, quantVal)
                    Next colC
                    targetRowBolt = targetRowBolt + 1
End If
            Next bKey
End If
        
        Call FlushAllBuffers(curWsAngle, curWsPlate, curWsBolt, curWsSum, targetRowAngle, targetRowPlate, targetRowBolt, sumRow, currentColAngle, currentColPlate, currentColBolt)
        If Not isDryRun Then
            ' þablon kapasitesi: ANGLE H (UNIT WEIGHT), PLATE G (UNIT WEIGHT) formülleri yeterli mi?
            capTxt = capTxt & TemplateCapacityWarning(curWsAngle, targetRowAngle - 1, 8)
            capTxt = capTxt & TemplateCapacityWarning(curWsPlate, targetRowPlate - 1, 7)
        End If
        If Not isDryRun Then
            Call WriteSumReconciliation(curWsSum, sumRow - 1)
            reconPrev = reconRow
            reconTxt = reconTxt & CollectReconciliation(curWsSum, sumRow - 1, reconOut, reconNoRef, reconRow)
            If reconPrev = 0 And reconRow > 0 Then reconSheet = curWsSum.Name
            Call UpdateFileHeaders(curWsSum, curWsAngle, curWsPlate, curWsBolt)
            Call CheckLibraryHealth(curWsAngle, targetRowAngle - 1)
        End If
        batchNum = batchNum + 1
    Next bStart

    ' Yazý uzunluðuna göre sütun geniþliklerini ayarla (dosya adlarý, notlar)
    If Not isDryRun Then
        Call AutoFitOutputColumns
        Call ApplyHealthFormatting
    End If

    Call RemoveProgress

    logText = logText & vbCrLf & "-----------------------------------" & vbCrLf
    logText = logText & "Toplam Atýlan Dosya: " & selCount & vbCrLf
    logText = logText & "Baþarýyla Ýþlenen Dosya: " & vfCount & vbCrLf
    logText = logText & "Þablona Ýþlenen Toplam Poz: " & totalProcessed & vbCrLf
    
    On Error Resume Next
    Set wsFeed = ThisWorkbook.Sheets("FEEDBACK")
    On Error GoTo 0
    If wsFeed Is Nothing Then
        Set wsFeed = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsFeed.Name = "FEEDBACK"
    End If
    wsFeed.Cells.Clear
    
    If isDryRun Then
        finalMsg = "TEST (DRY-RUN) BAÞARIYLA TAMAMLANDI!" & vbCrLf & vbCrLf & _
                   "Þablona hiçbir þey yazýlmadý. Sadece kütüphane eksikleri ve hatalar analiz edildi."
    Else
        finalMsg = vfCount & " adet dosya baþarýyla iþlendi." & vbCrLf & _
                   "Toplam " & totalProcessed & " parça þablona eklendi."
    End If

    ' ÖÐRENEN KÜTÜPHANE: tanýnmayanlarýn HEPSÝ tek sayfada ("OGRENME") listelenir
    Dim ogrenilenSayi As Long
    ogrenilenSayi = dictFeedback.Count
    If ogrenilenSayi > 0 Then
        Call WriteUnknownsSheet(dictFeedback)
        finalMsg = finalMsg & vbCrLf & vbCrLf & "DÝKKAT: " & ogrenilenSayi & " adet tanýnmayan parça 'OGRENME' sayfasýnda listelendi." & vbCrLf & _
                   "Her parça için TÜR'ü (ANGLE / PLATE / BOLTS&WASHER) seçip sarý alanlarý doldurun ve" & vbCrLf & _
                   "'KÜTÜPHANEYE AKTAR' düðmesine basýn; son iþlem ayný dosyalarla otomatik yenilenir."
    End If
    
    If conflictCount > 0 Then
        finalMsg = finalMsg & vbCrLf & vbCrLf & "DÝKKAT: " & conflictCount & " poz çakýþmasý bulundu (ayný poz, farklý profil / ölçü / boy)." & vbCrLf & _
                   "Adetler toplandý; NOT sütununda 'ÇAKIÞMA!' ve AUDIT'te detay var."
    End If

    If bufOverflowMsg <> "" Or capTxt <> "" Then
        finalMsg = finalMsg & vbCrLf & vbCrLf & "KAPASÝTE AÞILDI - bazý satýrlar eksik ya da aðýrlýksýz:" & vbCrLf & _
                   bufOverflowMsg & capTxt & "Listeyi bölerek (daha az dosya ile) iþleyin."
        AuditRecord "", "SYSTEM", 0, "", "Kapasite", "SYSTEM", "ERROR", 0, "", Replace(bufOverflowMsg & capTxt, vbCrLf, " ")
    End If
    If reconOut > 0 Then
        finalMsg = finalMsg & vbCrLf & vbCrLf & "AÐIRLIK MUTABAKATI: " & reconOut & " dosyada þablon aðýrlýðý listedekinden" & vbCrLf & _
                   "tolerans (±" & prmWeightTolPct & "%) dýþýnda farklý:" & vbCrLf & FirstLines(reconTxt, 10) & _
                   "Olasý nedenler: tanýnmayan/atlanan parça (OGRENME, AUDIT), kütüphanede eksik veya yanlýþ kg/m" & vbCrLf & _
                   "(Kutuphane_Kontrol), boy birimi. SUM sayfasýnda J sütunu kýrmýzý."
    End If
    If reconNoRef > 0 Then
        finalMsg = finalMsg & vbCrLf & vbCrLf & reconNoRef & " dosyanýn listesinde toplam aðýrlýk yok; aðýrlýk mutabakatý yapýlamadý."
    End If

    If Not isDryRun And Not runQuiet Then
        runRev = UpdateRevDate(appendMode)
        Call AppendRunHistory(validFiles, vfCount, totalProcessed)
        Call SaveLastRun(validFiles, vfCount, appendMode, extraPct, isMerge, runBackupPath)
    End If
    Call FinalizeAuditSheet
    logFilePath = IIf(ThisWorkbook.Path <> "", ThisWorkbook.Path & "\MTL_BOM_Log.txt", Environ("USERPROFILE") & "\Desktop\MTL_BOM_Log.txt")
    On Error Resume Next
    Set logFile = fso.CreateTextFile(logFilePath, True)
    If Not logFile Is Nothing Then
        logFile.WriteLine logText
        logFile.Close
End If
    On Error GoTo 0
    
    ' Dosya Arsivleme iptal edildi

    Set dictPosAngle = Nothing
    Set dictPosPlate = Nothing
    Set dictBoltQty = Nothing
    Set dictBoltList = Nothing
    Set dictDiameters = Nothing
    Set dictBoltWeights = Nothing
    Set dictProfileLib = Nothing
    Set dictSelectedFiles = Nothing
    Set dictFeedback = Nothing
    Set fso = Nothing
    Set globalRegBolt = Nothing
    Set qualRegexGlobal = Nothing

    Application.StatusBar = False
    Call ProtectAfterMacro
    Call ResetExcelState
    
    On Error Resume Next
    ThisWorkbook.Sheets("SUM").Activate
    ThisWorkbook.Sheets("SUM").Range("A5").Select
    On Error GoTo 0
    
    If Not runQuiet Then MsgBox finalMsg, IIf(ogrenilenSayi > 0 Or reconOut > 0 Or conflictCount > 0 Or bufOverflowMsg <> "" Or capTxt <> "", vbExclamation, vbInformation), "MTL Evrensel Motor"
    If ogrenilenSayi > 0 And Not runQuiet Then
        On Error Resume Next
        ThisWorkbook.Sheets("OGRENME").Activate
        ThisWorkbook.Sheets("OGRENME").Range("D2").Select
        On Error GoTo 0
    ElseIf reconOut > 0 And Not runQuiet And reconSheet <> "" Then
        ' tolerans dýþýndaki ilk dosyanýn fark hücresine git
        On Error Resume Next
        ThisWorkbook.Sheets(reconSheet).Activate
        ThisWorkbook.Sheets(reconSheet).Cells(reconRow, 10).Select
        On Error GoTo 0
    End If
    forcedRunMode = 0
    Exit Sub

Cikis:
    GoTo SafeExit

ErrorHandler:
    If Err.Number = 18 Then
        MsgBox "Sistem kullanýcý tarafýndan güvenli bir þekilde DURDURULDU!", vbExclamation, "Güvenli Çýkýþ"
    Else
        MsgBox "Kritik Sistem Hatasý: " & Err.Description, vbCritical, "Çökme Korumasý (Garbage Collector)"
    End If
    GoTo SafeExit

SafeExit:
    On Error Resume Next
    Call RemoveProgress
    Call ProtectAfterMacro
    Call ResetExcelState
    
    If Not dictPosAngle Is Nothing Then Set dictPosAngle = Nothing
    If Not dictPosPlate Is Nothing Then Set dictPosPlate = Nothing
    If Not dictBoltQty Is Nothing Then Set dictBoltQty = Nothing
    If Not dictBoltList Is Nothing Then Set dictBoltList = Nothing
    If Not dictDiameters Is Nothing Then Set dictDiameters = Nothing
    If Not dictBoltWeights Is Nothing Then Set dictBoltWeights = Nothing
    If Not dictProfileLib Is Nothing Then Set dictProfileLib = Nothing
    If Not dictSelectedFiles Is Nothing Then Set dictSelectedFiles = Nothing
    If Not dictFeedback Is Nothing Then Set dictFeedback = Nothing
    
    Set fso = Nothing
    Set globalRegBolt = Nothing
    Set qualRegexGlobal = Nothing

    Application.StatusBar = False
End Sub

Sub Paneli_Ac()
    ' Bu makroyu Excel'deki düðmenize atayýn
    UserForm1.Show
End Sub

' =========================================================================
' SUM SAYFASINDAKÝ DÜÐMELERE ATANACAK ÇAÐIRICILAR
' =========================================================================

Sub Dugme_Sifirdan_Basla()
    usrAppendMode = False
    UserForm1.OptionButton1.Value = True
    UserForm1.OptionButton2.Value = False
    UserForm1.Show
End Sub

Sub Dugme_Ustune_Ekle()
    usrAppendMode = True
    UserForm1.OptionButton1.Value = False
    UserForm1.OptionButton2.Value = True
    UserForm1.Show
End Sub

Sub Dugme_Temizle()
    Call Can_Temizleyici
End Sub

Sub Can_Temizleyici()
    Dim wsAngle As Worksheet, wsPlate As Worksheet, wsBolt As Worksheet, wsSum As Worksheet
    On Error Resume Next
    Set wsAngle = ThisWorkbook.Sheets("ANGLE")
    Set wsPlate = ThisWorkbook.Sheets("PLATE")
    Set wsBolt = ThisWorkbook.Sheets("BOLTS&WASHER")
    Set wsSum = ThisWorkbook.Sheets("SUM")
    On Error GoTo 0
    
    If wsAngle Is Nothing Or wsPlate Is Nothing Or wsBolt Is Nothing Or wsSum Is Nothing Then Exit Sub
    If MsgBox("Gömülü formülleriniz ve AZ sütunu korunarak veriler temizlenecek." & vbCrLf & _
              "AUDIT, FEEDBACK ve OGRENME sayfalarý da temizlenecek. Onaylýyor musunuz?", vbYesNo + vbQuestion, "Can Temizleyici") = vbNo Then Exit Sub
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Call UnprotectForMacro
    Call CleanTemplateRanges(wsAngle, wsPlate, wsBolt, wsSum)
    Call ClearReportSheets
    ' REV / DATE sýfýrla
    wsSum.Range("G1").Value = "REV:       0"
    wsSum.Range("G2").Value = "DATE: 00.00." & Year(Date)
    Call ProtectAfterMacro
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    
    wsSum.Activate
    wsSum.Range("A5").Select
    MsgBox "Temizlik tamamlandý! AUDIT, FEEDBACK ve OGRENME sayfalarý da temizlendi, REV sýfýrlandý." & vbCrLf & _
           "AZ sütunundaki formülünüz ve þablon formülleriniz korundu.", vbInformation, "Can Temizleyici"
End Sub

' OGRENME sayfasýndaki parçalarý TÜR'üne göre aktarýr:
'   ANGLE        -> L-U-I-O-Y-LIBRARY (A kod, B cins, D kg/m, E yüzey alan)
'   PLATE        -> kütüphane yok; kalýnlýk / geniþlik OGRENILEN'de tutulur
'   BOLTS&WASHER -> BOLT-LIBRARY (A kod, B cins, C birim aðýrlýk, D 1000 adet aðýrlýðý)
' Her aktarýlan parçanýn türü OGRENILEN sayfasýna yazýlýr; sonra son iþlem otomatik yenilenir.
Sub Kutuphaneye_Aktar()
    Dim ws As Worksheet, wsLib As Worksheet, wsBoltLib As Worksheet
    Dim dictLib As Object, dictBolt As Object, dictNew As Object
    Dim r As Long, lastR As Long, libRow As Long, boltRow As Long, i As Long
    Dim prof As String, code As String, cins As String, tp As String, eksik As String, durum As String, key As String
    Dim v1 As Double, v2 As Double, hitRow As Long, warnTxt As String
    Dim nAdded As Long, nSkipped As Long, nExists As Long, nMissing As Long, nLearned As Long, nWarned As Long

    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("OGRENME")
    Set wsLib = ThisWorkbook.Sheets("L-U-I-O-Y-LIBRARY")
    Set wsBoltLib = ThisWorkbook.Sheets("BOLT-LIBRARY")
    On Error GoTo 0
    If ws Is Nothing Or wsLib Is Nothing Then
        MsgBox "OGRENME veya L-U-I-O-Y-LIBRARY sayfasý bulunamadý.", vbCritical
        Exit Sub
    End If
    Call UnprotectForMacro

    ' Profil kütüphanesi: ad (kanonik) -> satýr
    Set dictLib = CreateObject("Scripting.Dictionary")
    dictLib.CompareMode = 1
    For r = 2 To wsLib.Cells(wsLib.Rows.Count, 2).End(xlUp).Row
        key = CanonicalSectionKey(wsLib.Cells(r, 2).Text)
        If key <> "" And Not dictLib.Exists(key) Then dictLib.Add key, r
    Next r
    libRow = LastUsedRow(wsLib, 1, 2) + 1

    ' Cývata kütüphanesi: kod ve ad -> satýr (veriler 3. satýrdan baþlar)
    Set dictBolt = CreateObject("Scripting.Dictionary")
    dictBolt.CompareMode = 1
    boltRow = 3
    If Not wsBoltLib Is Nothing Then
        For r = 3 To wsBoltLib.Cells(wsBoltLib.Rows.Count, 1).End(xlUp).Row
            For i = 1 To 2
                key = UCase$(Replace(Trim$(wsBoltLib.Cells(r, i).Text), " ", ""))
                If key <> "" And Not dictBolt.Exists(key) Then dictBolt.Add key, r
            Next i
        Next r
        boltRow = Application.WorksheetFunction.Max(3, LastUsedRow(wsBoltLib, 1, 2) + 1)
    End If

    lastR = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For r = 2 To lastR
        prof = Trim$(ws.Cells(r, 1).Text)
        If prof = "" Or Left$(CStr(ws.Cells(r, 8).Value), 9) = "AKTARILDI" Then GoTo NextUnk

        tp = OgrenmeTypeCode(ws.Cells(r, 3).Text)
        code = UCase$(Trim$(ws.Cells(r, 4).Text))
        cins = Trim$(ws.Cells(r, 5).Text)
        If cins = "" Then cins = prof
        v1 = CellNumber(ws.Cells(r, 6))
        v2 = CellNumber(ws.Cells(r, 7))
        ws.Cells(r, 8).Interior.ColorIndex = xlColorIndexNone

        ' Hiçbir þey girilmemiþse: kullanýcý tanýmlamak istemiyor
        If code = "" And v1 <= 0 And v2 <= 0 Then
            ws.Cells(r, 8).Value = "Boþ - atlandý"
            nSkipped = nSkipped + 1
            GoTo NextUnk
        End If
        If tp = "" Then
            ws.Cells(r, 8).Value = "TÜR seçin (ANGLE / PLATE / BOLTS&WASHER)"
            ws.Cells(r, 8).Interior.Color = RGB(255, 199, 206)
            nMissing = nMissing + 1
            GoTo NextUnk
        End If
        If tp = "BOLT" And wsBoltLib Is Nothing Then
            ws.Cells(r, 8).Value = "BOLT-LIBRARY sayfasý yok"
            ws.Cells(r, 8).Interior.Color = RGB(255, 199, 206)
            nMissing = nMissing + 1
            GoTo NextUnk
        End If

        ' Cývatada iki aðýrlýktan biri yeterli: birim aðýrlýk = 1000 adet / 1000
        If tp = "BOLT" Then
            If v1 <= 0 And v2 > 0 Then v1 = v2 / 1000#
            If v2 <= 0 And v1 > 0 Then v2 = v1 * 1000#
        End If

        eksik = ""
        Select Case tp
            Case "ANGLE"
                If code = "" Then eksik = eksik & ", MALZEME KODU"
                If v1 <= 0 Then eksik = eksik & ", kg/m"
                If v2 <= 0 Then eksik = eksik & ", YÜZEY ALAN"
            Case "PLATE"
                If v1 <= 0 Then eksik = eksik & ", KALINLIK"
                If v2 <= 0 Then eksik = eksik & ", GENÝÞLÝK"
            Case "BOLT"
                If code = "" Then eksik = eksik & ", MALZEME KODU"
                If v1 <= 0 Then eksik = eksik & ", BÝRÝM AÐIRLIK"
        End Select
        If eksik <> "" Then
            ws.Cells(r, 8).Value = "EKSÝK: " & Mid$(eksik, 3)
            ws.Cells(r, 8).Interior.Color = RGB(255, 199, 206)
            nMissing = nMissing + 1
            GoTo NextUnk
        End If

        durum = ""
        Select Case tp
            Case "PLATE"
                ' plakanýn kütüphanesi yok: ölçüler OGRENILEN'de
                durum = "AKTARILDI (PLATE " & NumToText(v1) & "x" & NumToText(v2) & ")"
                nAdded = nAdded + 1

            Case "BOLT"
                hitRow = 0
                key = UCase$(Replace(code, " ", ""))
                If dictBolt.Exists(key) Then hitRow = dictBolt(key)
                If hitRow > 0 Then
                    durum = "AKTARILDI (kod kütüphanede vardý, satýr " & hitRow & ")"
                    nExists = nExists + 1
                Else
                    wsBoltLib.Cells(boltRow, 1).Value = LibraryCodeValue(code)
                    wsBoltLib.Cells(boltRow, 2).Value = cins
                    wsBoltLib.Cells(boltRow, 3).Value = v1
                    wsBoltLib.Cells(boltRow, 4).Value = v2
                    wsBoltLib.Cells(boltRow, 1).Resize(1, 4).Interior.Color = RGB(198, 239, 206)
                    dictBolt(UCase$(Replace(code, " ", ""))) = boltRow
                    dictBolt(UCase$(Replace(cins, " ", ""))) = boltRow
                    boltRow = boltRow + 1
                    durum = "AKTARILDI"
                    nAdded = nAdded + 1
                End If

            Case "ANGLE"
                hitRow = 0
                key = CanonicalSectionKey(cins)
                If dictLib.Exists(key) Then hitRow = dictLib(key)
                key = CanonicalSectionKey(prof)
                If hitRow = 0 And dictLib.Exists(key) Then hitRow = dictLib(key)
                If hitRow > 0 Then
                    ' Eski "YENI_EKLE"/"TANIMSIZ" satýrý varsa 4 veriyi o satýra yaz; kodlu satýr varsa onun kodu kullanýlýr
                    If UCase$(Trim$(wsLib.Cells(hitRow, 1).Text)) = "YENI_EKLE" Or _
                       UCase$(Trim$(wsLib.Cells(hitRow, 1).Text)) = "TANIMSIZ" Or _
                       Trim$(wsLib.Cells(hitRow, 1).Text) = "" Then
                        Call WriteProfileLibRow(wsLib, hitRow, code, wsLib.Cells(hitRow, 2).Value, v1, v2)
                        durum = "AKTARILDI"
                        nAdded = nAdded + 1
                    Else
                        code = UCase$(Trim$(wsLib.Cells(hitRow, 1).Text))
                        durum = "AKTARILDI (kütüphanede vardý, kod " & code & ")"
                        nExists = nExists + 1
                    End If
                Else
                    Call WriteProfileLibRow(wsLib, libRow, code, cins, v1, v2)
                    dictLib(CanonicalSectionKey(cins)) = libRow
                    libRow = libRow + 1
                    durum = "AKTARILDI"

                    ' Cins adý deðiþtirildiyse listedeki ham ad bu satýrla eþleþiyor mu? Eþleþmiyorsa
                    ' ham adý da ayný 4 veriyle ekle.
                    Set dictNew = CreateObject("Scripting.Dictionary")
                    dictNew.CompareMode = 1
                    dictNew(CanonicalSectionKey(cins)) = code
                    dictNew(CanonicalSectionKey(code)) = code
                    dictNew(UCase$(Replace(cins, " ", ""))) = code
                    If StandardizeProfileName(cins) <> "" Then dictNew(StandardizeProfileName(cins)) = code
                    If Not ProfileKnown(UCase$(prof), dictNew) Then
                        Call WriteProfileLibRow(wsLib, libRow, code, prof, v1, v2)
                        dictLib(CanonicalSectionKey(prof)) = libRow
                        libRow = libRow + 1
                        durum = "AKTARILDI (+ ham ad da eklendi)"
                    End If
                    nAdded = nAdded + 1
                End If
        End Select

        ' Türü kaydet: bir sonraki iþlemde parça adýndan tanýnýr ve doðru sayfaya yazýlýr
        Call SaveLearnedItem(prof, tp, code, cins, v1, v2)
        nLearned = nLearned + 1
        ' Girilen kg/m / yüzey alan profil ölçüsüyle uyumlu mu? (aktarýlýr ama sarý uyarý)
        warnTxt = ""
        If tp = "ANGLE" Then
            warnTxt = SectionValueWarning(cins, v1, v2)
            If warnTxt = "" And UCase$(cins) <> UCase$(prof) Then warnTxt = SectionValueWarning(prof, v1, v2)
        End If
        If warnTxt <> "" Then
            ws.Cells(r, 8).Value = durum & " | KONTROL EDÝN: " & warnTxt
            ws.Cells(r, 8).Interior.Color = RGB(255, 235, 156)
            nWarned = nWarned + 1
        Else
            ws.Cells(r, 8).Value = durum
            ws.Cells(r, 8).Interior.Color = RGB(198, 239, 206)
        End If
NextUnk:
    Next r

    Call ProtectAfterMacro
    MsgBox nAdded & " parça kütüphaneye eklendi, " & nExists & " parçanýn kodu zaten vardý." & vbCrLf & _
           nLearned & " parçanýn türü OGRENILEN sayfasýna kaydedildi." & vbCrLf & _
           nSkipped & " parça boþ býrakýldý." & vbCrLf & _
           IIf(nWarned > 0, nWarned & " parçanýn kg/m / yüzey alaný ölçüden hesaplanandan çok farklý (DURUM'da sarý)." & vbCrLf, "") & _
           IIf(nMissing > 0, nMissing & " parça EKSÝK veri nedeniyle aktarýlmadý (DURUM sütununa bakýn)." & vbCrLf, ""), _
           IIf(nMissing > 0, vbExclamation, vbInformation), "Öðrenen Kütüphane"
    If nLearned > 0 Then Call Son_Islemi_Yenile
End Sub

' Kütüphane saðlýðý: iki kütüphaneyi tarar, sorunlarý KUTUPHANE_KONTROL sayfasýna yazar (deðiþtirmez)
Sub Kutuphane_Kontrol()
    Call RunLibraryCheck
End Sub

' Son iþlemi (ayný dosyalar, ayný mod) dosya seçmeden yeniden yapar.
' ÜSTÜNE EKLE modunda önce son iþlemden önceki yedek geri yüklenir (aksi halde adetler iki kez eklenirdi).
Sub Son_Islemi_Yenile()
    Dim ws As Worksheet, r As Long, n As Long, files() As String, nMiss As Long, missTxt As String
    Dim md As String, bkp As String, msg As String, fso2 As Object, p As String
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("SON_CALISMA")
    On Error GoTo 0
    If ws Is Nothing Then
        MsgBox "Kayýtlý son iþlem yok. Listeyi Eklemeye Baþla / Liste Üstüne Ekle ile iþleyin.", vbInformation
        Exit Sub
    End If
    md = UCase$(Trim$(ws.Range("B1").Text))
    bkp = Trim$(ws.Range("B4").Text)
    Set fso2 = CreateObject("Scripting.FileSystemObject")
    ReDim files(1 To 1)
    For r = 7 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        p = Trim$(ws.Cells(r, 1).Text)
        If p <> "" Then
            If fso2.FileExists(p) Then
                n = n + 1
                ReDim Preserve files(1 To n)
                files(n) = p
            Else
                nMiss = nMiss + 1
                If nMiss <= 5 Then missTxt = missTxt & vbCrLf & "  - " & p
            End If
        End If
    Next r
    If n = 0 Then
        MsgBox "Son iþlemin dosyalarý bulunamadý; listeyi elle yeniden iþleyin." & missTxt, vbExclamation
        Exit Sub
    End If

    msg = "Son iþlem ayný " & n & " dosya ile yeniden yapýlsýn mý?" & vbCrLf & _
          "Mod: " & IIf(md = "USTUNE", "LÝSTE ÜSTÜNE EKLE", "SIFIRDAN") & vbCrLf
    If nMiss > 0 Then msg = msg & vbCrLf & "Bulunamayan " & nMiss & " dosya atlanacak:" & missTxt & vbCrLf
    If md = "USTUNE" Then
        If bkp = "" Or bkp = "-" Or Not fso2.FileExists(bkp) Then
            MsgBox "Son iþlem LÝSTE ÜSTÜNE EKLE modundaydý ama iþlem öncesi yedek bulunamadý." & vbCrLf & _
                   "Adetler iki kez eklenmesin diye otomatik yenileme yapýlmadý; listeyi elle yeniden iþleyin.", vbExclamation
            Exit Sub
        End If
        msg = msg & vbCrLf & "ANGLE / PLATE / BOLTS&WASHER / SUM, son iþlemden önceki yedekten geri yüklenecek:" & vbCrLf & _
              "  " & bkp & vbCrLf & "(o iþlemden sonra bu sayfalarda elle yaptýðýnýz deðiþiklikler kaybolur)"
    End If
    If MsgBox(msg, vbYesNo + vbQuestion, "Son Ýþlemi Yenile") = vbNo Then Exit Sub

    If md = "USTUNE" Then
        Application.ScreenUpdating = False
        Call UnprotectForMacro
        If Not RestoreOutputsFromBackup(bkp) Then
            Call ProtectAfterMacro
            Application.ScreenUpdating = True
            MsgBox "Yedek geri yüklenemedi: " & bkp & vbCrLf & "Listeyi elle yeniden iþleyin.", vbCritical
            Exit Sub
        End If
        Call ProtectAfterMacro
    End If

    usrAppendMode = (md = "USTUNE")
    usrExtraPct = 0
    If IsNumeric(ws.Range("B2").Value) Then usrExtraPct = CDbl(ws.Range("B2").Value)
    usrIsMerge = (Trim$(ws.Range("B3").Text) = "1")
    runPresetFiles = files
    Call Evrensel_BOM_Cevirici_Core
End Sub


' L-U-I-O-Y-LIBRARY satýrýna 4 veriyi yazar: A kod, B cins, D kg/m, E yüzey alan
Private Sub WriteProfileLibRow(ByVal wsLib As Worksheet, ByVal r As Long, ByVal code As String, _
                               ByVal cins As String, ByVal kgm As Double, ByVal alan As Double)
    wsLib.Cells(r, 1).Value = LibraryCodeValue(code)
    wsLib.Cells(r, 2).Value = cins
    wsLib.Cells(r, 4).Value = kgm
    wsLib.Cells(r, 5).Value = alan
    wsLib.Cells(r, 1).Resize(1, 5).Interior.Color = RGB(198, 239, 206)
End Sub

' Verilen sütunlardaki son dolu satýr (A boþ, B dolu satýrlar da sayýlýr)
Private Function LastUsedRow(ByVal ws As Worksheet, ByVal col1 As Long, ByVal col2 As Long) As Long
    Dim a As Long, b As Long
    a = ws.Cells(ws.Rows.Count, col1).End(xlUp).Row
    b = ws.Cells(ws.Rows.Count, col2).End(xlUp).Row
    LastUsedRow = IIf(a > b, a, b)
End Function


' Çalýþma kitabý açýlýnca (ThisWorkbook > Workbook_Open) çaðrýlýr
Public Sub BOM_OnOpen()
    On Error Resume Next
    Call UnprotectForMacro
    Call FixButtonMacros
    Call ProtectAfterMacro
End Sub

' Düðmeler "Module6.Can_Temizleyici" gibi ESKÝ modül adýyla atanmýþsa,
' modül adýný atýp makroyu yeniden baðlar (dosya her açýldýðýnda kontrol edilir).
Public Sub FixButtonMacros()
    Dim ws As Worksheet, shp As Shape, act As String, nm As String, p As Long, known As String
    known = "|DUGME_SIFIRDAN_BASLA|DUGME_USTUNE_EKLE|DUGME_TEMIZLE|CAN_TEMIZLEYICI|PANELI_AC|KUTUPHANEYE_AKTAR|EVRENSEL_BOM_CEVIRICI_CORE|SON_ISLEMI_YENILE|TEST_CALISTIR|TEST_BEKLENEN_KAYDET|KUTUPHANE_KONTROL|"
    On Error Resume Next
    For Each ws In ThisWorkbook.Worksheets
        For Each shp In ws.Shapes
            act = ""
            act = shp.OnAction
            If act <> "" Then
                nm = act
                p = InStrRev(nm, "!")
                If p > 0 Then nm = Mid$(nm, p + 1)       ' 'Kitap.xlsm'!Makro
                p = InStrRev(nm, ".")
                If p > 0 Then nm = Mid$(nm, p + 1)       ' Module6.Makro
                nm = Replace(nm, "'", "")
                If InStr(known, "|" & UCase$(nm) & "|") > 0 And nm <> act Then shp.OnAction = nm
            End If
        Next shp
    Next ws
End Sub

' =========================================================================
' KÜTÜPHANE BAKIMI: eski InputBox döneminden kalan "YENI_EKLE" / "TANIMSIZ" satýrlarýný siler.
' Silinen satýrlar önce SILINEN_KAYITLAR sayfasýna yedeklenir.
' =========================================================================
Sub Kutuphane_Temizle()
    Dim ws As Worksheet, wsLog As Worksheet, r As Long, lastR As Long, n As Long, outR As Long
    Dim cd As String, delRng As Range
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("L-U-I-O-Y-LIBRARY")
    On Error GoTo 0
    If ws Is Nothing Then MsgBox "L-U-I-O-Y-LIBRARY bulunamadý.", vbCritical: Exit Sub

    lastR = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
    For r = 2 To lastR
        cd = UCase$(Trim$(ws.Cells(r, 1).Text))
        If cd = "YENI_EKLE" Or cd = "TANIMSIZ" Then n = n + 1
    Next r
    If n = 0 Then MsgBox "Temizlenecek kayýt yok.", vbInformation: Exit Sub
    If MsgBox(n & " adet 'YENI_EKLE' / 'TANIMSIZ' satýrý silinecek." & vbCrLf & _
              "Silinenler önce 'SILINEN_KAYITLAR' sayfasýna yedeklenecek. Devam edilsin mi?", _
              vbYesNo + vbQuestion, "Kütüphane Temizle") = vbNo Then Exit Sub

    Application.ScreenUpdating = False
    Call UnprotectForMacro
    On Error Resume Next
    Set wsLog = ThisWorkbook.Sheets("SILINEN_KAYITLAR")
    On Error GoTo 0
    If wsLog Is Nothing Then
        Set wsLog = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsLog.Name = "SILINEN_KAYITLAR"
        ws.Range("A1").Resize(1, 15).Copy wsLog.Range("B1")
        wsLog.Range("A1").Value = "SÝLÝNME TARÝHÝ"
        wsLog.Rows(1).Font.Bold = True
    End If
    outR = wsLog.Cells(wsLog.Rows.Count, 2).End(xlUp).Row + 1

    For r = 2 To lastR
        cd = UCase$(Trim$(ws.Cells(r, 1).Text))
        If cd = "YENI_EKLE" Or cd = "TANIMSIZ" Then
            wsLog.Cells(outR, 1).Value = Now
            wsLog.Cells(outR, 2).Resize(1, 15).Value = ws.Cells(r, 1).Resize(1, 15).Value
            outR = outR + 1
            If delRng Is Nothing Then Set delRng = ws.Rows(r) Else Set delRng = Union(delRng, ws.Rows(r))
        End If
    Next r
    If Not delRng Is Nothing Then delRng.Delete
    Call ProtectAfterMacro
    Application.ScreenUpdating = True
    MsgBox n & " satýr silindi ve SILINEN_KAYITLAR sayfasýna yedeklendi.", vbInformation, "Kütüphane Temizle"
End Sub

