Attribute VB_Name = "modState"
Option Explicit
Option Private Module

' =========================================================================
' ORTAK DURUM: tüm modüllerin paylaþtýðý deðiþkenler ve sabitler
' =========================================================================

' Tüm modüllerin ortak durumu (buffer'lar, sözlükler, parametreler, motor paylaþýmý)

Public usrExtraPct As Double
Public usrIsMerge As Boolean
Public usrAppendMode As Boolean
Public bufAngle() As Variant
Public bufPlate() As Variant
Public bufBolt() As Variant
Public bufSum() As Variant
Public curWsAngleRef As Worksheet
Public curWsPlateRef As Worksheet
Public curWsBoltRef As Worksheet
Public curWsSumRef As Worksheet
Public hwKeywords() As String
Public garbageKeywords() As String
Public plateKeywords() As String
Public angleKeywords() As String
Public globalRegBolt As Object
Public auditWs As Worksheet
Public auditNextRow As Long
Public auditProcessed As Long
Public auditExact As Long
Public auditWarning As Long
Public auditError As Long
Public auditUnmatched As Long
Public auditSkipped As Long
Public auditFiles As Object
Public forcedRunMode As Long
Public qualRegexGlobal As Object
Public rxHelper As Object
Public Const CODE_VERSION As String = "v3.5"
Public prmStepBoltAutoSet As Boolean   ' Basamak cývatasý için otomatik somun/pul/rondela seti
Public prmWeightTolPct As Double       ' SUM aðýrlýk farký toleransý (%)
Public prmQualH As String              ' Kalite yoksa, "H" iþaretli parça
Public prmQualDefault As String        ' Kalite yoksa, diðer
Public prmBoltGrade As String          ' Cývata kalitesi yazmýyorsa
Public prmBoltLen As Long              ' Cývata boyu yazmýyorsa (mm)
Public prmBackup As Boolean            ' Ýþlem öncesi otomatik yedek
Public prmBackupFolder As String       ' Yedek klasörü (boþ = otomatik)
Public prmBackupKeep As Long           ' Saklanacak yedek sayýsý
Public prmRevAuto As Boolean           ' Üstüne eklemede REV otomatik artsýn
Public runStart As Date
Public runMode As String
Public runBackupPath As String
Public runRev As String
Public auditCurrentPath As String
Public auditCurrentSheet As String

' --- Ana döngü ile motorlar (XSR / Excel) arasýnda paylaþýlan durum ---
Public currentColAngle As Long
Public currentColBolt As Long
Public currentColPlate As Long
Public curWsAngle As Worksheet
Public curWsPlate As Worksheet
Public dictBoltList As Object
Public dictBoltQty As Object
Public dictDiameters As Object
Public dictFeedback As Object
Public dictPosAngle As Object
Public dictPosPlate As Object
Public dictProfileLib As Object
Public ext As String
Public extraPct As Double
Public fileName As String
Public fileWeight As Double
Public fso As Object
Public isDryRun As Boolean
Public isMerge As Boolean
Public targetRowAngle As Long
Public targetRowPlate As Long
Public totalProcessed As Long
Public vItem As Variant

' --- Poz çakýþmasý sayacý ---
Public conflictCount As Long

' --- PDF motoru: geçici çalýþma kitabý (Power Query ile PDF'ten okunan tablolar) ---
Public pdfTempWb As Workbook

' --- Dosya seçme penceresi olmadan çalýþtýrma (son iþlemi yenile / test) ---
Public runPresetFiles As Variant
' --- Sessiz çalýþtýrma (test): son mesaj, yedek, geçmiþ ve REV yok ---
Public runQuiet As Boolean

' --- Öðrenilen parçalar (OGRENILEN sayfasý): ad -> Array(tür, kod, cins, deðer1, deðer2) ---
Public dictLearned As Object

' --- Kapasite: tampon dýþýna düþen yazmalar (sessiz satýr kaybýný raporlamak için) ---
Public bufOverflowMsg As String

' --- Bu çalýþtýrmada / SUM'da bulunan dosya iþaretleri (montaj çift sayým korumasý) ---
Public dictFileMarks As Object
