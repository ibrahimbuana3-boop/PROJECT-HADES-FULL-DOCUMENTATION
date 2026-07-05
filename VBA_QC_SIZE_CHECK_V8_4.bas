Option Explicit

'=========================================================
' HADES - QC SIZE CHECK V8.4 CORE BRIDGE + JACKET BODY-ONLY FILTER
' CorelDRAW 2021 VBA
'
' BASE:
' - QC SIZE CHECK V8.3 CORE BRIDGE
' - Phase 5 Core Foundation
'
' SUPPORT:
' 1. JERSEY  : 2 panel = 1 set
' 2. JAKET   : 3 panel = 1 set  (2 depan + 1 belakang)
' 3. CELANA  : 4 panel = 1 set
'
' SOURCE OF TRUTH:
' - Documents\Order.txt
' - Documents\SizeDB_*.txt
' - Red / Green outline pattern
'
' FITUR V8 DIPERTAHANKAN:
' - Membaca metadata @SIZEDB dari Order.txt
' - Auto pilih database
' - Fallback popup jika @SIZEDB tidak ada
' - Skip baris metadata @...
' - Support:
'   SizeDB_Pria.txt
'   SizeDB_Wanita.txt
'   SizeDB_Anak.txt
'   SizeDB_PriaSlimFit.txt
'   SizeDB_WanitaSlimFit.txt
'   SizeDB_Jaket.txt
'   SizeDB_JaketAnak.txt
'   SizeDB_CelanaPria.txt
'   SizeDB_CelanaWanita.txt
'   SizeDB_CelanaAnak.txt
'
' FIX BESAR V8.1 KHUSUS CELANA:
' - Rotation-safe:
'   Celana tidak lagi hanya membaca sisi terkecil.
'   VBA mengecek Width dan Height terhadap SizeDB.
'
' - Lebih tahan outline tambahan:
'   Outline merah kecil diabaikan dari matching celana.
'
' - Celana dihitung per panel global:
'   Tiap panel yang cocok langsung masuk size-nya.
'   Actual set tetap dihitung 4 panel = 1 set.
'
' - Debug unknown celana:
'   Jika ada outline merah besar tapi tidak cocok SizeDB,
'   dimensi W/H akan ditampilkan di report.
'
' MAIN MACRO:
' QC_SIZE_CHECK
'
' PHASE 5B CORE BRIDGE:
' - Path Documents memakai H5_OrderPath / H5_DocumentsFile
' - Read UTF-8 memakai H5_ReadTextUTF8
' - Normalize size memakai H5_NormalizeSizeKey
' - Numeric parsing memakai H5_ToDbl
' - Mode produk memakai H5_ProductModeFromDB
'
' V8.4 JACKET BODY-ONLY FILTER:
' - Mode Jaket tidak lagi menghitung semua outline merah yang cocok SizeDB.
' - Per top-level group, VBA mencari kombinasi 1 panel belakang + 2 panel depan.
' - Lengan/strip/manset diabaikan walaupun dimensinya mirip size lain.
'
' REPORT MODE:
' HADES_QC_SIZE_REPORT
' - Dipanggil oleh HADES_QC_FINAL_REPORT_V3A.
' - Tidak mengubah outline merah menjadi hijau.
' - Tidak menampilkan MsgBox hasil.
' - Mengirim status ke HADES_CORE_REPORT_PHASE2.
'=========================================================


'=========================
' GLOBAL VARIABLES
'=========================

Private Expected As Object
Private ActualPanels As Object
Private SizeDB As Object
Private OrderMeta As Object

Private CurrentDB As String
Private DBSource As String

Private isSplitFront As Boolean
Private isPants As Boolean

Private Const TOL As Double = 1#
Private Const PANTS_TOL As Double = 0.75

' Area minimum outline celana.
' Untuk mengabaikan red curve kecil, label, ornament, atau garis pendek.
Private Const PANTS_MIN_PANEL_AREA As Double = 80#

' Area minimum kandidat body jaket.
' Untuk mencegah strip kecil, manset, label, dan ornamen ikut dihitung.
Private Const JACKET_MIN_BODY_AREA As Double = 250#
Private Const JACKET_MAX_CANDIDATES As Long = 500

Private Const ORDER_FILE As String = "\Documents\Order.txt"

Private scanPanels As Long
Private matchedPanels As Long
Private unknownRedPanels As Long
Private ignoredSmallRedPanels As Long

Private unknownDimReport As String
Private unknownDimCount As Long
Private Const UNKNOWN_DIM_LIMIT As Long = 12

'=========================================================
' REPORT MODE STATE - PHASE 3A
'=========================================================
Private qscReportMode As Boolean
Private qscSuppressMark As Boolean

Private qscLastStatus As String
Private qscLastSummary As String
Private qscLastDetail As String
Private qscLastPass As Boolean



'=========================================================
' MAIN
'=========================================================
'=========================================================
' PUBLIC REPORT MODE - PHASE 3A
'=========================================================
Public Sub HADES_QC_SIZE_REPORT()

    On Error GoTo ERR_HANDLER

    qscReportMode = True
    qscSuppressMark = True

    qscLastStatus = ""
    qscLastSummary = ""
    qscLastDetail = ""
    qscLastPass = False

    Call QC_SIZE_CHECK

    If Trim$(qscLastStatus) = "" Then
        HADESR_AddResult _
            "SIZE & QUANTITY CHECK", _
            "FAIL", _
            "QC Size report mode tidak menghasilkan status.", _
            "Kemungkinan Order.txt / SizeDB gagal dibaca, selection kosong, atau QC_SIZE_CHECK berhenti sebelum ShowResult."
    Else
        HADESR_AddResult _
            "SIZE & QUANTITY CHECK", _
            qscLastStatus, _
            qscLastSummary, _
            qscLastDetail
    End If

SAFE_EXIT:
    qscReportMode = False
    qscSuppressMark = False
    Exit Sub

ERR_HANDLER:
    HADESR_AddResult _
        "SIZE & QUANTITY CHECK", _
        "FAIL", _
        "QC Size report mode error.", _
        "Error " & Err.Number & ": " & Err.Description

    Resume SAFE_EXIT

End Sub


Public Sub HADES5B_QC_SIZE_CORE_SMOKE_TEST()

    Dim reportText As String
    Dim failCount As Long
    Dim warnCount As Long
    Dim ok As Boolean

    On Error GoTo ERR_HANDLER

    ok = H5_RunCoreSelfTest(reportText, failCount, warnCount)
    H5_WriteCoreSelfTestReport reportText

    If ok Then
        MsgBox "QC Size V8.4 Core Bridge siap." & vbCrLf & vbCrLf & _
               "Core self-test: PASS" & vbCrLf & _
               "Macro HADES_QC_SIZE_REPORT memakai H5 core bridge dan Jacket Body-Only Filter V8.4.", _
               vbInformation, "HADES PHASE 5B"
    Else
        MsgBox "QC Size Core Bridge belum siap." & vbCrLf & vbCrLf & _
               "Core self-test: FAIL" & vbCrLf & _
               "Fail    : " & failCount & vbCrLf & _
               "Warning : " & warnCount & vbCrLf & vbCrLf & _
               "Buka HADES_CORE_SELF_TEST_LATEST.txt di Documents\HADES_REPORTS.", _
               vbCritical, "HADES PHASE 5B"
    End If

    Exit Sub

ERR_HANDLER:
    MsgBox "HADES5B QC SIZE CORE SMOKE TEST ERROR" & vbCrLf & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, _
           vbCritical, "HADES PHASE 5B"

End Sub



Sub QC_SIZE_CHECK()

    Dim oldUnit As Long
    Dim sr As ShapeRange
    Dim s As Shape
    Dim cmdStarted As Boolean

    On Error GoTo ErrHandler

    If ActiveSelection.Shapes.Count = 0 Then
        qscLastStatus = "FAIL"
        qscLastSummary = "Tidak ada objek dipilih."
        qscLastDetail = "Blok hasil layout terlebih dahulu."
        If Not qscReportMode Then
            MsgBox "Tidak ada objek dipilih." & vbCrLf & _
                   "Blok hasil layout terlebih dahulu.", _
                   vbExclamation, "QC SIZE CHECK"
        End If
        Exit Sub
    End If

    Set Expected = CreateObject("Scripting.Dictionary")
    Set ActualPanels = CreateObject("Scripting.Dictionary")
    Set SizeDB = CreateObject("Scripting.Dictionary")
    Set OrderMeta = CreateObject("Scripting.Dictionary")

    CurrentDB = ""
    DBSource = ""
    isSplitFront = False
    isPants = False

    scanPanels = 0
    matchedPanels = 0
    unknownRedPanels = 0
    ignoredSmallRedPanels = 0

    unknownDimReport = ""
    unknownDimCount = 0

    qscLastStatus = ""
    qscLastSummary = ""
    qscLastDetail = ""
    qscLastPass = False

    If Not qscReportMode Then
        qscSuppressMark = False
    End If

    cmdStarted = False

    oldUnit = ActiveDocument.Unit
    ActiveDocument.Unit = cdrCentimeter

    If Not LoadOrder() Then GoTo CleanExit

    '=====================================================
    ' DATABASE AUTO MODE
    '=====================================================
    If Len(Trim$(CurrentDB)) > 0 Then

        ConfigureModeFromDB
        DBSource = "AUTO dari Order.txt @SIZEDB"

    Else

        CurrentDB = InferDBFromMetadata()

        If Len(Trim$(CurrentDB)) > 0 Then
            ConfigureModeFromDB
            DBSource = "AUTO dari metadata spesifikasi Order.txt"
        Else
            If Not SelectDatabaseFallback() Then GoTo CleanExit
            DBSource = "MANUAL POPUP"
        End If

    End If

    If Not LoadDB() Then GoTo CleanExit

    Set sr = ActiveSelectionRange

    ActiveDocument.BeginCommandGroup "Hades QC Size Check V8.4"
    cmdStarted = True

    If isPants Then

        'V8.1:
        'Mode celana rotation-safe.
        'Scan semua panel secara global, bukan voting kaku per group.
        ScanPantsSelectionV81 sr

    ElseIf isSplitFront Then

        'V8.4:
        'Mode jaket body-only.
        'Scan per top-level group dan hanya hitung kombinasi 1 belakang + 2 depan.
        ScanJacketSelectionV84 sr

    Else

        'Jersey:
        'Scan semua panel merah/hijau secara recursive.
        For Each s In sr.Shapes
            ScanShape s
        Next s

    End If

    ActiveDocument.EndCommandGroup
    cmdStarted = False

    ShowResult

CleanExit:

    ActiveDocument.Unit = oldUnit
    ActiveWindow.Refresh
    Exit Sub

ErrHandler:

    On Error Resume Next

    If cmdStarted Then ActiveDocument.EndCommandGroup

    ActiveDocument.Unit = oldUnit
    ActiveWindow.Refresh

    qscLastStatus = "FAIL"
    qscLastSummary = "QC SIZE CHECK ERROR."
    qscLastDetail = "Error " & err.Number & ":" & vbCrLf & err.Description

    If Not qscReportMode Then
        MsgBox "QC SIZE CHECK ERROR" & vbCrLf & vbCrLf & _
               "Error " & err.Number & ":" & vbCrLf & err.Description, _
               vbCritical, "QC SIZE CHECK"
    End If

End Sub


'=========================================================
' DATABASE FALLBACK POPUP
'=========================================================
'=========================================================
' REPORT MODE FAIL HELPER
'=========================================================
Private Sub QSC_SetReportFail(ByVal summaryText As String, ByVal detailText As String)

    qscLastStatus = "FAIL"
    qscLastSummary = summaryText
    qscLastDetail = detailText

End Sub



Private Function SelectDatabaseFallback() As Boolean

    Dim a As String
    Dim b As String

    SelectDatabaseFallback = False

    isSplitFront = False
    isPants = False
    CurrentDB = ""

    a = InputBox( _
        "Order.txt belum memiliki @SIZEDB." & vbCrLf & vbCrLf & _
        "PILIH JENIS PRODUK:" & vbCrLf & vbCrLf & _
        "1 = JERSEY" & vbCrLf & _
        "2 = JAKET" & vbCrLf & _
        "3 = CELANA", _
        "QC SIZE CHECK")

    If Trim$(a) = "" Then Exit Function

    Select Case Trim$(a)

        Case "1"

            b = InputBox( _
                "PILIH DATABASE JERSEY:" & vbCrLf & vbCrLf & _
                "1 = PRIA REGULAR" & vbCrLf & _
                "2 = WANITA REGULAR" & vbCrLf & _
                "3 = ANAK" & vbCrLf & _
                "4 = PRIA SLIM FIT" & vbCrLf & _
                "5 = WANITA SLIM FIT", _
                "QC SIZE CHECK")

            Select Case Trim$(b)

                Case "1"
                    CurrentDB = "SizeDB_Pria.txt"

                Case "2"
                    CurrentDB = "SizeDB_Wanita.txt"

                Case "3"
                    CurrentDB = "SizeDB_Anak.txt"

                Case "4"
                    CurrentDB = "SizeDB_PriaSlimFit.txt"

                Case "5"
                    CurrentDB = "SizeDB_WanitaSlimFit.txt"

                Case Else
                    Exit Function

            End Select

            isSplitFront = False
            isPants = False

        Case "2"

            b = InputBox( _
                "PILIH DATABASE JAKET:" & vbCrLf & vbCrLf & _
                "1 = DEWASA" & vbCrLf & _
                "2 = ANAK", _
                "QC SIZE CHECK")

            Select Case Trim$(b)

                Case "1"
                    CurrentDB = "SizeDB_Jaket.txt"

                Case "2"
                    CurrentDB = "SizeDB_JaketAnak.txt"

                Case Else
                    Exit Function

            End Select

            isSplitFront = True
            isPants = False

        Case "3"

            b = InputBox( _
                "PILIH DATABASE CELANA:" & vbCrLf & vbCrLf & _
                "1 = CELANA PRIA" & vbCrLf & _
                "2 = CELANA WANITA" & vbCrLf & _
                "3 = CELANA ANAK", _
                "QC SIZE CHECK")

            Select Case Trim$(b)

                Case "1"
                    CurrentDB = "SizeDB_CelanaPria.txt"

                Case "2"
                    CurrentDB = "SizeDB_CelanaWanita.txt"

                Case "3"
                    CurrentDB = "SizeDB_CelanaAnak.txt"

                Case Else
                    Exit Function

            End Select

            isSplitFront = False
            isPants = True

        Case Else

            Exit Function

    End Select

    SelectDatabaseFallback = True

End Function


'=========================================================
' CONFIGURE MODE FROM DATABASE NAME
'=========================================================

Private Sub ConfigureModeFromDB()

    'Phase 5B:
    'Mode produk sekarang memakai core bersama agar aturan Jersey/Jaket/Celana
    'tidak tersebar berbeda-beda antar macro.
    isPants = False
    isSplitFront = False

    H5_ProductModeFromDB CurrentDB, isPants, isSplitFront

End Sub


Private Function InferDBFromMetadata() As String

    'Phase 5B:
    'Infer database dari metadata Order.txt memakai core bersama.
    If Not OrderMeta Is Nothing Then
        InferDBFromMetadata = H5_InferDBFromOrderMeta(OrderMeta)
    Else
        InferDBFromMetadata = ""
    End If

End Function


'=========================================================
' LOAD ORDER.TXT
'=========================================================

Private Function LoadOrder() As Boolean

    Dim path As String
    Dim allText As String
    Dim lines As Variant
    Dim i As Long
    Dim line As String
    Dim arr As Variant
    Dim sz As String

    On Error GoTo ErrHandler

    LoadOrder = False

    path = H5_OrderPath()

    If Dir(path) = "" Then
        MsgBox "Order.txt tidak ditemukan:" & vbCrLf & path, _
               vbCritical, "QC SIZE CHECK"
        Exit Function
    End If

    allText = H5_ReadTextUTF8(path)

    allText = Replace(allText, vbCrLf, vbLf)
    allText = Replace(allText, vbCr, vbLf)

    lines = Split(allText, vbLf)

    For i = LBound(lines) To UBound(lines)

        line = Trim$(CStr(lines(i)))

        If Len(line) = 0 Then GoTo NextLine

        If Left$(line, 1) = "@" Then
            ParseMetaLine line
            GoTo NextLine
        End If

        arr = Split(line, "|")

        'Format order wajib:
        'SIZE|NAMA|NOMOR|NICKNAME
        If UBound(arr) >= 3 Then

            sz = NormalizeSizeKey(CStr(arr(0)))

            If IsStandardSize(sz) Then

                If Not Expected.Exists(sz) Then
                    Expected.Add sz, 0
                End If

                Expected(sz) = CLng(Expected(sz)) + 1

            End If

        End If

NextLine:

    Next i

    If OrderMeta.Exists("SIZEDB") Then
        If Len(Trim$(CStr(OrderMeta("SIZEDB")))) > 0 Then
            CurrentDB = Trim$(CStr(OrderMeta("SIZEDB")))
        End If
    End If

    If Expected.Count = 0 Then
        MsgBox "Order.txt terbaca, tetapi tidak ada size valid." & vbCrLf & vbCrLf & _
               "Pastikan format data order tetap:" & vbCrLf & _
               "SIZE|NAMA|NOMOR|NICKNAME" & vbCrLf & vbCrLf & _
               "Baris metadata @... boleh ada dan akan diabaikan.", _
               vbCritical, "QC SIZE CHECK"
        Exit Function
    End If

    LoadOrder = True
    Exit Function

ErrHandler:

    MsgBox "Gagal membaca Order.txt." & vbCrLf & _
           "Error " & err.Number & ": " & err.Description, _
           vbCritical, "QC SIZE CHECK"

End Function


Private Sub ParseMetaLine(ByVal line As String)

    Dim p As Long
    Dim k As String
    Dim v As String

    line = Trim$(line)

    If Left$(line, 1) <> "@" Then Exit Sub

    p = InStr(1, line, "=", vbTextCompare)

    If p <= 2 Then Exit Sub

    k = Mid$(line, 2, p - 2)
    v = Mid$(line, p + 1)

    k = UCase$(Trim$(k))
    v = Trim$(v)

    If Len(k) = 0 Then Exit Sub

    If OrderMeta.Exists(k) Then
        OrderMeta(k) = v
    Else
        OrderMeta.Add k, v
    End If

End Sub


Private Function GetMeta(ByVal keyName As String) As String

    keyName = UCase$(Trim$(keyName))

    If OrderMeta Is Nothing Then
        GetMeta = ""
        Exit Function
    End If

    If OrderMeta.Exists(keyName) Then
        GetMeta = CStr(OrderMeta(keyName))
    Else
        GetMeta = ""
    End If

End Function


Private Function ReadTextFileUTF8(ByVal path As String) As String

    Dim stm As Object

    On Error GoTo FALLBACK

    Set stm = CreateObject("ADODB.Stream")

    stm.Type = 2
    stm.Charset = "utf-8"
    stm.Open
    stm.LoadFromFile path

    ReadTextFileUTF8 = stm.ReadText

    stm.Close

    Exit Function

FALLBACK:

    On Error Resume Next

    If Not stm Is Nothing Then stm.Close

    On Error GoTo 0

    Dim f As Integer
    Dim line As String
    Dim result As String

    f = FreeFile

    Open path For Input As #f

    Do Until EOF(f)
        Line Input #f, line
        result = result & line & vbLf
    Loop

    Close #f

    ReadTextFileUTF8 = result

End Function


'=========================================================
' LOAD SIZEDB
'=========================================================

Private Function LoadDB() As Boolean

    Dim path As String
    Dim f As Integer
    Dim line As String
    Dim arr As Variant
    Dim sz As String

    On Error GoTo ErrHandler

    LoadDB = False

    path = H5_DocumentsFile(CurrentDB)

    If Dir(path) = "" Then
        MsgBox "SizeDB tidak ditemukan:" & vbCrLf & path & vbCrLf & vbCrLf & _
               "Database dipilih dari: " & DBSource, _
               vbCritical, "QC SIZE CHECK"
        Exit Function
    End If

    f = FreeFile

    Open path For Input As #f

    Do Until EOF(f)

        Line Input #f, line
        line = Trim$(line)

        If Len(line) > 0 Then

            If Left$(line, 1) = "@" Then GoTo NextDBLine

            arr = Split(line, "|")

            If UBound(arr) >= 0 Then

                sz = NormalizeSizeKey(CStr(arr(0)))

                If IsStandardSize(sz) Then

                    If isPants Then

                        'Format celana:
                        'SIZE|L_DEPAN|L_BELAKANG
                        If UBound(arr) >= 2 Then
                            PutSizeDB sz, Array( _
                                ToDbl(arr(1)), _
                                ToDbl(arr(2)) _
                            )
                        End If

                    ElseIf isSplitFront Then

                        'Format jaket:
                        'SIZE|L_BELAKANG|L_DEPAN|T_DEPAN|T_BELAKANG
                        If UBound(arr) >= 4 Then
                            PutSizeDB sz, Array( _
                                ToDbl(arr(1)), _
                                ToDbl(arr(2)), _
                                ToDbl(arr(3)), _
                                ToDbl(arr(4)) _
                            )
                        End If

                    Else

                        'Format jersey:
                        'SIZE|LEBAR|TINGGI_DEPAN|TINGGI_BELAKANG
                        If UBound(arr) >= 3 Then
                            PutSizeDB sz, Array( _
                                ToDbl(arr(1)), _
                                ToDbl(arr(2)), _
                                ToDbl(arr(3)) _
                            )
                        End If

                    End If

                End If

            End If

        End If

NextDBLine:

    Loop

    Close #f

    If SizeDB.Count = 0 Then
        MsgBox "SizeDB terbaca, tetapi kosong / format tidak valid." & vbCrLf & vbCrLf & _
               "File: " & CurrentDB, _
               vbCritical, "QC SIZE CHECK"
        Exit Function
    End If

    LoadDB = True
    Exit Function

ErrHandler:

    On Error Resume Next
    Close #f

    MsgBox "Gagal membaca SizeDB." & vbCrLf & _
           "File: " & path & vbCrLf & vbCrLf & _
           "Error " & err.Number & ": " & err.Description, _
           vbCritical, "QC SIZE CHECK"

End Function


Private Sub PutSizeDB(ByVal sz As String, ByVal dataArr As Variant)

    If SizeDB.Exists(sz) Then
        SizeDB(sz) = dataArr
    Else
        SizeDB.Add sz, dataArr
    End If

End Sub


'=========================================================
' SCAN SHAPES - JERSEY / JAKET
'=========================================================

Private Sub ScanShape(ByVal s As Shape)

    Dim ch As Shape
    Dim detectedSize As String

    On Error Resume Next

    If s.Type = cdrGroupShape Then

        For Each ch In s.Shapes
            ScanShape ch
        Next ch

        Exit Sub

    End If

    ScanPowerClip s

    If s.Type <> cdrCurveShape Then Exit Sub

    If IsPanelOutline(s) Then

        scanPanels = scanPanels + 1

        detectedSize = DetectSize(s)

        If Len(detectedSize) > 0 Then

            matchedPanels = matchedPanels + 1
            AddActualPanel detectedSize, 1
            If Not qscSuppressMark Then MarkGreen s

        Else

            unknownRedPanels = unknownRedPanels + 1

        End If

    End If

End Sub


Private Sub ScanPowerClip(ByVal s As Shape)

    Dim pcShapes As Shapes
    Dim ch As Shape

    On Error Resume Next

    Set pcShapes = s.PowerClip.Shapes

    If Not pcShapes Is Nothing Then
        For Each ch In pcShapes
            ScanShape ch
        Next ch
    End If

End Sub



'=========================================================
' SCAN SHAPES - JAKET V8.4 BODY-ONLY FILTER
'=========================================================

Private Sub ScanJacketSelectionV84(ByVal sr As ShapeRange)

    Dim s As Shape
    Dim hasTopGroup As Boolean

    hasTopGroup = False

    'Jika selection berisi group jaket, proses tiap group sebagai 1 unit jaket.
    For Each s In sr.Shapes
        If s.Type = cdrGroupShape Then
            hasTopGroup = True
            ScanJacketUnitV84 s
        End If
    Next s

    'Fallback: jika operator menyeleksi komponen jaket yang sudah tidak dalam group,
    'proses seluruh selection sebagai 1 unit.
    If Not hasTopGroup Then
        ScanJacketRangeAsUnitV84 sr
    End If

End Sub


Private Sub ScanJacketRangeAsUnitV84(ByVal sr As ShapeRange)

    Dim s As Shape

    Dim candShape(1 To JACKET_MAX_CANDIDATES) As Shape
    Dim candSize(1 To JACKET_MAX_CANDIDATES) As String
    Dim candRole(1 To JACKET_MAX_CANDIDATES) As String
    Dim candErr(1 To JACKET_MAX_CANDIDATES) As Double

    Dim candCount As Long

    candCount = 0

    For Each s In sr.Shapes
        CollectJacketCandidatesRecursiveV84 s, candShape, candSize, candRole, candErr, candCount
    Next s

    AcceptBestJacketComboV84 candShape, candSize, candRole, candErr, candCount

End Sub


Private Sub ScanJacketUnitV84(ByVal unitShape As Shape)

    Dim candShape(1 To JACKET_MAX_CANDIDATES) As Shape
    Dim candSize(1 To JACKET_MAX_CANDIDATES) As String
    Dim candRole(1 To JACKET_MAX_CANDIDATES) As String
    Dim candErr(1 To JACKET_MAX_CANDIDATES) As Double

    Dim candCount As Long

    candCount = 0

    CollectJacketCandidatesRecursiveV84 unitShape, candShape, candSize, candRole, candErr, candCount

    AcceptBestJacketComboV84 candShape, candSize, candRole, candErr, candCount

End Sub


Private Sub CollectJacketCandidatesRecursiveV84( _
    ByVal s As Shape, _
    ByRef candShape() As Shape, _
    ByRef candSize() As String, _
    ByRef candRole() As String, _
    ByRef candErr() As Double, _
    ByRef candCount As Long)

    Dim ch As Shape
    Dim pcShapes As Shapes

    Dim detectedSize As String
    Dim detectedRole As String
    Dim bestErr As Double

    On Error Resume Next

    If s.Type = cdrGroupShape Then

        For Each ch In s.Shapes
            CollectJacketCandidatesRecursiveV84 ch, candShape, candSize, candRole, candErr, candCount
        Next ch

        Exit Sub

    End If

    Set pcShapes = s.PowerClip.Shapes

    If Not pcShapes Is Nothing Then
        For Each ch In pcShapes
            CollectJacketCandidatesRecursiveV84 ch, candShape, candSize, candRole, candErr, candCount
        Next ch
    End If

    If s.Type <> cdrCurveShape Then Exit Sub
    If Not IsPanelOutline(s) Then Exit Sub

    scanPanels = scanPanels + 1

    If JacketOutlineTooSmall(s) Then
        unknownRedPanels = unknownRedPanels + 1
        Exit Sub
    End If

    If DetectJacketBodyPanelFromShapeV84(s, detectedSize, detectedRole, bestErr) Then

        If candCount < JACKET_MAX_CANDIDATES Then
            candCount = candCount + 1
            Set candShape(candCount) = s
            candSize(candCount) = detectedSize
            candRole(candCount) = detectedRole
            candErr(candCount) = bestErr
        Else
            unknownRedPanels = unknownRedPanels + 1
        End If

    Else

        unknownRedPanels = unknownRedPanels + 1

    End If

End Sub


Private Function JacketOutlineTooSmall(ByVal s As Shape) As Boolean

    JacketOutlineTooSmall = False

    If Abs(s.SizeWidth * s.SizeHeight) < JACKET_MIN_BODY_AREA Then
        JacketOutlineTooSmall = True
    End If

End Function


Private Function DetectJacketBodyPanelFromShapeV84( _
    ByVal s As Shape, _
    ByRef detectedSize As String, _
    ByRef detectedRole As String, _
    ByRef bestErr As Double) As Boolean

    On Error GoTo SafeExit

    DetectJacketBodyPanelFromShapeV84 = DetectJacketBodyPanelFromDimensionsV84( _
        Abs(CDbl(s.SizeWidth)), _
        Abs(CDbl(s.SizeHeight)), _
        detectedSize, _
        detectedRole, _
        bestErr _
    )

    Exit Function

SafeExit:

    DetectJacketBodyPanelFromShapeV84 = False

End Function


Private Function DetectJacketBodyPanelFromDimensionsV84( _
    ByVal w As Double, _
    ByVal h As Double, _
    ByRef detectedSize As String, _
    ByRef detectedRole As String, _
    ByRef bestErr As Double) As Boolean

    Dim mn As Double
    Dim mx As Double

    Dim key As Variant
    Dim db As Variant

    Dim errBack As Double
    Dim errFront As Double

    detectedSize = ""
    detectedRole = ""
    bestErr = 999999#

    DetectJacketBodyPanelFromDimensionsV84 = False

    If w <= 0 Or h <= 0 Then Exit Function

    mn = MinD(w, h)
    mx = MaxD(w, h)

    For Each key In SizeDB.keys

        db = SizeDB(key)

        'Format jaket:
        'SIZE|L_BELAKANG|L_DEPAN|T_DEPAN|T_BELAKANG

        errBack = Abs(mn - CDbl(db(0))) + Abs(mx - CDbl(db(3)))

        If Abs(mn - CDbl(db(0))) <= TOL And Abs(mx - CDbl(db(3))) <= TOL Then
            If errBack < bestErr Then
                bestErr = errBack
                detectedSize = CStr(key)
                detectedRole = "BACK"
                DetectJacketBodyPanelFromDimensionsV84 = True
            End If
        End If

        errFront = Abs(mn - CDbl(db(1))) + Abs(mx - CDbl(db(2)))

        If Abs(mn - CDbl(db(1))) <= TOL And Abs(mx - CDbl(db(2))) <= TOL Then
            If errFront < bestErr Then
                bestErr = errFront
                detectedSize = CStr(key)
                detectedRole = "FRONT"
                DetectJacketBodyPanelFromDimensionsV84 = True
            End If
        End If

    Next key

End Function


Private Sub AcceptBestJacketComboV84( _
    ByRef candShape() As Shape, _
    ByRef candSize() As String, _
    ByRef candRole() As String, _
    ByRef candErr() As Double, _
    ByVal candCount As Long)

    Dim key As Variant

    Dim bestSize As String
    Dim bestBackIdx As Long
    Dim bestFront1Idx As Long
    Dim bestFront2Idx As Long
    Dim bestScore As Double

    Dim backIdx As Long
    Dim front1Idx As Long
    Dim front2Idx As Long
    Dim score As Double

    If candCount <= 0 Then Exit Sub

    bestSize = ""
    bestBackIdx = 0
    bestFront1Idx = 0
    bestFront2Idx = 0
    bestScore = 999999#

    For Each key In SizeDB.keys

        If FindJacketComboForSizeV84(CStr(key), candSize, candRole, candErr, candCount, backIdx, front1Idx, front2Idx, score) Then

            If score < bestScore Then
                bestScore = score
                bestSize = CStr(key)
                bestBackIdx = backIdx
                bestFront1Idx = front1Idx
                bestFront2Idx = front2Idx
            End If

        End If

    Next key

    If Len(bestSize) > 0 Then

        AddActualPanel bestSize, 3
        matchedPanels = matchedPanels + 3

        If Not qscSuppressMark Then
            If bestBackIdx > 0 Then MarkGreen candShape(bestBackIdx)
            If bestFront1Idx > 0 Then MarkGreen candShape(bestFront1Idx)
            If bestFront2Idx > 0 Then MarkGreen candShape(bestFront2Idx)
        End If

        'Kandidat lain dalam group dianggap komponen jaket non-body:
        'lengan, strip, manset, label, dll. Mereka sengaja tidak dihitung sebagai FAIL size lain.

    Else

        unknownRedPanels = unknownRedPanels + candCount

    End If

End Sub


Private Function FindJacketComboForSizeV84( _
    ByVal targetSize As String, _
    ByRef candSize() As String, _
    ByRef candRole() As String, _
    ByRef candErr() As Double, _
    ByVal candCount As Long, _
    ByRef backIdx As Long, _
    ByRef front1Idx As Long, _
    ByRef front2Idx As Long, _
    ByRef score As Double) As Boolean

    Dim i As Long

    backIdx = 0
    front1Idx = 0
    front2Idx = 0
    score = 999999#

    FindJacketComboForSizeV84 = False

    'Ambil 1 kandidat BACK dengan error paling kecil.
    For i = 1 To candCount

        If candSize(i) = targetSize And candRole(i) = "BACK" Then

            If backIdx = 0 Then
                backIdx = i
            ElseIf candErr(i) < candErr(backIdx) Then
                backIdx = i
            End If

        End If

    Next i

    If backIdx = 0 Then Exit Function

    'Ambil 2 kandidat FRONT dengan error paling kecil.
    For i = 1 To candCount

        If candSize(i) = targetSize And candRole(i) = "FRONT" Then

            If front1Idx = 0 Then

                front1Idx = i

            ElseIf candErr(i) < candErr(front1Idx) Then

                front2Idx = front1Idx
                front1Idx = i

            ElseIf front2Idx = 0 Then

                front2Idx = i

            ElseIf candErr(i) < candErr(front2Idx) Then

                front2Idx = i

            End If

        End If

    Next i

    If front1Idx = 0 Or front2Idx = 0 Then Exit Function

    score = candErr(backIdx) + candErr(front1Idx) + candErr(front2Idx)

    FindJacketComboForSizeV84 = True

End Function


'=========================================================
' SCAN SHAPES - CELANA V8.1 ROTATION SAFE
'=========================================================

Private Sub ScanPantsSelectionV81(ByVal sr As ShapeRange)

    Dim s As Shape

    For Each s In sr.Shapes
        ScanPantsPanelsRecursiveV81 s
    Next s

End Sub


Private Sub ScanPantsPanelsRecursiveV81(ByVal s As Shape)

    Dim ch As Shape
    Dim pcShapes As Shapes

    Dim detectedSize As String
    Dim bestErr As Double
    Dim usedSide As String

    On Error Resume Next

    If s.Type = cdrGroupShape Then

        For Each ch In s.Shapes
            ScanPantsPanelsRecursiveV81 ch
        Next ch

        Exit Sub

    End If

    Set pcShapes = s.PowerClip.Shapes

    If Not pcShapes Is Nothing Then
        For Each ch In pcShapes
            ScanPantsPanelsRecursiveV81 ch
        Next ch
    End If

    If s.Type <> cdrCurveShape Then Exit Sub

    If Not IsPanelOutline(s) Then Exit Sub

    scanPanels = scanPanels + 1

    If PantsOutlineTooSmall(s) Then
        ignoredSmallRedPanels = ignoredSmallRedPanels + 1
        Exit Sub
    End If

    If DetectPantsPanelFromShape(s, detectedSize, bestErr, usedSide) Then

        matchedPanels = matchedPanels + 1
        AddActualPanel detectedSize, 1
        If Not qscSuppressMark Then MarkGreen s

    Else

        unknownRedPanels = unknownRedPanels + 1
        AddUnknownPantsDimension s

    End If

End Sub


Private Function PantsOutlineTooSmall(ByVal s As Shape) As Boolean

    PantsOutlineTooSmall = False

    If Abs(s.SizeWidth * s.SizeHeight) < PANTS_MIN_PANEL_AREA Then
        PantsOutlineTooSmall = True
    End If

End Function


Private Function DetectPantsPanelFromShape( _
    ByVal s As Shape, _
    ByRef detectedSize As String, _
    ByRef bestErr As Double, _
    ByRef usedSide As String) As Boolean

    Dim w As Double
    Dim h As Double

    On Error GoTo SafeExit

    w = Abs(CDbl(s.SizeWidth))
    h = Abs(CDbl(s.SizeHeight))

    DetectPantsPanelFromShape = DetectPantsPanelFromDimensions( _
        w, _
        h, _
        detectedSize, _
        bestErr, _
        usedSide _
    )

    Exit Function

SafeExit:

    DetectPantsPanelFromShape = False

End Function


Private Function DetectPantsPanelFromDimensions( _
    ByVal w As Double, _
    ByVal h As Double, _
    ByRef detectedSize As String, _
    ByRef bestErr As Double, _
    ByRef usedSide As String) As Boolean

    Dim key As Variant
    Dim db As Variant

    Dim dbFront As Double
    Dim dbBack As Double

    Dim errWF As Double
    Dim errWB As Double
    Dim errHF As Double
    Dim errHB As Double

    Dim localErr As Double
    Dim localSide As String

    detectedSize = ""
    usedSide = ""
    bestErr = 999999#

    DetectPantsPanelFromDimensions = False

    If w <= 0 Or h <= 0 Then Exit Function

    For Each key In SizeDB.keys

        db = SizeDB(key)

        dbFront = CDbl(db(0))
        dbBack = CDbl(db(1))

        'Rotation-safe:
        'Cek Width dan Height terhadap L_DEPAN / L_BELAKANG.
        errWF = Abs(w - dbFront)
        errWB = Abs(w - dbBack)
        errHF = Abs(h - dbFront)
        errHB = Abs(h - dbBack)

        localErr = errWF
        localSide = "W~DEPAN"

        If errWB < localErr Then
            localErr = errWB
            localSide = "W~BELAKANG"
        End If

        If errHF < localErr Then
            localErr = errHF
            localSide = "H~DEPAN"
        End If

        If errHB < localErr Then
            localErr = errHB
            localSide = "H~BELAKANG"
        End If

        If localErr <= PANTS_TOL Then

            If localErr < bestErr Then
                bestErr = localErr
                detectedSize = CStr(key)
                usedSide = localSide
                DetectPantsPanelFromDimensions = True
            End If

        End If

    Next key

End Function


Private Sub AddUnknownPantsDimension(ByVal s As Shape)

    If unknownDimCount >= UNKNOWN_DIM_LIMIT Then Exit Sub

    unknownDimCount = unknownDimCount + 1

    unknownDimReport = unknownDimReport & _
        "- W=" & FormatNumber(Abs(s.SizeWidth), 3) & _
        " | H=" & FormatNumber(Abs(s.SizeHeight), 3) & _
        " | Area=" & FormatNumber(Abs(s.SizeWidth * s.SizeHeight), 2) & _
        vbCrLf

End Sub


'=========================================================
' OUTLINE DETECTION
'=========================================================

Private Function IsPanelOutline(ByVal s As Shape) As Boolean

    If IsRedOutline(s) Then
        IsPanelOutline = True
        Exit Function
    End If

    If IsGreenOutline(s) Then
        IsPanelOutline = True
        Exit Function
    End If

    IsPanelOutline = False

End Function


Private Function IsRedOutline(ByVal s As Shape) As Boolean

    Dim r As Long
    Dim g As Long
    Dim b As Long

    On Error GoTo SafeExit

    IsRedOutline = False

    If s.Outline.Width <= 0 Then Exit Function

    r = s.Outline.Color.RGBRed
    g = s.Outline.Color.RGBGreen
    b = s.Outline.Color.RGBBlue

    If r >= 230 And g <= 80 And b <= 80 Then
        IsRedOutline = True
    End If

SafeExit:

End Function


Private Function IsGreenOutline(ByVal s As Shape) As Boolean

    Dim r As Long
    Dim g As Long
    Dim b As Long

    On Error GoTo SafeExit

    IsGreenOutline = False

    If s.Outline.Width <= 0 Then Exit Function

    r = s.Outline.Color.RGBRed
    g = s.Outline.Color.RGBGreen
    b = s.Outline.Color.RGBBlue

    'Support hijau murni dari QC_SIZE_CHECK
    If r <= 80 And g >= 180 And b <= 80 Then
        IsGreenOutline = True
        Exit Function
    End If

    'Support hijau lama / hijau alternatif
    If Abs(r - 97) <= 25 And Abs(g - 186) <= 25 And Abs(b - 12) <= 25 Then
        IsGreenOutline = True
        Exit Function
    End If

SafeExit:

End Function


Private Sub MarkGreen(ByVal s As Shape)

    Dim c As Color

    On Error Resume Next

    Set c = CreateColor
    c.RGBAssign 0, 255, 0

    s.Outline.Color.CopyAssign c

End Sub


'=========================================================
' SIZE DETECTION - JERSEY / JAKET / FALLBACK
'=========================================================

Private Function DetectSize(ByVal s As Shape) As String

    On Error GoTo SafeExit

    DetectSize = DetectSizeFromDimensions( _
        Abs(CDbl(s.SizeWidth)), _
        Abs(CDbl(s.SizeHeight)) _
    )

SafeExit:

End Function


Private Function DetectSizeFromDimensions(ByVal w As Double, ByVal h As Double) As String

    Dim mn As Double
    Dim mx As Double

    Dim key As Variant
    Dim db As Variant

    Dim err As Double
    Dim errA As Double
    Dim errB As Double

    Dim bestSize As String
    Dim bestErr As Double

    mn = MinD(w, h)
    mx = MaxD(w, h)

    bestSize = ""
    bestErr = 999999#

    For Each key In SizeDB.keys

        db = SizeDB(key)

        If isPants Then

            'Fallback celana jika fungsi ini dipakai.
            'V8.1: cek sisi kecil dan besar.
            errA = MinD( _
                Abs(mn - CDbl(db(0))), _
                Abs(mx - CDbl(db(0))) _
            )

            errB = MinD( _
                Abs(mn - CDbl(db(1))), _
                Abs(mx - CDbl(db(1))) _
            )

            err = MinD(errA, errB)

            If err <= PANTS_TOL Then
                If err < bestErr Then
                    bestErr = err
                    bestSize = CStr(key)
                End If
            End If

        ElseIf isSplitFront Then

            'Jaket split-front:
            'SIZE|L_BELAKANG|L_DEPAN|T_DEPAN|T_BELAKANG

            'Panel belakang:
            errA = Abs(mn - CDbl(db(0))) + Abs(mx - CDbl(db(3)))

            If errA <= (TOL * 2) Then
                If errA < bestErr Then
                    bestErr = errA
                    bestSize = CStr(key)
                End If
            End If

            'Panel depan:
            errB = Abs(mn - CDbl(db(1))) + Abs(mx - CDbl(db(2)))

            If errB <= (TOL * 2) Then
                If errB < bestErr Then
                    bestErr = errB
                    bestSize = CStr(key)
                End If
            End If

        Else

            'Jersey normal:
            'SIZE|LEBAR|TINGGI_DEPAN|TINGGI_BELAKANG

            If Abs(mn - CDbl(db(0))) <= TOL Then

                errA = Abs(mx - CDbl(db(1)))
                errB = Abs(mx - CDbl(db(2)))

                err = Abs(mn - CDbl(db(0))) + MinD(errA, errB)

                If errA <= TOL Or errB <= TOL Then
                    If err < bestErr Then
                        bestErr = err
                        bestSize = CStr(key)
                    End If
                End If

            End If

        End If

    Next key

    DetectSizeFromDimensions = bestSize

End Function


'=========================================================
' ACTUAL COUNT
'=========================================================

Private Sub AddActualPanel(ByVal sz As String, ByVal amount As Long)

    sz = NormalizeSizeKey(sz)

    If amount <= 0 Then Exit Sub

    If Not ActualPanels.Exists(sz) Then
        ActualPanels.Add sz, 0
    End If

    ActualPanels(sz) = CLng(ActualPanels(sz)) + amount

End Sub


'=========================================================
' RESULT
'=========================================================

Private Sub ShowResult()

    Dim msg As String
    Dim passAll As Boolean

    Dim totalOrder As Long
    Dim totalLayout As Long

    Dim key As Variant
    Dim sz As Variant

    Dim expCount As Long
    Dim panelCount As Long
    Dim setCount As Long

    Dim allSizes As Object
    Dim orderedSizes As Variant

    Dim i As Long

    Set allSizes = CreateObject("Scripting.Dictionary")

    For Each key In Expected.keys
        If Not allSizes.Exists(CStr(key)) Then allSizes.Add CStr(key), True
    Next key

    For Each key In ActualPanels.keys
        If Not allSizes.Exists(CStr(key)) Then allSizes.Add CStr(key), True
    Next key

    totalOrder = 0

    For Each key In Expected.keys
        totalOrder = totalOrder + CLng(Expected(key))
    Next key

    totalLayout = 0

    For Each key In ActualPanels.keys
        totalLayout = totalLayout + PanelsToSetCount(CLng(ActualPanels(key)))
    Next key

    passAll = True

    msg = ""

    If totalOrder = totalLayout Then
        msg = msg & "QC PASSED" & vbCrLf
    Else
        msg = msg & "QC FAILED" & vbCrLf
        passAll = False
    End If

    msg = msg & String(45, "-") & vbCrLf
    msg = msg & "Database : " & CurrentDB & vbCrLf
    msg = msg & "DB Source: " & DBSource & vbCrLf
    msg = msg & "Mode     : " & ProductModeText() & vbCrLf
    msg = msg & "Core     : H5 Phase 5B bridge aktif" & vbCrLf
    msg = msg & String(45, "-") & vbCrLf

    If Len(GetMeta("JENIS_PESANAN")) > 0 Or _
       Len(GetMeta("JENIS_POLA")) > 0 Or _
       Len(GetMeta("MODEL_JAHIT")) > 0 Then

        msg = msg & "Metadata Order.txt:" & vbCrLf
        msg = msg & "Jenis Pesanan : " & GetMeta("JENIS_PESANAN") & vbCrLf
        msg = msg & "Jenis Pola    : " & GetMeta("JENIS_POLA") & vbCrLf
        msg = msg & "Model Jahit   : " & GetMeta("MODEL_JAHIT") & vbCrLf
        msg = msg & String(45, "-") & vbCrLf

    End If

    msg = msg & "TOTAL SET ORDER  : " & totalOrder & vbCrLf
    msg = msg & "TOTAL SET LAYOUT : " & totalLayout & vbCrLf
    msg = msg & String(45, "-") & vbCrLf

    orderedSizes = Array("XXS", "XS", "S", "M", "L", "XL", "2XL", "3XL", "4XL", "5XL", "6XL")

    For i = LBound(orderedSizes) To UBound(orderedSizes)

        sz = orderedSizes(i)

        If allSizes.Exists(CStr(sz)) Then

            expCount = 0
            panelCount = 0

            If Expected.Exists(CStr(sz)) Then expCount = CLng(Expected(CStr(sz)))
            If ActualPanels.Exists(CStr(sz)) Then panelCount = CLng(ActualPanels(CStr(sz)))

            setCount = PanelsToSetCount(panelCount)

            msg = msg & CStr(sz) & " : Expected " & expCount & _
                  " | Found " & setCount

            If expCount = setCount Then
                msg = msg & " | OK"
            Else
                msg = msg & " | FAIL"
                passAll = False
            End If

            msg = msg & vbCrLf

        End If

    Next i

    'Tampilkan size tambahan jika ada key di luar daftar standar.
    For Each key In allSizes.keys

        If Not IsStandardSize(CStr(key)) Then

            expCount = 0
            panelCount = 0

            If Expected.Exists(CStr(key)) Then expCount = CLng(Expected(CStr(key)))
            If ActualPanels.Exists(CStr(key)) Then panelCount = CLng(ActualPanels(CStr(key)))

            setCount = PanelsToSetCount(panelCount)

            msg = msg & CStr(key) & " : Expected " & expCount & _
                  " | Found " & setCount

            If expCount = setCount Then
                msg = msg & " | OK"
            Else
                msg = msg & " | FAIL"
                passAll = False
            End If

            msg = msg & vbCrLf

        End If

    Next key

    msg = msg & String(45, "-") & vbCrLf
    msg = msg & "Panel outlines scanned : " & scanPanels & vbCrLf
    msg = msg & "Matched panels         : " & matchedPanels & vbCrLf
    msg = msg & "Unknown panel outlines : " & unknownRedPanels & vbCrLf

    If isPants Then
        msg = msg & "Ignored small outlines : " & ignoredSmallRedPanels & vbCrLf
    End If

    If isSplitFront Then
        msg = msg & vbCrLf & "Mode Jaket V8.4: 3 panel body-only = 1 set" & vbCrLf
    ElseIf isPants Then
        msg = msg & vbCrLf & "Mode Celana V8.1: 4 panel = 1 set" & vbCrLf
        msg = msg & "Celana rotation-safe: cek Width dan Height." & vbCrLf
        msg = msg & "Pants tolerance: " & FormatNumber(PANTS_TOL, 2) & " cm" & vbCrLf
    Else
        msg = msg & vbCrLf & "Mode Jersey: 2 panel = 1 set" & vbCrLf
    End If

    If isPants And unknownDimReport <> "" Then
        msg = msg & vbCrLf & "Unknown pants outline dimensions:" & vbCrLf
        msg = msg & unknownDimReport
        If unknownDimCount >= UNKNOWN_DIM_LIMIT Then
            msg = msg & "... unknown lainnya tidak ditampilkan." & vbCrLf
        End If
    End If

    qscLastPass = passAll

    If passAll Then
        qscLastStatus = "PASS"
        qscLastSummary = ProductModeText() & " | Expected " & totalOrder & " set | Found " & totalLayout & " set."
    Else
        qscLastStatus = "FAIL"
        qscLastSummary = ProductModeText() & " | Expected " & totalOrder & " set | Found " & totalLayout & " set."
    End If

    qscLastDetail = msg & vbCrLf & "CORE_BRIDGE=H5_PHASE5B" & vbCrLf & "JACKET_FILTER=V8_4_BODY_ONLY" & vbCrLf

    If qscReportMode Then Exit Sub

    If passAll Then
        MsgBox msg, vbInformation, "QC SIZE CHECK"
    Else
        MsgBox msg, vbExclamation, "QC SIZE CHECK"
    End If

End Sub


Private Function PanelsToSetCount(ByVal panelCount As Long) As Long

    Dim divisor As Long

    If panelCount <= 0 Then
        PanelsToSetCount = 0
        Exit Function
    End If

    If isPants Then
        divisor = 4
    ElseIf isSplitFront Then
        divisor = 3
    Else
        divisor = 2
    End If

    If panelCount Mod divisor = 0 Then
        PanelsToSetCount = panelCount \ divisor
    Else
        'Jika panel tidak genap sesuai struktur set,
        'biarkan angka panel mentah agar terlihat FAIL.
        PanelsToSetCount = panelCount
    End If

End Function


Private Function ProductModeText() As String

    If isPants Then
        ProductModeText = "CELANA"
    ElseIf isSplitFront Then
        ProductModeText = "JAKET"
    Else
        ProductModeText = "JERSEY"
    End If

End Function


'=========================================================
' HELPERS
'=========================================================

Private Function MinD(ByVal a As Double, ByVal b As Double) As Double
    MinD = H5_MinD(a, b)
End Function


Private Function MaxD(ByVal a As Double, ByVal b As Double) As Double
    MaxD = H5_MaxD(a, b)
End Function


Private Function ToDbl(ByVal v As Variant) As Double
    ToDbl = H5_ToDbl(v)
End Function


Private Function NormalizeSizeKey(ByVal sz As String) As String
    NormalizeSizeKey = H5_NormalizeSizeKey(sz)
End Function


Private Function IsStandardSize(ByVal sz As String) As Boolean
    IsStandardSize = H5_IsStandardSize(sz)
End Function

