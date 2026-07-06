'=========================================================
' AUTO-GENERATED PATCH: NORMAL ENGINE RENAMED FOR SMART DISPATCHER
' Do not import together with old QC Typo module that still exposes QC_TYPO_CHECK/HADES_QC_TYPO_REPORT.
'=========================================================
Option Explicit

'=========================================================
' PROJECT HADES — QC TYPO CHECK NORMAL ENGINE V13.2 SMART PAIR REPORT
' CorelDRAW 2021 VBA
'
' BASE:
' - Revisi dari QC TYPO CHECK V11.2 SAFE
'
' FIX UTAMA V12:
' - Hard validation tetap exact.
' - Fuzzy / Levenshtein hanya dipakai untuk memilih
'   EXPECTED kandidat pada report saat FAIL.
'
' UPDATE V12.1 / PHASE 3B:
' - Menambahkan HADES_QC_TYPO_REPORT.
' - QC_TYPO_CHECK lama tetap bisa dipakai manual.
' - Report mode tidak meminta operator memilih P / W / F.
' - Status PASS / FAIL dikirim langsung ke HADES_CORE_REPORT.
'
' UPDATE V12.2 / PHASE 5C CORE BRIDGE:
' - Mulai memakai HADES Core Phase 5 untuk path, UTF-8, template,
'   Order.txt, SizeDB, normalize text, numeric text, red detection,
'   dan mode produk Jersey/Jaket/Celana.
' - Hard exact validation TETAP dipertahankan.
' - Fuzzy tetap hanya untuk memilih expected candidate saat FAIL.
' - Matching role/template tidak diubah agar behavior produksi tetap stabil.
'
' CONTOH:
' Order.txt:
'   L|ZALDI||
'
' Layout:
'   ZALDIE
'
' HASIL:
'   Tetap FAIL.
'   Report expected memilih ZALDI, bukan ZUL/order pertama.
'
' UPDATE:
' - Skip metadata Order.txt yang diawali "@"
' - Bisa membaca @SIZEDB= dari Order.txt jika template belum punya DB
' - IsRed membaca outline merah atau fill merah
' - Support database tambahan:
'   SizeDB_PriaSlimFit.txt
'   SizeDB_WanitaSlimFit.txt
'   SizeDB_CelanaAnak.txt
'
' SOURCE:
' - Documents\Order.txt
' - Documents\TypoTemplate_Current.txt
' - Documents\SizeDB_*.txt
'

' UPDATE V13.2 SMART PAIR REPORT:
' - Report expected candidate dipilih dengan pair-aware scoring.
' - Jika nama typo tetapi nomor benar, expected mengikuti nomor + fuzzy nama.
' - Jika nama benar tetapi nomor tertukar, expected mengikuti nama dan nomor dianggap mismatch.
' - Report mencantumkan Size, Expected Pair, Found Pair, dan indikasi pemilik nomor/nama yang tertukar.
' - ENGINE VERSION: shortcut publik dipindahkan ke dispatcher QC_TYPO_CHECK / HADES_QC_TYPO_REPORT.
'
' UPDATE V13.2G GREEN FAIL PANEL MARKER:
' - ENGINE VERSION: public entry normal adalah QTC_NORMAL_CHECK dan QTC_NORMAL_REPORT.
' - Saat group FAIL typo, panel pola badan depan/belakang yang cocok SizeDB
'   di dalam group tersebut diberi outline hijau.
' - Validasi typo, report, Order.txt, SizeDB, dan QC Final Menu tidak diubah.
' CARA PAKAI:
' 1. Select MASTER placeholder 1 size
' 2. Run BUILD_TYPO_TEMPLATE
' 3. Select ALL HASIL LAYOUT
' 4. Run QC_TYPO_CHECK lewat dispatcher otomatis
'=========================================================


'=========================================================
' GLOBAL
'=========================================================

Dim orders As Collection
Dim used() As Boolean

Dim SizeDB As Object
Dim Tpl As Object

Dim CurrentDB As String
Dim isSplitFront As Boolean
Dim isPants As Boolean

Dim report As String
Dim PassCount As Long
Dim failCount As Long
Dim qtcMarkedPanels As Long

'=========================================================
' REPORT MODE STATE - PHASE 3B
'=========================================================
Private qtcReportMode As Boolean
Private qtcLastStatus As String
Private qtcLastSummary As String
Private qtcLastDetail As String


'=========================================================
' CONSTANTS
'=========================================================

Const SIZE_TOL As Double = 1#
Const PANTS_TOL As Double = 0.35
Const TEXT_TOL As Double = 0.2

Const ID_MIN As Double = 0.28
Const ID_MAX As Double = 0.65

Const CHECK_EMPTY_FIELDS As Boolean = True

' Smart report score
Const FUZZY_MIN_NAME_SCORE As Double = 45#
Const FUZZY_MIN_NUMBER_SCORE As Double = 60#

' Green marker untuk panel body pada group yang FAIL typo.
' Shortcut tidak berubah. Behavior validasi tidak diubah.
Const QTC_MARK_FAIL_PANEL_GREEN As Boolean = True
Const QTC_FAIL_PANEL_GREEN_R As Long = 0
Const QTC_FAIL_PANEL_GREEN_G As Long = 255
Const QTC_FAIL_PANEL_GREEN_B As Long = 0


'=========================================================
' MAIN
'=========================================================

Public Sub QTC_NORMAL_REPORT()

    On Error GoTo ERR_HANDLER

    qtcReportMode = True
    qtcLastStatus = ""
    qtcLastSummary = ""
    qtcLastDetail = ""

    Call QTC_NORMAL_CHECK

    If Trim$(qtcLastStatus) = "" Then
        HADESR_AddResult _
            "TYPO CHECK", _
            "FAIL", _
            "QC Typo report mode tidak menghasilkan status.", _
            "Kemungkinan template, Order.txt, SizeDB, atau selection gagal sebelum report terbentuk."
    Else
        HADESR_AddResult _
            "TYPO CHECK", _
            qtcLastStatus, _
            qtcLastSummary, _
            qtcLastDetail
    End If

SAFE_EXIT:
    qtcReportMode = False
    Exit Sub

ERR_HANDLER:
    HADESR_AddResult _
        "TYPO CHECK", _
        "FAIL", _
        "QC Typo report mode error.", _
        "Error " & Err.Number & ": " & Err.Description

    Resume SAFE_EXIT

End Sub

Public Sub HADES5C_QC_TYPO_CORE_SMOKE_TEST()

    Dim reportText As String
    Dim failCount As Long
    Dim warnCount As Long
    Dim ok As Boolean
    Dim dbName As String
    Dim tpl As Object
    Dim rows As Collection

    On Error GoTo ERR_HANDLER

    ok = H5_RunCoreSelfTest(reportText, failCount, warnCount)
    H5_WriteCoreSelfTestReport reportText

    dbName = H5_DetectCurrentSizeDBFileName()
    Set tpl = H5_LoadTypoTemplate()
    Set rows = H5_LoadOrderRows()

    If ok Then
        MsgBox "QC Typo Core Bridge siap." & vbCrLf & vbCrLf & _
               "Core self-test : PASS" & vbCrLf & _
               "Template keys  : " & tpl.Count & vbCrLf & _
               "Order rows     : " & rows.Count & vbCrLf & _
               "SizeDB         : " & dbName & vbCrLf & vbCrLf & _
               "HADES_QC_TYPO_REPORT akan memakai H5 core bridge untuk path, UTF-8, normalize, DB mode, dan red detection.", _
               vbInformation, "HADES PHASE 5C"
    Else
        MsgBox "QC Typo Core Bridge belum siap." & vbCrLf & vbCrLf & _
               "Core self-test: FAIL" & vbCrLf & _
               "Fail    : " & failCount & vbCrLf & _
               "Warning : " & warnCount & vbCrLf & vbCrLf & _
               "Buka HADES_CORE_SELF_TEST_LATEST.txt di Documents\HADES_REPORTS.", _
               vbCritical, "HADES PHASE 5C"
    End If

    Exit Sub

ERR_HANDLER:
    MsgBox "HADES5C QC TYPO CORE SMOKE TEST ERROR" & vbCrLf & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, _
           vbCritical, "HADES PHASE 5C"

End Sub


Public Sub QTC_NORMAL_CHECK()

    qtcLastStatus = ""
    qtcLastSummary = ""
    qtcLastDetail = ""

    If ActiveSelection Is Nothing Then
        qtcLastStatus = "FAIL"
        qtcLastSummary = "Tidak ada hasil layout yang dipilih."
        qtcLastDetail = "Pilih HASIL LAYOUT terlebih dahulu sebelum menjalankan QC Typo."
        If Not qtcReportMode Then
            MsgBox "Pilih HASIL LAYOUT terlebih dahulu.", vbExclamation
        End If
        Exit Sub
    End If

    If ActiveSelection.Shapes.Count = 0 Then
        qtcLastStatus = "FAIL"
        qtcLastSummary = "Tidak ada hasil layout yang dipilih."
        qtcLastDetail = "Pilih HASIL LAYOUT terlebih dahulu sebelum menjalankan QC Typo."
        If Not qtcReportMode Then
            MsgBox "Pilih HASIL LAYOUT terlebih dahulu.", vbExclamation
        End If
        Exit Sub
    End If

    Set orders = New Collection
    Set SizeDB = CreateObject("Scripting.Dictionary")
    Set Tpl = CreateObject("Scripting.Dictionary")

    report = ""
    PassCount = 0
    failCount = 0
    qtcMarkedPanels = 0

    CurrentDB = ""
    isSplitFront = False
    isPants = False

    Dim oldUnit As cdrUnit
    oldUnit = ActiveDocument.Unit

    On Error GoTo ERR_HANDLER

    ActiveDocument.Unit = cdrCentimeter

    LoadTemplate

    ' V12:
    ' Kalau template belum punya DB, coba baca @SIZEDB dari Order.txt.
    LoadOrderMetadata

    If CurrentDB = "" Then
        If Not SelectDatabaseFallback Then
            qtcLastStatus = "FAIL"
            qtcLastSummary = "Database SizeDB belum dipilih / belum terbaca."
            qtcLastDetail = "TypoTemplate_Current.txt atau Order.txt belum menyimpan DB/SIZEDB, dan operator membatalkan pilihan fallback."
            GoTo EXIT_CLEAN
        End If
    End If

    ApplyModeFromCurrentDB

    LoadDB CurrentDB
    LoadOrders

    If orders.Count = 0 Then
        qtcLastStatus = "FAIL"
        qtcLastSummary = "Order.txt kosong / tidak terbaca."
        qtcLastDetail = "Pastikan Documents\Order.txt ada dan berformat SIZE|NAMA|NOMOR|NICKNAME. Metadata @... boleh ada, tetapi data order harus tetap ada."
        If Not qtcReportMode Then
            MsgBox "Order.txt kosong / tidak terbaca.", vbExclamation
        End If
        GoTo EXIT_CLEAN
    End If

    ReDim used(1 To orders.Count)

    Dim groups As Collection
    Set groups = CollectTopGroups(ActiveSelection)

    If groups.Count = 0 Then
        qtcLastStatus = "FAIL"
        qtcLastSummary = "Tidak ada group utama pada selection."
        qtcLastDetail = "Pastikan yang diselect adalah HASIL LAYOUT yang sudah digroup per jersey/set."
        If Not qtcReportMode Then
            MsgBox _
                "Tidak ada group utama pada selection." & vbCrLf & vbCrLf & _
                "Pastikan yang diselect adalah HASIL LAYOUT yang sudah digroup per jersey.", _
                vbExclamation
        End If
        GoTo EXIT_CLEAN
    End If

    Dim i As Long

    For i = 1 To groups.Count
        CheckGroup groups(i), i
    Next i

    ActiveDocument.Unit = oldUnit

    ShowFinalResult

    Exit Sub

EXIT_CLEAN:

    On Error Resume Next
    ActiveDocument.Unit = oldUnit
    On Error GoTo 0

    Exit Sub

ERR_HANDLER:

    Dim eNo As Long
    Dim eDesc As String

    eNo = Err.Number
    eDesc = Err.Description

    On Error Resume Next
    ActiveDocument.Unit = oldUnit
    On Error GoTo 0

    If eNo = 0 And Trim(eDesc) = "" Then
        eDesc = "Error tidak teridentifikasi. Cek selection, template, Order.txt, dan SizeDB."
    End If

    qtcLastStatus = "FAIL"
    qtcLastSummary = "System error pada QC Typo Check."
    qtcLastDetail = "Error " & eNo & ": " & eDesc

    If Not qtcReportMode Then
        MsgBox _
            "SYSTEM ERROR - QC TYPO CHECK V13.2" & vbCrLf & vbCrLf & _
            "No : " & eNo & vbCrLf & _
            eDesc, _
            vbCritical
    End If

End Sub


'=========================================================
' LOAD TEMPLATE
'=========================================================

Sub LoadTemplate()

    On Error GoTo FAIL

    Set Tpl = H5_LoadTypoTemplate()

    If Tpl Is Nothing Then
        Err.Raise vbObjectError + 100, , _
            "TypoTemplate_Current.txt tidak terbaca." & vbCrLf & _
            "Jalankan BUILD_TYPO_TEMPLATE terlebih dahulu."
    End If

    If Tpl.Count = 0 Then
        Err.Raise vbObjectError + 100, , _
            "TypoTemplate_Current.txt tidak ditemukan / kosong." & vbCrLf & _
            "Jalankan BUILD_TYPO_TEMPLATE terlebih dahulu."
    End If

    If Not Tpl.Exists("MASTER_PANEL") Then
        Err.Raise vbObjectError + 101, , _
            "TypoTemplate tidak memiliki MASTER_PANEL." & vbCrLf & _
            "Build template ulang dari master."
    End If

    If H5_ToDbl(Tpl("MASTER_PANEL")) <= 0 Then
        Err.Raise vbObjectError + 102, , _
            "MASTER_PANEL pada TypoTemplate tidak valid."
    End If

    If Tpl.Exists("DB") Then
        CurrentDB = Trim$(CStr(Tpl("DB")))
    End If

    If Len(Trim$(CurrentDB)) = 0 Then
        If Tpl.Exists("SIZEDB") Then
            CurrentDB = Trim$(CStr(Tpl("SIZEDB")))
        End If
    End If

    If Len(Trim$(CurrentDB)) > 0 Then
        If InStr(1, UCase$(CurrentDB), ".TXT", vbTextCompare) = 0 Then
            CurrentDB = CurrentDB & ".txt"
        End If
    End If

    If Tpl.Exists("MODE") Then

        Select Case UCase$(Trim$(CStr(Tpl("MODE"))))

            Case "SPLIT_FRONT"
                isSplitFront = True
                isPants = False

            Case "PANTS", "CELANA"
                isSplitFront = False
                isPants = True

        End Select

    End If

    ApplyModeFromCurrentDB

    Exit Sub

FAIL:
    Err.Raise vbObjectError + 103, , _
        "Gagal membaca TypoTemplate_Current.txt melalui HADES Core." & vbCrLf & _
        Err.Description

End Sub


'=========================================================
' LOAD ORDER METADATA
'=========================================================

Sub LoadOrderMetadata()

    Dim meta As Object
    Dim dbName As String
    Dim jenis As String

    On Error GoTo SAFE_EXIT

    Set meta = H5_LoadOrderMeta()

    If meta Is Nothing Then GoTo SAFE_EXIT

    If Len(Trim$(CurrentDB)) = 0 Then

        If meta.Exists("SIZEDB") Then dbName = Trim$(CStr(meta("SIZEDB")))

        If Len(Trim$(dbName)) = 0 Then
            If meta.Exists("DB") Then dbName = Trim$(CStr(meta("DB")))
        End If

        If Len(Trim$(dbName)) = 0 Then
            dbName = H5_InferDBFromOrderMeta(meta)
        End If

        If Len(Trim$(dbName)) > 0 Then
            If InStr(1, UCase$(dbName), ".TXT", vbTextCompare) = 0 Then
                dbName = dbName & ".txt"
            End If
            CurrentDB = dbName
        End If

    End If

    If meta.Exists("JENIS_PESANAN") Then
        jenis = UCase$(CStr(meta("JENIS_PESANAN")))

        If InStr(1, jenis, "CELANA", vbTextCompare) > 0 Then
            isPants = True
            isSplitFront = False
        End If

        If InStr(1, jenis, "JAKET", vbTextCompare) > 0 Then
            isPants = False
            isSplitFront = True
        End If
    End If

SAFE_EXIT:
    ApplyModeFromCurrentDB

End Sub


'=========================================================
' APPLY MODE FROM DB
'=========================================================

Sub ApplyModeFromCurrentDB()

    If Len(Trim$(CurrentDB)) = 0 Then Exit Sub

    H5_ProductModeFromDB CurrentDB, isPants, isSplitFront

End Sub


'=========================================================
' DATABASE POPUP FALLBACK
'=========================================================

Function SelectDatabaseFallback() As Boolean

    SelectDatabaseFallback = False

    Dim a As String
    Dim b As String

    a = InputBox( _
        "TypoTemplate / Order.txt belum menyimpan DB." & vbCrLf & vbCrLf & _
        "PILIH PRODUK" & vbCrLf & vbCrLf & _
        "1 = JERSEY" & vbCrLf & _
        "2 = JAKET" & vbCrLf & _
        "3 = CELANA", _
        "QC TYPO CHECK V13.2")

    If a = "" Then Exit Function

    Select Case Trim(a)

        Case "1"

            isSplitFront = False
            isPants = False

            b = InputBox( _
                "PILIH JERSEY" & vbCrLf & vbCrLf & _
                "1 = PRIA REGULAR" & vbCrLf & _
                "2 = WANITA REGULAR" & vbCrLf & _
                "3 = ANAK" & vbCrLf & _
                "4 = PRIA SLIM FIT" & vbCrLf & _
                "5 = WANITA SLIM FIT", _
                "QC TYPO CHECK V13.2")

            If b = "" Then Exit Function

            Select Case Trim(b)

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
                    MsgBox "Pilihan tidak valid.", vbExclamation
                    Exit Function

            End Select

        Case "2"

            isSplitFront = True
            isPants = False

            b = InputBox( _
                "PILIH JAKET" & vbCrLf & vbCrLf & _
                "1 = DEWASA" & vbCrLf & _
                "2 = ANAK", _
                "QC TYPO CHECK V13.2")

            If b = "" Then Exit Function

            Select Case Trim(b)

                Case "1"
                    CurrentDB = "SizeDB_Jaket.txt"

                Case "2"
                    CurrentDB = "SizeDB_JaketAnak.txt"

                Case Else
                    MsgBox "Pilihan tidak valid.", vbExclamation
                    Exit Function

            End Select

        Case "3"

            isSplitFront = False
            isPants = True

            b = InputBox( _
                "PILIH CELANA" & vbCrLf & vbCrLf & _
                "1 = CELANA PRIA" & vbCrLf & _
                "2 = CELANA WANITA" & vbCrLf & _
                "3 = CELANA ANAK", _
                "QC TYPO CHECK V13.2")

            If b = "" Then Exit Function

            Select Case Trim(b)

                Case "1"
                    CurrentDB = "SizeDB_CelanaPria.txt"

                Case "2"
                    CurrentDB = "SizeDB_CelanaWanita.txt"

                Case "3"
                    CurrentDB = "SizeDB_CelanaAnak.txt"

                Case Else
                    MsgBox "Pilihan tidak valid.", vbExclamation
                    Exit Function

            End Select

        Case Else

            MsgBox "Pilihan tidak valid.", vbExclamation
            Exit Function

    End Select

    ApplyModeFromCurrentDB

    SelectDatabaseFallback = True

End Function


'=========================================================
' LOAD ORDER
'=========================================================

Sub LoadOrders()

    On Error GoTo FAIL

    Set orders = H5_LoadOrderRows()

    Exit Sub

FAIL:
    Err.Raise vbObjectError + 201, , _
        "Gagal membaca Order.txt melalui HADES Core." & vbCrLf & _
        Err.Description

End Sub


'=========================================================
' LOAD SIZE DATABASE
'=========================================================

Sub LoadDB(ByVal fileName As String)

    On Error GoTo FAIL

    Set SizeDB = H5_LoadSizeDB(fileName, isPants, isSplitFront)

    If SizeDB Is Nothing Then
        Err.Raise vbObjectError + 302, , _
            "SizeDB tidak terbaca melalui HADES Core."
    End If

    If SizeDB.Count = 0 Then
        Err.Raise vbObjectError + 302, , _
            "SizeDB kosong / tidak ditemukan / format tidak valid:" & vbCrLf & _
            H5_DocumentsFile(fileName)
    End If

    Exit Sub

FAIL:
    Err.Raise vbObjectError + 301, , _
        "Gagal membaca database melalui HADES Core: " & fileName & vbCrLf & _
        Err.Description

End Sub


'=========================================================
' CHECK GROUP
'=========================================================

Sub CheckGroup(ByVal g As Shape, ByVal groupIndex As Long)

    Dim detectedSize As String
    Dim currentPanel As Double

    DetectGroupPanel g, detectedSize, currentPanel

    If detectedSize = "" Or currentPanel <= 0 Then

        failCount = failCount + 1

        report = report & _
            "GROUP #" & groupIndex & vbCrLf & _
            "STATUS : SIZE / PANEL MERAH TIDAK TERDETEKSI" & vbCrLf & _
            String(40, "-") & vbCrLf & vbCrLf

        Exit Sub

    End If

    Dim masterPanel As Double
    masterPanel = H5_ToDbl(Tpl("MASTER_PANEL"))

    Dim scaleFactor As Double
    scaleFactor = currentPanel / masterPanel

    Dim roles As Object
    Dim allTexts As Collection

    Set roles = CreateObject("Scripting.Dictionary")
    Set allTexts = New Collection

    ReadTextRoles g, roles, allTexts, scaleFactor

    Dim best As Long
    best = FindBestOrder(roles, allTexts, detectedSize)

    If best > 0 Then

        used(best) = True
        PassCount = PassCount + 1

        Exit Sub

    End If

    failCount = failCount + 1

    Dim nearest As Long
    nearest = FindNearestOrder(roles, allTexts, detectedSize)

    If QTC_MARK_FAIL_PANEL_GREEN Then
        MarkTypoFailPanelsGreen g, detectedSize
    End If

    BuildFailReport _
        groupIndex, _
        detectedSize, _
        currentPanel, _
        scaleFactor, _
        roles, _
        allTexts, _
        nearest

End Sub


'=========================================================
' DETECT PANEL / SIZE
'=========================================================

Sub DetectGroupPanel( _
    ByVal s As Shape, _
    ByRef detectedSize As String, _
    ByRef panelHeight As Double)

    detectedSize = ""
    panelHeight = 0

    Dim bestArea As Double
    bestArea = 0

    ScanPanelRecursive s, detectedSize, panelHeight, bestArea

End Sub


Sub ScanPanelRecursive( _
    ByVal s As Shape, _
    ByRef detectedSize As String, _
    ByRef panelHeight As Double, _
    ByRef bestArea As Double)

    If s.Type = cdrGroupShape Then

        Dim c As Shape

        For Each c In s.Shapes
            ScanPanelRecursive c, detectedSize, panelHeight, bestArea
        Next c

        Exit Sub

    End If

    If s.Type <> cdrCurveShape Then Exit Sub
    If Not IsRed(s) Then Exit Sub

    Dim sz As String
    Dim ph As Double
    Dim area As Double

    sz = DetectSizeFromShape(s, ph, area)

    If sz <> "" Then

        If area > bestArea Then

            bestArea = area
            panelHeight = ph

            If InStr(1, sz, "_", vbTextCompare) > 0 Then
                detectedSize = Left(sz, InStr(1, sz, "_", vbTextCompare) - 1)
            Else
                detectedSize = sz
            End If

        End If

    End If

End Sub


Function DetectSizeFromShape( _
    ByVal shp As Shape, _
    ByRef panelH As Double, _
    ByRef area As Double) As String

    Dim w As Double
    Dim h As Double

    w = Round(shp.SizeWidth, 3)
    h = Round(shp.SizeHeight, 3)

    If w <= 0 Or h <= 0 Then Exit Function

    Dim mx As Double
    Dim mn As Double

    If w > h Then
        mx = w
        mn = h
    Else
        mx = h
        mn = w
    End If

    panelH = mx
    area = w * h

    Dim k As Variant
    Dim db As Variant

    For Each k In SizeDB.keys

        db = SizeDB(k)

        If isPants Then

            If UBound(db) >= 2 Then

                If Abs(mn - H5_ToDbl(db(1))) <= PANTS_TOL Or _
                   Abs(mn - H5_ToDbl(db(2))) <= PANTS_TOL Then

                    DetectSizeFromShape = CStr(k)
                    Exit Function

                End If

            End If

        ElseIf isSplitFront Then

            If UBound(db) >= 4 Then

                If Abs(mn - H5_ToDbl(db(1))) <= SIZE_TOL And _
                   Abs(mx - H5_ToDbl(db(4))) <= SIZE_TOL Then

                    DetectSizeFromShape = CStr(k) & "_BACK"
                    Exit Function

                End If

                If Abs(mn - H5_ToDbl(db(2))) <= SIZE_TOL And _
                   Abs(mx - H5_ToDbl(db(3))) <= SIZE_TOL Then

                    DetectSizeFromShape = CStr(k) & "_FRONT"
                    Exit Function

                End If

            End If

        Else

            If UBound(db) >= 3 Then

                If Abs(mn - H5_ToDbl(db(1))) <= SIZE_TOL And _
                   (Abs(mx - H5_ToDbl(db(2))) <= SIZE_TOL Or _
                    Abs(mx - H5_ToDbl(db(3))) <= SIZE_TOL) Then

                    DetectSizeFromShape = CStr(k)
                    Exit Function

                End If

            End If

        End If

    Next k

End Function


Function IsRed(ByVal shp As Shape) As Boolean

    IsRed = H5_IsRedShape(shp)

End Function


Function IsRedRGB( _
    ByVal r As Long, _
    ByVal g As Long, _
    ByVal b As Long) As Boolean

    IsRedRGB = H5_IsRedRGB(r, g, b)

End Function


'=========================================================
' READ TEXT + CLASSIFY ROLE
'=========================================================

Sub ReadTextRoles( _
    ByVal s As Shape, _
    ByRef roles As Object, _
    ByRef allTexts As Collection, _
    ByVal scaleFactor As Double)

    If s.Type = cdrGroupShape Then

        Dim c As Shape

        For Each c In s.Shapes
            ReadTextRoles c, roles, allTexts, scaleFactor
        Next c

        Exit Sub

    End If

    If s.Type <> cdrTextShape Then Exit Sub

    Dim raw As String
    raw = Trim(s.Text.Story.Text)

    If raw = "" Then Exit Sub

    If IgnoreSmallID(s, raw) Then Exit Sub

    Dim clean As String
    clean = Normalize(raw)

    If clean = "" Then Exit Sub

    allTexts.Add clean

    Dim role As String
    role = ClassifyTextRole(s, clean, scaleFactor)

    If role <> "" Then
        AddRoleText roles, role, clean
    End If

End Sub


Function ClassifyTextRole( _
    ByVal t As Shape, _
    ByVal txt As String, _
    ByVal scaleFactor As Double) As String

    Dim bestRole As String
    Dim bestScore As Double

    bestRole = ""
    bestScore = 9999#

    If IsNumberText(txt) Then

        TryRole t, "NUMBER", scaleFactor, bestRole, bestScore

    Else

        If Not isPants Then
            TryRole t, "NAMA_ATLIT", scaleFactor, bestRole, bestScore
            TryRole t, "NAMA", scaleFactor, bestRole, bestScore
            TryRole t, "NICKNAME", scaleFactor, bestRole, bestScore
        End If

    End If

    ClassifyTextRole = bestRole

End Function


Sub TryRole( _
    ByVal t As Shape, _
    ByVal role As String, _
    ByVal scaleFactor As Double, _
    ByRef bestRole As String, _
    ByRef bestScore As Double)

    Dim cnt As Long
    cnt = ExpectedRoleCount(role)

    If cnt <= 0 Then Exit Sub

    Dim i As Long

    For i = 1 To cnt

        Dim hMaster As Double
        hMaster = RoleSlotHeight(role, i)

        If hMaster <= 0 Then
            hMaster = LegacyRoleHeight(role)
        End If

        If hMaster <= 0 Then GoTo NEXTI

        Dim targetH As Double
        targetH = hMaster * scaleFactor

        If targetH <= 0 Then GoTo NEXTI

        Dim hNow As Double
        hNow = Round(t.SizeHeight, 3)

        Dim diffRatio As Double
        diffRatio = Abs(hNow - targetH) / targetH

        If diffRatio > TEXT_TOL Then GoTo NEXTI

        Dim tplAlign As String
        Dim curAlign As String

        tplAlign = RoleSlotAlign(role, i)
        curAlign = GetAlign(t)

        If tplAlign <> "" And curAlign <> "" Then
            If tplAlign <> curAlign Then
                diffRatio = diffRatio + 0.05
            End If
        End If

        If diffRatio < bestScore Then
            bestScore = diffRatio
            bestRole = role
        End If

NEXTI:

    Next i

End Sub


Sub AddRoleText( _
    ByRef roles As Object, _
    ByVal role As String, _
    ByVal txt As String)

    Dim d As Object

    If roles.Exists(role) Then
        Set d = roles(role)
    Else
        Set d = CreateObject("Scripting.Dictionary")
        roles.Add role, d
    End If

    If d.Exists(txt) Then
        d(txt) = CLng(d(txt)) + 1
    Else
        d.Add txt, 1
    End If

End Sub


'=========================================================
' TEMPLATE HELPERS
'=========================================================

Function ExpectedRoleCount(ByVal role As String) As Long

    role = UCase(role)

    If Tpl.Exists(role & "_COUNT") Then
        ExpectedRoleCount = CLng(H5_ToDbl(Tpl(role & "_COUNT")))
        Exit Function
    End If

    If Tpl.Exists(role & "_H") Then
        ExpectedRoleCount = 1
        Exit Function
    End If

    If Tpl.Exists(role) Then
        ExpectedRoleCount = 1
        Exit Function
    End If

End Function


Function RoleSlotHeight(ByVal role As String, ByVal idx As Long) As Double

    role = UCase(role)

    If Tpl.Exists(role & "_" & idx & "_H") Then
        RoleSlotHeight = H5_ToDbl(Tpl(role & "_" & idx & "_H"))
        Exit Function
    End If

    If idx = 1 Then
        RoleSlotHeight = LegacyRoleHeight(role)
    End If

End Function


Function RoleSlotAlign(ByVal role As String, ByVal idx As Long) As String

    role = UCase(role)

    If Tpl.Exists(role & "_" & idx & "_ALIGN") Then
        RoleSlotAlign = UCase(Trim(Tpl(role & "_" & idx & "_ALIGN")))
        Exit Function
    End If

    If idx = 1 Then
        If Tpl.Exists(role & "_ALIGN") Then
            RoleSlotAlign = UCase(Trim(Tpl(role & "_ALIGN")))
            Exit Function
        End If
    End If

End Function


Function LegacyRoleHeight(ByVal role As String) As Double

    role = UCase(role)

    If Tpl.Exists(role & "_H") Then
        LegacyRoleHeight = H5_ToDbl(Tpl(role & "_H"))
        Exit Function
    End If

    If Tpl.Exists(role) Then
        LegacyRoleHeight = H5_ToDbl(Tpl(role))
        Exit Function
    End If

    If role = "NAMA_ATLIT" Then
        If Tpl.Exists("NAMAATLIT") Then
            LegacyRoleHeight = H5_ToDbl(Tpl("NAMAATLIT"))
            Exit Function
        End If
    End If

    If role = "NICKNAME" Then
        If Tpl.Exists("NICK") Then
            LegacyRoleHeight = H5_ToDbl(Tpl("NICK"))
            Exit Function
        End If
    End If

End Function


'=========================================================
' ORDER MATCHING — HARD VALIDATION
'=========================================================

Function FindBestOrder( _
    ByVal roles As Object, _
    ByVal allTexts As Collection, _
    ByVal detectedSize As String) As Long

    Dim i As Long

    For i = 1 To orders.Count

        If used(i) Then GoTo NEXTI

        Dim ord As Variant
        ord = orders(i)

        If H5_NormalizeSizeKey(OrderField(ord, 0)) <> H5_NormalizeSizeKey(detectedSize) Then
            GoTo NEXTI
        End If

        If ValidateOrder(roles, allTexts, ord) Then
            FindBestOrder = i
            Exit Function
        End If

NEXTI:

    Next i

End Function


Function ValidateOrder( _
    ByVal roles As Object, _
    ByVal allTexts As Collection, _
    ByVal ord As Variant) As Boolean

    ValidateOrder = True

    If isPants Then

        If Normalize(OrderField(ord, 2)) <> "" Then
            If AllTextExactCount(allTexts, OrderField(ord, 2)) = 0 Then
                ValidateOrder = False
                Exit Function
            End If
        End If

        If Not ValidateRoleField(roles, "NUMBER", OrderField(ord, 2)) Then
            ValidateOrder = False
            Exit Function
        End If

        Exit Function

    End If

    If Normalize(OrderField(ord, 1)) <> "" Then
        If AllTextExactCount(allTexts, OrderField(ord, 1)) = 0 Then
            ValidateOrder = False
            Exit Function
        End If
    End If

    If Normalize(OrderField(ord, 2)) <> "" Then
        If AllTextExactCount(allTexts, OrderField(ord, 2)) = 0 Then
            ValidateOrder = False
            Exit Function
        End If
    End If

    If Normalize(OrderField(ord, 3)) <> "" Then
        If AllTextExactCount(allTexts, OrderField(ord, 3)) = 0 Then
            ValidateOrder = False
            Exit Function
        End If
    End If

    If Not ValidateRoleField(roles, "NAMA_ATLIT", OrderField(ord, 1)) Then
        ValidateOrder = False
        Exit Function
    End If

    If Not ValidateRoleField(roles, "NAMA", OrderField(ord, 1)) Then
        ValidateOrder = False
        Exit Function
    End If

    If Not ValidateRoleField(roles, "NUMBER", OrderField(ord, 2)) Then
        ValidateOrder = False
        Exit Function
    End If

    If Not ValidateRoleField(roles, "NICKNAME", OrderField(ord, 3)) Then
        ValidateOrder = False
        Exit Function
    End If

End Function


Function ValidateRoleField( _
    ByVal roles As Object, _
    ByVal role As String, _
    ByVal Expected As String) As Boolean

    Dim cntExpected As Long
    cntExpected = ExpectedRoleCount(role)

    If cntExpected <= 0 Then
        ValidateRoleField = True
        Exit Function
    End If

    Dim exp As String
    exp = Normalize(Expected)

    Dim foundTotal As Long
    foundTotal = RoleTotalCount(roles, role)

    If exp = "" Then

        If CHECK_EMPTY_FIELDS Then
            ValidateRoleField = (foundTotal = 0)
        Else
            ValidateRoleField = True
        End If

        Exit Function

    End If

    If RoleExactCount(roles, role, exp) <> cntExpected Then
        ValidateRoleField = False
        Exit Function
    End If

    If foundTotal <> cntExpected Then
        ValidateRoleField = False
        Exit Function
    End If

    ValidateRoleField = True

End Function


'=========================================================
' ORDER MATCHING — SMART REPORT ONLY
'=========================================================

Function FindNearestOrder( _
    ByVal roles As Object, _
    ByVal allTexts As Collection, _
    ByVal detectedSize As String) As Long

    Dim best As Long
    Dim bestScore As Double
    Dim i As Long
    Dim sc As Double

    best = 0
    bestScore = -1

    ' V13.2:
    ' Nearest candidate untuk report dibuat pair-aware.
    ' Tujuan:
    ' - ZALDIE + 30 memilih expected ZALDI|30, bukan ZUL|21.
    ' - ZUL + 30 memilih expected ZUL|21 dan melaporkan nomor tertukar.
    For i = 1 To orders.Count

        If used(i) Then GoTo NEXTI

        Dim ord As Variant
        ord = orders(i)

        If H5_NormalizeSizeKey(OrderField(ord, 0)) <> H5_NormalizeSizeKey(detectedSize) Then
            GoTo NEXTI
        End If

        sc = ScoreOrder(roles, allTexts, ord)

        If best = 0 Or sc > bestScore Then
            best = i
            bestScore = sc
        End If

NEXTI:

    Next i

    If best > 0 Then
        FindNearestOrder = best
        Exit Function
    End If

    ' Fallback terakhir: semua size, hanya untuk report darurat.
    best = 0
    bestScore = -1

    For i = 1 To orders.Count

        If used(i) Then GoTo NEXTJ

        ord = orders(i)
        sc = ScoreOrder(roles, allTexts, ord)

        If best = 0 Or sc > bestScore Then
            best = i
            bestScore = sc
        End If

NEXTJ:

    Next i

    FindNearestOrder = best

End Function


Function ScoreOrder( _
    ByVal roles As Object, _
    ByVal allTexts As Collection, _
    ByVal ord As Variant) As Double

    Dim sc As Double
    Dim expName As String
    Dim expNo As String
    Dim expNick As String
    Dim simName As Double
    Dim simNick As Double
    Dim simNo As Double

    sc = 0

    expName = Normalize(OrderField(ord, 1))
    expNo = Normalize(OrderField(ord, 2))
    expNick = Normalize(OrderField(ord, 3))

    If isPants Then
        If expNo <> "" Then
            If AllTextExactCount(allTexts, expNo) > 0 Then sc = sc + 250000
            If RoleExactCount(roles, "NUMBER", expNo) > 0 Then sc = sc + 80000
            simNo = BestAllTextSimilarity(expNo, allTexts)
            If simNo >= FUZZY_MIN_NUMBER_SCORE Then sc = sc + simNo * 900
        End If
        ScoreOrder = sc
        Exit Function
    End If

    ' Prioritas 1: nama exact. Ini membuat kasus nomor tertukar tetap diarahkan ke nama yang benar.
    If expName <> "" Then
        If AllTextExactCount(allTexts, expName) > 0 Then sc = sc + 260000
        If RoleExactCount(roles, "NAMA_ATLIT", expName) > 0 Then sc = sc + 90000
        If RoleExactCount(roles, "NAMA", expName) > 0 Then sc = sc + 85000
    End If

    ' Prioritas 2: nomor exact. Ini membuat kasus ZALDIE + 30 memilih ZALDI|30.
    If expNo <> "" Then
        If AllTextExactCount(allTexts, expNo) > 0 Then sc = sc + 190000
        If RoleExactCount(roles, "NUMBER", expNo) > 0 Then sc = sc + 85000
    End If

    ' Prioritas 3: nickname exact.
    If expNick <> "" Then
        If AllTextExactCount(allTexts, expNick) > 0 Then sc = sc + 70000
        If RoleExactCount(roles, "NICKNAME", expNick) > 0 Then sc = sc + 50000
    End If

    ' Fuzzy hanya untuk report candidate, bukan untuk PASS.
    If expName <> "" Then
        simName = BestNameLikeSimilarity(expName, allTexts)
        If simName >= FUZZY_MIN_NAME_SCORE Then sc = sc + simName * 1200

        simName = BestRoleSimilarity(roles, "NAMA_ATLIT", expName)
        If simName >= FUZZY_MIN_NAME_SCORE Then sc = sc + simName * 750

        simName = BestRoleSimilarity(roles, "NAMA", expName)
        If simName >= FUZZY_MIN_NAME_SCORE Then sc = sc + simName * 650
    End If

    If expNo <> "" Then
        simNo = BestAllTextSimilarity(expNo, allTexts)
        If simNo >= FUZZY_MIN_NUMBER_SCORE Then sc = sc + simNo * 500
        simNo = BestRoleSimilarity(roles, "NUMBER", expNo)
        If simNo >= FUZZY_MIN_NUMBER_SCORE Then sc = sc + simNo * 600
    End If

    If expNick <> "" Then
        simNick = BestNameLikeSimilarity(expNick, allTexts)
        If simNick >= FUZZY_MIN_NAME_SCORE Then sc = sc + simNick * 350
        simNick = BestRoleSimilarity(roles, "NICKNAME", expNick)
        If simNick >= FUZZY_MIN_NAME_SCORE Then sc = sc + simNick * 450
    End If

    ScoreOrder = sc

End Function


Function BestNameLikeSimilarity( _
    ByVal expectedValue As String, _
    ByVal allTexts As Collection) As Double

    Dim best As Double
    Dim v As Variant
    Dim s As String
    Dim sc As Double

    best = 0
    expectedValue = Normalize(expectedValue)

    If expectedValue = "" Then Exit Function

    For Each v In allTexts

        s = Normalize(CStr(v))

        ' Hindari angka ikut dianggap kandidat nama.
        If Not IsNumberText(s) Then
            sc = StringSimilarityPercent(expectedValue, s)
            If sc > best Then best = sc
        End If

    Next v

    BestNameLikeSimilarity = best

End Function


Function ScoreNameField( _
    ByVal roles As Object, _
    ByVal allTexts As Collection, _
    ByVal expectedValue As String, _
    ByVal weight As Double) As Double

    Dim exp As String
    exp = Normalize(expectedValue)

    If exp = "" Then Exit Function

    Dim sc As Double
    sc = 0

    ' Exact tetap paling kuat.
    If AllTextExactCount(allTexts, exp) > 0 Then sc = sc + 10000 * weight
    If RoleExactCount(roles, "NAMA_ATLIT", exp) > 0 Then sc = sc + 7000 * weight
    If RoleExactCount(roles, "NAMA", exp) > 0 Then sc = sc + 6500 * weight

    ' Fuzzy hanya untuk report.
    Dim simAll As Double
    Dim simRole1 As Double
    Dim simRole2 As Double

    simAll = BestAllTextSimilarity(exp, allTexts)
    simRole1 = BestRoleSimilarity(roles, "NAMA_ATLIT", exp)
    simRole2 = BestRoleSimilarity(roles, "NAMA", exp)

    If simAll >= FUZZY_MIN_NAME_SCORE Then sc = sc + simAll * 35 * weight
    If simRole1 >= FUZZY_MIN_NAME_SCORE Then sc = sc + simRole1 * 25 * weight
    If simRole2 >= FUZZY_MIN_NAME_SCORE Then sc = sc + simRole2 * 22 * weight

    ScoreNameField = sc

End Function


Function ScoreNumberField( _
    ByVal roles As Object, _
    ByVal allTexts As Collection, _
    ByVal expectedValue As String, _
    ByVal weight As Double) As Double

    Dim exp As String
    exp = Normalize(expectedValue)

    If exp = "" Then Exit Function

    Dim sc As Double
    sc = 0

    If AllTextExactCount(allTexts, exp) > 0 Then sc = sc + 8500 * weight
    If RoleExactCount(roles, "NUMBER", exp) > 0 Then sc = sc + 7500 * weight

    Dim simAll As Double
    Dim simRole As Double

    simAll = BestAllTextSimilarity(exp, allTexts)
    simRole = BestRoleSimilarity(roles, "NUMBER", exp)

    If simAll >= FUZZY_MIN_NUMBER_SCORE Then sc = sc + simAll * 12 * weight
    If simRole >= FUZZY_MIN_NUMBER_SCORE Then sc = sc + simRole * 15 * weight

    ScoreNumberField = sc

End Function


Function ScoreNicknameField( _
    ByVal roles As Object, _
    ByVal allTexts As Collection, _
    ByVal expectedValue As String, _
    ByVal weight As Double) As Double

    Dim exp As String
    exp = Normalize(expectedValue)

    If exp = "" Then Exit Function

    Dim sc As Double
    sc = 0

    If AllTextExactCount(allTexts, exp) > 0 Then sc = sc + 5000 * weight
    If RoleExactCount(roles, "NICKNAME", exp) > 0 Then sc = sc + 4500 * weight

    Dim simAll As Double
    Dim simRole As Double

    simAll = BestAllTextSimilarity(exp, allTexts)
    simRole = BestRoleSimilarity(roles, "NICKNAME", exp)

    If simAll >= FUZZY_MIN_NAME_SCORE Then sc = sc + simAll * 12 * weight
    If simRole >= FUZZY_MIN_NAME_SCORE Then sc = sc + simRole * 10 * weight

    ScoreNicknameField = sc

End Function


Function BestAllTextSimilarity( _
    ByVal expectedValue As String, _
    ByVal allTexts As Collection) As Double

    Dim best As Double
    Dim v As Variant
    Dim sc As Double

    best = 0
    expectedValue = Normalize(expectedValue)

    If expectedValue = "" Then Exit Function

    For Each v In allTexts

        sc = StringSimilarityPercent(expectedValue, Normalize(CStr(v)))

        If sc > best Then best = sc

    Next v

    BestAllTextSimilarity = best

End Function


Function BestRoleSimilarity( _
    ByVal roles As Object, _
    ByVal role As String, _
    ByVal expectedValue As String) As Double

    Dim best As Double
    Dim k As Variant
    Dim sc As Double

    best = 0
    expectedValue = Normalize(expectedValue)

    If expectedValue = "" Then Exit Function
    If Not roles.Exists(role) Then Exit Function

    Dim d As Object
    Set d = roles(role)

    For Each k In d.keys

        sc = StringSimilarityPercent(expectedValue, Normalize(CStr(k)))

        If sc > best Then best = sc

    Next k

    BestRoleSimilarity = best

End Function


Function StringSimilarityPercent( _
    ByVal a As String, _
    ByVal b As String) As Double

    a = Normalize(a)
    b = Normalize(b)

    If a = "" Or b = "" Then
        StringSimilarityPercent = 0
        Exit Function
    End If

    If a = b Then
        StringSimilarityPercent = 100
        Exit Function
    End If

    Dim maxLen As Long
    maxLen = Len(a)
    If Len(b) > maxLen Then maxLen = Len(b)

    If maxLen <= 0 Then
        StringSimilarityPercent = 0
        Exit Function
    End If

    Dim dist As Long
    dist = LevenshteinDistance(a, b)

    StringSimilarityPercent = ((maxLen - dist) / maxLen) * 100

    If StringSimilarityPercent < 0 Then StringSimilarityPercent = 0

End Function


Function LevenshteinDistance( _
    ByVal s As String, _
    ByVal t As String) As Long

    Dim n As Long
    Dim m As Long

    n = Len(s)
    m = Len(t)

    If n = 0 Then
        LevenshteinDistance = m
        Exit Function
    End If

    If m = 0 Then
        LevenshteinDistance = n
        Exit Function
    End If

    Dim d() As Long
    ReDim d(0 To n, 0 To m)

    Dim i As Long
    Dim j As Long
    Dim cost As Long

    For i = 0 To n
        d(i, 0) = i
    Next i

    For j = 0 To m
        d(0, j) = j
    Next j

    For i = 1 To n

        For j = 1 To m

            If Mid(s, i, 1) = Mid(t, j, 1) Then
                cost = 0
            Else
                cost = 1
            End If

            d(i, j) = Min3Long( _
                        d(i - 1, j) + 1, _
                        d(i, j - 1) + 1, _
                        d(i - 1, j - 1) + cost)

        Next j

    Next i

    LevenshteinDistance = d(n, m)

End Function


Function Min3Long( _
    ByVal a As Long, _
    ByVal b As Long, _
    ByVal c As Long) As Long

    Dim x As Long

    x = a

    If b < x Then x = b
    If c < x Then x = c

    Min3Long = x

End Function


Function OrderField(ByVal ord As Variant, ByVal idx As Long) As String

    On Error Resume Next

    If IsArray(ord) Then
        If UBound(ord) >= idx Then
            OrderField = CStr(ord(idx))
        End If
    End If

    On Error GoTo 0

End Function


'=========================================================
' ROLE / TEXT COUNT HELPERS
'=========================================================

Function RoleTotalCount( _
    ByVal roles As Object, _
    ByVal role As String) As Long

    If Not roles.Exists(role) Then Exit Function

    Dim d As Object
    Set d = roles(role)

    Dim k As Variant

    For Each k In d.keys
        RoleTotalCount = RoleTotalCount + CLng(d(k))
    Next k

End Function


Function RoleExactCount( _
    ByVal roles As Object, _
    ByVal role As String, _
    ByVal Expected As String) As Long

    Expected = Normalize(Expected)

    If Expected = "" Then Exit Function
    If Not roles.Exists(role) Then Exit Function

    Dim d As Object
    Set d = roles(role)

    If d.Exists(Expected) Then
        RoleExactCount = CLng(d(Expected))
    End If

End Function


Function AllTextExactCount( _
    ByVal allTexts As Collection, _
    ByVal Expected As String) As Long

    Dim exp As String
    exp = Normalize(Expected)

    If exp = "" Then Exit Function

    Dim v As Variant

    For Each v In allTexts

        If Normalize(CStr(v)) = exp Then
            AllTextExactCount = AllTextExactCount + 1
        End If

    Next v

End Function



'=========================================================
' GREEN FAIL PANEL MARKER
'=========================================================

Private Sub MarkTypoFailPanelsGreen( _
    ByVal groupShape As Shape, _
    ByVal targetSize As String)

    On Error Resume Next

    If groupShape Is Nothing Then Exit Sub
    If Trim$(targetSize) = "" Then Exit Sub

    ScanAndMarkTypoPanels groupShape, H5_NormalizeSizeKey(targetSize)

    On Error GoTo 0

End Sub

Private Sub ScanAndMarkTypoPanels( _
    ByVal s As Shape, _
    ByVal targetSize As String)

    On Error Resume Next

    Dim c As Shape

    If s.Type = cdrGroupShape Then
        For Each c In s.Shapes
            ScanAndMarkTypoPanels c, targetSize
        Next c
        Exit Sub
    End If

    If s.Type <> cdrCurveShape Then Exit Sub
    If Not IsRed(s) Then Exit Sub

    Dim detected As String
    Dim baseSize As String
    Dim ph As Double
    Dim area As Double

    detected = DetectSizeFromShape(s, ph, area)

    If detected = "" Then Exit Sub

    baseSize = detected
    If InStr(1, baseSize, "_", vbTextCompare) > 0 Then
        baseSize = Left$(baseSize, InStr(1, baseSize, "_", vbTextCompare) - 1)
    End If

    If H5_NormalizeSizeKey(baseSize) <> targetSize Then Exit Sub

    ApplyTypoFailGreenOutline s

    On Error GoTo 0

End Sub

Private Sub ApplyTypoFailGreenOutline(ByVal shp As Shape)

    On Error Resume Next

    If shp Is Nothing Then Exit Sub
    If shp.Outline Is Nothing Then Exit Sub
    If shp.Outline.Type = cdrNoOutline Then Exit Sub

    shp.Outline.Color.RGBAssign _
        QTC_FAIL_PANEL_GREEN_R, _
        QTC_FAIL_PANEL_GREEN_G, _
        QTC_FAIL_PANEL_GREEN_B

    qtcMarkedPanels = qtcMarkedPanels + 1

    On Error GoTo 0

End Sub

'=========================================================
' REPORT
'=========================================================

Sub BuildFailReport( _
    ByVal idx As Long, _
    ByVal detectedSize As String, _
    ByVal currentPanel As Double, _
    ByVal scaleFactor As Double, _
    ByVal roles As Object, _
    ByVal allTexts As Collection, _
    ByVal nearest As Long)

    report = report & _
        "GROUP #" & idx & "  |  SIZE " & detectedSize & vbCrLf & _
        "STATUS : FAIL - TYPO / DATA PAIR MISMATCH" & vbCrLf & _
        "Panel  : " & Format(currentPanel, "0.000") & " cm" & vbCrLf & _
        "Scale  : " & Format(scaleFactor, "0.000") & vbCrLf

    If nearest > 0 Then

        Dim ord As Variant
        Dim expName As String
        Dim expNo As String
        Dim expNick As String
        Dim foundName As String
        Dim foundNo As String
        Dim foundNick As String

        ord = orders(nearest)

        expName = OrderField(ord, 1)
        expNo = OrderField(ord, 2)
        expNick = OrderField(ord, 3)

        foundName = SmartFoundName(roles, allTexts, expName)
        foundNo = SmartFoundNumber(roles, allTexts, expNo)
        foundNick = SmartFoundNickname(roles, allTexts, expNick)

        If isPants Then

            report = report & vbCrLf & _
                "EXPECTED : " & SafeDash(expNo) & vbCrLf & _
                "FOUND    : " & SafeDash(foundNo) & vbCrLf & vbCrLf

        Else

            report = report & vbCrLf & _
                "EXPECTED : " & FormatPair(expName, expNo, expNick) & vbCrLf & _
                "FOUND    : " & FormatPair(foundName, foundNo, foundNick) & vbCrLf & vbCrLf

        End If

        report = report & _
            "ORDER CANDIDATE" & vbCrLf & _
            "Size     : " & OrderField(ord, 0) & vbCrLf & _
            "Nama     : " & SafeDash(expName) & vbCrLf & _
            "Nomor    : " & SafeDash(expNo) & vbCrLf & _
            "Nickname : " & SafeDash(expNick) & vbCrLf & vbCrLf

        report = report & _
            "FOUND BY SENSOR" & vbCrLf & _
            "Nama kandidat : " & SafeDash(foundName) & vbCrLf & _
            "Nomor kandidat: " & SafeDash(foundNo) & vbCrLf & _
            "Nick kandidat : " & SafeDash(foundNick) & vbCrLf & vbCrLf

        report = report & _
            "FOUND BY ROLE" & vbCrLf & _
            "NAMA_ATLIT : " & RoleList(roles, "NAMA_ATLIT") & vbCrLf & _
            "NAMA       : " & RoleList(roles, "NAMA") & vbCrLf & _
            "NUMBER     : " & RoleList(roles, "NUMBER") & vbCrLf & _
            "NICKNAME   : " & RoleList(roles, "NICKNAME") & vbCrLf & vbCrLf

        report = report & _
            "ALL ACTIVE TEXT FOUND" & vbCrLf & _
            JoinCollectionInline(allTexts) & vbCrLf & vbCrLf

        report = report & _
            "DETAIL FAIL" & vbCrLf

        If isPants Then

            AppendRoleDiff roles, "NUMBER", "Nomor", expNo
            AppendOwnerHint detectedSize, "Nomor", foundNo, 2, nearest

        Else

            AppendPairDiff "Nama", expName, foundName
            AppendPairDiff "Nomor", expNo, foundNo
            AppendPairDiff "Nickname", expNick, foundNick

            AppendRoleDiff roles, "NAMA_ATLIT", "Nama Atlit", expName
            AppendRoleDiff roles, "NAMA", "Nama", expName
            AppendRoleDiff roles, "NUMBER", "Nomor", expNo
            AppendRoleDiff roles, "NICKNAME", "Nickname", expNick

            AppendOwnerHint detectedSize, "Nama", foundName, 1, nearest
            AppendOwnerHint detectedSize, "Nomor", foundNo, 2, nearest
            AppendOwnerHint detectedSize, "Nickname", foundNick, 3, nearest

        End If

    Else

        report = report & vbCrLf & _
            "EXPECTED : tidak ada kandidat Order.txt" & vbCrLf & _
            "FOUND    : " & JoinCollectionInline(allTexts) & vbCrLf & vbCrLf

    End If

    report = report & _
        String(48, "-") & vbCrLf & vbCrLf

End Sub


Private Sub AppendPairDiff( _
    ByVal label As String, _
    ByVal expectedValue As String, _
    ByVal foundValue As String)

    Dim exp As String
    Dim fnd As String

    exp = Normalize(expectedValue)
    fnd = Normalize(foundValue)

    If exp = "" And fnd = "" Then Exit Sub

    If exp <> fnd Then
        report = report & _
            label & " MISMATCH" & vbCrLf & _
            "Expected : " & SafeDash(expectedValue) & vbCrLf & _
            "Found    : " & SafeDash(foundValue) & vbCrLf & vbCrLf
    End If

End Sub


Private Sub AppendOwnerHint( _
    ByVal detectedSize As String, _
    ByVal label As String, _
    ByVal foundValue As String, _
    ByVal fieldIndex As Long, _
    ByVal expectedOrderIndex As Long)

    Dim fnd As String
    fnd = Normalize(foundValue)

    If fnd = "" Then Exit Sub

    Dim owner As Long
    owner = FindOrderOwnerByField(detectedSize, fieldIndex, fnd, expectedOrderIndex)

    If owner > 0 Then
        report = report & _
            "INDIKASI TERTUKAR" & vbCrLf & _
            label & " found '" & foundValue & "' cocok dengan order lain:" & vbCrLf & _
            "Owner    : " & FormatPair(OrderField(orders(owner), 1), OrderField(orders(owner), 2), OrderField(orders(owner), 3)) & vbCrLf & vbCrLf
    End If

End Sub


Private Function FindOrderOwnerByField( _
    ByVal detectedSize As String, _
    ByVal fieldIndex As Long, _
    ByVal foundValue As String, _
    ByVal expectedOrderIndex As Long) As Long

    Dim i As Long
    Dim ord As Variant
    Dim fnd As String

    fnd = Normalize(foundValue)

    If fnd = "" Then Exit Function

    For i = 1 To orders.Count

        If i = expectedOrderIndex Then GoTo NEXTI

        ord = orders(i)

        If H5_NormalizeSizeKey(OrderField(ord, 0)) <> H5_NormalizeSizeKey(detectedSize) Then
            GoTo NEXTI
        End If

        If Normalize(OrderField(ord, fieldIndex)) = fnd Then
            FindOrderOwnerByField = i
            Exit Function
        End If

NEXTI:

    Next i

End Function


Private Function FormatPair( _
    ByVal nm As String, _
    ByVal no As String, _
    ByVal nick As String) As String

    Dim r As String

    r = SafeDash(nm)

    If Trim$(no) <> "" Then
        r = r & ", " & no
    Else
        r = r & ", -"
    End If

    If Trim$(nick) <> "" Then
        r = r & " | " & nick
    End If

    FormatPair = r

End Function


Private Function SafeDash(ByVal s As String) As String

    If Trim$(s) = "" Then
        SafeDash = "-"
    Else
        SafeDash = s
    End If

End Function


Private Function SmartFoundName( _
    ByVal roles As Object, _
    ByVal allTexts As Collection, _
    ByVal expectedName As String) As String

    Dim s As String

    s = BestRoleTextForExpected(roles, "NAMA_ATLIT", expectedName)
    If s <> "" Then
        SmartFoundName = s
        Exit Function
    End If

    s = BestRoleTextForExpected(roles, "NAMA", expectedName)
    If s <> "" Then
        SmartFoundName = s
        Exit Function
    End If

    SmartFoundName = BestAllTextForExpected(expectedName, allTexts, False)

End Function


Private Function SmartFoundNumber( _
    ByVal roles As Object, _
    ByVal allTexts As Collection, _
    ByVal expectedNo As String) As String

    Dim s As String

    s = BestRoleTextForExpected(roles, "NUMBER", expectedNo)
    If s <> "" Then
        SmartFoundNumber = s
        Exit Function
    End If

    SmartFoundNumber = BestAllTextForExpected(expectedNo, allTexts, True)

End Function


Private Function SmartFoundNickname( _
    ByVal roles As Object, _
    ByVal allTexts As Collection, _
    ByVal expectedNick As String) As String

    Dim s As String

    s = BestRoleTextForExpected(roles, "NICKNAME", expectedNick)
    If s <> "" Then
        SmartFoundNickname = s
        Exit Function
    End If

    SmartFoundNickname = BestAllTextForExpected(expectedNick, allTexts, False)

End Function


Private Function BestRoleTextForExpected( _
    ByVal roles As Object, _
    ByVal role As String, _
    ByVal expectedValue As String) As String

    If Not roles.Exists(role) Then Exit Function

    Dim d As Object
    Dim k As Variant
    Dim best As String
    Dim bestScore As Double
    Dim sc As Double
    Dim exp As String

    Set d = roles(role)
    exp = Normalize(expectedValue)

    For Each k In d.keys

        If exp <> "" Then
            sc = StringSimilarityPercent(exp, Normalize(CStr(k)))
        Else
            sc = 1
        End If

        If best = "" Or sc > bestScore Then
            best = CStr(k)
            bestScore = sc
        End If

    Next k

    BestRoleTextForExpected = best

End Function


Private Function BestAllTextForExpected( _
    ByVal expectedValue As String, _
    ByVal allTexts As Collection, _
    ByVal numericOnly As Boolean) As String

    Dim v As Variant
    Dim s As String
    Dim exp As String
    Dim best As String
    Dim bestScore As Double
    Dim sc As Double

    exp = Normalize(expectedValue)

    For Each v In allTexts

        s = Normalize(CStr(v))

        If numericOnly Then
            If Not IsNumberText(s) Then GoTo NEXTV
        Else
            If IsNumberText(s) Then GoTo NEXTV
        End If

        If exp <> "" Then
            sc = StringSimilarityPercent(exp, s)
        Else
            sc = 1
        End If

        If best = "" Or sc > bestScore Then
            best = CStr(v)
            bestScore = sc
        End If

NEXTV:

    Next v

    BestAllTextForExpected = best

End Function


Sub AppendHardTextDiff( _
    ByVal allTexts As Collection, _
    ByVal label As String, _
    ByVal Expected As String)

    Dim exp As String
    exp = Normalize(Expected)

    If exp = "" Then Exit Sub

    If AllTextExactCount(allTexts, exp) = 0 Then

        report = report & _
            label & " TIDAK DITEMUKAN PERSIS" & vbCrLf & _
            "Expected : " & Expected & vbCrLf & _
            "Found    : " & JoinCollectionInline(allTexts) & vbCrLf & vbCrLf

    End If

End Sub


Sub AppendRoleDiff( _
    ByVal roles As Object, _
    ByVal role As String, _
    ByVal label As String, _
    ByVal Expected As String)

    Dim cntExpected As Long
    cntExpected = ExpectedRoleCount(role)

    If cntExpected <= 0 Then Exit Sub

    Dim exp As String
    exp = Normalize(Expected)

    Dim foundTotal As Long
    foundTotal = RoleTotalCount(roles, role)

    If exp = "" Then

        If CHECK_EMPTY_FIELDS And foundTotal > 0 Then

            report = report & _
                label & " SEHARUSNYA KOSONG" & vbCrLf & _
                "Expected : kosong" & vbCrLf & _
                "Found    : " & RoleList(roles, role) & vbCrLf & vbCrLf

        End If

        Exit Sub

    End If

    Dim exactCnt As Long
    exactCnt = RoleExactCount(roles, role, exp)

    If exactCnt <> cntExpected Or foundTotal <> cntExpected Then

        report = report & _
            label & " SALAH" & vbCrLf & _
            "Expected : " & Expected & " x" & cntExpected & vbCrLf & _
            "Found    : " & RoleList(roles, role) & vbCrLf & vbCrLf

    End If

End Sub


Function RoleList( _
    ByVal roles As Object, _
    ByVal role As String) As String

    If Not roles.Exists(role) Then
        RoleList = "-"
        Exit Function
    End If

    Dim d As Object
    Set d = roles(role)

    If d.Count = 0 Then
        RoleList = "-"
        Exit Function
    End If

    Dim r As String
    Dim k As Variant

    For Each k In d.keys

        If r <> "" Then
            r = r & ", "
        End If

        r = r & CStr(k) & " x" & d(k)

    Next k

    RoleList = r

End Function


Sub ShowFinalResult()

    If failCount = 0 Then

        qtcLastStatus = "PASS"

        If isPants Then
            qtcLastSummary = "QC Typo passed untuk mode celana."
            qtcLastDetail = _
                "PASS : " & PassCount & vbCrLf & _
                "FAIL : 0" & vbCrLf & _
                "Mode : CELANA" & vbCrLf & _
                "Validasi nomor aktif berjalan."
        Else
            qtcLastSummary = "QC Typo passed."
            qtcLastDetail = _
                "PASS : " & PassCount & vbCrLf & _
                "FAIL : 0" & vbCrLf & _
                "Mode : " & QTC_ProductModeText() & vbCrLf & _
                "Hard exact validation + template expected count aktif."
        End If

        If qtcReportMode Then Exit Sub

        If isPants Then

            MsgBox _
                "QC TYPO PASSED" & vbCrLf & vbCrLf & _
                "PASS : " & PassCount & vbCrLf & _
                "FAIL : 0" & vbCrLf & vbCrLf & _
                "Mode Celana: validasi nomor aktif.", _
                vbInformation

        Else

            MsgBox _
                "QC TYPO PASSED" & vbCrLf & vbCrLf & _
                "PASS : " & PassCount & vbCrLf & _
                "FAIL : 0" & vbCrLf & vbCrLf & _
                "Hard exact validation + template expected count aktif.", _
                vbInformation

        End If

    Else

        qtcLastStatus = "FAIL"
        qtcLastSummary = "Ada typo / mismatch data order pada hasil layout."
        qtcLastDetail = _
            "PASS : " & PassCount & vbCrLf & _
            "FAIL : " & failCount & vbCrLf & _
            "Marked panel hijau : " & qtcMarkedPanels & vbCrLf & _
            "Mode : " & QTC_ProductModeText() & vbCrLf & vbCrLf & _
            report

        If qtcReportMode Then Exit Sub

        MsgBox _
            "QC TYPO FAILED" & vbCrLf & vbCrLf & _
            "PASS : " & PassCount & vbCrLf & _
            "FAIL : " & failCount & vbCrLf & _
            "Marked panel hijau : " & qtcMarkedPanels & vbCrLf & vbCrLf & _
            report, _
            vbCritical

    End If

End Sub

Private Function QTC_ProductModeText() As String

    If isPants Then
        QTC_ProductModeText = "CELANA"
    ElseIf isSplitFront Then
        QTC_ProductModeText = "JAKET"
    Else
        QTC_ProductModeText = "JERSEY"
    End If

End Function


'=========================================================
' GROUP / UTIL
'=========================================================

Function CollectTopGroups(ByVal selectionObj As Object) As Collection

    Dim col As New Collection
    Dim s As Shape

    For Each s In selectionObj.Shapes

        If s.Type = cdrGroupShape Then
            col.Add s
        End If

    Next s

    Set CollectTopGroups = col

End Function


Function Normalize(ByVal s As String) As String

    Normalize = H5_NormalizeText(s)

End Function


Function IsNumberText(ByVal s As String) As Boolean

    IsNumberText = H5_IsNumericText(s, 3)

End Function


Function IgnoreSmallID( _
    ByVal t As Shape, _
    ByVal txt As String) As Boolean

    txt = Trim(txt)

    ' IDPO adalah wilayah IDPO Check, bukan QC Typo.
    If Normalize(txt) = "IDPO" Then
        If t.SizeHeight >= ID_MIN And _
           t.SizeHeight <= ID_MAX Then
            IgnoreSmallID = True
        End If
        Exit Function
    End If

    If Len(txt) <> 6 Then Exit Function
    If Not IsNumeric(txt) Then Exit Function

    If t.SizeHeight >= ID_MIN And _
       t.SizeHeight <= ID_MAX Then

        IgnoreSmallID = True

    End If

End Function


Function GetAlign(ByVal t As Shape) As String

    On Error Resume Next

    Dim a As Long
    a = t.Text.AlignProperties.Alignment

    Select Case a

        Case cdrLeftAlignment
            GetAlign = "LEFT"

        Case cdrRightAlignment
            GetAlign = "RIGHT"

        Case Else
            GetAlign = "CENTER"

    End Select

    On Error GoTo 0

End Function


Function JoinCollection(ByVal c As Collection) As String

    Dim r As String
    Dim v As Variant

    For Each v In c
        r = r & CStr(v) & vbCrLf
    Next v

    JoinCollection = r

End Function


Function JoinCollectionInline(ByVal c As Collection) As String

    Dim r As String
    Dim v As Variant

    For Each v In c

        If r <> "" Then
            r = r & ", "
        End If

        r = r & CStr(v)

    Next v

    If r = "" Then
        r = "-"
    End If

    JoinCollectionInline = r

End Function



'=========================================================
' HADES5 LOCAL COMPATIBILITY LAYER — V13.2 HOTFIX
'=========================================================
' Alasan:
' V13.1 memanggil helper H5_* dari Phase5 Core. Jika module
' Phase5 Core belum ter-import, VBA akan compile error:
'   Sub or Function not defined: H5_LoadTypoTemplate
'
' Layer ini membuat QC_TYPO_CHECK tetap self-contained.
' Nama shortcut lama tidak berubah.
'=========================================================

Private Function H5_DocumentsFile(ByVal fileName As String) As String
    H5_DocumentsFile = Environ$("USERPROFILE") & "\Documents\" & fileName
End Function

Private Function H5_ReadTextFileUTF8(ByVal path As String) As String

    On Error GoTo FALLBACK

    Dim stm As Object
    Set stm = CreateObject("ADODB.Stream")

    stm.Type = 2
    stm.CharSet = "utf-8"
    stm.Open
    stm.LoadFromFile path

    H5_ReadTextFileUTF8 = stm.ReadText(-1)

    stm.Close
    Set stm = Nothing

    Exit Function

FALLBACK:

    On Error Resume Next
    If Not stm Is Nothing Then
        stm.Close
        Set stm = Nothing
    End If
    On Error GoTo FAIL_ANSI

    H5_ReadTextFileUTF8 = H5_ReadTextFileANSI(path)
    Exit Function

FAIL_ANSI:
    Err.Raise vbObjectError + 5910, , _
        "Gagal membaca file sebagai UTF-8 maupun ANSI." & vbCrLf & path

End Function

Private Function H5_ReadTextFileANSI(ByVal path As String) As String

    Dim f As Integer
    Dim ln As String
    Dim buf As String

    f = FreeFile
    Open path For Input As #f

    Do Until EOF(f)
        Line Input #f, ln
        buf = buf & ln & vbLf
    Loop

    Close #f

    H5_ReadTextFileANSI = buf

End Function

Private Function H5_RemoveBOM(ByVal s As String) As String

    If Len(s) > 0 Then
        If AscW(Left$(s, 1)) = &HFEFF Then
            H5_RemoveBOM = Mid$(s, 2)
            Exit Function
        End If
    End If

    H5_RemoveBOM = s

End Function

Private Function H5_LoadKeyValueFile(ByVal path As String) As Object

    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")

    If Dir(path) = "" Then
        Set H5_LoadKeyValueFile = d
        Exit Function
    End If

    Dim content As String
    Dim lines As Variant
    Dim i As Long
    Dim ln As String
    Dim p As Long
    Dim k As String
    Dim v As String

    content = H5_ReadTextFileUTF8(path)
    content = Replace(content, vbCrLf, vbLf)
    content = Replace(content, vbCr, vbLf)

    lines = Split(content, vbLf)

    For i = LBound(lines) To UBound(lines)

        ln = H5_RemoveBOM(CStr(lines(i)))
        ln = Trim$(ln)

        If ln <> "" Then
            p = InStr(1, ln, "=", vbTextCompare)
            If p > 1 Then
                k = UCase$(Trim$(Left$(ln, p - 1)))
                v = Trim$(Mid$(ln, p + 1))
                If Len(k) > 0 Then
                    If d.Exists(k) Then
                        d(k) = v
                    Else
                        d.Add k, v
                    End If
                End If
            End If
        End If

    Next i

    Set H5_LoadKeyValueFile = d

End Function

Private Function H5_LoadTypoTemplate() As Object
    Set H5_LoadTypoTemplate = H5_LoadKeyValueFile(H5_DocumentsFile("TypoTemplate_Current.txt"))
End Function

Private Function H5_LoadOrderMeta() As Object

    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")

    Dim path As String
    path = H5_DocumentsFile("Order.txt")

    If Dir(path) = "" Then
        Set H5_LoadOrderMeta = d
        Exit Function
    End If

    Dim content As String
    Dim lines As Variant
    Dim i As Long
    Dim ln As String
    Dim p As Long
    Dim k As String
    Dim v As String

    content = H5_ReadTextFileUTF8(path)
    content = Replace(content, vbCrLf, vbLf)
    content = Replace(content, vbCr, vbLf)

    lines = Split(content, vbLf)

    For i = LBound(lines) To UBound(lines)

        ln = H5_RemoveBOM(CStr(lines(i)))
        ln = Trim$(ln)

        If Left$(ln, 1) = "@" Then
            p = InStr(1, ln, "=", vbTextCompare)
            If p > 2 Then
                k = UCase$(Trim$(Mid$(ln, 2, p - 2)))
                v = Trim$(Mid$(ln, p + 1))
                If Len(k) > 0 Then
                    If d.Exists(k) Then
                        d(k) = v
                    Else
                        d.Add k, v
                    End If
                End If
            End If
        End If

    Next i

    Set H5_LoadOrderMeta = d

End Function

Private Function H5_LoadOrderRows() As Collection

    Dim rows As New Collection
    Dim path As String
    path = H5_DocumentsFile("Order.txt")

    If Dir(path) = "" Then
        Err.Raise vbObjectError + 5920, , _
            "Order.txt tidak ditemukan di Documents." & vbCrLf & path
    End If

    Dim content As String
    Dim lines As Variant
    Dim i As Long
    Dim ln As String
    Dim arr As Variant
    Dim sz As String

    content = H5_ReadTextFileUTF8(path)
    content = Replace(content, vbCrLf, vbLf)
    content = Replace(content, vbCr, vbLf)

    lines = Split(content, vbLf)

    For i = LBound(lines) To UBound(lines)

        ln = H5_RemoveBOM(CStr(lines(i)))
        ln = Trim$(ln)

        If ln <> "" Then
            If Left$(ln, 1) <> "@" Then
                arr = Split(ln, "|")
                If UBound(arr) >= 3 Then
                    sz = H5_NormalizeSizeKey(CStr(arr(0)))
                    If sz <> "" Then rows.Add arr
                End If
            End If
        End If

    Next i

    Set H5_LoadOrderRows = rows

End Function

Private Function H5_LoadSizeDB( _
    ByVal fileName As String, _
    ByVal pantsMode As Boolean, _
    ByVal splitFrontMode As Boolean) As Object

    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")

    If Trim$(fileName) = "" Then
        Set H5_LoadSizeDB = d
        Exit Function
    End If

    Dim path As String
    path = H5_DocumentsFile(fileName)

    If Dir(path) = "" Then
        Set H5_LoadSizeDB = d
        Exit Function
    End If

    Dim f As Integer
    Dim ln As String
    Dim arr As Variant
    Dim sz As String

    f = FreeFile

    On Error GoTo FAIL

    Open path For Input As #f

    Do Until EOF(f)

        Line Input #f, ln
        ln = Trim$(H5_RemoveBOM(ln))

        If ln <> "" Then

            arr = Split(ln, "|")
            sz = ""

            If pantsMode Then
                If UBound(arr) >= 2 Then sz = H5_NormalizeSizeKey(CStr(arr(0)))
            ElseIf splitFrontMode Then
                If UBound(arr) >= 4 Then sz = H5_NormalizeSizeKey(CStr(arr(0)))
            Else
                If UBound(arr) >= 3 Then sz = H5_NormalizeSizeKey(CStr(arr(0)))
            End If

            If sz <> "" Then
                If d.Exists(sz) Then
                    d(sz) = arr
                Else
                    d.Add sz, arr
                End If
            End If

        End If

    Loop

    Close #f

    Set H5_LoadSizeDB = d
    Exit Function

FAIL:
    On Error Resume Next
    Close #f
    On Error GoTo 0
    Err.Raise vbObjectError + 5930, , _
        "Gagal membaca SizeDB: " & fileName & vbCrLf & Err.Description

End Function

Private Function H5_DetectCurrentSizeDBFileName() As String

    Dim meta As Object
    Dim tplLocal As Object
    Dim dbName As String

    On Error Resume Next

    Set tplLocal = H5_LoadTypoTemplate()
    If Not tplLocal Is Nothing Then
        If tplLocal.Exists("SIZEDB") Then dbName = Trim$(CStr(tplLocal("SIZEDB")))
        If dbName = "" And tplLocal.Exists("DB") Then dbName = Trim$(CStr(tplLocal("DB")))
    End If

    If dbName = "" Then
        Set meta = H5_LoadOrderMeta()
        If Not meta Is Nothing Then
            If meta.Exists("SIZEDB") Then dbName = Trim$(CStr(meta("SIZEDB")))
            If dbName = "" And meta.Exists("DB") Then dbName = Trim$(CStr(meta("DB")))
            If dbName = "" Then dbName = H5_InferDBFromOrderMeta(meta)
        End If
    End If

    If dbName <> "" Then
        If InStr(1, UCase$(dbName), ".TXT", vbTextCompare) = 0 Then dbName = dbName & ".txt"
    End If

    H5_DetectCurrentSizeDBFileName = dbName

    On Error GoTo 0

End Function

Private Function H5_InferDBFromOrderMeta(ByVal meta As Object) As String

    If meta Is Nothing Then Exit Function

    Dim jenis As String
    Dim pola As String
    Dim model As String
    Dim allText As String

    If meta.Exists("JENIS_PESANAN") Then jenis = UCase$(CStr(meta("JENIS_PESANAN")))
    If meta.Exists("JENIS_POLA") Then pola = UCase$(CStr(meta("JENIS_POLA")))
    If meta.Exists("MODEL_JAHIT") Then model = UCase$(CStr(meta("MODEL_JAHIT")))

    allText = jenis & " " & pola & " " & model

    If InStr(1, allText, "CELANA", vbTextCompare) > 0 Then
        If InStr(1, allText, "ANAK", vbTextCompare) > 0 Then
            H5_InferDBFromOrderMeta = "SizeDB_CelanaAnak.txt"
        ElseIf InStr(1, allText, "WANITA", vbTextCompare) > 0 Or _
               InStr(1, allText, "PEREMPUAN", vbTextCompare) > 0 Or _
               InStr(1, allText, "CEWEK", vbTextCompare) > 0 Then
            H5_InferDBFromOrderMeta = "SizeDB_CelanaWanita.txt"
        Else
            H5_InferDBFromOrderMeta = "SizeDB_CelanaPria.txt"
        End If
        Exit Function
    End If

    If InStr(1, allText, "JAKET", vbTextCompare) > 0 Then
        If InStr(1, allText, "ANAK", vbTextCompare) > 0 Then
            H5_InferDBFromOrderMeta = "SizeDB_JaketAnak.txt"
        Else
            H5_InferDBFromOrderMeta = "SizeDB_Jaket.txt"
        End If
        Exit Function
    End If

    If InStr(1, allText, "JERSEY", vbTextCompare) > 0 Then
        If InStr(1, allText, "ANAK", vbTextCompare) > 0 Then
            H5_InferDBFromOrderMeta = "SizeDB_Anak.txt"
            Exit Function
        End If

        If InStr(1, allText, "SLIM", vbTextCompare) > 0 Then
            If InStr(1, allText, "WANITA", vbTextCompare) > 0 Or _
               InStr(1, allText, "PEREMPUAN", vbTextCompare) > 0 Or _
               InStr(1, allText, "CEWEK", vbTextCompare) > 0 Then
                H5_InferDBFromOrderMeta = "SizeDB_WanitaSlimFit.txt"
            Else
                H5_InferDBFromOrderMeta = "SizeDB_PriaSlimFit.txt"
            End If
            Exit Function
        End If

        If InStr(1, allText, "WANITA", vbTextCompare) > 0 Or _
           InStr(1, allText, "PEREMPUAN", vbTextCompare) > 0 Or _
           InStr(1, allText, "CEWEK", vbTextCompare) > 0 Then
            H5_InferDBFromOrderMeta = "SizeDB_Wanita.txt"
        Else
            H5_InferDBFromOrderMeta = "SizeDB_Pria.txt"
        End If
        Exit Function
    End If

End Function

Private Sub H5_ProductModeFromDB( _
    ByVal dbName As String, _
    ByRef pantsMode As Boolean, _
    ByRef splitFrontMode As Boolean)

    Dim db As String
    db = UCase$(Trim$(dbName))

    pantsMode = False
    splitFrontMode = False

    If InStr(1, db, "CELANA", vbTextCompare) > 0 Then
        pantsMode = True
        splitFrontMode = False
        Exit Sub
    End If

    If InStr(1, db, "JAKET", vbTextCompare) > 0 Then
        pantsMode = False
        splitFrontMode = True
        Exit Sub
    End If

End Sub

Private Function H5_NormalizeSizeKey(ByVal sz As String) As String

    Dim s As String
    s = UCase$(Trim$(sz))

    Select Case s
        Case "XXL"
            H5_NormalizeSizeKey = "2XL"
        Case "XXXL"
            H5_NormalizeSizeKey = "3XL"
        Case "XXXXL"
            H5_NormalizeSizeKey = "4XL"
        Case "XXXXXL"
            H5_NormalizeSizeKey = "5XL"
        Case "XXXXXXL"
            H5_NormalizeSizeKey = "6XL"
        Case Else
            H5_NormalizeSizeKey = s
    End Select

End Function

Private Function H5_NormalizeText(ByVal s As String) As String

    s = Replace(s, vbCr, "")
    s = Replace(s, vbLf, "")
    s = Replace(s, Chr(160), " ")

    On Error Resume Next
    s = Replace(s, ChrW(&HFEFF), "")
    s = Replace(s, ChrW(&H200B), "")
    s = Replace(s, ChrW(&H200C), "")
    s = Replace(s, ChrW(&H200D), "")
    s = Replace(s, ChrW(&HFB00), "FF")
    s = Replace(s, ChrW(&HFB01), "FI")
    s = Replace(s, ChrW(&HFB02), "FL")
    s = Replace(s, ChrW(&HFB03), "FFI")
    s = Replace(s, ChrW(&HFB04), "FFL")
    s = Replace(s, ChrW(&HFB05), "ST")
    s = Replace(s, ChrW(&HFB06), "ST")
    On Error GoTo 0

    Do While InStr(1, s, "  ", vbTextCompare) > 0
        s = Replace(s, "  ", " ")
    Loop

    H5_NormalizeText = UCase$(Trim$(s))

End Function

Private Function H5_IsNumericText( _
    ByVal s As String, _
    Optional ByVal maxLen As Long = 0) As Boolean

    s = Trim$(s)

    If s = "" Then Exit Function
    If maxLen > 0 Then
        If Len(s) > maxLen Then Exit Function
    End If

    Dim i As Long
    Dim ch As String

    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch < "0" Or ch > "9" Then Exit Function
    Next i

    H5_IsNumericText = True

End Function

Private Function H5_ToDbl(ByVal v As Variant) As Double

    Dim s As String
    s = Trim$(CStr(v))
    s = Replace(s, ",", ".")

    H5_ToDbl = Val(s)

End Function

Private Function H5_IsRedRGB( _
    ByVal r As Long, _
    ByVal g As Long, _
    ByVal b As Long) As Boolean

    H5_IsRedRGB = (r > 200 And g < 90 And b < 90)

End Function

Private Function H5_IsGreenRGB( _
    ByVal r As Long, _
    ByVal g As Long, _
    ByVal b As Long) As Boolean

    H5_IsGreenRGB = (r <= 100 And g >= 160 And b <= 100)

End Function

Private Function H5_IsRedShape(ByVal shp As Shape) As Boolean

    On Error Resume Next

    Dim r As Long
    Dim g As Long
    Dim b As Long

    H5_IsRedShape = False

    If Not shp.Outline Is Nothing Then
        If shp.Outline.Type <> cdrNoOutline Then
            r = shp.Outline.Color.RGBRed
            g = shp.Outline.Color.RGBGreen
            b = shp.Outline.Color.RGBBlue

            If H5_IsRedRGB(r, g, b) Or H5_IsGreenRGB(r, g, b) Then
                H5_IsRedShape = True
                On Error GoTo 0
                Exit Function
            End If
        End If
    End If

    If Not shp.Fill Is Nothing Then
        If shp.Fill.Type <> cdrNoFill Then
            r = shp.Fill.UniformColor.RGBRed
            g = shp.Fill.UniformColor.RGBGreen
            b = shp.Fill.UniformColor.RGBBlue

            If H5_IsRedRGB(r, g, b) Or H5_IsGreenRGB(r, g, b) Then
                H5_IsRedShape = True
                On Error GoTo 0
                Exit Function
            End If
        End If
    End If

    On Error GoTo 0

End Function

Private Function H5_RunCoreSelfTest( _
    ByRef reportText As String, _
    ByRef failCountLocal As Long, _
    ByRef warnCount As Long) As Boolean

    reportText = "QC Typo V13.2 local compatibility self-test" & vbCrLf
    failCountLocal = 0
    warnCount = 0

    If Dir(H5_DocumentsFile("Order.txt")) = "" Then
        warnCount = warnCount + 1
        reportText = reportText & "WARNING: Order.txt belum ditemukan." & vbCrLf
    Else
        reportText = reportText & "OK: Order.txt ditemukan." & vbCrLf
    End If

    If Dir(H5_DocumentsFile("TypoTemplate_Current.txt")) = "" Then
        warnCount = warnCount + 1
        reportText = reportText & "WARNING: TypoTemplate_Current.txt belum ditemukan." & vbCrLf
    Else
        reportText = reportText & "OK: TypoTemplate_Current.txt ditemukan." & vbCrLf
    End If

    H5_RunCoreSelfTest = True

End Function

Private Sub H5_WriteCoreSelfTestReport(ByVal reportText As String)

    On Error Resume Next

    Dim folder As String
    Dim path As String
    Dim f As Integer

    folder = Environ$("USERPROFILE") & "\Documents\HADES_REPORTS"
    If Dir(folder, vbDirectory) = "" Then MkDir folder

    path = folder & "\HADES_CORE_SELF_TEST_LATEST.txt"

    f = FreeFile
    Open path For Output As #f
    Print #f, reportText
    Close #f

    On Error GoTo 0

End Sub