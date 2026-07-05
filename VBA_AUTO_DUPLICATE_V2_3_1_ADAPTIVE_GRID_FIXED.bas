Option Explicit

'=========================================================
' PROJECT HADES — AUTO DUPLICATE V2.3.1 ADAPTIVE GRID FIXED
'
' BASE:
' - HADES AUTO DUPLICATE V2.1
'
' PURPOSE:
' - Membuat HASIL LAYOUT dari MASTER LAYOUT.
' - Membaca Order.txt untuk quantity per size.
' - Membaca SizeDB untuk mendeteksi size source/master layout.
' - Menggandakan source sesuai quantity Order.txt.
' - Menata output otomatis di ATAS ACTIVE PAGE / PAGE PUTIH:
'   - duplicate nomor 1 tiap size center terhadap ActivePage
'   - duplicate berikutnya ke kanan
'   - maksimal 8 pcs ke samping
'   - jika lebih dari 8, turun baris
'   - size kecil di atas, size besar di bawah
'
' FITUR V2.2:
' - Membaca metadata Order.txt:
'
'   @JENIS_PESANAN=JERSEY
'   @JENIS_POLA=JERSEY REGULER
'   @MODEL_JAHIT=DEWASA PRIA
'   @SIZEDB=SizeDB_Pria.txt
'
' - Jika @SIZEDB ada:
'   popup database dilewati.
'
' - Jika @SIZEDB tidak ada:
'   popup lama tetap muncul.
'
' - Baris metadata @... diabaikan saat menghitung quantity.
'
' - SELECT ALL MASTER SIZE SAFE:
'   Boleh select master S, M, L, XL, 2XL.
'   Jika Order.txt hanya M, L, XL,
'   maka S dan 2XL diabaikan.
'
' - Source yang wajib ada hanya size yang muncul di Order.txt.
'
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
' MAIN MACRO:
' HADES_AUTO_DUPLICATE_V21
'
' ALIAS:
' AUTO_DUPLICATE_V2
' AUTO_DUPLICATE_V21
' AUTO_DUPLICATE_V22
'=========================================================


'=========================================================
' GLOBAL VARIABLES
'=========================================================

Private AD_OrderQtyBySize As Object
Private AD_SizeDB As Object
Private AD_SourceBySize As Object
Private AD_OrderMeta As Object

Private AD_CurrentDB As String
Private AD_DBSource As String

Private AD_IsSplitFront As Boolean
Private AD_IsPants As Boolean

Private AD_Report As String
Private AD_Warning As String
Private AD_PageAnchorMethod As String

Private Const AD_SIZE_TOL As Double = 1#
Private Const AD_PANTS_TOL As Double = 0.35
Private Const AD_MAX_COL As Long = 8

'Adaptive grid thresholds - FIX V2.3.1
'V2.3 sebelumnya memakai konstanta ini di AD_CalcAdaptiveGrid,
'tetapi konstanta belum dideklarasikan sehingga compile error:
'Variable not defined.
Private Const AD_GRID_SINGLE_ROW_MAX As Long = 7
Private Const AD_GRID_TWO_ROW_MAX As Long = 14
Private Const AD_GRID_BALANCED_MAX As Long = 30
Private Const AD_GRID_BALANCED_TARGET_COL As Long = 6

'Jarak antar baju dalam satu row
Private Const AD_GAP_X As Double = 15.5

'Jarak antar row dalam size yang sama
Private Const AD_GAP_Y As Double = 15.5

'Jarak antar blok size, misal M ke L
Private Const AD_BLOCK_GAP_Y As Double = 20#

'Jarak bottom Hasil Layout dari atas page putih
Private Const AD_PAGE_TOP_OFFSET As Double = 210#

'Koreksi manual jika masih perlu geser sedikit.
'Positif X = geser kanan. Negatif X = geser kiri.
'Positif Y = geser atas.  Negatif Y = geser bawah.
Private Const AD_PAGE_CENTER_X_CORRECTION As Double = 0#
Private Const AD_PAGE_TOP_Y_CORRECTION As Double = 0#

'Warna hijau lama / alternatif
Private Const AD_GREEN_R As Long = 97
Private Const AD_GREEN_G As Long = 186
Private Const AD_GREEN_B As Long = 12
Private Const AD_GREEN_TOL As Long = 18


'=========================================================
' PUBLIC ENTRY
'=========================================================

Sub AUTO_DUPLICATE_V2()
    Call HADES_AUTO_DUPLICATE_V21
End Sub

Sub AUTO_DUPLICATE_V21()
    Call HADES_AUTO_DUPLICATE_V21
End Sub

Sub AUTO_DUPLICATE_V22()
    Call HADES_AUTO_DUPLICATE_V21
End Sub

Sub AUTO_DUPLICATE_V23()
    Call HADES_AUTO_DUPLICATE_V21
End Sub

Sub HADES_AUTO_DUPLICATE_V21()

    Dim oldUnit As Long
    Dim cmdStarted As Boolean
    Dim sr As ShapeRange

    oldUnit = ActiveDocument.Unit
    cmdStarted = False

    On Error GoTo ERR_HANDLER

    ActiveDocument.Unit = cdrCentimeter

    On Error Resume Next
    Set sr = ActiveSelectionRange
    On Error GoTo ERR_HANDLER

    If sr Is Nothing Then
        MsgBox "Pilih MASTER LAYOUT source terlebih dahulu.", vbExclamation, "HADES AUTO DUPLICATE V2.3"
        GoTo EXIT_CLEAN
    End If

    If sr.Count = 0 Then
        MsgBox "Pilih MASTER LAYOUT source terlebih dahulu.", vbExclamation, "HADES AUTO DUPLICATE V2.3"
        GoTo EXIT_CLEAN
    End If

    Set AD_OrderQtyBySize = CreateObject("Scripting.Dictionary")
    Set AD_SizeDB = CreateObject("Scripting.Dictionary")
    Set AD_SourceBySize = CreateObject("Scripting.Dictionary")
    Set AD_OrderMeta = CreateObject("Scripting.Dictionary")

    AD_CurrentDB = ""
    AD_DBSource = ""
    AD_IsSplitFront = False
    AD_IsPants = False

    AD_Report = ""
    AD_Warning = ""
    AD_PageAnchorMethod = ""

    '=====================================================
    ' V2.2:
    ' Order dibaca dulu supaya @SIZEDB bisa dipakai
    ' untuk menentukan SizeDB tanpa popup.
    '=====================================================
    AD_LoadOrders

    If Len(Trim$(AD_CurrentDB)) > 0 Then

        AD_ConfigureModeFromDB
        AD_DBSource = "AUTO dari Order.txt @SIZEDB"

    Else

        AD_CurrentDB = AD_InferDBFromMetadata()

        If Len(Trim$(AD_CurrentDB)) > 0 Then

            AD_ConfigureModeFromDB
            AD_DBSource = "AUTO dari metadata spesifikasi Order.txt"

        Else

            If Not AD_SelectDatabaseFallback Then
                GoTo EXIT_CLEAN
            End If

            AD_DBSource = "MANUAL POPUP"

        End If

    End If

    AD_LoadSizeDB AD_CurrentDB

    Dim sources As Collection
    Set sources = AD_CollectSelectedSources(sr)

    If sources.Count = 0 Then

        MsgBox _
            "Tidak ada source yang valid pada selection." & vbCrLf & vbCrLf & _
            "Pilih group MASTER LAYOUT yang akan digandakan.", _
            vbExclamation, _
            "HADES AUTO DUPLICATE V2.3"

        GoTo EXIT_CLEAN

    End If

    If Not AD_PreflightSources(sources) Then

        MsgBox _
            "AUTO DUPLICATE DIBATALKAN" & vbCrLf & vbCrLf & _
            "Preflight gagal. Tidak ada objek yang digandakan." & vbCrLf & vbCrLf & _
            AD_Report, _
            vbCritical, _
            "HADES AUTO DUPLICATE V2.3"

        GoTo EXIT_CLEAN

    End If

    Dim orderedSizes As Variant
    orderedSizes = AD_GetSortedOrderSizes()

    Dim pageCenterX As Double
    Dim pageTopY As Double

    AD_GetAccuratePageAnchor pageCenterX, pageTopY

    pageCenterX = pageCenterX + AD_PAGE_CENTER_X_CORRECTION
    pageTopY = pageTopY + AD_PAGE_TOP_Y_CORRECTION

    If Not AD_ShowPreview(orderedSizes, pageCenterX, pageTopY, sources.Count) Then
        GoTo EXIT_CLEAN
    End If

    Dim totalLayoutHeight As Double
    totalLayoutHeight = AD_TotalLayoutHeight(orderedSizes)

    Dim currentTop As Double

    '=====================================================
    ' OUTPUT ANCHOR:
    '
    ' Hasil Layout diletakkan di atas ActivePage.
    ' Bottom keseluruhan hasil layout berada:
    ' pageTopY + AD_PAGE_TOP_OFFSET
    '
    ' Karena size kecil di atas dan size besar di bawah,
    ' currentTop dimulai dari titik paling atas keseluruhan layout.
    '=====================================================

    currentTop = pageTopY + AD_PAGE_TOP_OFFSET + totalLayoutHeight

    Dim createdCount As Long
    createdCount = 0

    ActiveDocument.BeginCommandGroup "HADES Auto Duplicate V2.3"
    cmdStarted = True

    Dim k As Variant
    Dim sz As String
    Dim q As Long
    Dim src As Shape

    For Each k In orderedSizes

        sz = CStr(k)
        q = CLng(AD_OrderQtyBySize(sz))

        Set src = AD_SourceBySize(sz)

        AD_DuplicateOneSizeBlock _
            src, _
            q, _
            pageCenterX, _
            currentTop, _
            createdCount

        currentTop = currentTop - AD_BlockHeight(src, q) - AD_BLOCK_GAP_Y

    Next k

    ActiveDocument.EndCommandGroup
    cmdStarted = False

    ActiveWindow.Refresh

    Dim msg As String

    msg = "AUTO DUPLICATE SELESAI" & vbCrLf & vbCrLf & _
          "Database       : " & AD_CurrentDB & vbCrLf & _
          "DB Source      : " & AD_DBSource & vbCrLf & _
          "Mode           : " & AD_ProductModeText() & vbCrLf & _
          "Source dipilih : " & sources.Count & vbCrLf & _
          "Source dipakai : " & AD_SourceBySize.Count & vbCrLf & _
          "Output dibuat  : " & createdCount & " duplicate" & vbCrLf & _
          "Anchor method  : " & AD_PageAnchorMethod & vbCrLf & vbCrLf & _
          "Posisi output:" & vbCrLf & _
          "- Hasil Layout dibuat di atas ActivePage / page putih." & vbCrLf & _
          "- Duplicate nomor 1 tiap size center terhadap page." & vbCrLf & vbCrLf & _
          "Aturan layout:" & vbCrLf & _
          "- Grid adaptif per quantity size." & vbCrLf & _
          "- Qty 1-7: satu baris." & vbCrLf & _
          "- Qty 8-14: dua baris seimbang." & vbCrLf & _
          "- Qty 15-30: balanced grid target 5-6 kolom." & vbCrLf & _
          "- Qty 31-60: maksimal " & AD_MAX_COL & " pcs ke samping, lalu turun." & vbCrLf & _
          "- Source/master asli tidak dipindah dan tidak dihitung." & vbCrLf & _
          "- Source size yang tidak ada di Order.txt otomatis diabaikan." & vbCrLf & vbCrLf & _
          "Langkah berikutnya:" & vbCrLf & _
          "1. Select HASIL LAYOUT output." & vbCrLf & _
          "2. Run QC_AUTO_RENAME." & vbCrLf & _
          "3. Run QC_SIZE_CHECK untuk validasi final."

    If Len(AD_Warning) > 0 Then
        msg = msg & vbCrLf & vbCrLf & _
              "WARNING:" & vbCrLf & _
              AD_Warning
        MsgBox msg, vbExclamation, "HADES AUTO DUPLICATE V2.3"
    Else
        MsgBox msg, vbInformation, "HADES AUTO DUPLICATE V2.3"
    End If

EXIT_CLEAN:

    On Error Resume Next

    If cmdStarted Then ActiveDocument.EndCommandGroup

    ActiveDocument.Unit = oldUnit
    ActiveWindow.Refresh

    On Error GoTo 0
    Exit Sub

ERR_HANDLER:

    Dim eNo As Long
    Dim eDesc As String

    eNo = Err.Number
    eDesc = Err.Description

    On Error Resume Next

    If cmdStarted Then ActiveDocument.EndCommandGroup

    ActiveDocument.Unit = oldUnit
    ActiveWindow.Refresh

    On Error GoTo 0

    MsgBox _
        "SYSTEM ERROR - HADES AUTO DUPLICATE V2.3" & vbCrLf & vbCrLf & _
        "No : " & eNo & vbCrLf & _
        eDesc, _
        vbCritical, _
        "HADES AUTO DUPLICATE V2.3"

End Sub


'=========================================================
' DATABASE AUTO / FALLBACK
'=========================================================

Private Sub AD_ConfigureModeFromDB()

    Dim db As String

    db = UCase$(Trim$(AD_CurrentDB))

    AD_IsSplitFront = False
    AD_IsPants = False

    If InStr(1, db, "CELANA", vbTextCompare) > 0 Then
        AD_IsPants = True
        AD_IsSplitFront = False
        Exit Sub
    End If

    If InStr(1, db, "JAKET", vbTextCompare) > 0 Then
        AD_IsSplitFront = True
        AD_IsPants = False
        Exit Sub
    End If

    'Default selain Celana/Jaket dianggap Jersey
    AD_IsSplitFront = False
    AD_IsPants = False

End Sub

Private Function AD_InferDBFromMetadata() As String

    Dim jenis As String
    Dim pola As String
    Dim model As String
    Dim allText As String

    jenis = UCase$(AD_GetMeta("JENIS_PESANAN"))
    pola = UCase$(AD_GetMeta("JENIS_POLA"))
    model = UCase$(AD_GetMeta("MODEL_JAHIT"))

    allText = jenis & " " & pola & " " & model

    AD_InferDBFromMetadata = ""

    If InStr(1, allText, "CELANA", vbTextCompare) > 0 Then

        If InStr(1, allText, "ANAK", vbTextCompare) > 0 Then
            AD_InferDBFromMetadata = "SizeDB_CelanaAnak.txt"
        ElseIf InStr(1, allText, "WANITA", vbTextCompare) > 0 Or _
               InStr(1, allText, "PEREMPUAN", vbTextCompare) > 0 Or _
               InStr(1, allText, "CEWEK", vbTextCompare) > 0 Then
            AD_InferDBFromMetadata = "SizeDB_CelanaWanita.txt"
        Else
            AD_InferDBFromMetadata = "SizeDB_CelanaPria.txt"
        End If

        Exit Function

    End If

    If InStr(1, allText, "JAKET", vbTextCompare) > 0 Then

        If InStr(1, allText, "ANAK", vbTextCompare) > 0 Then
            AD_InferDBFromMetadata = "SizeDB_JaketAnak.txt"
        Else
            AD_InferDBFromMetadata = "SizeDB_Jaket.txt"
        End If

        Exit Function

    End If

    If InStr(1, allText, "JERSEY", vbTextCompare) > 0 Then

        If InStr(1, allText, "ANAK", vbTextCompare) > 0 Then
            AD_InferDBFromMetadata = "SizeDB_Anak.txt"
            Exit Function
        End If

        If InStr(1, allText, "SLIM", vbTextCompare) > 0 Then

            If InStr(1, allText, "WANITA", vbTextCompare) > 0 Or _
               InStr(1, allText, "PEREMPUAN", vbTextCompare) > 0 Or _
               InStr(1, allText, "CEWEK", vbTextCompare) > 0 Then
                AD_InferDBFromMetadata = "SizeDB_WanitaSlimFit.txt"
            Else
                AD_InferDBFromMetadata = "SizeDB_PriaSlimFit.txt"
            End If

            Exit Function

        End If

        If InStr(1, allText, "WANITA", vbTextCompare) > 0 Or _
           InStr(1, allText, "PEREMPUAN", vbTextCompare) > 0 Or _
           InStr(1, allText, "CEWEK", vbTextCompare) > 0 Then
            AD_InferDBFromMetadata = "SizeDB_Wanita.txt"
        Else
            AD_InferDBFromMetadata = "SizeDB_Pria.txt"
        End If

        Exit Function

    End If

End Function

Private Function AD_SelectDatabaseFallback() As Boolean

    AD_SelectDatabaseFallback = False

    Dim a As String
    Dim b As String

    AD_IsSplitFront = False
    AD_IsPants = False
    AD_CurrentDB = ""

    a = InputBox( _
        "Order.txt belum memiliki @SIZEDB." & vbCrLf & vbCrLf & _
        "PILIH PRODUK / POLA" & vbCrLf & vbCrLf & _
        "1 = JERSEY" & vbCrLf & _
        "2 = JAKET" & vbCrLf & _
        "3 = CELANA", _
        "HADES AUTO DUPLICATE V2.3")

    If Trim$(a) = "" Then Exit Function

    Select Case Trim$(a)

        Case "1"

            AD_IsSplitFront = False
            AD_IsPants = False

            b = InputBox( _
                "PILIH DATABASE JERSEY" & vbCrLf & vbCrLf & _
                "1 = PRIA REGULAR" & vbCrLf & _
                "2 = WANITA REGULAR" & vbCrLf & _
                "3 = ANAK" & vbCrLf & _
                "4 = PRIA SLIM FIT" & vbCrLf & _
                "5 = WANITA SLIM FIT", _
                "HADES AUTO DUPLICATE V2.3")

            If Trim$(b) = "" Then Exit Function

            Select Case Trim$(b)

                Case "1"
                    AD_CurrentDB = "SizeDB_Pria.txt"

                Case "2"
                    AD_CurrentDB = "SizeDB_Wanita.txt"

                Case "3"
                    AD_CurrentDB = "SizeDB_Anak.txt"

                Case "4"
                    AD_CurrentDB = "SizeDB_PriaSlimFit.txt"

                Case "5"
                    AD_CurrentDB = "SizeDB_WanitaSlimFit.txt"

                Case Else
                    MsgBox "Pilihan database tidak valid.", vbExclamation, "HADES AUTO DUPLICATE V2.3"
                    Exit Function

            End Select

        Case "2"

            AD_IsSplitFront = True
            AD_IsPants = False

            b = InputBox( _
                "PILIH DATABASE JAKET" & vbCrLf & vbCrLf & _
                "1 = JAKET DEWASA" & vbCrLf & _
                "2 = JAKET ANAK", _
                "HADES AUTO DUPLICATE V2.3")

            If Trim$(b) = "" Then Exit Function

            Select Case Trim$(b)

                Case "1"
                    AD_CurrentDB = "SizeDB_Jaket.txt"

                Case "2"
                    AD_CurrentDB = "SizeDB_JaketAnak.txt"

                Case Else
                    MsgBox "Pilihan database tidak valid.", vbExclamation, "HADES AUTO DUPLICATE V2.3"
                    Exit Function

            End Select

        Case "3"

            AD_IsSplitFront = False
            AD_IsPants = True

            b = InputBox( _
                "PILIH DATABASE CELANA" & vbCrLf & vbCrLf & _
                "1 = CELANA PRIA" & vbCrLf & _
                "2 = CELANA WANITA" & vbCrLf & _
                "3 = CELANA ANAK", _
                "HADES AUTO DUPLICATE V2.3")

            If Trim$(b) = "" Then Exit Function

            Select Case Trim$(b)

                Case "1"
                    AD_CurrentDB = "SizeDB_CelanaPria.txt"

                Case "2"
                    AD_CurrentDB = "SizeDB_CelanaWanita.txt"

                Case "3"
                    AD_CurrentDB = "SizeDB_CelanaAnak.txt"

                Case Else
                    MsgBox "Pilihan database tidak valid.", vbExclamation, "HADES AUTO DUPLICATE V2.3"
                    Exit Function

            End Select

        Case Else

            MsgBox "Pilihan produk tidak valid.", vbExclamation, "HADES AUTO DUPLICATE V2.3"
            Exit Function

    End Select

    AD_SelectDatabaseFallback = True

End Function


'=========================================================
' LOAD ORDER — UTF-8 SAFE + METADATA
'=========================================================

Private Sub AD_LoadOrders()

    Set AD_OrderQtyBySize = CreateObject("Scripting.Dictionary")
    Set AD_OrderMeta = CreateObject("Scripting.Dictionary")

    Dim path As String
    path = Environ$("USERPROFILE") & "\Documents\Order.txt"

    If Dir(path) = "" Then
        Err.Raise vbObjectError + 100, , _
            "Order.txt tidak ditemukan di Documents." & vbCrLf & path
    End If

    On Error GoTo FAIL

    Dim content As String
    content = AD_ReadTextFileUTF8(path)

    content = Replace(content, vbCrLf, vbLf)
    content = Replace(content, vbCr, vbLf)

    Dim lines As Variant
    lines = Split(content, vbLf)

    Dim i As Long
    Dim ln As String
    Dim arr As Variant
    Dim sz As String

    For i = LBound(lines) To UBound(lines)

        ln = CStr(lines(i))
        ln = AD_RemoveBOM(ln)
        ln = Trim$(ln)

        If ln <> "" Then

            If Left$(ln, 1) = "@" Then

                AD_ParseMetaLine ln

            Else

                arr = Split(ln, "|")

                If UBound(arr) >= 3 Then

                    sz = AD_NormalizeSizeKey(CStr(arr(0)))

                    If AD_IsStandardSize(sz) Then
                        AD_AddOrderQty sz
                    End If

                End If

            End If

        End If

    Next i

    If AD_OrderMeta.Exists("SIZEDB") Then
        If Len(Trim$(CStr(AD_OrderMeta("SIZEDB")))) > 0 Then
            AD_CurrentDB = Trim$(CStr(AD_OrderMeta("SIZEDB")))
        End If
    End If

    If AD_OrderQtyBySize.Count = 0 Then
        Err.Raise vbObjectError + 101, , _
            "Order.txt kosong atau format tidak valid." & vbCrLf & _
            "Format wajib: SIZE|NAMA|NOMOR|NICKNAME" & vbCrLf & vbCrLf & _
            "Baris metadata @... boleh ada dan akan diabaikan."
    End If

    Exit Sub

FAIL:

    Err.Raise vbObjectError + 102, , _
        "Gagal membaca Order.txt." & vbCrLf & Err.Description

End Sub

Private Sub AD_ParseMetaLine(ByVal line As String)

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

    If AD_OrderMeta.Exists(k) Then
        AD_OrderMeta(k) = v
    Else
        AD_OrderMeta.Add k, v
    End If

End Sub

Private Function AD_GetMeta(ByVal keyName As String) As String

    keyName = UCase$(Trim$(keyName))

    If AD_OrderMeta Is Nothing Then
        AD_GetMeta = ""
        Exit Function
    End If

    If AD_OrderMeta.Exists(keyName) Then
        AD_GetMeta = CStr(AD_OrderMeta(keyName))
    Else
        AD_GetMeta = ""
    End If

End Function

Private Function AD_ReadTextFileUTF8(ByVal path As String) As String

    On Error GoTo Fallback

    Dim stm As Object
    Set stm = CreateObject("ADODB.Stream")

    stm.Type = 2
    stm.Charset = "utf-8"
    stm.Open
    stm.LoadFromFile path

    AD_ReadTextFileUTF8 = stm.ReadText(-1)

    stm.Close
    Set stm = Nothing

    Exit Function

Fallback:

    On Error Resume Next

    If Not stm Is Nothing Then
        stm.Close
        Set stm = Nothing
    End If

    On Error GoTo FAIL_ANSI

    AD_ReadTextFileUTF8 = AD_ReadTextFileANSI(path)
    Exit Function

FAIL_ANSI:

    Err.Raise vbObjectError + 900, , _
        "Gagal membaca file sebagai UTF-8 maupun ANSI." & vbCrLf & path

End Function

Private Function AD_ReadTextFileANSI(ByVal path As String) As String

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

    AD_ReadTextFileANSI = buf

End Function

Private Function AD_RemoveBOM(ByVal s As String) As String

    If Len(s) > 0 Then
        If AscW(Left$(s, 1)) = &HFEFF Then
            AD_RemoveBOM = Mid$(s, 2)
            Exit Function
        End If
    End If

    AD_RemoveBOM = s

End Function

Private Sub AD_AddOrderQty(ByVal sz As String)

    sz = AD_NormalizeSizeKey(sz)

    If AD_OrderQtyBySize.Exists(sz) Then
        AD_OrderQtyBySize(sz) = CLng(AD_OrderQtyBySize(sz)) + 1
    Else
        AD_OrderQtyBySize.Add sz, 1
    End If

End Sub


'=========================================================
' LOAD SIZE DB
'=========================================================

Private Sub AD_LoadSizeDB(ByVal fileName As String)

    Set AD_SizeDB = CreateObject("Scripting.Dictionary")

    Dim path As String
    path = Environ$("USERPROFILE") & "\Documents\" & fileName

    If Dir(path) = "" Then
        Err.Raise vbObjectError + 200, , _
            "SizeDB tidak ditemukan:" & vbCrLf & path & vbCrLf & vbCrLf & _
            "Database dipilih dari: " & AD_DBSource
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
        ln = Trim$(ln)

        If ln <> "" Then

            arr = Split(ln, "|")

            If AD_IsPants Then

                If UBound(arr) >= 2 Then
                    sz = AD_NormalizeSizeKey(CStr(arr(0)))
                    If AD_IsStandardSize(sz) Then AD_PutSizeDB sz, arr
                End If

            ElseIf AD_IsSplitFront Then

                If UBound(arr) >= 4 Then
                    sz = AD_NormalizeSizeKey(CStr(arr(0)))
                    If AD_IsStandardSize(sz) Then AD_PutSizeDB sz, arr
                End If

            Else

                If UBound(arr) >= 3 Then
                    sz = AD_NormalizeSizeKey(CStr(arr(0)))
                    If AD_IsStandardSize(sz) Then AD_PutSizeDB sz, arr
                End If

            End If

        End If

    Loop

    Close #f

    If AD_SizeDB.Count = 0 Then

        If AD_IsPants Then
            Err.Raise vbObjectError + 201, , _
                "SizeDB Celana kosong atau format tidak valid." & vbCrLf & _
                "Format wajib: SIZE|L_DEPAN|L_BELAKANG"
        ElseIf AD_IsSplitFront Then
            Err.Raise vbObjectError + 201, , _
                "SizeDB Jaket kosong atau format tidak valid." & vbCrLf & _
                "Format wajib: SIZE|L_BELAKANG|L_DEPAN|T_DEPAN|T_BELAKANG"
        Else
            Err.Raise vbObjectError + 201, , _
                "SizeDB Jersey kosong atau format tidak valid." & vbCrLf & _
                "Format wajib: SIZE|LEBAR|TINGGI_DEPAN|TINGGI_BELAKANG"
        End If

    End If

    Exit Sub

FAIL:

    On Error Resume Next
    Close #f
    On Error GoTo 0

    Err.Raise vbObjectError + 202, , _
        "Gagal membaca SizeDB: " & fileName & vbCrLf & Err.Description

End Sub

Private Sub AD_PutSizeDB(ByVal sz As String, ByVal arr As Variant)

    sz = AD_NormalizeSizeKey(sz)

    If AD_SizeDB.Exists(sz) Then
        AD_SizeDB(sz) = arr
    Else
        AD_SizeDB.Add sz, arr
    End If

End Sub


'=========================================================
' COLLECT SOURCE SELECTION
'=========================================================

Private Function AD_CollectSelectedSources(ByVal sr As ShapeRange) As Collection

    Dim col As New Collection
    Dim s As Shape

    For Each s In sr

        If s.Type = cdrGroupShape Then
            col.Add s
        Else
            'Tetap boleh ditambahkan kalau user select shape tunggal,
            'tetapi workflow normal adalah group master per size.
            col.Add s
        End If

    Next s

    Set AD_CollectSelectedSources = col

End Function


'=========================================================
' PREFLIGHT SOURCE
'=========================================================

Private Function AD_PreflightSources(ByVal sources As Collection) As Boolean

    AD_PreflightSources = False

    Set AD_SourceBySize = CreateObject("Scripting.Dictionary")
    AD_Report = ""
    AD_Warning = ""

    Dim i As Long
    Dim src As Shape
    Dim sz As String

    For i = 1 To sources.Count

        Set src = sources(i)
        sz = AD_DetectSourceSize(src)

        If sz = "" Then

            'Select all safe:
            'Source yang tidak terdeteksi tidak langsung menggagalkan macro.
            'Jika memang size itu dibutuhkan Order.txt, error akan muncul
            'di bagian "SOURCE MASTER LAYOUT TIDAK DITEMUKAN".
            AD_Warning = AD_Warning & _
                "SOURCE #" & i & " tidak terdeteksi size-nya dan diabaikan." & vbCrLf

        Else

            sz = AD_NormalizeSizeKey(sz)

            '=================================================
            ' SELECT ALL MASTER SIZE SAFE:
            '
            ' Source yang tidak ada di Order.txt diabaikan.
            ' Contoh:
            ' Master selected : S, M, L, XL, 2XL
            ' Order.txt       : M, L, XL
            ' S dan 2XL       : ignore
            '=================================================
            If AD_OrderQtyBySize.Exists(sz) Then

                If AD_SourceBySize.Exists(sz) Then

                    AD_Report = AD_Report & _
                        "SIZE : " & sz & vbCrLf & _
                        "STATUS : SOURCE DUPLICATE" & vbCrLf & _
                        "Keterangan: ditemukan lebih dari 1 source untuk size yang sama dan size ini dibutuhkan Order.txt." & vbCrLf & _
                        String(45, "-") & vbCrLf & vbCrLf

                Else

                    AD_SourceBySize.Add sz, src

                End If

            Else

                AD_Warning = AD_Warning & _
                    "SOURCE size " & sz & " diabaikan karena tidak ada di Order.txt." & vbCrLf

            End If

        End If

    Next i

    Dim k As Variant

    'Yang wajib ada hanyalah source untuk size yang muncul di Order.txt.
    For Each k In AD_OrderQtyBySize.Keys

        If Not AD_SourceBySize.Exists(CStr(k)) Then

            AD_Report = AD_Report & _
                "SIZE : " & CStr(k) & vbCrLf & _
                "ORDER : " & CLng(AD_OrderQtyBySize(k)) & " pcs" & vbCrLf & _
                "STATUS: SOURCE MASTER LAYOUT TIDAK DITEMUKAN" & vbCrLf & _
                String(45, "-") & vbCrLf & vbCrLf

        End If

    Next k

    If AD_Report = "" Then
        AD_PreflightSources = True
    End If

End Function


'=========================================================
' DETECT SIZE FROM OUTLINE
'=========================================================

Private Function AD_DetectSourceSize(ByVal src As Shape) As String

    If AD_IsPants Then
        AD_DetectSourceSize = AD_DetectPantsSourceSize(src)
        Exit Function
    End If

    Dim bestSize As String
    Dim bestArea As Double

    bestSize = ""
    bestArea = 0

    AD_ScanPanelRecursive src, bestSize, bestArea

    AD_DetectSourceSize = bestSize

End Function

Private Sub AD_ScanPanelRecursive( _
    ByVal s As Shape, _
    ByRef bestSize As String, _
    ByRef bestArea As Double)

    Dim c As Shape
    Dim pcShapes As Shapes

    On Error Resume Next

    If s.Type = cdrGroupShape Then

        For Each c In s.Shapes
            AD_ScanPanelRecursive c, bestSize, bestArea
        Next c

        Exit Sub

    End If

    Set pcShapes = s.PowerClip.Shapes

    If Not pcShapes Is Nothing Then
        For Each c In pcShapes
            AD_ScanPanelRecursive c, bestSize, bestArea
        Next c
    End If

    On Error GoTo 0

    If s.Type <> cdrCurveShape Then Exit Sub
    If Not AD_IsPanelOutline(s) Then Exit Sub

    Dim detected As String
    Dim area As Double

    detected = AD_DetectSizeFromShape(s, area)

    If detected <> "" Then

        If area > bestArea Then

            bestArea = area

            If InStr(1, detected, "_", vbTextCompare) > 0 Then
                bestSize = Left$(detected, InStr(1, detected, "_", vbTextCompare) - 1)
            Else
                bestSize = detected
            End If

        End If

    End If

End Sub


'=========================================================
' SPECIAL CELANA SIZE DETECTION
'=========================================================

Private Function AD_DetectPantsSourceSize(ByVal src As Shape) As String

    Dim vote As Object
    Dim errSum As Object

    Set vote = CreateObject("Scripting.Dictionary")
    Set errSum = CreateObject("Scripting.Dictionary")

    Dim k As Variant

    For Each k In AD_SizeDB.Keys
        vote.Add CStr(k), 0
        errSum.Add CStr(k), 0#
    Next k

    Dim panelCount As Long
    panelCount = 0

    AD_ScanPantsPanelsRecursive src, vote, errSum, panelCount

    Dim bestSize As String
    Dim bestVote As Long
    Dim bestErr As Double

    bestSize = ""
    bestVote = 0
    bestErr = 999999#

    For Each k In vote.Keys

        If CLng(vote(k)) > bestVote Then

            bestVote = CLng(vote(k))
            bestErr = CDbl(errSum(k))
            bestSize = CStr(k)

        ElseIf CLng(vote(k)) = bestVote And CLng(vote(k)) > 0 Then

            If CDbl(errSum(k)) < bestErr Then
                bestErr = CDbl(errSum(k))
                bestSize = CStr(k)
            End If

        End If

    Next k

    If bestVote >= 2 Then
        AD_DetectPantsSourceSize = bestSize
    End If

End Function

Private Sub AD_ScanPantsPanelsRecursive( _
    ByVal s As Shape, _
    ByRef vote As Object, _
    ByRef errSum As Object, _
    ByRef panelCount As Long)

    Dim c As Shape
    Dim pcShapes As Shapes

    On Error Resume Next

    If s.Type = cdrGroupShape Then

        For Each c In s.Shapes
            AD_ScanPantsPanelsRecursive c, vote, errSum, panelCount
        Next c

        Exit Sub

    End If

    Set pcShapes = s.PowerClip.Shapes

    If Not pcShapes Is Nothing Then
        For Each c In pcShapes
            AD_ScanPantsPanelsRecursive c, vote, errSum, panelCount
        Next c
    End If

    On Error GoTo 0

    If s.Type <> cdrCurveShape Then Exit Sub
    If Not AD_IsPanelOutline(s) Then Exit Sub

    Dim w As Double
    Dim h As Double
    Dim mn As Double

    w = Round(s.SizeWidth, 3)
    h = Round(s.SizeHeight, 3)

    If w <= 0 Or h <= 0 Then Exit Sub

    If w < h Then
        mn = w
    Else
        mn = h
    End If

    panelCount = panelCount + 1

    AD_AddPantsPanelVote mn, vote, errSum

End Sub

Private Sub AD_AddPantsPanelVote( _
    ByVal panelWidth As Double, _
    ByRef vote As Object, _
    ByRef errSum As Object)

    Dim k As Variant
    Dim db As Variant
    Dim errFront As Double
    Dim errBack As Double
    Dim errBestForSize As Double

    Dim closestSize As String
    Dim closestErr As Double

    closestSize = ""
    closestErr = 999999#

    For Each k In AD_SizeDB.Keys

        db = AD_SizeDB(k)

        If UBound(db) >= 2 Then

            errFront = Abs(panelWidth - AD_ToDbl(db(1)))
            errBack = Abs(panelWidth - AD_ToDbl(db(2)))

            If errFront < errBack Then
                errBestForSize = errFront
            Else
                errBestForSize = errBack
            End If

            If errBestForSize < closestErr Then
                closestErr = errBestForSize
                closestSize = CStr(k)
            End If

        End If

    Next k

    If closestSize <> "" Then

        If closestErr <= AD_PANTS_TOL Then
            vote(closestSize) = CLng(vote(closestSize)) + 1
            errSum(closestSize) = CDbl(errSum(closestSize)) + closestErr
        End If

    End If

End Sub

Private Function AD_DetectSizeFromShape(ByVal shp As Shape, ByRef area As Double) As String

    Dim w As Double
    Dim h As Double

    w = Round(shp.SizeWidth, 3)
    h = Round(shp.SizeHeight, 3)

    If w <= 0 Or h <= 0 Then Exit Function

    Dim mn As Double
    Dim mx As Double

    If w > h Then
        mx = w
        mn = h
    Else
        mx = h
        mn = w
    End If

    area = w * h

    Dim k As Variant
    Dim db As Variant

    For Each k In AD_SizeDB.Keys

        db = AD_SizeDB(k)

        If AD_IsPants Then

            If UBound(db) >= 2 Then

                If Abs(mn - AD_ToDbl(db(1))) <= AD_PANTS_TOL Or _
                   Abs(mn - AD_ToDbl(db(2))) <= AD_PANTS_TOL Then

                    AD_DetectSizeFromShape = CStr(k)
                    Exit Function

                End If

            End If

        ElseIf AD_IsSplitFront Then

            If UBound(db) >= 4 Then

                If Abs(mn - AD_ToDbl(db(1))) <= AD_SIZE_TOL And _
                   Abs(mx - AD_ToDbl(db(4))) <= AD_SIZE_TOL Then

                    AD_DetectSizeFromShape = CStr(k) & "_BACK"
                    Exit Function

                End If

                If Abs(mn - AD_ToDbl(db(2))) <= AD_SIZE_TOL And _
                   Abs(mx - AD_ToDbl(db(3))) <= AD_SIZE_TOL Then

                    AD_DetectSizeFromShape = CStr(k) & "_FRONT"
                    Exit Function

                End If

            End If

        Else

            If UBound(db) >= 3 Then

                If Abs(mn - AD_ToDbl(db(1))) <= AD_SIZE_TOL And _
                   (Abs(mx - AD_ToDbl(db(2))) <= AD_SIZE_TOL Or _
                    Abs(mx - AD_ToDbl(db(3))) <= AD_SIZE_TOL) Then

                    AD_DetectSizeFromShape = CStr(k)
                    Exit Function

                End If

            End If

        End If

    Next k

End Function


'=========================================================
' OUTLINE DETECTION
'=========================================================

Private Function AD_IsPanelOutline(ByVal shp As Shape) As Boolean

    AD_IsPanelOutline = False

    If AD_IsRedOutline(shp) Then
        AD_IsPanelOutline = True
        Exit Function
    End If

    If AD_IsGreenOutline(shp) Then
        AD_IsPanelOutline = True
        Exit Function
    End If

End Function

Private Function AD_IsRedOutline(ByVal shp As Shape) As Boolean

    On Error Resume Next

    If shp.Outline Is Nothing Then Exit Function
    If shp.Outline.Type = cdrNoOutline Then Exit Function

    AD_IsRedOutline = _
        shp.Outline.Color.RGBRed > 200 And _
        shp.Outline.Color.RGBGreen < 80 And _
        shp.Outline.Color.RGBBlue < 80

    On Error GoTo 0

End Function

Private Function AD_IsGreenOutline(ByVal shp As Shape) As Boolean

    Dim r As Long
    Dim g As Long
    Dim b As Long

    On Error Resume Next

    If shp.Outline Is Nothing Then Exit Function
    If shp.Outline.Type = cdrNoOutline Then Exit Function

    r = shp.Outline.Color.RGBRed
    g = shp.Outline.Color.RGBGreen
    b = shp.Outline.Color.RGBBlue

    'Hijau murni dari QC_SIZE_CHECK baru
    If r <= 80 And g >= 180 And b <= 80 Then
        AD_IsGreenOutline = True
        On Error GoTo 0
        Exit Function
    End If

    'Hijau lama / alternatif
    AD_IsGreenOutline = _
        Abs(r - AD_GREEN_R) <= AD_GREEN_TOL And _
        Abs(g - AD_GREEN_G) <= AD_GREEN_TOL And _
        Abs(b - AD_GREEN_B) <= AD_GREEN_TOL

    On Error GoTo 0

End Function


'=========================================================
' PREVIEW
'=========================================================

Private Function AD_ShowPreview( _
    ByVal orderedSizes As Variant, _
    ByVal pageCenterX As Double, _
    ByVal pageTopY As Double, _
    ByVal sourceSelectedCount As Long) As Boolean

    AD_ShowPreview = False

    Dim msg As String
    Dim k As Variant
    Dim totalQty As Long
    Dim q As Long
    Dim gridCols As Long
    Dim gridRows As Long
    Dim emptySlots As Long

    totalQty = 0

    msg = "PREVIEW AUTO DUPLICATE V2.3 ADAPTIVE GRID" & vbCrLf & vbCrLf
    msg = msg & "Database : " & AD_CurrentDB & vbCrLf
    msg = msg & "DB Source: " & AD_DBSource & vbCrLf
    msg = msg & "Mode     : " & AD_ProductModeText() & vbCrLf
    msg = msg & "Output   : Di atas ActivePage / page putih" & vbCrLf
    msg = msg & "Anchor   : Duplicate nomor 1 center terhadap page" & vbCrLf
    msg = msg & "Method   : " & AD_PageAnchorMethod & vbCrLf
    msg = msg & String(45, "-") & vbCrLf
    msg = msg & "Source dipilih : " & sourceSelectedCount & vbCrLf
    msg = msg & "Source dipakai : " & AD_SourceBySize.Count & vbCrLf
    msg = msg & "Page Center X  : " & FormatNumber(pageCenterX, 3) & " cm" & vbCrLf
    msg = msg & "Page Top Y     : " & FormatNumber(pageTopY, 3) & " cm" & vbCrLf
    msg = msg & String(45, "-") & vbCrLf
    msg = msg & "Grid mode          : Adaptive Balanced Grid" & vbCrLf
    msg = msg & "Large qty max kolom: " & AD_MAX_COL & vbCrLf
    msg = msg & "Jarak antar baju   : " & FormatNumber(AD_GAP_X, 1) & " cm" & vbCrLf
    msg = msg & "Jarak antar row    : " & FormatNumber(AD_GAP_Y, 1) & " cm" & vbCrLf
    msg = msg & "Jarak antar size   : " & FormatNumber(AD_BLOCK_GAP_Y, 1) & " cm" & vbCrLf
    msg = msg & "Jarak dari page    : " & FormatNumber(AD_PAGE_TOP_OFFSET, 1) & " cm" & vbCrLf
    msg = msg & String(45, "-") & vbCrLf

    For Each k In orderedSizes

        q = CLng(AD_OrderQtyBySize(CStr(k)))
        totalQty = totalQty + q

        AD_CalcAdaptiveGrid q, gridCols, gridRows
        emptySlots = (gridCols * gridRows) - q

        msg = msg & _
            "SIZE " & CStr(k) & _
            " -> " & q & " duplicate" & _
            " | grid " & gridCols & " x " & gridRows

        If emptySlots > 0 Then
            msg = msg & " -" & emptySlots
        End If

        msg = msg & vbCrLf

    Next k

    msg = msg & String(45, "-") & vbCrLf
    msg = msg & "Total output : " & totalQty & " duplicate" & vbCrLf & vbCrLf
    msg = msg & "Source/master asli tidak akan dipindah." & vbCrLf
    msg = msg & "Source size yang tidak ada di Order.txt akan diabaikan." & vbCrLf

    If Len(AD_Warning) > 0 Then
        msg = msg & vbCrLf & "Catatan:" & vbCrLf & AD_Warning
    End If

    msg = msg & vbCrLf & "Lanjut membuat HASIL LAYOUT?"

    If MsgBox(msg, vbQuestion + vbYesNo, "HADES AUTO DUPLICATE V2.3") = vbYes Then
        AD_ShowPreview = True
    End If

End Function


Private Sub AD_CalcAdaptiveGrid( _
    ByVal qty As Long, _
    ByRef cols As Long, _
    ByRef rows As Long)

    cols = 0
    rows = 0

    If qty <= 0 Then Exit Sub

    If qty <= AD_GRID_SINGLE_ROW_MAX Then

        rows = 1
        cols = qty

    ElseIf qty <= AD_GRID_TWO_ROW_MAX Then

        rows = 2
        cols = AD_CeilDivLong(qty, rows)

    ElseIf qty <= AD_GRID_BALANCED_MAX Then

        'Balanced grid menengah.
        'Target awal 5-6 kolom agar blok tidak terlalu memanjang.
        rows = AD_CeilDivLong(qty, AD_GRID_BALANCED_TARGET_COL)
        If rows < 1 Then rows = 1

        cols = AD_CeilDivLong(qty, rows)

        If cols < 1 Then cols = 1

    Else

        'Large quantity mode.
        'Untuk qty besar, lebih aman memanjang sampai 8 kolom baru turun.
        cols = AD_MAX_COL
        rows = AD_CeilDivLong(qty, cols)

    End If

    If cols < 1 Then cols = 1
    If rows < 1 Then rows = 1

End Sub

Private Function AD_CeilDivLong(ByVal a As Long, ByVal b As Long) As Long

    If b <= 0 Then
        AD_CeilDivLong = 0
        Exit Function
    End If

    AD_CeilDivLong = (a + b - 1) \ b

End Function


Private Sub AD_DuplicateOneSizeBlock( _
    ByVal sourceShape As Shape, _
    ByVal qty As Long, _
    ByVal anchorCenterX As Double, _
    ByVal blockTop As Double, _
    ByRef createdCount As Long)

    If sourceShape Is Nothing Then Exit Sub
    If qty <= 0 Then Exit Sub

    Dim gridCols As Long
    Dim gridRows As Long

    AD_CalcAdaptiveGrid qty, gridCols, gridRows

    If gridCols <= 0 Then gridCols = 1

    Dim i As Long
    Dim col As Long
    Dim row As Long

    Dim stepX As Double
    Dim stepY As Double

    stepX = sourceShape.SizeWidth + AD_GAP_X
    stepY = sourceShape.SizeHeight + AD_GAP_Y

    Dim targetX As Double
    Dim targetY As Double

    Dim dup As Shape

    For i = 1 To qty

        col = (i - 1) Mod gridCols
        row = Int((i - 1) / gridCols)

        'Duplicate nomor 1:
        'col = 0, targetX = anchorCenterX.
        'Baris berikutnya selalu mulai lagi dari kiri / anchorCenterX.
        targetX = anchorCenterX + (col * stepX)

        targetY = blockTop - (sourceShape.SizeHeight / 2) - (row * stepY)

        Set dup = sourceShape.Duplicate

        dup.PositionX = targetX
        dup.PositionY = targetY

        createdCount = createdCount + 1

    Next i

End Sub

Private Function AD_BlockHeight(ByVal sourceShape As Shape, ByVal qty As Long) As Double

    If sourceShape Is Nothing Then Exit Function
    If qty <= 0 Then Exit Function

    Dim cols As Long
    Dim rows As Long

    AD_CalcAdaptiveGrid qty, cols, rows

    If rows <= 0 Then rows = 1

    AD_BlockHeight = _
        (rows * sourceShape.SizeHeight) + _
        ((rows - 1) * AD_GAP_Y)

End Function

Private Function AD_TotalLayoutHeight(ByVal orderedSizes As Variant) As Double

    Dim totalH As Double
    Dim countBlock As Long
    Dim k As Variant
    Dim sz As String
    Dim q As Long
    Dim src As Shape

    totalH = 0
    countBlock = 0

    For Each k In orderedSizes

        sz = CStr(k)
        q = CLng(AD_OrderQtyBySize(sz))
        Set src = AD_SourceBySize(sz)

        If q > 0 Then

            totalH = totalH + AD_BlockHeight(src, q)
            countBlock = countBlock + 1

        End If

    Next k

    If countBlock > 1 Then
        totalH = totalH + ((countBlock - 1) * AD_BLOCK_GAP_Y)
    End If

    AD_TotalLayoutHeight = totalH

End Function


'=========================================================
' ACCURATE PAGE ANCHOR
'=========================================================

Private Sub AD_GetAccuratePageAnchor( _
    ByRef outCenterX As Double, _
    ByRef outTopY As Double)

    'Fallback lama.
    outCenterX = ActivePage.SizeWidth / 2
    outTopY = ActivePage.SizeHeight
    AD_PageAnchorMethod = "Fallback Width/Height"

    Dim tmp As Shape
    Dim centerY As Double
    Dim failed As Boolean

    failed = False

    On Error Resume Next

    'Buat objek sementara kecil.
    Set tmp = ActiveLayer.CreateRectangle2(0, 0, 0.2, 0.2)

    If tmp Is Nothing Then
        On Error GoTo 0
        Exit Sub
    End If

    Err.Clear

    'Gunakan CallByName agar lebih aman terhadap variasi object model.
    CallByName tmp, "AlignToPageCenter", VbMethod, cdrAlignHCenter
    If Err.Number <> 0 Then failed = True

    Err.Clear

    CallByName tmp, "AlignToPageCenter", VbMethod, cdrAlignVCenter
    If Err.Number <> 0 Then failed = True

    If failed = False Then

        outCenterX = tmp.PositionX
        centerY = tmp.PositionY
        outTopY = centerY + (ActivePage.SizeHeight / 2)

        AD_PageAnchorMethod = "Smart Temporary Shape Align"

    End If

    tmp.Delete

    On Error GoTo 0

End Sub


'=========================================================
' SORT SIZE
'=========================================================

Private Function AD_GetSortedOrderSizes() As Variant

    Dim keys As Variant
    keys = AD_OrderQtyBySize.Keys

    Dim i As Long
    Dim j As Long
    Dim tmp As Variant

    If AD_OrderQtyBySize.Count > 1 Then

        For i = LBound(keys) To UBound(keys) - 1

            For j = i + 1 To UBound(keys)

                If AD_SizeRank(CStr(keys(i))) > AD_SizeRank(CStr(keys(j))) Then

                    tmp = keys(i)
                    keys(i) = keys(j)
                    keys(j) = tmp

                End If

            Next j

        Next i

    End If

    AD_GetSortedOrderSizes = keys

End Function

Private Function AD_SizeRank(ByVal sz As String) As Long

    Dim s As String
    s = AD_NormalizeSizeKey(sz)

    Select Case s

        Case "XXS"
            AD_SizeRank = 1

        Case "XS"
            AD_SizeRank = 2

        Case "S"
            AD_SizeRank = 3

        Case "M"
            AD_SizeRank = 4

        Case "L"
            AD_SizeRank = 5

        Case "XL"
            AD_SizeRank = 6

        Case "2XL"
            AD_SizeRank = 7

        Case "3XL"
            AD_SizeRank = 8

        Case "4XL"
            AD_SizeRank = 9

        Case "5XL"
            AD_SizeRank = 10

        Case "6XL"
            AD_SizeRank = 11

        Case Else
            AD_SizeRank = 999

    End Select

End Function


'=========================================================
' HELPERS
'=========================================================

Private Function AD_Ceiling(ByVal x As Double) As Long

    Dim i As Long
    i = Int(x)

    If x > i Then
        AD_Ceiling = i + 1
    Else
        AD_Ceiling = i
    End If

End Function

Private Function AD_ToDbl(ByVal v As Variant) As Double

    Dim s As String
    s = Trim$(CStr(v))
    s = Replace(s, ",", ".")

    AD_ToDbl = Val(s)

End Function

Private Function AD_NormalizeSizeKey(ByVal sz As String) As String

    Dim s As String

    s = UCase$(Trim$(sz))

    Select Case s

        Case "XXL"
            AD_NormalizeSizeKey = "2XL"

        Case "XXXL"
            AD_NormalizeSizeKey = "3XL"

        Case "XXXXL"
            AD_NormalizeSizeKey = "4XL"

        Case "XXXXXL"
            AD_NormalizeSizeKey = "5XL"

        Case "XXXXXXL"
            AD_NormalizeSizeKey = "6XL"

        Case Else
            AD_NormalizeSizeKey = s

    End Select

End Function

Private Function AD_IsStandardSize(ByVal sz As String) As Boolean

    Dim s As String

    s = AD_NormalizeSizeKey(sz)

    Select Case s

        Case "XXS", "XS", "S", "M", "L", "XL", "2XL", "3XL", "4XL", "5XL", "6XL"
            AD_IsStandardSize = True

        Case Else
            AD_IsStandardSize = False

    End Select

End Function

Private Function AD_ProductModeText() As String

    If AD_IsPants Then
        AD_ProductModeText = "CELANA"
    ElseIf AD_IsSplitFront Then
        AD_ProductModeText = "JAKET"
    Else
        AD_ProductModeText = "JERSEY"
    End If

End Function