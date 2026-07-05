Option Explicit

'=========================================================
' PROJECT HADES — QC AUTO RENAME V5.3 VISUAL LIGATURE BREAKER
'
' BASE:
' - QC AUTO RENAME V5.1
'
' FITUR UTAMA:
' - Rename teks hasil layout berdasarkan Documents\Order.txt
' - Group hasil layout dideteksi size-nya dari outline merah / hijau
' - Urutan rename berdasarkan posisi visual:
'   atas dulu, lalu kiri ke kanan
'
' FORMAT ORDER:
' Metadata optional:
'
'   @JENIS_PESANAN=JERSEY
'   @JENIS_POLA=JERSEY REGULER
'   @MODEL_JAHIT=DEWASA PRIA
'   @SIZEDB=SizeDB_Pria.txt
'
' Data order:
'
'   SIZE|NAMA|NOMOR|NICKNAME
'
' FITUR BARU V5.2:
' - Membaca @SIZEDB dari Order.txt
' - Popup database dilewati jika @SIZEDB tersedia
' - Baris metadata @... diabaikan saat membaca order
' - Fallback popup tetap ada jika metadata belum tersedia
' - Bisa infer DB dari @JENIS_PESANAN / @JENIS_POLA / @MODEL_JAHIT
' - Support hijau murni 0,255,0 dari QC_SIZE_CHECK V8
'
' MODE:
' - JERSEY : rename NAMA + NOMOR + NICKNAME
' - JAKET  : rename NAMA + NOMOR + NICKNAME
' - CELANA : rename NOMOR saja
'
' MAIN MACRO:
' QC_AUTO_RENAME
'
' ALIAS:
' AUTO_RENAME_V4
' AUTO_RENAME_V5
' AUTO_RENAME_V51
' AUTO_RENAME_V52
' AUTO_RENAME_V53
'=========================================================


'=========================================================
' GLOBAL VARIABLES
'=========================================================

Private AR_OrdersBySize As Object
Private AR_GroupsBySize As Object
Private AR_SizeDB As Object
Private AR_OrderMeta As Object

Private AR_CurrentDB As String
Private AR_DBSource As String

Private AR_IsSplitFront As Boolean
Private AR_IsPants As Boolean

Private AR_Report As String
Private AR_Warning As String

Private AR_DeleteQueue As Collection

Private AR_CJKDetectedCount As Long
Private AR_CJKFontAppliedCount As Long
Private AR_CJKFontFailedCount As Long

Private Const AR_SIZE_TOL As Double = 1#
Private Const AR_PANTS_TOL As Double = 0.35
Private Const AR_ROW_TOL As Double = 5#

Private Const AR_MIN_TEXT_H As Double = 1#
Private Const AR_ID_MIN_H As Double = 0.28
Private Const AR_ID_MAX_H As Double = 0.65

Private Const AR_CJK_FONT_1 As String = "Meiryo"
Private Const AR_CJK_FONT_2 As String = "Yu Gothic"
Private Const AR_CJK_FONT_3 As String = "Noto Sans CJK JP"
Private Const AR_CJK_FONT_4 As String = "Noto Sans JP"
Private Const AR_CJK_FONT_5 As String = "MS Gothic"

Private Const AR_ENABLE_VISUAL_LIGATURE_BREAKER As Boolean = True

Private Const AR_GREEN_R As Long = 97
Private Const AR_GREEN_G As Long = 186
Private Const AR_GREEN_B As Long = 12
Private Const AR_GREEN_TOL As Long = 18


'=========================================================
' PUBLIC ENTRY
'=========================================================

Sub AUTO_RENAME_V4()
    Call QC_AUTO_RENAME
End Sub

Sub AUTO_RENAME_V5()
    Call QC_AUTO_RENAME
End Sub

Sub AUTO_RENAME_V51()
    Call QC_AUTO_RENAME
End Sub

Sub AUTO_RENAME_V52()
    Call QC_AUTO_RENAME
End Sub

Sub AUTO_RENAME_V53()
    Call QC_AUTO_RENAME
End Sub

Sub QC_AUTO_RENAME()

    Dim oldUnit As Long
    Dim sr As ShapeRange
    Dim topGroups As Collection
    Dim cmdStarted As Boolean

    oldUnit = ActiveDocument.Unit
    cmdStarted = False

    On Error GoTo ERR_HANDLER

    On Error Resume Next
    Set sr = ActiveSelectionRange
    On Error GoTo ERR_HANDLER

    If sr Is Nothing Then
        MsgBox "Pilih HASIL LAYOUT terlebih dahulu.", vbExclamation, "QC AUTO RENAME"
        Exit Sub
    End If

    If sr.Count = 0 Then
        MsgBox "Pilih HASIL LAYOUT terlebih dahulu.", vbExclamation, "QC AUTO RENAME"
        Exit Sub
    End If

    Set AR_OrdersBySize = CreateObject("Scripting.Dictionary")
    Set AR_GroupsBySize = CreateObject("Scripting.Dictionary")
    Set AR_SizeDB = CreateObject("Scripting.Dictionary")
    Set AR_OrderMeta = CreateObject("Scripting.Dictionary")
    Set AR_DeleteQueue = New Collection

    AR_CurrentDB = ""
    AR_DBSource = ""
    AR_IsSplitFront = False
    AR_IsPants = False

    AR_Report = ""
    AR_Warning = ""

    AR_CJKDetectedCount = 0
    AR_CJKFontAppliedCount = 0
    AR_CJKFontFailedCount = 0

    ActiveDocument.Unit = cdrCentimeter

    '=====================================================
    ' V5.2:
    ' Order dibaca dulu agar metadata @SIZEDB bisa dipakai
    ' untuk menentukan database tanpa popup.
    '=====================================================
    AR_LoadOrders

    If Len(Trim$(AR_CurrentDB)) > 0 Then

        AR_ConfigureModeFromDB
        AR_DBSource = "AUTO dari Order.txt @SIZEDB"

    Else

        AR_CurrentDB = AR_InferDBFromMetadata()

        If Len(Trim$(AR_CurrentDB)) > 0 Then

            AR_ConfigureModeFromDB
            AR_DBSource = "AUTO dari metadata spesifikasi Order.txt"

        Else

            If Not AR_SelectDatabaseFallback Then
                GoTo EXIT_CLEAN
            End If

            AR_DBSource = "MANUAL POPUP"

        End If

    End If

    AR_LoadSizeDB AR_CurrentDB

    Set topGroups = AR_CollectTopGroups(sr)

    If topGroups.Count = 0 Then

        MsgBox _
            "Tidak ada group utama pada selection." & vbCrLf & vbCrLf & _
            "Pastikan hasil layout sudah digroup per set, lalu select semua group.", _
            vbExclamation, _
            "QC AUTO RENAME"

        GoTo EXIT_CLEAN

    End If

    If Not AR_Preflight(topGroups) Then

        MsgBox _
            "QC AUTO RENAME DIBATALKAN" & vbCrLf & vbCrLf & _
            "Preflight gagal. Tidak ada teks yang diubah." & vbCrLf & vbCrLf & _
            AR_Report, _
            vbCritical, _
            "QC AUTO RENAME"

        GoTo EXIT_CLEAN

    End If

    ActiveDocument.BeginCommandGroup "QC Auto Rename V5.2"
    cmdStarted = True

    AR_ApplyRename

    ActiveDocument.EndCommandGroup
    cmdStarted = False

    If AR_CJKDetectedCount > 0 Then

        AR_Warning = AR_Warning & _
            "JAPANESE / CJK FONT CHECK" & vbCrLf & _
            "Teks Jepang/CJK terdeteksi : " & AR_CJKDetectedCount & vbCrLf & _
            "Font berhasil diterapkan   : " & AR_CJKFontAppliedCount & vbCrLf & _
            "Font gagal diterapkan      : " & AR_CJKFontFailedCount & vbCrLf & _
            "Fallback font              : " & AR_CJK_FONT_1 & ", " & AR_CJK_FONT_2 & ", " & AR_CJK_FONT_3 & ", " & AR_CJK_FONT_4 & ", " & AR_CJK_FONT_5 & vbCrLf & _
            String(35, "-") & vbCrLf

    End If

    ActiveDocument.Unit = oldUnit
    ActiveWindow.Refresh

    If AR_Warning <> "" Then

        MsgBox _
            "QC AUTO RENAME SELESAI DENGAN WARNING" & vbCrLf & vbCrLf & _
            "Database : " & AR_CurrentDB & vbCrLf & _
            "DB Source: " & AR_DBSource & vbCrLf & _
            "Mode     : " & AR_ProductModeText() & vbCrLf & vbCrLf & _
            AR_Warning, _
            vbExclamation, _
            "QC AUTO RENAME"

    Else

        MsgBox _
            "QC AUTO RENAME BERHASIL" & vbCrLf & vbCrLf & _
            "Database : " & AR_CurrentDB & vbCrLf & _
            "DB Source: " & AR_DBSource & vbCrLf & _
            "Mode     : " & AR_ProductModeText() & vbCrLf & vbCrLf & _
            "Semua group berhasil diproses.", _
            vbInformation, _
            "QC AUTO RENAME"

    End If

    Exit Sub

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

    eNo = err.Number
    eDesc = err.Description

    On Error Resume Next

    If cmdStarted Then ActiveDocument.EndCommandGroup

    ActiveDocument.Unit = oldUnit
    ActiveWindow.Refresh

    On Error GoTo 0

    If eNo = 0 And Trim$(eDesc) = "" Then
        eDesc = "Error tidak teridentifikasi. Kemungkinan terjadi saat membaca selection, group, Order.txt, atau SizeDB."
    End If

    MsgBox _
        "SYSTEM ERROR - QC AUTO RENAME V5.2" & vbCrLf & vbCrLf & _
        "No : " & eNo & vbCrLf & _
        eDesc, _
        vbCritical, _
        "QC AUTO RENAME"

End Sub


'=========================================================
' DATABASE AUTO / FALLBACK
'=========================================================

Private Sub AR_ConfigureModeFromDB()

    Dim db As String

    db = UCase$(Trim$(AR_CurrentDB))

    AR_IsSplitFront = False
    AR_IsPants = False

    If InStr(1, db, "CELANA", vbTextCompare) > 0 Then
        AR_IsPants = True
        AR_IsSplitFront = False
        Exit Sub
    End If

    If InStr(1, db, "JAKET", vbTextCompare) > 0 Then
        AR_IsSplitFront = True
        AR_IsPants = False
        Exit Sub
    End If

    'Default selain Celana/Jaket dianggap Jersey
    AR_IsSplitFront = False
    AR_IsPants = False

End Sub

Private Function AR_InferDBFromMetadata() As String

    Dim jenis As String
    Dim pola As String
    Dim model As String
    Dim allText As String

    jenis = UCase$(AR_GetMeta("JENIS_PESANAN"))
    pola = UCase$(AR_GetMeta("JENIS_POLA"))
    model = UCase$(AR_GetMeta("MODEL_JAHIT"))

    allText = jenis & " " & pola & " " & model

    AR_InferDBFromMetadata = ""

    If InStr(1, allText, "CELANA", vbTextCompare) > 0 Then

        If InStr(1, allText, "ANAK", vbTextCompare) > 0 Then
            AR_InferDBFromMetadata = "SizeDB_CelanaAnak.txt"
        ElseIf InStr(1, allText, "WANITA", vbTextCompare) > 0 Or _
               InStr(1, allText, "PEREMPUAN", vbTextCompare) > 0 Or _
               InStr(1, allText, "CEWEK", vbTextCompare) > 0 Then
            AR_InferDBFromMetadata = "SizeDB_CelanaWanita.txt"
        Else
            AR_InferDBFromMetadata = "SizeDB_CelanaPria.txt"
        End If

        Exit Function

    End If

    If InStr(1, allText, "JAKET", vbTextCompare) > 0 Then

        If InStr(1, allText, "ANAK", vbTextCompare) > 0 Then
            AR_InferDBFromMetadata = "SizeDB_JaketAnak.txt"
        Else
            AR_InferDBFromMetadata = "SizeDB_Jaket.txt"
        End If

        Exit Function

    End If

    If InStr(1, allText, "JERSEY", vbTextCompare) > 0 Then

        If InStr(1, allText, "ANAK", vbTextCompare) > 0 Then
            AR_InferDBFromMetadata = "SizeDB_Anak.txt"
            Exit Function
        End If

        If InStr(1, allText, "SLIM", vbTextCompare) > 0 Then

            If InStr(1, allText, "WANITA", vbTextCompare) > 0 Or _
               InStr(1, allText, "PEREMPUAN", vbTextCompare) > 0 Or _
               InStr(1, allText, "CEWEK", vbTextCompare) > 0 Then
                AR_InferDBFromMetadata = "SizeDB_WanitaSlimFit.txt"
            Else
                AR_InferDBFromMetadata = "SizeDB_PriaSlimFit.txt"
            End If

            Exit Function

        End If

        If InStr(1, allText, "WANITA", vbTextCompare) > 0 Or _
           InStr(1, allText, "PEREMPUAN", vbTextCompare) > 0 Or _
           InStr(1, allText, "CEWEK", vbTextCompare) > 0 Then
            AR_InferDBFromMetadata = "SizeDB_Wanita.txt"
        Else
            AR_InferDBFromMetadata = "SizeDB_Pria.txt"
        End If

        Exit Function

    End If

End Function

Private Function AR_SelectDatabaseFallback() As Boolean

    AR_SelectDatabaseFallback = False

    Dim a As String
    Dim b As String

    AR_IsSplitFront = False
    AR_IsPants = False
    AR_CurrentDB = ""

    a = InputBox( _
        "Order.txt belum memiliki @SIZEDB." & vbCrLf & vbCrLf & _
        "PILIH PRODUK / POLA" & vbCrLf & vbCrLf & _
        "1 = JERSEY" & vbCrLf & _
        "2 = JAKET" & vbCrLf & _
        "3 = CELANA", _
        "QC AUTO RENAME")

    If Trim$(a) = "" Then Exit Function

    Select Case Trim$(a)

        Case "1"

            AR_IsSplitFront = False
            AR_IsPants = False

            b = InputBox( _
                "PILIH DATABASE JERSEY" & vbCrLf & vbCrLf & _
                "1 = PRIA REGULAR" & vbCrLf & _
                "2 = WANITA REGULAR" & vbCrLf & _
                "3 = ANAK" & vbCrLf & _
                "4 = PRIA SLIM FIT" & vbCrLf & _
                "5 = WANITA SLIM FIT", _
                "QC AUTO RENAME")

            If Trim$(b) = "" Then Exit Function

            Select Case Trim$(b)

                Case "1"
                    AR_CurrentDB = "SizeDB_Pria.txt"

                Case "2"
                    AR_CurrentDB = "SizeDB_Wanita.txt"

                Case "3"
                    AR_CurrentDB = "SizeDB_Anak.txt"

                Case "4"
                    AR_CurrentDB = "SizeDB_PriaSlimFit.txt"

                Case "5"
                    AR_CurrentDB = "SizeDB_WanitaSlimFit.txt"

                Case Else
                    MsgBox "Pilihan database tidak valid.", vbExclamation, "QC AUTO RENAME"
                    Exit Function

            End Select

        Case "2"

            AR_IsSplitFront = True
            AR_IsPants = False

            b = InputBox( _
                "PILIH DATABASE JAKET" & vbCrLf & vbCrLf & _
                "1 = JAKET DEWASA" & vbCrLf & _
                "2 = JAKET ANAK", _
                "QC AUTO RENAME")

            If Trim$(b) = "" Then Exit Function

            Select Case Trim$(b)

                Case "1"
                    AR_CurrentDB = "SizeDB_Jaket.txt"

                Case "2"
                    AR_CurrentDB = "SizeDB_JaketAnak.txt"

                Case Else
                    MsgBox "Pilihan database tidak valid.", vbExclamation, "QC AUTO RENAME"
                    Exit Function

            End Select

        Case "3"

            AR_IsSplitFront = False
            AR_IsPants = True

            b = InputBox( _
                "PILIH DATABASE CELANA" & vbCrLf & vbCrLf & _
                "1 = CELANA PRIA" & vbCrLf & _
                "2 = CELANA WANITA" & vbCrLf & _
                "3 = CELANA ANAK", _
                "QC AUTO RENAME")

            If Trim$(b) = "" Then Exit Function

            Select Case Trim$(b)

                Case "1"
                    AR_CurrentDB = "SizeDB_CelanaPria.txt"

                Case "2"
                    AR_CurrentDB = "SizeDB_CelanaWanita.txt"

                Case "3"
                    AR_CurrentDB = "SizeDB_CelanaAnak.txt"

                Case Else
                    MsgBox "Pilihan database tidak valid.", vbExclamation, "QC AUTO RENAME"
                    Exit Function

            End Select

        Case Else

            MsgBox "Pilihan produk tidak valid.", vbExclamation, "QC AUTO RENAME"
            Exit Function

    End Select

    AR_SelectDatabaseFallback = True

End Function


'=========================================================
' LOAD ORDER — UTF-8 SAFE + METADATA
'=========================================================

Private Sub AR_LoadOrders()

    Set AR_OrdersBySize = CreateObject("Scripting.Dictionary")
    Set AR_OrderMeta = CreateObject("Scripting.Dictionary")

    Dim path As String
    path = Environ$("USERPROFILE") & "\Documents\Order.txt"

    If Dir(path) = "" Then
        err.Raise vbObjectError + 100, , _
            "Order.txt tidak ditemukan di Documents." & vbCrLf & path
    End If

    On Error GoTo FAIL

    Dim content As String
    content = AR_ReadTextFileUTF8(path)

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
        ln = AR_RemoveBOM(ln)
        ln = Trim$(ln)

        If ln <> "" Then

            If Left$(ln, 1) = "@" Then

                AR_ParseMetaLine ln

            Else

                arr = Split(ln, "|")

                If UBound(arr) >= 3 Then

                    sz = AR_NormalizeSizeKey(CStr(arr(0)))

                    If sz <> "" Then
                        AR_AddOrderBySize sz, arr
                    End If

                End If

            End If

        End If

    Next i

    If AR_OrderMeta.Exists("SIZEDB") Then
        If Len(Trim$(CStr(AR_OrderMeta("SIZEDB")))) > 0 Then
            AR_CurrentDB = Trim$(CStr(AR_OrderMeta("SIZEDB")))
        End If
    End If

    If AR_OrdersBySize.Count = 0 Then
        err.Raise vbObjectError + 101, , _
            "Order.txt kosong atau format tidak valid." & vbCrLf & _
            "Format wajib: SIZE|NAMA|NOMOR|NICKNAME" & vbCrLf & vbCrLf & _
            "Baris metadata @... boleh ada dan akan diabaikan."
    End If

    Exit Sub

FAIL:

    err.Raise vbObjectError + 102, , _
        "Gagal membaca Order.txt." & vbCrLf & err.Description

End Sub

Private Sub AR_ParseMetaLine(ByVal line As String)

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

    If AR_OrderMeta.Exists(k) Then
        AR_OrderMeta(k) = v
    Else
        AR_OrderMeta.Add k, v
    End If

End Sub

Private Function AR_GetMeta(ByVal keyName As String) As String

    keyName = UCase$(Trim$(keyName))

    If AR_OrderMeta Is Nothing Then
        AR_GetMeta = ""
        Exit Function
    End If

    If AR_OrderMeta.Exists(keyName) Then
        AR_GetMeta = CStr(AR_OrderMeta(keyName))
    Else
        AR_GetMeta = ""
    End If

End Function

Private Function AR_ReadTextFileUTF8(ByVal path As String) As String

    On Error GoTo Fallback

    Dim stm As Object
    Set stm = CreateObject("ADODB.Stream")

    stm.Type = 2
    stm.CharSet = "utf-8"
    stm.Open
    stm.LoadFromFile path

    AR_ReadTextFileUTF8 = stm.ReadText(-1)

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

    AR_ReadTextFileUTF8 = AR_ReadTextFileANSI(path)
    Exit Function

FAIL_ANSI:

    err.Raise vbObjectError + 900, , _
        "Gagal membaca file sebagai UTF-8 maupun ANSI." & vbCrLf & path

End Function

Private Function AR_ReadTextFileANSI(ByVal path As String) As String

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

    AR_ReadTextFileANSI = buf

End Function

Private Function AR_RemoveBOM(ByVal s As String) As String

    If Len(s) > 0 Then
        If AscW(Left$(s, 1)) = &HFEFF Then
            AR_RemoveBOM = Mid$(s, 2)
            Exit Function
        End If
    End If

    AR_RemoveBOM = s

End Function

Private Sub AR_AddOrderBySize(ByVal sz As String, ByVal arr As Variant)

    Dim col As Collection

    sz = AR_NormalizeSizeKey(sz)

    If AR_OrdersBySize.Exists(sz) Then
        Set col = AR_OrdersBySize(sz)
    Else
        Set col = New Collection
        AR_OrdersBySize.Add sz, col
    End If

    col.Add arr

End Sub


'=========================================================
' LOAD SIZE DB
'=========================================================

Private Sub AR_LoadSizeDB(ByVal fileName As String)

    Set AR_SizeDB = CreateObject("Scripting.Dictionary")

    Dim path As String
    path = Environ$("USERPROFILE") & "\Documents\" & fileName

    If Dir(path) = "" Then
        err.Raise vbObjectError + 200, , _
            "SizeDB tidak ditemukan:" & vbCrLf & path & vbCrLf & vbCrLf & _
            "Database dipilih dari: " & AR_DBSource
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

            If AR_IsPants Then

                If UBound(arr) >= 2 Then
                    sz = AR_NormalizeSizeKey(CStr(arr(0)))
                    If sz <> "" Then AR_PutSizeDB sz, arr
                End If

            ElseIf AR_IsSplitFront Then

                If UBound(arr) >= 4 Then
                    sz = AR_NormalizeSizeKey(CStr(arr(0)))
                    If sz <> "" Then AR_PutSizeDB sz, arr
                End If

            Else

                If UBound(arr) >= 3 Then
                    sz = AR_NormalizeSizeKey(CStr(arr(0)))
                    If sz <> "" Then AR_PutSizeDB sz, arr
                End If

            End If

        End If

    Loop

    Close #f

    If AR_SizeDB.Count = 0 Then

        If AR_IsPants Then
            err.Raise vbObjectError + 201, , _
                "SizeDB Celana kosong atau format tidak valid." & vbCrLf & _
                "Format wajib: SIZE|L_DEPAN|L_BELAKANG"
        ElseIf AR_IsSplitFront Then
            err.Raise vbObjectError + 201, , _
                "SizeDB Jaket kosong atau format tidak valid." & vbCrLf & _
                "Format wajib: SIZE|L_BELAKANG|L_DEPAN|T_DEPAN|T_BELAKANG"
        Else
            err.Raise vbObjectError + 201, , _
                "SizeDB Jersey kosong atau format tidak valid." & vbCrLf & _
                "Format wajib: SIZE|LEBAR|TINGGI_DEPAN|TINGGI_BELAKANG"
        End If

    End If

    Exit Sub

FAIL:

    On Error Resume Next
    Close #f
    On Error GoTo 0

    err.Raise vbObjectError + 202, , _
        "Gagal membaca SizeDB: " & fileName & vbCrLf & err.Description

End Sub

Private Sub AR_PutSizeDB(ByVal sz As String, ByVal arr As Variant)

    sz = AR_NormalizeSizeKey(sz)

    If AR_SizeDB.Exists(sz) Then
        AR_SizeDB(sz) = arr
    Else
        AR_SizeDB.Add sz, arr
    End If

End Sub


'=========================================================
' COLLECT GROUPS
'=========================================================

Private Function AR_CollectTopGroups(ByVal sr As ShapeRange) As Collection

    Dim col As New Collection
    Dim s As Shape

    For Each s In sr

        If s.Type = cdrGroupShape Then
            col.Add s
        End If

    Next s

    Set AR_CollectTopGroups = col

End Function


'=========================================================
' PREFLIGHT
'=========================================================

Private Function AR_Preflight(ByVal groups As Collection) As Boolean

    AR_Preflight = False

    Set AR_GroupsBySize = CreateObject("Scripting.Dictionary")

    Dim i As Long
    Dim g As Shape
    Dim sz As String

    For i = 1 To groups.Count

        Set g = groups(i)

        sz = AR_DetectGroupSize(g)

        If sz = "" Then

            AR_Report = AR_Report & _
                "GROUP #" & i & vbCrLf & _
                "STATUS : SIZE / PANEL MERAH-HIJAU TIDAK TERDETEKSI" & vbCrLf & _
                String(40, "-") & vbCrLf & vbCrLf

        Else

            AR_AddGroupBySize sz, g

        End If

    Next i

    Dim k As Variant
    Dim expectedCount As Long
    Dim found As Long
    Dim hasError As Boolean

    hasError = False

    For Each k In AR_OrdersBySize.keys

        expectedCount = AR_OrdersBySize(k).Count
        found = 0

        If AR_GroupsBySize.Exists(k) Then
            found = AR_GroupsBySize(k).Count
        End If

        If expectedCount <> found Then

            hasError = True

            AR_Report = AR_Report & _
                "SIZE : " & CStr(k) & vbCrLf & _
                "ORDER  : " & expectedCount & vbCrLf & _
                "LAYOUT : " & found & vbCrLf & _
                "STATUS : JUMLAH TIDAK COCOK" & vbCrLf & _
                String(40, "-") & vbCrLf & vbCrLf

        End If

    Next k

    For Each k In AR_GroupsBySize.keys

        If Not AR_OrdersBySize.Exists(k) Then

            hasError = True

            AR_Report = AR_Report & _
                "SIZE : " & CStr(k) & vbCrLf & _
                "ORDER  : 0" & vbCrLf & _
                "LAYOUT : " & AR_GroupsBySize(k).Count & vbCrLf & _
                "STATUS : SIZE ADA DI LAYOUT TAPI TIDAK ADA DI ORDER" & vbCrLf & _
                String(40, "-") & vbCrLf & vbCrLf

        End If

    Next k

    If Not hasError Then
        AR_Preflight = True
    End If

End Function

Private Sub AR_AddGroupBySize(ByVal sz As String, ByVal g As Shape)

    Dim col As Collection

    sz = AR_NormalizeSizeKey(sz)

    If AR_GroupsBySize.Exists(sz) Then
        Set col = AR_GroupsBySize(sz)
    Else
        Set col = New Collection
        AR_GroupsBySize.Add sz, col
    End If

    col.Add g

End Sub


'=========================================================
' DETECT SIZE FROM OUTLINE
'=========================================================

Private Function AR_DetectGroupSize(ByVal g As Shape) As String

    If AR_IsPants Then
        AR_DetectGroupSize = AR_DetectPantsGroupSize(g)
        Exit Function
    End If

    Dim bestSize As String
    Dim bestArea As Double

    bestSize = ""
    bestArea = 0

    AR_ScanPanelRecursive g, bestSize, bestArea

    AR_DetectGroupSize = bestSize

End Function

Private Sub AR_ScanPanelRecursive( _
    ByVal s As Shape, _
    ByRef bestSize As String, _
    ByRef bestArea As Double)

    Dim c As Shape
    Dim pcShapes As Shapes

    On Error Resume Next

    If s.Type = cdrGroupShape Then

        For Each c In s.Shapes
            AR_ScanPanelRecursive c, bestSize, bestArea
        Next c

        Exit Sub

    End If

    Set pcShapes = s.PowerClip.Shapes

    If Not pcShapes Is Nothing Then
        For Each c In pcShapes
            AR_ScanPanelRecursive c, bestSize, bestArea
        Next c
    End If

    On Error GoTo 0

    If s.Type <> cdrCurveShape Then Exit Sub
    If Not AR_IsPanelOutline(s) Then Exit Sub

    Dim detected As String
    Dim area As Double

    detected = AR_DetectSizeFromShape(s, area)

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

Private Function AR_DetectPantsGroupSize(ByVal g As Shape) As String

    Dim vote As Object
    Dim errSum As Object

    Set vote = CreateObject("Scripting.Dictionary")
    Set errSum = CreateObject("Scripting.Dictionary")

    Dim k As Variant

    For Each k In AR_SizeDB.keys
        vote.Add CStr(k), 0
        errSum.Add CStr(k), 0#
    Next k

    Dim panelCount As Long
    panelCount = 0

    AR_ScanPantsPanelsRecursive g, vote, errSum, panelCount

    Dim bestSize As String
    Dim bestVote As Long
    Dim bestErr As Double

    bestSize = ""
    bestVote = 0
    bestErr = 999999#

    For Each k In vote.keys

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
        AR_DetectPantsGroupSize = bestSize
    End If

End Function

Private Sub AR_ScanPantsPanelsRecursive( _
    ByVal s As Shape, _
    ByRef vote As Object, _
    ByRef errSum As Object, _
    ByRef panelCount As Long)

    Dim c As Shape
    Dim pcShapes As Shapes

    On Error Resume Next

    If s.Type = cdrGroupShape Then

        For Each c In s.Shapes
            AR_ScanPantsPanelsRecursive c, vote, errSum, panelCount
        Next c

        Exit Sub

    End If

    Set pcShapes = s.PowerClip.Shapes

    If Not pcShapes Is Nothing Then
        For Each c In pcShapes
            AR_ScanPantsPanelsRecursive c, vote, errSum, panelCount
        Next c
    End If

    On Error GoTo 0

    If s.Type <> cdrCurveShape Then Exit Sub
    If Not AR_IsPanelOutline(s) Then Exit Sub

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

    AR_AddPantsPanelVote mn, vote, errSum

End Sub

Private Sub AR_AddPantsPanelVote( _
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

    For Each k In AR_SizeDB.keys

        db = AR_SizeDB(k)

        If UBound(db) >= 2 Then

            errFront = Abs(panelWidth - AR_ToDbl(db(1)))
            errBack = Abs(panelWidth - AR_ToDbl(db(2)))

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

        If closestErr <= AR_PANTS_TOL Then
            vote(closestSize) = CLng(vote(closestSize)) + 1
            errSum(closestSize) = CDbl(errSum(closestSize)) + closestErr
        End If

    End If

End Sub

Private Function AR_DetectSizeFromShape(ByVal shp As Shape, ByRef area As Double) As String

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

    For Each k In AR_SizeDB.keys

        db = AR_SizeDB(k)

        If AR_IsPants Then

            If UBound(db) >= 2 Then

                If Abs(mn - AR_ToDbl(db(1))) <= AR_PANTS_TOL Or _
                   Abs(mn - AR_ToDbl(db(2))) <= AR_PANTS_TOL Then

                    AR_DetectSizeFromShape = CStr(k)
                    Exit Function

                End If

            End If

        ElseIf AR_IsSplitFront Then

            If UBound(db) >= 4 Then

                If Abs(mn - AR_ToDbl(db(1))) <= AR_SIZE_TOL And _
                   Abs(mx - AR_ToDbl(db(4))) <= AR_SIZE_TOL Then

                    AR_DetectSizeFromShape = CStr(k) & "_BACK"
                    Exit Function

                End If

                If Abs(mn - AR_ToDbl(db(2))) <= AR_SIZE_TOL And _
                   Abs(mx - AR_ToDbl(db(3))) <= AR_SIZE_TOL Then

                    AR_DetectSizeFromShape = CStr(k) & "_FRONT"
                    Exit Function

                End If

            End If

        Else

            If UBound(db) >= 3 Then

                If Abs(mn - AR_ToDbl(db(1))) <= AR_SIZE_TOL And _
                   (Abs(mx - AR_ToDbl(db(2))) <= AR_SIZE_TOL Or _
                    Abs(mx - AR_ToDbl(db(3))) <= AR_SIZE_TOL) Then

                    AR_DetectSizeFromShape = CStr(k)
                    Exit Function

                End If

            End If

        End If

    Next k

End Function


'=========================================================
' OUTLINE DETECTION
'=========================================================

Private Function AR_IsPanelOutline(ByVal shp As Shape) As Boolean

    AR_IsPanelOutline = False

    If AR_IsRedOutline(shp) Then
        AR_IsPanelOutline = True
        Exit Function
    End If

    If AR_IsGreenOutline(shp) Then
        AR_IsPanelOutline = True
        Exit Function
    End If

End Function

Private Function AR_IsRedOutline(ByVal shp As Shape) As Boolean

    On Error Resume Next

    If shp.Outline Is Nothing Then Exit Function
    If shp.Outline.Type = cdrNoOutline Then Exit Function

    AR_IsRedOutline = _
        shp.Outline.Color.RGBRed > 200 And _
        shp.Outline.Color.RGBGreen < 80 And _
        shp.Outline.Color.RGBBlue < 80

    On Error GoTo 0

End Function

Private Function AR_IsGreenOutline(ByVal shp As Shape) As Boolean

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
        AR_IsGreenOutline = True
        On Error GoTo 0
        Exit Function
    End If

    'Hijau lama / alternatif
    AR_IsGreenOutline = _
        Abs(r - AR_GREEN_R) <= AR_GREEN_TOL And _
        Abs(g - AR_GREEN_G) <= AR_GREEN_TOL And _
        Abs(b - AR_GREEN_B) <= AR_GREEN_TOL

    On Error GoTo 0

End Function


'=========================================================
' APPLY RENAME
'=========================================================

Private Sub AR_ApplyRename()

    Dim k As Variant

    For Each k In AR_OrdersBySize.keys

        If AR_GroupsBySize.Exists(k) Then
            AR_RenameSizeBatch CStr(k), AR_GroupsBySize(k), AR_OrdersBySize(k)
        End If

    Next k

End Sub

Private Sub AR_RenameSizeBatch( _
    ByVal sz As String, _
    ByVal groups As Collection, _
    ByVal orders As Collection)

    Dim n As Long
    n = groups.Count

    If n = 0 Then Exit Sub

    Dim arr() As Shape
    ReDim arr(1 To n)

    Dim i As Long

    For i = 1 To n
        Set arr(i) = groups(i)
    Next i

    AR_SortGroupsVisual arr, n

    For i = 1 To n
        AR_RenameOneGroup arr(i), orders(i), sz, i
    Next i

End Sub

Private Sub AR_RenameOneGroup( _
    ByVal g As Shape, _
    ByVal ord As Variant, _
    ByVal sz As String, _
    ByVal idxInSize As Long)

    Dim nm As String
    Dim no As String
    Dim nick As String

    nm = AR_OrderField(ord, 1)
    no = AR_OrderField(ord, 2)
    nick = AR_OrderField(ord, 3)

    Dim nameR As Long
    Dim numR As Long
    Dim nickR As Long

    nameR = 0
    numR = 0
    nickR = 0

    Set AR_DeleteQueue = New Collection

    AR_RenameTextRecursive g, nm, no, nick, nameR, numR, nickR

    AR_DeleteQueuedTextShapes
    AR_CleanupEmptyGroupsRecursive g

    If AR_IsPants Then

        If Trim$(no) <> "" And numR = 0 Then

            AR_Warning = AR_Warning & _
                "SIZE " & sz & " #" & idxInSize & vbCrLf & _
                "Nomor celana tidak terganti. Placeholder nomor tidak ditemukan." & vbCrLf & _
                "Expected : " & no & vbCrLf & _
                String(35, "-") & vbCrLf

        End If

    Else

        If Trim$(nm) <> "" And nameR = 0 Then

            AR_Warning = AR_Warning & _
                "SIZE " & sz & " #" & idxInSize & vbCrLf & _
                "Nama tidak terganti. Placeholder nama tidak ditemukan." & vbCrLf & _
                "Expected : " & nm & vbCrLf & _
                String(35, "-") & vbCrLf

        End If

        If Trim$(no) <> "" And numR = 0 Then

            AR_Warning = AR_Warning & _
                "SIZE " & sz & " #" & idxInSize & vbCrLf & _
                "Nomor tidak terganti. Placeholder nomor tidak ditemukan." & vbCrLf & _
                "Expected : " & no & vbCrLf & _
                String(35, "-") & vbCrLf

        End If

        If Trim$(nick) <> "" And nickR = 0 Then

            AR_Warning = AR_Warning & _
                "SIZE " & sz & " #" & idxInSize & vbCrLf & _
                "Nickname tidak terganti. Placeholder nickname tidak ditemukan." & vbCrLf & _
                "Expected : " & nick & vbCrLf & _
                String(35, "-") & vbCrLf

        End If

    End If

End Sub

Private Sub AR_RenameTextRecursive( _
    ByVal s As Shape, _
    ByVal nm As String, _
    ByVal no As String, _
    ByVal nick As String, _
    ByRef nameR As Long, _
    ByRef numR As Long, _
    ByRef nickR As Long)

    Dim c As Shape
    Dim pcShapes As Shapes

    On Error Resume Next

    If s.Type = cdrGroupShape Then

        For Each c In s.Shapes
            AR_RenameTextRecursive c, nm, no, nick, nameR, numR, nickR
        Next c

        Exit Sub

    End If

    Set pcShapes = s.PowerClip.Shapes

    If Not pcShapes Is Nothing Then
        For Each c In pcShapes
            AR_RenameTextRecursive c, nm, no, nick, nameR, numR, nickR
        Next c
    End If

    On Error GoTo 0

    If s.Type <> cdrTextShape Then Exit Sub

    Dim raw As String
    raw = ""

    On Error Resume Next
    raw = s.Text.Story.Text
    On Error GoTo 0

    raw = Trim$(raw)

    If raw = "" Then
        AR_QueueDeleteTextShape s
        Exit Sub
    End If

    If AR_IgnoreSmallID(s, raw) Then Exit Sub

    If AR_IsPants Then

        If AR_IsNumberPlaceholder(s, raw) Then

            If AR_SetTextSafe(s, no) Then
                numR = numR + 1
            End If

            Exit Sub

        End If

        Exit Sub

    End If

    If AR_IsNamePlaceholder(raw) Then

        If AR_SetTextSafe(s, nm) Then
            nameR = nameR + 1
        End If

        Exit Sub

    End If

    If AR_IsNicknamePlaceholder(raw) Then

        If AR_SetTextSafe(s, nick) Then
            nickR = nickR + 1
        End If

        Exit Sub

    End If

    If AR_IsNumberPlaceholder(s, raw) Then

        If AR_SetTextSafe(s, no) Then
            numR = numR + 1
        End If

        Exit Sub

    End If

End Sub

Private Function AR_SetTextSafe(ByVal t As Shape, ByVal newText As String) As Boolean

    On Error Resume Next

    newText = CStr(newText)
    newText = AR_RemoveLigatures(newText)
    newText = AR_BreakVisualLigatures(newText)

    If Trim$(newText) = "" Then
        AR_QueueDeleteTextShape t
        AR_SetTextSafe = True
        Exit Function
    End If

    Dim hasCJK As Boolean
    Dim fontReadyBefore As Boolean
    Dim fontReadyAfter As Boolean

    hasCJK = AR_ContainsJapaneseOrCJK(newText)

    If hasCJK Then
        AR_CJKDetectedCount = AR_CJKDetectedCount + 1
        fontReadyBefore = AR_ApplyJapaneseFont(t)
    End If

    err.Clear
    t.Text.Story.Text = newText

    If err.Number <> 0 Then
        AR_SetTextSafe = False
        err.Clear
        On Error GoTo 0
        Exit Function
    End If

    If hasCJK Then

        fontReadyAfter = AR_ApplyJapaneseFont(t)

        If Trim$(t.Text.Story.Text) = "" Then
            err.Clear
            t.Text.Story.Text = newText
            fontReadyAfter = AR_ApplyJapaneseFont(t)
        End If

        If fontReadyBefore Or fontReadyAfter Then
            AR_CJKFontAppliedCount = AR_CJKFontAppliedCount + 1
        Else
            AR_CJKFontFailedCount = AR_CJKFontFailedCount + 1
        End If

    End If

    AR_SetTextSafe = True

    On Error GoTo 0

End Function


'=========================================================
' LIGATURE CLEANER — ACTIVE REPAIR BEFORE WRITE
'=========================================================

Private Function AR_RemoveLigatures(ByVal s As String) As String

    On Error Resume Next

    s = Replace(s, ChrW(&HFB00), "ff")   ' ﬀ
    s = Replace(s, ChrW(&HFB01), "fi")   ' ﬁ
    s = Replace(s, ChrW(&HFB02), "fl")   ' ﬂ
    s = Replace(s, ChrW(&HFB03), "ffi")  ' ﬃ
    s = Replace(s, ChrW(&HFB04), "ffl")  ' ﬄ
    s = Replace(s, ChrW(&HFB05), "st")   ' ﬅ
    s = Replace(s, ChrW(&HFB06), "st")   ' ﬆ

    On Error GoTo 0

    AR_RemoveLigatures = s

End Function


'=========================================================
' VISUAL LIGATURE BREAKER — FONT RENDERING FIX
'=========================================================
' Catatan penting:
' AR_RemoveLigatures hanya memperbaiki karakter Unicode seperti "ﬁ".
' Jika font jersey memakai OpenType ligature, teks biasa "fi" masih bisa
' digambar menyatu oleh font. Fungsi ini menyisipkan Zero Width Non-Joiner
' di antara F + I/L/F agar bentuk visual "Fi" tidak otomatis bergabung.

Private Function AR_BreakVisualLigatures(ByVal s As String) As String

    If Not AR_ENABLE_VISUAL_LIGATURE_BREAKER Then
        AR_BreakVisualLigatures = s
        Exit Function
    End If

    On Error GoTo FAIL

    Dim z As String
    Dim r As String
    Dim i As Long
    Dim ch As String
    Dim nx As String

    z = ChrW(&H200C) ' Zero Width Non-Joiner

    For i = 1 To Len(s)

        ch = Mid$(s, i, 1)
        r = r & ch

        If i < Len(s) Then
            nx = Mid$(s, i + 1, 1)

            If AR_IsVisualLigaturePair(ch, nx) Then
                r = r & z
            End If
        End If

    Next i

    AR_BreakVisualLigatures = r
    Exit Function

FAIL:
    AR_BreakVisualLigatures = s

End Function

Private Function AR_IsVisualLigaturePair(ByVal a As String, ByVal b As String) As Boolean

    a = UCase$(a)
    b = UCase$(b)

    If a = "F" Then
        Select Case b
            Case "F", "I", "L"
                AR_IsVisualLigaturePair = True
        End Select
    End If

End Function


'=========================================================
' JAPANESE / CJK FONT PATCH
'=========================================================

Private Function AR_ContainsJapaneseOrCJK(ByVal s As String) As Boolean

    Dim i As Long
    Dim ch As String
    Dim code As Long

    For i = 1 To Len(s)

        ch = Mid$(s, i, 1)
        code = AscW(ch)

        If code < 0 Then code = code + 65536

        If AR_IsJapaneseOrCJKCode(code) Then
            AR_ContainsJapaneseOrCJK = True
            Exit Function
        End If

    Next i

End Function

Private Function AR_IsJapaneseOrCJKCode(ByVal code As Long) As Boolean

    Select Case code

        Case &H3000 To &H303F
            AR_IsJapaneseOrCJKCode = True

        Case &H3040 To &H309F
            AR_IsJapaneseOrCJKCode = True

        Case &H30A0 To &H30FF
            AR_IsJapaneseOrCJKCode = True

        Case &H31F0 To &H31FF
            AR_IsJapaneseOrCJKCode = True

        Case &H3400 To &H4DBF
            AR_IsJapaneseOrCJKCode = True

        Case &H4E00 To &H9FFF
            AR_IsJapaneseOrCJKCode = True

        Case &HF900 To &HFAFF
            AR_IsJapaneseOrCJKCode = True

        Case &HFF00 To &HFFEF
            AR_IsJapaneseOrCJKCode = True

    End Select

End Function

Private Function AR_ApplyJapaneseFont(ByVal t As Shape) As Boolean

    AR_ApplyJapaneseFont = False

    If t Is Nothing Then Exit Function

    If AR_TrySetFontName(t, AR_CJK_FONT_1, True) Then
        AR_ApplyJapaneseFont = True
        Exit Function
    End If

    If AR_TrySetFontName(t, AR_CJK_FONT_2, True) Then
        AR_ApplyJapaneseFont = True
        Exit Function
    End If

    If AR_TrySetFontName(t, AR_CJK_FONT_3, True) Then
        AR_ApplyJapaneseFont = True
        Exit Function
    End If

    If AR_TrySetFontName(t, AR_CJK_FONT_4, True) Then
        AR_ApplyJapaneseFont = True
        Exit Function
    End If

    If AR_TrySetFontName(t, AR_CJK_FONT_5, True) Then
        AR_ApplyJapaneseFont = True
        Exit Function
    End If

End Function

Private Function AR_TrySetFontName( _
    ByVal t As Shape, _
    ByVal fontName As String, _
    Optional ByVal makeBold As Boolean = True) As Boolean

    AR_TrySetFontName = False

    If t Is Nothing Then Exit Function
    If Trim$(fontName) = "" Then Exit Function

    On Error Resume Next

    err.Clear

    t.Text.Story.Font = fontName

    If err.Number <> 0 Then
        err.Clear
        On Error GoTo 0
        Exit Function
    End If

    If makeBold Then
        err.Clear
        t.Text.Story.Bold = True
        err.Clear
    End If

    AR_TrySetFontName = True

    On Error GoTo 0

End Function


'=========================================================
' DELETE QUEUE / EMPTY TEXT CLEANUP
'=========================================================

Private Sub AR_QueueDeleteTextShape(ByVal t As Shape)

    On Error Resume Next

    If AR_DeleteQueue Is Nothing Then
        Set AR_DeleteQueue = New Collection
    End If

    If Not t Is Nothing Then
        AR_DeleteQueue.Add t
    End If

    On Error GoTo 0

End Sub

Private Sub AR_DeleteQueuedTextShapes()

    Dim i As Long
    Dim t As Shape

    On Error Resume Next

    If AR_DeleteQueue Is Nothing Then Exit Sub

    For i = AR_DeleteQueue.Count To 1 Step -1

        Set t = AR_DeleteQueue(i)

        If Not t Is Nothing Then
            t.Locked = False
            t.Delete
        End If

    Next i

    Set AR_DeleteQueue = New Collection

    On Error GoTo 0

End Sub

Private Sub AR_CleanupEmptyGroupsRecursive(ByVal g As Shape)

    Dim i As Long
    Dim ch As Shape

    On Error Resume Next

    If g Is Nothing Then Exit Sub
    If g.Type <> cdrGroupShape Then Exit Sub

    For i = g.Shapes.Count To 1 Step -1

        Set ch = g.Shapes(i)

        If ch.Type = cdrGroupShape Then

            AR_CleanupEmptyGroupsRecursive ch

            If ch.Shapes.Count = 0 Then
                ch.Locked = False
                ch.Delete
            End If

        End If

    Next i

    On Error GoTo 0

End Sub


'=========================================================
' SORT GROUPS VISUAL
'=========================================================

Private Sub AR_SortGroupsVisual(ByRef arr() As Shape, ByVal n As Long)

    Dim i As Long
    Dim j As Long
    Dim tmp As Shape

    For i = 1 To n - 1

        For j = i + 1 To n

            If AR_NeedSwap(arr(i), arr(j)) Then

                Set tmp = arr(i)
                Set arr(i) = arr(j)
                Set arr(j) = tmp

            End If

        Next j

    Next i

End Sub

Private Function AR_NeedSwap(ByVal a As Shape, ByVal b As Shape) As Boolean

    Dim ax As Double
    Dim ay As Double
    Dim bx As Double
    Dim by As Double

    ax = a.PositionX
    ay = a.PositionY
    bx = b.PositionX
    by = b.PositionY

    If Abs(ay - by) > AR_ROW_TOL Then

        AR_NeedSwap = (ay < by)

    Else

        AR_NeedSwap = (ax > bx)

    End If

End Function


'=========================================================
' PLACEHOLDER RULES
'=========================================================

Private Function AR_IsNamePlaceholder(ByVal txt As String) As Boolean

    Dim s As String
    s = AR_Normalize(txt)

    Select Case s

        Case "NAMA ATLIT", _
             "NAMA ATLET", _
             "NAMA", _
             "PLAYER", _
             "PLAYERS", _
             "PLAYER NAME", _
             "NAMA PEMAIN"

            AR_IsNamePlaceholder = True

    End Select

End Function

Private Function AR_IsNicknamePlaceholder(ByVal txt As String) As Boolean

    Dim s As String
    s = AR_Normalize(txt)

    Select Case s

        Case "NICKNAME", _
             "NICK NAME", _
             "NICK", _
             "NAMA PANGGILAN"

            AR_IsNicknamePlaceholder = True

    End Select

End Function

Private Function AR_IsNumberPlaceholder(ByVal t As Shape, ByVal txt As String) As Boolean

    Dim s As String
    s = Trim$(txt)

    If s = "" Then Exit Function
    If Len(s) > 3 Then Exit Function
    If Not IsNumeric(s) Then Exit Function

    If t.SizeHeight < AR_MIN_TEXT_H Then Exit Function

    AR_IsNumberPlaceholder = True

End Function

Private Function AR_IgnoreSmallID(ByVal t As Shape, ByVal txt As String) As Boolean

    Dim s As String
    s = Trim$(txt)

    If AR_Normalize(s) = "IDPO" Then

        If t.SizeHeight >= AR_ID_MIN_H And t.SizeHeight <= AR_ID_MAX_H Then
            AR_IgnoreSmallID = True
        End If

        Exit Function

    End If

    If Len(s) = 6 And IsNumeric(s) Then

        If t.SizeHeight >= AR_ID_MIN_H And t.SizeHeight <= AR_ID_MAX_H Then
            AR_IgnoreSmallID = True
        End If

    End If

End Function


'=========================================================
' TEXT / ORDER HELPERS
'=========================================================

Private Function AR_Normalize(ByVal s As String) As String

    s = Replace(s, vbCr, "")
    s = Replace(s, vbLf, "")
    s = Replace(s, Chr(160), " ")

    On Error Resume Next
    s = Replace(s, ChrW(&H200C), "")
    s = Replace(s, ChrW(&H200D), "")
    s = Replace(s, ChrW(&HFB01), "FI")
    s = Replace(s, ChrW(&HFB02), "FL")
    On Error GoTo 0

    Do While InStr(1, s, "  ", vbTextCompare) > 0
        s = Replace(s, "  ", " ")
    Loop

    AR_Normalize = UCase$(Trim$(s))

End Function

Private Function AR_OrderField(ByVal ord As Variant, ByVal idx As Long) As String

    On Error Resume Next

    If IsArray(ord) Then
        If UBound(ord) >= idx Then
            AR_OrderField = CStr(ord(idx))
        End If
    End If

    On Error GoTo 0

End Function

Private Function AR_ToDbl(ByVal v As Variant) As Double

    Dim s As String
    s = Trim$(CStr(v))
    s = Replace(s, ",", ".")

    AR_ToDbl = val(s)

End Function

Private Function AR_NormalizeSizeKey(ByVal sz As String) As String

    Dim s As String

    s = UCase$(Trim$(sz))

    Select Case s

        Case "XXL"
            AR_NormalizeSizeKey = "2XL"

        Case "XXXL"
            AR_NormalizeSizeKey = "3XL"

        Case "XXXXL"
            AR_NormalizeSizeKey = "4XL"

        Case "XXXXXL"
            AR_NormalizeSizeKey = "5XL"

        Case "XXXXXXL"
            AR_NormalizeSizeKey = "6XL"

        Case Else
            AR_NormalizeSizeKey = s

    End Select

End Function

Private Function AR_ProductModeText() As String

    If AR_IsPants Then
        AR_ProductModeText = "CELANA"
    ElseIf AR_IsSplitFront Then
        AR_ProductModeText = "JAKET"
    Else
        AR_ProductModeText = "JERSEY"
    End If

End Function

