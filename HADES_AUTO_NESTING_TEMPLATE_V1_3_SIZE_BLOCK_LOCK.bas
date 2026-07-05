Option Explicit

'=========================================================
' PROJECT HADES - AUTO NESTING TEMPLATE V1.3 SIZE BLOCK LOCK
' CorelDRAW 2021 VBA
'
' BASIS:
' - Dibangun dari mesin lama V1.2D ROTATION SAFE.
' - Tetap template-based nesting, bukan nesting bebas.
' - Tetap memakai width/height/area + position tie-breaker + rotation-safe.
'
' PERKEMBANGAN V1.3:
' - SIZE BLOCK LOCK.
' - Tidak memakai warna outline sebagai identitas size.
' - Setiap komponen dikunci ke blok horizontal size dari kiri ke kanan.
' - Cocok untuk kasus kerah/tulangan yang dimensinya sama antar size.
'
' PRINSIP PRODUKSI:
' - Paling kiri = size terkecil.
' - Paling kanan = size terbesar.
' - Template LRP boleh masih memakai warna pola lama.
' - Master layout produksi tetap disarankan memakai outline merah.
'
' MACRO UTAMA:
' - HADES_NESTING_SELF_TEST
' - HADES_BUILD_NESTING_TEMPLATE
' - HADES_APPLY_AUTO_NESTING
' - HADES_OPEN_NESTING_TEMPLATE_FILE
' - HADES_OPEN_NESTING_LOG_FILE
'=========================================================


'=========================
' CONFIG
'=========================
Private Const TEMPLATE_FILE As String = "\Documents\HADES_NESTING_TEMPLATE_CURRENT.txt"
Private Const LOG_FILE As String = "\Documents\HADES_NESTING_LOG_LATEST.txt"
Private Const ORDER_FILE As String = "\Documents\Order.txt"

Private Const VERSION_TEXT As String = "1.3 SIZE BLOCK LOCK"

Private Const MAX_ITEMS As Long = 250
Private Const MAX_BLOCKS As Long = 60

' V1.3 lebih aman daripada V1.2D:
' jumlah komponen harus sama supaya tidak ada komponen template/target yang tertinggal.
Private Const STRICT_COUNT_MATCH As Boolean = True

' Jika True, slot template blok 1 hanya boleh mengambil target blok 1, dst.
Private Const ENABLE_SIZE_BLOCK_LOCK As Boolean = True

' Jika True, jumlah komponen per blok template dan target harus sama.
Private Const STRICT_BLOCK_POPULATION As Boolean = True

' Shape matching lama tetap menjadi dasar utama.
Private Const POSITION_WEIGHT As Double = 0.45
Private Const ROTATION_WEIGHT As Double = 0.18

' Aktifkan auto-rotate agar target mengikuti orientasi template.
Private Const ENABLE_AUTO_ROTATE As Boolean = True
Private Const ROTATION_TOLERANCE_DEG As Double = 3#

' Filter objek terlalu kecil. Jangan terlalu besar agar tulangan tetap ikut.
Private Const MIN_COMPONENT_AREA_CM2 As Double = 1#


'=========================================================
' SELF TEST
'=========================================================
Public Sub HADES_NESTING_SELF_TEST()

    On Error GoTo ErrHandler

    WriteLog "SELF TEST START"

    MsgBox _
        "HADES AUTO NESTING V" & VERSION_TEXT & vbCrLf & vbCrLf & _
        "Macro berhasil terpanggil." & vbCrLf & vbCrLf & _
        "Fitur aktif:" & vbCrLf & _
        "- Template-based nesting" & vbCrLf & _
        "- Hybrid shape matching" & vbCrLf & _
        "- Rotation safe" & vbCrLf & _
        "- Size Block Lock kiri ke kanan", _
        vbInformation, _
        "HADES NESTING SELF TEST"

    WriteLog "SELF TEST OK"

    Exit Sub

ErrHandler:

    MsgBox _
        "SELF TEST ERROR" & vbCrLf & vbCrLf & _
        "Error " & Err.Number & ":" & vbCrLf & Err.Description, _
        vbCritical, _
        "HADES NESTING SELF TEST"

End Sub


'=========================================================
' BUILD TEMPLATE
'=========================================================
Public Sub HADES_BUILD_NESTING_TEMPLATE()

    Dim oldUnit As cdrUnit
    Dim sr As ShapeRange
    Dim cnt As Long

    Dim names(1 To MAX_ITEMS) As String
    Dim cx(1 To MAX_ITEMS) As Double
    Dim cy(1 To MAX_ITEMS) As Double
    Dim w(1 To MAX_ITEMS) As Double
    Dim h(1 To MAX_ITEMS) As Double
    Dim area(1 To MAX_ITEMS) As Double
    Dim ang(1 To MAX_ITEMS) As Double
    Dim blockIdx(1 To MAX_ITEMS) As Long
    Dim localX(1 To MAX_ITEMS) As Double
    Dim localY(1 To MAX_ITEMS) As Double

    Dim blockLeft(1 To MAX_BLOCKS) As Double
    Dim blockTop(1 To MAX_BLOCKS) As Double
    Dim blockRight(1 To MAX_BLOCKS) As Double
    Dim blockBottom(1 To MAX_BLOCKS) As Double

    Dim baseLeft As Double
    Dim baseTop As Double
    Dim baseRight As Double
    Dim baseBottom As Double
    Dim blockCount As Long

    On Error GoTo ErrHandler

    If ActiveSelection.Shapes.Count = 0 Then
        MsgBox _
            "Tidak ada objek dipilih." & vbCrLf & vbCrLf & _
            "Select template LRP putih yang sudah di-ungroup sekali." & vbCrLf & _
            "Jangan deep ungroup sampai teks/atribut kecil pecah semua.", _
            vbExclamation, _
            "BUILD NESTING TEMPLATE"
        Exit Sub
    End If

    oldUnit = ActiveDocument.Unit
    ActiveDocument.Unit = cdrCentimeter

    Set sr = ActiveSelectionRange

    cnt = CollectTopLevelShapes(sr, names, cx, cy, w, h, area, ang)

    If cnt <= 0 Then
        MsgBox _
            "Tidak ada komponen valid pada selection.", _
            vbExclamation, _
            "BUILD NESTING TEMPLATE"
        GoTo CleanExit
    End If

    GetSelectionBounds sr, baseLeft, baseTop, baseRight, baseBottom

    If ENABLE_SIZE_BLOCK_LOCK Then

        blockCount = ResolveSizeBlockCount(cnt)

        If blockCount <= 0 Then
            MsgBox _
                "Jumlah blok size tidak bisa ditentukan." & vbCrLf & vbCrLf & _
                "Pastikan Order.txt ada, atau masukkan jumlah blok size saat diminta.", _
                vbExclamation, _
                "BUILD NESTING TEMPLATE"
            GoTo CleanExit
        End If

        If blockCount > cnt Then
            MsgBox _
                "Jumlah blok size lebih banyak daripada jumlah komponen." & vbCrLf & vbCrLf & _
                "Blok size : " & blockCount & vbCrLf & _
                "Komponen  : " & cnt, _
                vbCritical, _
                "BUILD NESTING TEMPLATE"
            GoTo CleanExit
        End If

        AssignSizeBlocks _
            cnt, cx, cy, w, h, blockCount, _
            blockIdx, blockLeft, blockTop, blockRight, blockBottom

        FillLocalPositions _
            cnt, cx, cy, blockIdx, blockLeft, blockTop, localX, localY

    Else

        blockCount = 1
        AssignSingleBlock cnt, cx, cy, w, h, blockIdx, blockLeft, blockTop, blockRight, blockBottom
        FillLocalPositions cnt, cx, cy, blockIdx, blockLeft, blockTop, localX, localY

    End If

    SortItemsByAreaDescV13 cnt, names, cx, cy, w, h, area, ang, blockIdx, localX, localY

    SaveTemplate _
        cnt, names, cx, cy, w, h, area, ang, blockIdx, localX, localY, _
        blockCount, baseLeft, baseTop, baseRight, baseBottom

    MsgBox _
        "NESTING TEMPLATE SAVED" & vbCrLf & vbCrLf & _
        "Versi          : " & VERSION_TEXT & vbCrLf & _
        "Jumlah komponen: " & cnt & vbCrLf & _
        "Blok size      : " & blockCount & vbCrLf & vbCrLf & _
        "File:" & vbCrLf & _
        Environ$("USERPROFILE") & TEMPLATE_FILE & vbCrLf & vbCrLf & _
        "Template siap dipakai oleh HADES_APPLY_AUTO_NESTING.", _
        vbInformation, _
        "BUILD NESTING TEMPLATE"

CleanExit:

    On Error Resume Next
    ActiveDocument.Unit = oldUnit
    On Error GoTo 0

    Exit Sub

ErrHandler:

    On Error Resume Next
    ActiveDocument.Unit = oldUnit
    On Error GoTo 0

    MsgBox _
        "BUILD NESTING TEMPLATE ERROR" & vbCrLf & vbCrLf & _
        "Error " & Err.Number & ":" & vbCrLf & Err.Description, _
        vbCritical, _
        "BUILD NESTING TEMPLATE"

End Sub


'=========================================================
' APPLY TEMPLATE
'=========================================================
Public Sub HADES_APPLY_AUTO_NESTING()

    Dim oldUnit As cdrUnit
    Dim sr As ShapeRange

    Dim slotCount As Long
    Dim itemCount As Long
    Dim blockCount As Long

    Dim slotName(1 To MAX_ITEMS) As String
    Dim slotRelX(1 To MAX_ITEMS) As Double
    Dim slotRelY(1 To MAX_ITEMS) As Double
    Dim slotW(1 To MAX_ITEMS) As Double
    Dim slotH(1 To MAX_ITEMS) As Double
    Dim slotArea(1 To MAX_ITEMS) As Double
    Dim slotAngle(1 To MAX_ITEMS) As Double
    Dim slotBlock(1 To MAX_ITEMS) As Long
    Dim slotLocalX(1 To MAX_ITEMS) As Double
    Dim slotLocalY(1 To MAX_ITEMS) As Double

    Dim targetShape(1 To MAX_ITEMS) As Shape
    Dim targetName(1 To MAX_ITEMS) As String
    Dim targetCX(1 To MAX_ITEMS) As Double
    Dim targetCY(1 To MAX_ITEMS) As Double
    Dim targetW(1 To MAX_ITEMS) As Double
    Dim targetH(1 To MAX_ITEMS) As Double
    Dim targetArea(1 To MAX_ITEMS) As Double
    Dim targetAngle(1 To MAX_ITEMS) As Double
    Dim targetRelX(1 To MAX_ITEMS) As Double
    Dim targetRelY(1 To MAX_ITEMS) As Double
    Dim targetBlock(1 To MAX_ITEMS) As Long
    Dim targetLocalX(1 To MAX_ITEMS) As Double
    Dim targetLocalY(1 To MAX_ITEMS) As Double

    Dim blockLeft(1 To MAX_BLOCKS) As Double
    Dim blockTop(1 To MAX_BLOCKS) As Double
    Dim blockRight(1 To MAX_BLOCKS) As Double
    Dim blockBottom(1 To MAX_BLOCKS) As Double

    Dim used(1 To MAX_ITEMS) As Boolean
    Dim matchIdx(1 To MAX_ITEMS) As Long

    Dim baseLeft As Double
    Dim baseTop As Double
    Dim baseRight As Double
    Dim baseBottom As Double

    Dim i As Long
    Dim bestIdx As Long
    Dim targetX As Double
    Dim targetY As Double
    Dim dx As Double
    Dim dy As Double

    On Error GoTo ErrHandler

    If ActiveSelection.Shapes.Count = 0 Then
        MsgBox _
            "Tidak ada objek dipilih." & vbCrLf & vbCrLf & _
            "Select master layout target yang ingin ditata.", _
            vbExclamation, _
            "APPLY AUTO NESTING"
        Exit Sub
    End If

    If Dir(Environ$("USERPROFILE") & TEMPLATE_FILE) = "" Then
        MsgBox _
            "Template nesting belum ada." & vbCrLf & vbCrLf & _
            "Jalankan HADES_BUILD_NESTING_TEMPLATE terlebih dahulu.", _
            vbExclamation, _
            "APPLY AUTO NESTING"
        Exit Sub
    End If

    oldUnit = ActiveDocument.Unit
    ActiveDocument.Unit = cdrCentimeter

    LoadTemplate _
        slotCount, slotName, slotRelX, slotRelY, slotW, slotH, slotArea, slotAngle, _
        slotBlock, slotLocalX, slotLocalY, blockCount

    If slotCount <= 0 Then
        MsgBox _
            "Template nesting kosong / tidak valid.", _
            vbExclamation, _
            "APPLY AUTO NESTING"
        GoTo CleanExit
    End If

    Set sr = ActiveSelectionRange

    itemCount = CollectTopLevelShapesWithRef( _
                    sr, targetShape, targetName, targetCX, targetCY, _
                    targetW, targetH, targetArea, targetAngle)

    If itemCount <= 0 Then
        MsgBox _
            "Tidak ada komponen target valid.", _
            vbExclamation, _
            "APPLY AUTO NESTING"
        GoTo CleanExit
    End If

    If STRICT_COUNT_MATCH Then
        If itemCount <> slotCount Then
            MsgBox _
                "AUTO NESTING DIBATALKAN" & vbCrLf & vbCrLf & _
                "Jumlah komponen target tidak sama dengan template." & vbCrLf & vbCrLf & _
                "Template : " & slotCount & vbCrLf & _
                "Target   : " & itemCount & vbCrLf & vbCrLf & _
                "V1.3 dibuat strict agar kerah/tulangan tidak tertukar atau tertinggal.", _
                vbCritical, _
                "APPLY AUTO NESTING"
            GoTo CleanExit
        End If
    End If

    If blockCount <= 0 Then blockCount = 1
    If blockCount > MAX_BLOCKS Then blockCount = MAX_BLOCKS

    GetSelectionBounds sr, baseLeft, baseTop, baseRight, baseBottom

    If ENABLE_SIZE_BLOCK_LOCK Then

        AssignSizeBlocks _
            itemCount, targetCX, targetCY, targetW, targetH, blockCount, _
            targetBlock, blockLeft, blockTop, blockRight, blockBottom

        FillLocalPositions _
            itemCount, targetCX, targetCY, targetBlock, _
            blockLeft, blockTop, targetLocalX, targetLocalY

        If STRICT_BLOCK_POPULATION Then
            If Not ValidateBlockPopulation(slotCount, slotBlock, itemCount, targetBlock, blockCount) Then
                MsgBox _
                    "AUTO NESTING DIBATALKAN" & vbCrLf & vbCrLf & _
                    "Jumlah komponen per blok size tidak sama." & vbCrLf & _
                    "Kemungkinan selection template atau target tidak sepadan." & vbCrLf & vbCrLf & _
                    BuildBlockPopulationReport(slotCount, slotBlock, itemCount, targetBlock, blockCount), _
                    vbCritical, _
                    "APPLY AUTO NESTING"
                GoTo CleanExit
            End If
        End If

    Else

        AssignSingleBlock itemCount, targetCX, targetCY, targetW, targetH, targetBlock, blockLeft, blockTop, blockRight, blockBottom
        FillLocalPositions itemCount, targetCX, targetCY, targetBlock, blockLeft, blockTop, targetLocalX, targetLocalY

    End If

    ' Simpan posisi relatif target SEBELUM objek dipindah.
    FillTargetRelativePositions itemCount, targetShape, baseLeft, baseTop, targetRelX, targetRelY

    ' Preflight matching dulu. Jangan memindah objek sebelum semua slot mendapat pasangan.
    For i = 1 To slotCount

        bestIdx = FindBestTargetForSlot( _
                    i, itemCount, used, _
                    slotRelX, slotRelY, slotW, slotH, slotArea, slotAngle, slotBlock, slotLocalX, slotLocalY, _
                    targetRelX, targetRelY, targetW, targetH, targetArea, targetAngle, targetBlock, targetLocalX, targetLocalY, _
                    baseRight - baseLeft, baseTop - baseBottom)

        If bestIdx <= 0 Then
            MsgBox _
                "AUTO NESTING DIBATALKAN" & vbCrLf & vbCrLf & _
                "Ada slot template yang tidak mendapat pasangan aman." & vbCrLf & vbCrLf & _
                "Slot      : " & i & vbCrLf & _
                "Blok size : " & slotBlock(i) & vbCrLf & vbCrLf & _
                "Penyebab paling umum:" & vbCrLf & _
                "- target belum di-ungroup sekali," & vbCrLf & _
                "- jumlah komponen per size berbeda," & vbCrLf & _
                "- template yang dipakai bukan template untuk layout ini.", _
                vbCritical, _
                "APPLY AUTO NESTING"
            GoTo CleanExit
        End If

        used(bestIdx) = True
        matchIdx(i) = bestIdx

    Next i

    ActiveDocument.BeginCommandGroup "Hades Auto Nesting V1.3 Size Block Lock"

    For i = 1 To slotCount

        bestIdx = matchIdx(i)

        If bestIdx > 0 Then

            ApplyRotationIfNeeded targetShape(bestIdx), slotAngle(i), slotW(i), slotH(i)

            targetX = baseLeft + slotRelX(i)
            targetY = baseTop - slotRelY(i)

            dx = targetX - ShapeCenterX(targetShape(bestIdx))
            dy = targetY - ShapeCenterY(targetShape(bestIdx))

            targetShape(bestIdx).Move dx, dy

        End If

    Next i

    ActiveDocument.EndCommandGroup

    ActiveWindow.Refresh

    MsgBox _
        "AUTO NESTING SELESAI" & vbCrLf & vbCrLf & _
        "Versi          : " & VERSION_TEXT & vbCrLf & _
        "Template slot  : " & slotCount & vbCrLf & _
        "Target item    : " & itemCount & vbCrLf & _
        "Blok size      : " & blockCount & vbCrLf & vbCrLf & _
        "Size Block Lock aktif: kerah/tulangan tidak boleh silang blok size." & vbCrLf & _
        "Tetap periksa hasil secara visual sebelum convert.", _
        vbInformation, _
        "APPLY AUTO NESTING"

CleanExit:

    On Error Resume Next
    ActiveDocument.Unit = oldUnit
    On Error GoTo 0

    Exit Sub

ErrHandler:

    On Error Resume Next
    ActiveDocument.EndCommandGroup
    ActiveDocument.Unit = oldUnit
    On Error GoTo 0

    MsgBox _
        "APPLY AUTO NESTING ERROR" & vbCrLf & vbCrLf & _
        "Error " & Err.Number & ":" & vbCrLf & Err.Description, _
        vbCritical, _
        "APPLY AUTO NESTING"

End Sub


'=========================================================
' OPEN FILES
'=========================================================
Public Sub HADES_OPEN_NESTING_TEMPLATE_FILE()

    Dim path As String

    path = Environ$("USERPROFILE") & TEMPLATE_FILE

    If Dir(path) = "" Then
        MsgBox _
            "Template file belum ada:" & vbCrLf & path, _
            vbExclamation, _
            "OPEN NESTING TEMPLATE"
        Exit Sub
    End If

    Shell "notepad.exe " & Chr$(34) & path & Chr$(34), vbNormalFocus

End Sub


Public Sub HADES_OPEN_NESTING_LOG_FILE()

    Dim path As String

    path = Environ$("USERPROFILE") & LOG_FILE

    If Dir(path) = "" Then
        MsgBox _
            "Log file belum ada:" & vbCrLf & path, _
            vbExclamation, _
            "OPEN NESTING LOG"
        Exit Sub
    End If

    Shell "notepad.exe " & Chr$(34) & path & Chr$(34), vbNormalFocus

End Sub


'=========================================================
' COLLECT SHAPES
'=========================================================
Private Function CollectTopLevelShapes( _
    ByVal sr As ShapeRange, _
    ByRef names() As String, _
    ByRef cx() As Double, _
    ByRef cy() As Double, _
    ByRef w() As Double, _
    ByRef h() As Double, _
    ByRef area() As Double, _
    ByRef ang() As Double) As Long

    Dim s As Shape
    Dim n As Long

    n = 0

    For Each s In sr.Shapes

        If n >= MAX_ITEMS Then Exit For

        If IsValidNestingComponent(s) Then

            n = n + 1

            names(n) = "ITEM_" & CStr(n)
            cx(n) = ShapeCenterX(s)
            cy(n) = ShapeCenterY(s)
            w(n) = Abs(CDbl(s.SizeWidth))
            h(n) = Abs(CDbl(s.SizeHeight))
            area(n) = w(n) * h(n)
            ang(n) = ShapeAngleSafe(s)

        End If

    Next s

    CollectTopLevelShapes = n

End Function


Private Function CollectTopLevelShapesWithRef( _
    ByVal sr As ShapeRange, _
    ByRef shp() As Shape, _
    ByRef names() As String, _
    ByRef cx() As Double, _
    ByRef cy() As Double, _
    ByRef w() As Double, _
    ByRef h() As Double, _
    ByRef area() As Double, _
    ByRef ang() As Double) As Long

    Dim s As Shape
    Dim n As Long

    n = 0

    For Each s In sr.Shapes

        If n >= MAX_ITEMS Then Exit For

        If IsValidNestingComponent(s) Then

            n = n + 1

            Set shp(n) = s
            names(n) = "ITEM_" & CStr(n)
            cx(n) = ShapeCenterX(s)
            cy(n) = ShapeCenterY(s)
            w(n) = Abs(CDbl(s.SizeWidth))
            h(n) = Abs(CDbl(s.SizeHeight))
            area(n) = w(n) * h(n)
            ang(n) = ShapeAngleSafe(s)

        End If

    Next s

    CollectTopLevelShapesWithRef = n

End Function


Private Function IsValidNestingComponent(ByVal s As Shape) As Boolean

    On Error GoTo SafeExit

    IsValidNestingComponent = False

    If s.SizeWidth <= 0 Then Exit Function
    If s.SizeHeight <= 0 Then Exit Function

    If Abs(s.SizeWidth * s.SizeHeight) < MIN_COMPONENT_AREA_CM2 Then Exit Function

    IsValidNestingComponent = True

SafeExit:

End Function


'=========================================================
' SIZE BLOCK LOCK
'=========================================================
Private Function ResolveSizeBlockCount(ByVal componentCount As Long) As Long

    Dim n As Long
    Dim answer As String

    n = GetOrderUniqueSizeCount()

    If n > 0 And n <= componentCount And n <= MAX_BLOCKS Then
        ResolveSizeBlockCount = n
        WriteLog "BLOCK COUNT FROM ORDER.TXT = " & n
        Exit Function
    End If

    answer = InputBox( _
        "Jumlah blok size tidak bisa dibaca dari Order.txt." & vbCrLf & vbCrLf & _
        "Masukkan jumlah blok size dari kiri ke kanan." & vbCrLf & _
        "Contoh: S, M, L, XL = 4 blok." & vbCrLf & vbCrLf & _
        "Catatan: ini hanya fallback. Alur normal tetap dari Order.txt.", _
        "HADES AUTO NESTING - SIZE BLOCK COUNT", _
        "")

    answer = Trim$(answer)

    If Len(answer) = 0 Then
        ResolveSizeBlockCount = 0
        Exit Function
    End If

    n = CLng(Val(answer))

    If n < 1 Then n = 0
    If n > componentCount Then n = 0
    If n > MAX_BLOCKS Then n = 0

    ResolveSizeBlockCount = n
    WriteLog "BLOCK COUNT FROM INPUTBOX = " & n

End Function


Private Function GetOrderUniqueSizeCount() As Long

    Dim path As String
    Dim f As Integer
    Dim ln As String
    Dim p As Long
    Dim sz As String
    Dim sizes(1 To 200) As String
    Dim n As Long

    On Error GoTo SafeExit

    path = Environ$("USERPROFILE") & ORDER_FILE

    If Dir(path) = "" Then Exit Function

    f = FreeFile
    Open path For Input As #f

    Do Until EOF(f)

        Line Input #f, ln
        ln = Trim$(ln)

        If Len(ln) = 0 Then GoTo NextLine
        If Left$(ln, 1) = "@" Then GoTo NextLine

        p = InStr(1, ln, "|", vbTextCompare)
        If p <= 1 Then GoTo NextLine

        sz = UCase$(Trim$(Left$(ln, p - 1)))

        If IsLikelySizeCode(sz) Then
            If Not ExistsInStringArray(sizes, n, sz) Then
                If n < 200 Then
                    n = n + 1
                    sizes(n) = sz
                End If
            End If
        End If

NextLine:

    Loop

    Close #f

    GetOrderUniqueSizeCount = n
    Exit Function

SafeExit:

    On Error Resume Next
    If f > 0 Then Close #f
    GetOrderUniqueSizeCount = 0

End Function


Private Function IsLikelySizeCode(ByVal s As String) As Boolean

    s = UCase$(Trim$(s))

    Select Case s
        Case "XXS", "XS", "S", "M", "L", "XL", "2XL", "3XL", "4XL", "5XL", "6XL", "7XL", "8XL"
            IsLikelySizeCode = True
        Case Else
            IsLikelySizeCode = False
    End Select

End Function


Private Function ExistsInStringArray( _
    ByRef arr() As String, _
    ByVal cnt As Long, _
    ByVal value As String) As Boolean

    Dim i As Long

    For i = 1 To cnt
        If UCase$(Trim$(arr(i))) = UCase$(Trim$(value)) Then
            ExistsInStringArray = True
            Exit Function
        End If
    Next i

    ExistsInStringArray = False

End Function


Private Sub AssignSingleBlock( _
    ByVal cnt As Long, _
    ByRef cx() As Double, _
    ByRef cy() As Double, _
    ByRef w() As Double, _
    ByRef h() As Double, _
    ByRef blockIdx() As Long, _
    ByRef blockLeft() As Double, _
    ByRef blockTop() As Double, _
    ByRef blockRight() As Double, _
    ByRef blockBottom() As Double)

    Dim i As Long
    Dim first As Boolean

    first = True

    For i = 1 To cnt

        blockIdx(i) = 1

        If first Then
            blockLeft(1) = cx(i) - (w(i) / 2#)
            blockRight(1) = cx(i) + (w(i) / 2#)
            blockTop(1) = cy(i) + (h(i) / 2#)
            blockBottom(1) = cy(i) - (h(i) / 2#)
            first = False
        Else
            If cx(i) - (w(i) / 2#) < blockLeft(1) Then blockLeft(1) = cx(i) - (w(i) / 2#)
            If cx(i) + (w(i) / 2#) > blockRight(1) Then blockRight(1) = cx(i) + (w(i) / 2#)
            If cy(i) + (h(i) / 2#) > blockTop(1) Then blockTop(1) = cy(i) + (h(i) / 2#)
            If cy(i) - (h(i) / 2#) < blockBottom(1) Then blockBottom(1) = cy(i) - (h(i) / 2#)
        End If

    Next i

End Sub


Private Sub AssignSizeBlocks( _
    ByVal cnt As Long, _
    ByRef cx() As Double, _
    ByRef cy() As Double, _
    ByRef w() As Double, _
    ByRef h() As Double, _
    ByVal blockCount As Long, _
    ByRef blockIdx() As Long, _
    ByRef blockLeft() As Double, _
    ByRef blockTop() As Double, _
    ByRef blockRight() As Double, _
    ByRef blockBottom() As Double)

    Dim centroid(1 To MAX_BLOCKS) As Double
    Dim rawBlock(1 To MAX_ITEMS) As Long
    Dim mapRank(1 To MAX_BLOCKS) As Long
    Dim sumX(1 To MAX_BLOCKS) As Double
    Dim cntB(1 To MAX_BLOCKS) As Long
    Dim minX As Double
    Dim maxX As Double
    Dim spanX As Double
    Dim i As Long
    Dim b As Long
    Dim iter As Long
    Dim bestB As Long
    Dim bestD As Double
    Dim d As Double
    Dim first As Boolean

    If blockCount <= 1 Then
        AssignSingleBlock cnt, cx, cy, w, h, blockIdx, blockLeft, blockTop, blockRight, blockBottom
        Exit Sub
    End If

    minX = cx(1)
    maxX = cx(1)

    For i = 2 To cnt
        If cx(i) < minX Then minX = cx(i)
        If cx(i) > maxX Then maxX = cx(i)
    Next i

    spanX = maxX - minX
    If spanX <= 0 Then spanX = 1#

    For b = 1 To blockCount
        centroid(b) = minX + ((CDbl(b) - 0.5) * spanX / CDbl(blockCount))
    Next b

    For iter = 1 To 15

        For b = 1 To blockCount
            sumX(b) = 0#
            cntB(b) = 0
        Next b

        For i = 1 To cnt

            bestB = 1
            bestD = Abs(cx(i) - centroid(1))

            For b = 2 To blockCount
                d = Abs(cx(i) - centroid(b))
                If d < bestD Then
                    bestD = d
                    bestB = b
                End If
            Next b

            rawBlock(i) = bestB
            sumX(bestB) = sumX(bestB) + cx(i)
            cntB(bestB) = cntB(bestB) + 1

        Next i

        For b = 1 To blockCount
            If cntB(b) > 0 Then
                centroid(b) = sumX(b) / CDbl(cntB(b))
            End If
        Next b

    Next iter

    ' Mapping centroid kiri-ke-kanan menjadi block 1..N.
    For b = 1 To blockCount
        mapRank(b) = 1
        For i = 1 To blockCount
            If centroid(i) < centroid(b) Then
                mapRank(b) = mapRank(b) + 1
            ElseIf centroid(i) = centroid(b) And i < b Then
                mapRank(b) = mapRank(b) + 1
            End If
        Next i
    Next b

    For b = 1 To blockCount
        cntB(b) = 0
    Next b

    For i = 1 To cnt
        blockIdx(i) = mapRank(rawBlock(i))
        If blockIdx(i) < 1 Then blockIdx(i) = 1
        If blockIdx(i) > blockCount Then blockIdx(i) = blockCount
        cntB(blockIdx(i)) = cntB(blockIdx(i)) + 1
    Next i

    For b = 1 To blockCount
        If cntB(b) = 0 Then
            Err.Raise vbObjectError + 1301, , _
                "Size Block Lock gagal: blok " & b & " kosong. Periksa jumlah blok size / selection."
        End If
    Next b

    For b = 1 To blockCount
        first = True

        For i = 1 To cnt

            If blockIdx(i) = b Then

                If first Then
                    blockLeft(b) = cx(i) - (w(i) / 2#)
                    blockRight(b) = cx(i) + (w(i) / 2#)
                    blockTop(b) = cy(i) + (h(i) / 2#)
                    blockBottom(b) = cy(i) - (h(i) / 2#)
                    first = False
                Else
                    If cx(i) - (w(i) / 2#) < blockLeft(b) Then blockLeft(b) = cx(i) - (w(i) / 2#)
                    If cx(i) + (w(i) / 2#) > blockRight(b) Then blockRight(b) = cx(i) + (w(i) / 2#)
                    If cy(i) + (h(i) / 2#) > blockTop(b) Then blockTop(b) = cy(i) + (h(i) / 2#)
                    If cy(i) - (h(i) / 2#) < blockBottom(b) Then blockBottom(b) = cy(i) - (h(i) / 2#)
                End If

            End If

        Next i
    Next b

End Sub


Private Sub FillLocalPositions( _
    ByVal cnt As Long, _
    ByRef cx() As Double, _
    ByRef cy() As Double, _
    ByRef blockIdx() As Long, _
    ByRef blockLeft() As Double, _
    ByRef blockTop() As Double, _
    ByRef localX() As Double, _
    ByRef localY() As Double)

    Dim i As Long
    Dim b As Long

    For i = 1 To cnt
        b = blockIdx(i)
        If b < 1 Then b = 1
        If b > MAX_BLOCKS Then b = MAX_BLOCKS
        localX(i) = cx(i) - blockLeft(b)
        localY(i) = blockTop(b) - cy(i)
    Next i

End Sub


Private Function ValidateBlockPopulation( _
    ByVal slotCount As Long, _
    ByRef slotBlock() As Long, _
    ByVal itemCount As Long, _
    ByRef targetBlock() As Long, _
    ByVal blockCount As Long) As Boolean

    Dim sCnt(1 To MAX_BLOCKS) As Long
    Dim tCnt(1 To MAX_BLOCKS) As Long
    Dim i As Long
    Dim b As Long

    For i = 1 To slotCount
        b = slotBlock(i)
        If b >= 1 And b <= MAX_BLOCKS Then sCnt(b) = sCnt(b) + 1
    Next i

    For i = 1 To itemCount
        b = targetBlock(i)
        If b >= 1 And b <= MAX_BLOCKS Then tCnt(b) = tCnt(b) + 1
    Next i

    ValidateBlockPopulation = True

    For b = 1 To blockCount
        If sCnt(b) <> tCnt(b) Then
            ValidateBlockPopulation = False
            Exit Function
        End If
    Next b

End Function


Private Function BuildBlockPopulationReport( _
    ByVal slotCount As Long, _
    ByRef slotBlock() As Long, _
    ByVal itemCount As Long, _
    ByRef targetBlock() As Long, _
    ByVal blockCount As Long) As String

    Dim sCnt(1 To MAX_BLOCKS) As Long
    Dim tCnt(1 To MAX_BLOCKS) As Long
    Dim i As Long
    Dim b As Long
    Dim msg As String

    For i = 1 To slotCount
        b = slotBlock(i)
        If b >= 1 And b <= MAX_BLOCKS Then sCnt(b) = sCnt(b) + 1
    Next i

    For i = 1 To itemCount
        b = targetBlock(i)
        If b >= 1 And b <= MAX_BLOCKS Then tCnt(b) = tCnt(b) + 1
    Next i

    msg = ""

    For b = 1 To blockCount
        msg = msg & "Blok " & b & " | Template: " & sCnt(b) & " | Target: " & tCnt(b) & vbCrLf
    Next b

    BuildBlockPopulationReport = msg

End Function


'=========================================================
' BOUNDS / POSITION HELPERS
'=========================================================
Private Sub GetSelectionBounds( _
    ByVal sr As ShapeRange, _
    ByRef leftX As Double, _
    ByRef topY As Double, _
    ByRef rightX As Double, _
    ByRef bottomY As Double)

    Dim s As Shape
    Dim first As Boolean

    first = True

    For Each s In sr.Shapes

        If IsValidNestingComponent(s) Then

            If first Then

                leftX = ShapeLeftX(s)
                rightX = ShapeRightX(s)
                topY = ShapeTopY(s)
                bottomY = ShapeBottomY(s)

                first = False

            Else

                If ShapeLeftX(s) < leftX Then leftX = ShapeLeftX(s)
                If ShapeRightX(s) > rightX Then rightX = ShapeRightX(s)
                If ShapeTopY(s) > topY Then topY = ShapeTopY(s)
                If ShapeBottomY(s) < bottomY Then bottomY = ShapeBottomY(s)

            End If

        End If

    Next s

End Sub


Private Function ShapeLeftX(ByVal s As Shape) As Double
    ShapeLeftX = CDbl(s.leftX)
End Function


Private Function ShapeRightX(ByVal s As Shape) As Double
    ShapeRightX = CDbl(s.rightX)
End Function


Private Function ShapeTopY(ByVal s As Shape) As Double
    ShapeTopY = CDbl(s.topY)
End Function


Private Function ShapeBottomY(ByVal s As Shape) As Double
    ShapeBottomY = CDbl(s.bottomY)
End Function


Private Function ShapeCenterX(ByVal s As Shape) As Double
    ShapeCenterX = (ShapeLeftX(s) + ShapeRightX(s)) / 2#
End Function


Private Function ShapeCenterY(ByVal s As Shape) As Double
    ShapeCenterY = (ShapeTopY(s) + ShapeBottomY(s)) / 2#
End Function


'=========================================================
' TEMPLATE SAVE / LOAD
'=========================================================
Private Sub SaveTemplate( _
    ByVal cnt As Long, _
    ByRef names() As String, _
    ByRef cx() As Double, _
    ByRef cy() As Double, _
    ByRef w() As Double, _
    ByRef h() As Double, _
    ByRef area() As Double, _
    ByRef ang() As Double, _
    ByRef blockIdx() As Long, _
    ByRef localX() As Double, _
    ByRef localY() As Double, _
    ByVal blockCount As Long, _
    ByVal baseLeft As Double, _
    ByVal baseTop As Double, _
    ByVal baseRight As Double, _
    ByVal baseBottom As Double)

    Dim path As String
    Dim f As Integer
    Dim i As Long

    path = Environ$("USERPROFILE") & TEMPLATE_FILE

    f = FreeFile

    Open path For Output As #f

    Print #f, "HADES_NESTING_TEMPLATE_VERSION=" & VERSION_TEXT
    Print #f, "COUNT=" & cnt
    Print #f, "BLOCK_COUNT=" & blockCount
    Print #f, "BASE_WIDTH=" & FormatNum(baseRight - baseLeft)
    Print #f, "BASE_HEIGHT=" & FormatNum(baseTop - baseBottom)

    For i = 1 To cnt

        Print #f, "SLOT_" & i & "_NAME=" & names(i)
        Print #f, "SLOT_" & i & "_BLOCK=" & CStr(blockIdx(i))
        Print #f, "SLOT_" & i & "_REL_X=" & FormatNum(cx(i) - baseLeft)
        Print #f, "SLOT_" & i & "_REL_Y=" & FormatNum(baseTop - cy(i))
        Print #f, "SLOT_" & i & "_LOCAL_X=" & FormatNum(localX(i))
        Print #f, "SLOT_" & i & "_LOCAL_Y=" & FormatNum(localY(i))
        Print #f, "SLOT_" & i & "_W=" & FormatNum(w(i))
        Print #f, "SLOT_" & i & "_H=" & FormatNum(h(i))
        Print #f, "SLOT_" & i & "_AREA=" & FormatNum(area(i))
        Print #f, "SLOT_" & i & "_ANGLE=" & FormatNum(NormalizeAngleDeg(ang(i)))

    Next i

    Close #f

    WriteLog "SAVE TEMPLATE OK | VERSION=" & VERSION_TEXT & " | COUNT=" & cnt & " | BLOCKS=" & blockCount & " | FILE=" & path

End Sub


Private Sub LoadTemplate( _
    ByRef cnt As Long, _
    ByRef names() As String, _
    ByRef relX() As Double, _
    ByRef relY() As Double, _
    ByRef w() As Double, _
    ByRef h() As Double, _
    ByRef area() As Double, _
    ByRef ang() As Double, _
    ByRef blockIdx() As Long, _
    ByRef localX() As Double, _
    ByRef localY() As Double, _
    ByRef blockCount As Long)

    Dim path As String
    Dim f As Integer
    Dim ln As String
    Dim p As Long
    Dim k As String
    Dim v As String
    Dim idx As Long
    Dim fieldName As String
    Dim i As Long

    path = Environ$("USERPROFILE") & TEMPLATE_FILE

    cnt = 0
    blockCount = 0

    f = FreeFile

    Open path For Input As #f

    Do Until EOF(f)

        Line Input #f, ln
        ln = Trim$(ln)

        If Len(ln) = 0 Then GoTo NextLine

        p = InStr(1, ln, "=", vbTextCompare)

        If p <= 0 Then GoTo NextLine

        k = UCase$(Trim$(Left$(ln, p - 1)))
        v = Trim$(Mid$(ln, p + 1))

        If k = "COUNT" Then
            cnt = CLng(Val(v))
            If cnt > MAX_ITEMS Then cnt = MAX_ITEMS
            GoTo NextLine
        End If

        If k = "BLOCK_COUNT" Then
            blockCount = CLng(Val(v))
            If blockCount > MAX_BLOCKS Then blockCount = MAX_BLOCKS
            GoTo NextLine
        End If

        If Left$(k, 5) = "SLOT_" Then

            idx = ExtractSlotIndex(k)

            If idx >= 1 And idx <= MAX_ITEMS Then

                fieldName = ExtractSlotField(k)

                Select Case fieldName

                    Case "NAME"
                        names(idx) = v

                    Case "BLOCK"
                        blockIdx(idx) = CLng(Val(v))

                    Case "REL_X"
                        relX(idx) = ToDbl(v)

                    Case "REL_Y"
                        relY(idx) = ToDbl(v)

                    Case "LOCAL_X"
                        localX(idx) = ToDbl(v)

                    Case "LOCAL_Y"
                        localY(idx) = ToDbl(v)

                    Case "W"
                        w(idx) = ToDbl(v)

                    Case "H"
                        h(idx) = ToDbl(v)

                    Case "AREA"
                        area(idx) = ToDbl(v)

                    Case "ANGLE"
                        ang(idx) = ToDbl(v)

                End Select

            End If

        End If

NextLine:

    Loop

    Close #f

    If cnt <= 0 Then
        Err.Raise vbObjectError + 1201, , "COUNT pada template tidak valid."
    End If

    If blockCount <= 0 Then
        blockCount = 1
        For i = 1 To cnt
            blockIdx(i) = 1
        Next i
    End If

    SortSlotsByAreaDescV13 cnt, names, relX, relY, w, h, area, ang, blockIdx, localX, localY

    WriteLog "LOAD TEMPLATE OK | COUNT=" & cnt & " | BLOCKS=" & blockCount

End Sub


Private Function ExtractSlotIndex(ByVal keyName As String) As Long

    Dim s As String
    Dim p As Long

    keyName = UCase$(Trim$(keyName))

    s = Mid$(keyName, 6)
    p = InStr(1, s, "_", vbTextCompare)

    If p <= 1 Then Exit Function

    ExtractSlotIndex = CLng(Val(Left$(s, p - 1)))

End Function


Private Function ExtractSlotField(ByVal keyName As String) As String

    Dim s As String
    Dim p As Long

    keyName = UCase$(Trim$(keyName))

    s = Mid$(keyName, 6)
    p = InStr(1, s, "_", vbTextCompare)

    If p <= 0 Then Exit Function

    ExtractSlotField = Mid$(s, p + 1)

End Function


'=========================================================
' TARGET RELATIVE POSITION
'=========================================================
Private Sub FillTargetRelativePositions( _
    ByVal itemCount As Long, _
    ByRef shp() As Shape, _
    ByVal baseLeft As Double, _
    ByVal baseTop As Double, _
    ByRef relX() As Double, _
    ByRef relY() As Double)

    Dim i As Long

    For i = 1 To itemCount

        If Not shp(i) Is Nothing Then
            relX(i) = ShapeCenterX(shp(i)) - baseLeft
            relY(i) = baseTop - ShapeCenterY(shp(i))
        Else
            relX(i) = 0
            relY(i) = 0
        End If

    Next i

End Sub


Private Function PositionMatchError( _
    ByVal sx As Double, _
    ByVal sy As Double, _
    ByVal tx As Double, _
    ByVal ty As Double, _
    ByVal baseW As Double, _
    ByVal baseH As Double) As Double

    Dim ex As Double
    Dim ey As Double

    If baseW <= 0 Then baseW = 1#
    If baseH <= 0 Then baseH = 1#

    ex = Abs(tx - sx) / baseW
    ey = Abs(ty - sy) / baseH

    PositionMatchError = ex + ey

End Function


Private Function LocalPositionMatchError( _
    ByVal sx As Double, _
    ByVal sy As Double, _
    ByVal tx As Double, _
    ByVal ty As Double, _
    ByVal baseW As Double, _
    ByVal baseH As Double) As Double

    ' Local position dipakai sebagai penguat tambahan untuk komponen kembar di blok yang sama.
    LocalPositionMatchError = PositionMatchError(sx, sy, tx, ty, baseW, baseH)

End Function


'=========================================================
' MATCHING
'=========================================================
Private Function FindBestTargetForSlot( _
    ByVal slotIdx As Long, _
    ByVal itemCount As Long, _
    ByRef used() As Boolean, _
    ByRef slotRelX() As Double, _
    ByRef slotRelY() As Double, _
    ByRef slotW() As Double, _
    ByRef slotH() As Double, _
    ByRef slotArea() As Double, _
    ByRef slotAngle() As Double, _
    ByRef slotBlock() As Long, _
    ByRef slotLocalX() As Double, _
    ByRef slotLocalY() As Double, _
    ByRef targetRelX() As Double, _
    ByRef targetRelY() As Double, _
    ByRef targetW() As Double, _
    ByRef targetH() As Double, _
    ByRef targetArea() As Double, _
    ByRef targetAngle() As Double, _
    ByRef targetBlock() As Long, _
    ByRef targetLocalX() As Double, _
    ByRef targetLocalY() As Double, _
    ByVal baseW As Double, _
    ByVal baseH As Double) As Long

    Dim i As Long
    Dim best As Long
    Dim bestErr As Double
    Dim errVal As Double

    best = 0
    bestErr = 999999#

    For i = 1 To itemCount

        If Not used(i) Then

            If ENABLE_SIZE_BLOCK_LOCK Then
                If slotBlock(slotIdx) > 0 And targetBlock(i) > 0 Then
                    If slotBlock(slotIdx) <> targetBlock(i) Then GoTo NextCandidate
                End If
            End If

            errVal = ShapeMatchError( _
                        slotW(slotIdx), _
                        slotH(slotIdx), _
                        slotArea(slotIdx), _
                        targetW(i), _
                        targetH(i), _
                        targetArea(i))

            errVal = errVal + ( _
                        PositionMatchError( _
                            slotRelX(slotIdx), _
                            slotRelY(slotIdx), _
                            targetRelX(i), _
                            targetRelY(i), _
                            baseW, _
                            baseH) _
                        * POSITION_WEIGHT)

            errVal = errVal + ( _
                        LocalPositionMatchError( _
                            slotLocalX(slotIdx), _
                            slotLocalY(slotIdx), _
                            targetLocalX(i), _
                            targetLocalY(i), _
                            baseW, _
                            baseH) _
                        * 0.25)

            errVal = errVal + ( _
                        RotationMatchError( _
                            slotAngle(slotIdx), _
                            targetAngle(i)) _
                        * ROTATION_WEIGHT)

            If errVal < bestErr Then
                bestErr = errVal
                best = i
            End If

        End If

NextCandidate:

    Next i

    FindBestTargetForSlot = best

End Function


Private Function ShapeMatchError( _
    ByVal sw As Double, _
    ByVal sh As Double, _
    ByVal sa As Double, _
    ByVal tw As Double, _
    ByVal th As Double, _
    ByVal ta As Double) As Double

    Dim e1 As Double
    Dim e2 As Double
    Dim ea As Double

    If sw <= 0 Or sh <= 0 Or sa <= 0 Or tw <= 0 Or th <= 0 Or ta <= 0 Then
        ShapeMatchError = 999999#
        Exit Function
    End If

    e1 = Abs(tw - sw) / sw + Abs(th - sh) / sh
    e2 = Abs(th - sw) / sw + Abs(tw - sh) / sh
    ea = Abs(ta - sa) / sa

    If e2 < e1 Then e1 = e2

    ShapeMatchError = e1 + (ea * 0.35)

End Function


'=========================================================
' SORTING
'=========================================================
Private Sub SortItemsByAreaDescV13( _
    ByVal cnt As Long, _
    ByRef names() As String, _
    ByRef cx() As Double, _
    ByRef cy() As Double, _
    ByRef w() As Double, _
    ByRef h() As Double, _
    ByRef area() As Double, _
    ByRef ang() As Double, _
    ByRef blockIdx() As Long, _
    ByRef localX() As Double, _
    ByRef localY() As Double)

    Dim i As Long
    Dim j As Long

    For i = 1 To cnt - 1
        For j = i + 1 To cnt

            If area(j) > area(i) Then
                SwapString names(i), names(j)
                SwapDouble cx(i), cx(j)
                SwapDouble cy(i), cy(j)
                SwapDouble w(i), w(j)
                SwapDouble h(i), h(j)
                SwapDouble area(i), area(j)
                SwapDouble ang(i), ang(j)
                SwapLong blockIdx(i), blockIdx(j)
                SwapDouble localX(i), localX(j)
                SwapDouble localY(i), localY(j)
            End If

        Next j
    Next i

End Sub


Private Sub SortSlotsByAreaDescV13( _
    ByVal cnt As Long, _
    ByRef names() As String, _
    ByRef relX() As Double, _
    ByRef relY() As Double, _
    ByRef w() As Double, _
    ByRef h() As Double, _
    ByRef area() As Double, _
    ByRef ang() As Double, _
    ByRef blockIdx() As Long, _
    ByRef localX() As Double, _
    ByRef localY() As Double)

    Dim i As Long
    Dim j As Long

    For i = 1 To cnt - 1
        For j = i + 1 To cnt

            If area(j) > area(i) Then
                SwapString names(i), names(j)
                SwapDouble relX(i), relX(j)
                SwapDouble relY(i), relY(j)
                SwapDouble w(i), w(j)
                SwapDouble h(i), h(j)
                SwapDouble area(i), area(j)
                SwapDouble ang(i), ang(j)
                SwapLong blockIdx(i), blockIdx(j)
                SwapDouble localX(i), localX(j)
                SwapDouble localY(i), localY(j)
            End If

        Next j
    Next i

End Sub


Private Sub SwapDouble(ByRef a As Double, ByRef b As Double)

    Dim t As Double

    t = a
    a = b
    b = t

End Sub


Private Sub SwapString(ByRef a As String, ByRef b As String)

    Dim t As String

    t = a
    a = b
    b = t

End Sub


Private Sub SwapLong(ByRef a As Long, ByRef b As Long)

    Dim t As Long

    t = a
    a = b
    b = t

End Sub


'=========================================================
' ROTATION HELPERS
'=========================================================
Private Function ShapeAngleSafe(ByVal s As Shape) As Double

    On Error GoTo SafeExit

    ShapeAngleSafe = CDbl(CallByName(s, "RotationAngle", VbGet))
    ShapeAngleSafe = NormalizeAngleDeg(ShapeAngleSafe)

    Exit Function

SafeExit:

    ShapeAngleSafe = 0#

End Function


Private Function NormalizeAngleDeg(ByVal a As Double) As Double

    Do While a > 180#
        a = a - 360#
    Loop

    Do While a <= -180#
        a = a + 360#
    Loop

    NormalizeAngleDeg = a

End Function


Private Function RotationMatchError( _
    ByVal slotAngle As Double, _
    ByVal targetAngle As Double) As Double

    Dim d As Double

    d = Abs(NormalizeAngleDeg(targetAngle - slotAngle))

    If d > 90# Then
        d = 180# - d
    End If

    RotationMatchError = d / 90#

End Function


Private Sub ApplyRotationIfNeeded( _
    ByVal s As Shape, _
    ByVal slotAngle As Double, _
    ByVal slotW As Double, _
    ByVal slotH As Double)

    If Not ENABLE_AUTO_ROTATE Then Exit Sub

    Dim curAngle As Double
    Dim delta As Double
    Dim tw As Double
    Dim th As Double

    On Error Resume Next

    curAngle = ShapeAngleSafe(s)
    delta = NormalizeAngleDeg(slotAngle - curAngle)

    If Abs(delta) > ROTATION_TOLERANCE_DEG Then
        s.Rotate delta
        If Err.Number = 0 Then Exit Sub
        Err.Clear
    End If

    tw = Abs(CDbl(s.SizeWidth))
    th = Abs(CDbl(s.SizeHeight))

    If slotW > 0 And slotH > 0 And tw > 0 And th > 0 Then

        If (slotW - slotH) * (tw - th) < 0 Then
            s.Rotate 90#
            Err.Clear
        End If

    End If

    On Error GoTo 0

End Sub


'=========================================================
' UTIL
'=========================================================
Private Function ToDbl(ByVal v As String) As Double

    Dim s As String

    s = Trim$(CStr(v))
    s = Replace(s, ",", ".")

    ToDbl = CDbl(Val(s))

End Function


Private Function FormatNum(ByVal v As Double) As String

    FormatNum = Replace(Format$(v, "0.000"), ",", ".")

End Function


Private Sub WriteLog(ByVal msg As String)

    Dim path As String
    Dim f As Integer

    On Error Resume Next

    path = Environ$("USERPROFILE") & LOG_FILE

    f = FreeFile

    Open path For Append As #f

    Print #f, Format$(Now, "yyyy-mm-dd hh:nn:ss") & " | " & msg

    Close #f

    On Error GoTo 0

End Sub
