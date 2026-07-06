Option Explicit

'==========================================================
' PROJECT HADES
' HADES AUTO MASS NESTING — RECORD PATTERN CATALOG V3.3 AUTO-ONLY + UID MARKER
'
' SHORTCUT UTAMA:
' HADES_RECORD_PATTERN_CATALOG
'
' ALIAS:
' HADES_RECORD_PATTERN_CATALOG_V2
' HADES_RECORD_PATTERN_CATALOG_V3
' HADES_RECORD_PATTERN_CATALOG_V33
'
' FUNGSI:
' - Jalankan saat master layout masih tergroup 1 group per size.
' - Macro membaca Documents\Order.txt.
' - Macro membaca @SIZEDB jika ada.
' - Jika @SIZEDB tidak ada, macro scan semua SizeDB*.txt di Documents.
' - Macro mendeteksi size group dari body panel yang cocok dengan SizeDB.
' - TANPA input size manual.
' - Jika auto-detect gagal, macro berhenti dan memberi detail penyebab.
' - Memberi UID marker sementara pada Shape.Name setiap panel source.
' - UID marker dibersihkan otomatis oleh Auto Mass Nesting V4 setelah sukses.
'
' OUTPUT:
' Documents\HADES_PATTERN_CATALOG_CURRENT.txt
'
' CATALOG FORMAT kompatibel dengan Auto Mass Nesting V3.1:
' PANEL|SIZE|PANEL_NO|WIDTH_CM|HEIGHT_CM|COREL_TYPE|NAME|BUCKET|...|PANEL_UID|ORIGINAL_NAME
'
' Field setelah BUCKET adalah tambahan coordinate signature,
' aman diabaikan oleh Auto Mass Nesting V3.1 lama.
'==========================================================

Private Const HAMNREC_CATALOG_FILE As String = "HADES_PATTERN_CATALOG_CURRENT.txt"
Private Const HAMNREC_ORDER_FILE As String = "Order.txt"
Private Const HAMNREC_REPORT_FILE As String = "HADES_RECORD_PATTERN_CATALOG_REPORT_LATEST.txt"

Private Const HAMNREC_MIN_PANEL_W As Double = 3#
Private Const HAMNREC_MIN_PANEL_H As Double = 3#
Private Const HAMNREC_VISUAL_Y_TOL As Double = 2#
Private Const HAMNREC_SIZE_MATCH_TOL As Double = 1#

Private Const HAMNREC_BUCKET_BODY As String = "BODY"
Private Const HAMNREC_BUCKET_SLEEVE As String = "SLEEVE"
Private Const HAMNREC_BUCKET_SMALL As String = "SMALL"
Private Const HAMNREC_MARK_PREFIX As String = "HADES_PC_UID|"


Public Sub HADES_RECORD_PATTERN_CATALOG()
    HADES_RECORD_PATTERN_CATALOG_V33
End Sub

Public Sub HADES_RECORD_PATTERN_CATALOG_V2()
    HADES_RECORD_PATTERN_CATALOG_V33
End Sub

Public Sub HADES_RECORD_PATTERN_CATALOG_V3()
    HADES_RECORD_PATTERN_CATALOG_V33
End Sub


Public Sub HADES_RECORD_PATTERN_CATALOG_V32()
    HADES_RECORD_PATTERN_CATALOG_V33
End Sub

Public Sub HADES_RECORD_PATTERN_CATALOG_V33()
    HADES_RECORD_PATTERN_CATALOG_V3_3_AUTOONLY
End Sub


Public Sub HADES_RECORD_PATTERN_CATALOG_V3_3_AUTOONLY()

    On Error GoTo ErrHandler

    Dim oldUnit As Long
    Dim sr As ShapeRange
    Dim n As Long
    Dim i As Long
    Dim j As Long

    Dim arrShapes() As Shape
    Dim tmp As Shape

    Dim docs As String
    Dim orderPath As String
    Dim outPath As String
    Dim reportPath As String

    Dim orderSizes As Object
    Dim orderRows As Long
    Dim metaSizeDB As String
    Dim orderInfo As String

    Dim candN As Long
    Dim candSize() As String
    Dim candW() As Double
    Dim candH() As Double
    Dim candDB() As String
    Dim dbSource As String

    Dim resolvedSizes() As String
    Dim resolvedDB() As String
    Dim detectLog As String
    Dim hardError As String

    Dim f As Integer
    Dim grp As Shape
    Dim ch As Shape
    Dim sz As String

    Dim pN As Long
    Dim pW() As Double
    Dim pH() As Double
    Dim pType() As String
    Dim pName() As String
    Dim pBucket() As String
    Dim pLeft() As Double
    Dim pTop() As Double
    Dim pCX() As Double
    Dim pCY() As Double
    Dim pShape() As Shape
    Dim pUID() As String
    Dim pOrigName() As String
    Dim markerCount As Long

    Dim totalPanels As Long
    Dim missingOrderSizes As String

    If ActiveDocument Is Nothing Then
        MsgBox "Tidak ada dokumen aktif.", vbExclamation, "Hades Record Pattern Catalog V3.3"
        Exit Sub
    End If

    If ActiveSelection.Shapes.Count = 0 Then
        MsgBox "Select group master per size dulu." & vbCrLf & vbCrLf & _
               "Contoh selection:" & vbCrLf & _
               "Group S + Group M + Group L", _
               vbExclamation, "Hades Record Pattern Catalog V3.3"
        Exit Sub
    End If

    oldUnit = ActiveDocument.Unit
    ActiveDocument.Unit = cdrCentimeter

    docs = HAMNREC_GetDocumentsPath()
    orderPath = docs & "\" & HAMNREC_ORDER_FILE
    outPath = docs & "\" & HAMNREC_CATALOG_FILE
    reportPath = docs & "\" & HAMNREC_REPORT_FILE

    Set orderSizes = CreateObject("Scripting.Dictionary")
    orderSizes.CompareMode = 1

    If Not HAMNREC_LoadOrderInfo(orderPath, metaSizeDB, orderSizes, orderRows, orderInfo) Then
        hardError = "Order.txt tidak ditemukan / gagal dibaca:" & vbCrLf & orderPath
        GoTo FailStop
    End If

    If orderRows <= 0 Or orderSizes.Count <= 0 Then
        hardError = "Order.txt tidak memiliki baris order valid." & vbCrLf & _
                    "Format wajib: SIZE|NAMA|NOMOR|NICKNAME"
        GoTo FailStop
    End If

    If Not HAMNREC_LoadSizeDBAuto(docs, metaSizeDB, orderSizes, candN, candSize, candW, candH, candDB, dbSource, detectLog) Then
        hardError = detectLog
        GoTo FailStop
    End If

    If candN <= 0 Then
        hardError = "Tidak ada kandidat ukuran yang berhasil dibaca dari SizeDB." & vbCrLf & detectLog
        GoTo FailStop
    End If

    Set sr = ActiveSelectionRange
    n = sr.Shapes.Count

    ReDim arrShapes(1 To n)
    ReDim resolvedSizes(1 To n)
    ReDim resolvedDB(1 To n)

    For i = 1 To n
        Set arrShapes(i) = sr.Shapes(i)
    Next i

    'Sort visual: atas ke bawah, lalu kiri ke kanan.
    For i = 1 To n - 1
        For j = i + 1 To n
            If HAMNREC_ShouldSwap(arrShapes(i), arrShapes(j)) Then
                Set tmp = arrShapes(i)
                Set arrShapes(i) = arrShapes(j)
                Set arrShapes(j) = tmp
            End If
        Next j
    Next i

    detectLog = detectLog & vbCrLf & "[GROUP AUTO DETECT]" & vbCrLf

    For i = 1 To n
        If Not HAMNREC_DetectOneGroupSizeStrong(arrShapes(i), candN, candSize, candW, candH, candDB, resolvedSizes(i), resolvedDB(i), detectLog, i) Then
            hardError = "Auto-detect size group gagal." & vbCrLf & vbCrLf & detectLog
            GoTo FailStop
        End If
    Next i

    If Not HAMNREC_CheckDuplicateResolvedSizes(n, resolvedSizes, hardError) Then
        hardError = hardError & vbCrLf & vbCrLf & detectLog
        GoTo FailStop
    End If

    missingOrderSizes = HAMNREC_FindMissingOrderSizes(orderSizes, n, resolvedSizes)
    If missingOrderSizes <> "" Then
        hardError = "Ada size di Order.txt yang tidak ditemukan pada selection group master:" & vbCrLf & _
                    missingOrderSizes & vbCrLf & vbCrLf & _
                    "Solusi: select semua group size yang dibutuhkan oleh Order.txt, lalu jalankan ulang." & vbCrLf & vbCrLf & detectLog
        GoTo FailStop
    End If

    f = FreeFile
    Open outPath For Output As #f

    Print #f, "@VERSION=HADES_PATTERN_CATALOG_V3_3_AUTOONLY_UID_MARKER"
    Print #f, "@CREATED=" & Format$(Now, "yyyy-mm-dd hh:nn:ss")
    Print #f, "@UNIT=CM"
    Print #f, "@SOURCE=CORELDRAW_SELECTION_GROUP_PER_SIZE"
    Print #f, "@SIZE_SOURCE=" & HAMNREC_CleanField(dbSource)
    Print #f, "@ORDER_FILE=" & HAMNREC_ORDER_FILE
    Print #f, "@ORDER_ROWS=" & CStr(orderRows)
    Print #f, "@ORDER_SIZES=" & HAMNREC_CleanField(HAMNREC_DictKeysCsv(orderSizes))
    Print #f, "@NOTE=Auto-only. No manual size input. Direct child inside each selected size group is treated as one panel."
    Print #f, "@PANEL_FORMAT=PANEL|SIZE|PANEL_NO|WIDTH_CM|HEIGHT_CM|COREL_TYPE|NAME|BUCKET|PANEL_LEFT|PANEL_TOP|PANEL_CENTER_X|PANEL_CENTER_Y|LOCAL_CENTER_X|LOCAL_CENTER_Y|SIZEDB_SOURCE|PANEL_UID|ORIGINAL_NAME"
    Print #f, "@BUCKET=BODY/SLEEVE/SMALL"
    Print #f, ""

    totalPanels = 0

    For i = 1 To n

        Set grp = arrShapes(i)
        sz = HAMNREC_NormalizeSize(CStr(resolvedSizes(i)))

        Print #f, "[SIZE=" & sz & "]"
        Print #f, "GROUP_INDEX|" & CStr(i)
        Print #f, "GROUP_BOUNDS|" & _
                  HAMNREC_DblToStr(grp.SizeWidth) & "|" & _
                  HAMNREC_DblToStr(grp.SizeHeight)
        Print #f, "GROUP_LEFT_TOP|" & _
                  HAMNREC_DblToStr(grp.LeftX) & "|" & _
                  HAMNREC_DblToStr(grp.TopY)
        Print #f, "GROUP_CENTER|" & _
                  HAMNREC_DblToStr(HAMNREC_CenterX(grp)) & "|" & _
                  HAMNREC_DblToStr(HAMNREC_CenterY(grp))
        Print #f, "GROUP_SIZEDB|" & HAMNREC_CleanField(resolvedDB(i))

        pN = 0

        If grp.Type <> cdrGroupShape Then
            HAMNREC_AddPanel pN, pW, pH, pType, pName, pBucket, pLeft, pTop, pCX, pCY, pShape, pUID, pOrigName, _
                              grp.SizeWidth, grp.SizeHeight, CStr(grp.Type), grp.Name, HAMNREC_BUCKET_BODY, _
                              grp.LeftX, grp.TopY, HAMNREC_CenterX(grp), HAMNREC_CenterY(grp), grp
        Else
            For j = 1 To grp.Shapes.Count
                Set ch = grp.Shapes(j)

                If ch.SizeWidth >= HAMNREC_MIN_PANEL_W And ch.SizeHeight >= HAMNREC_MIN_PANEL_H Then
                    HAMNREC_AddPanel pN, pW, pH, pType, pName, pBucket, pLeft, pTop, pCX, pCY, pShape, pUID, pOrigName, _
                                      ch.SizeWidth, ch.SizeHeight, CStr(ch.Type), ch.Name, "", _
                                      ch.LeftX, ch.TopY, HAMNREC_CenterX(ch), HAMNREC_CenterY(ch), ch
                End If
            Next j

            HAMNREC_ClassifyBuckets pN, pW, pH, pName, pBucket
        End If

        For j = 1 To pN
            totalPanels = totalPanels + 1

            pUID(j) = HAMNREC_BuildPanelUID(sz, pBucket(j), j)
            pOrigName(j) = pName(j)
            If HAMNREC_SetPanelMarker(pShape(j), pUID(j), pOrigName(j)) Then markerCount = markerCount + 1

            Print #f, "PANEL|" & sz & "|" & CStr(j) & "|" & _
                      HAMNREC_DblToStr(pW(j)) & "|" & _
                      HAMNREC_DblToStr(pH(j)) & "|" & _
                      pType(j) & "|" & _
                      HAMNREC_CleanField(pName(j)) & "|" & _
                      HAMNREC_NormalizeBucket(pBucket(j)) & "|" & _
                      HAMNREC_DblToStr(pLeft(j)) & "|" & _
                      HAMNREC_DblToStr(pTop(j)) & "|" & _
                      HAMNREC_DblToStr(pCX(j)) & "|" & _
                      HAMNREC_DblToStr(pCY(j)) & "|" & _
                      HAMNREC_DblToStr(pCX(j) - HAMNREC_CenterX(grp)) & "|" & _
                      HAMNREC_DblToStr(pCY(j) - HAMNREC_CenterY(grp)) & "|" & _
                      HAMNREC_CleanField(resolvedDB(i)) & "|" & _
                      HAMNREC_CleanField(pUID(j)) & "|" & _
                      HAMNREC_CleanField(pOrigName(j))
        Next j

        Print #f, "SIZE_PANEL_COUNT|" & sz & "|" & CStr(pN)
        Print #f, ""

    Next i

    Print #f, "@TOTAL_SIZE_GROUP=" & CStr(n)
    Print #f, "@TOTAL_PANEL=" & CStr(totalPanels)
    Print #f, "@TOTAL_MARKER=" & CStr(markerCount)

    Close #f

    HAMNREC_WriteReport reportPath, "PASS", orderInfo, dbSource, detectLog, outPath, ""

    ActiveDocument.Unit = oldUnit

    MsgBox "Pattern Catalog V3.3 AUTO-ONLY + UID MARKER berhasil dibuat." & vbCrLf & vbCrLf & _
           "File:" & vbCrLf & outPath & vbCrLf & vbCrLf & _
           "Jumlah group size : " & n & vbCrLf & _
           "Jumlah panel      : " & totalPanels & vbCrLf & _
           "UID marker        : " & markerCount & vbCrLf & _
           "Size source       : " & dbSource & vbCrLf & vbCrLf & _
           "Tidak ada input size manual." & vbCrLf & vbCrLf & _
           "Langkah berikutnya:" & vbCrLf & _
           "1. Ungroup sekali group size master." & vbCrLf & _
           "2. Jangan geser panel source." & vbCrLf & _
           "3. Select semua panel master loose." & vbCrLf & _
           "4. Jalankan HADES_AUTO_MASS_NESTING.", _
           vbInformation, "Hades Record Pattern Catalog V3.3"

    Exit Sub

FailStop:
    On Error Resume Next
    Close #f
    HAMNREC_WriteReport reportPath, "FAIL", orderInfo, dbSource, detectLog, "", hardError
    ActiveDocument.Unit = oldUnit
    MsgBox "RECORD PATTERN CATALOG GAGAL." & vbCrLf & vbCrLf & _
           hardError & vbCrLf & vbCrLf & _
           "Report:" & vbCrLf & reportPath, _
           vbCritical, "Hades Record Pattern Catalog V3.3"
    Exit Sub

ErrHandler:
    On Error Resume Next
    Close #f
    ActiveDocument.Unit = oldUnit
    MsgBox "ERROR HADES_RECORD_PATTERN_CATALOG_V3_3:" & vbCrLf & _
           Err.Description, vbCritical, "Hades Record Pattern Catalog V3.3"

End Sub


'==========================================================
' ORDER.TXT READER
'==========================================================

Private Function HAMNREC_LoadOrderInfo( _
    ByVal orderPath As String, _
    ByRef metaSizeDB As String, _
    ByRef orderSizes As Object, _
    ByRef orderRows As Long, _
    ByRef orderInfo As String _
) As Boolean

    On Error GoTo Fail

    Dim f As Integer
    Dim lineText As String
    Dim parts() As String
    Dim sz As String
    Dim p As Long

    metaSizeDB = ""
    orderRows = 0
    orderInfo = ""

    If Dir$(orderPath) = "" Then
        HAMNREC_LoadOrderInfo = False
        Exit Function
    End If

    f = FreeFile
    Open orderPath For Input As #f

    Do While Not EOF(f)
        Line Input #f, lineText
        lineText = Trim$(lineText)

        If lineText = "" Then GoTo NextLine

        If UCase$(Left$(lineText, 8)) = "@SIZEDB=" Then
            p = InStr(1, lineText, "=", vbBinaryCompare)
            If p > 0 Then metaSizeDB = Trim$(Mid$(lineText, p + 1))
            GoTo NextLine
        End If

        If Left$(lineText, 1) = "@" Then GoTo NextLine
        If InStr(1, lineText, "|", vbBinaryCompare) = 0 Then GoTo NextLine

        parts = Split(lineText, "|")
        If UBound(parts) < 0 Then GoTo NextLine

        sz = HAMNREC_NormalizeSize(parts(0))

        If sz <> "" Then
            orderRows = orderRows + 1
            If orderSizes.Exists(sz) Then
                orderSizes(sz) = CLng(orderSizes(sz)) + 1
            Else
                orderSizes.Add sz, 1
            End If
        End If

NextLine:
    Loop

    Close #f

    orderInfo = "ORDER_PATH=" & orderPath & vbCrLf & _
                "ORDER_ROWS=" & CStr(orderRows) & vbCrLf & _
                "ORDER_SIZES=" & HAMNREC_DictKeysCsv(orderSizes) & vbCrLf & _
                "META_SIZEDB=" & metaSizeDB & vbCrLf

    HAMNREC_LoadOrderInfo = True
    Exit Function

Fail:
    On Error Resume Next
    Close #f
    HAMNREC_LoadOrderInfo = False

End Function


'==========================================================
' SIZEDB LOADER
'==========================================================

Private Function HAMNREC_LoadSizeDBAuto( _
    ByVal docs As String, _
    ByVal metaSizeDB As String, _
    ByVal orderSizes As Object, _
    ByRef candN As Long, _
    ByRef candSize() As String, _
    ByRef candW() As Double, _
    ByRef candH() As Double, _
    ByRef candDB() As String, _
    ByRef dbSource As String, _
    ByRef detectLog As String _
) As Boolean

    Dim dbPath As String
    Dim dbFile As String
    Dim loadedFiles As Long

    candN = 0
    detectLog = "[SIZEDB AUTO LOAD]" & vbCrLf
    dbSource = ""

    If Trim$(metaSizeDB) <> "" Then
        dbPath = docs & "\" & metaSizeDB

        If Dir$(dbPath) = "" Then
            detectLog = detectLog & "@SIZEDB ditemukan tetapi file tidak ada: " & dbPath & vbCrLf
            HAMNREC_LoadSizeDBAuto = False
            Exit Function
        End If

        If Not HAMNREC_LoadOneSizeDB(dbPath, metaSizeDB, orderSizes, candN, candSize, candW, candH, candDB, detectLog) Then
            HAMNREC_LoadSizeDBAuto = False
            Exit Function
        End If

        dbSource = "AUTO_META_SIZEDB:" & metaSizeDB
        HAMNREC_LoadSizeDBAuto = (candN > 0)
        Exit Function
    End If

    dbFile = Dir$(docs & "\SizeDB*.txt")

    Do While dbFile <> ""
        dbPath = docs & "\" & dbFile

        If HAMNREC_LoadOneSizeDB(dbPath, dbFile, orderSizes, candN, candSize, candW, candH, candDB, detectLog) Then
            loadedFiles = loadedFiles + 1
        End If

        dbFile = Dir$()
    Loop

    If loadedFiles <= 0 Then
        detectLog = detectLog & "Tidak ada file SizeDB*.txt di Documents." & vbCrLf
        HAMNREC_LoadSizeDBAuto = False
        Exit Function
    End If

    dbSource = "AUTO_SCAN_DOCUMENTS_SizeDB*.txt|FILES=" & CStr(loadedFiles)
    HAMNREC_LoadSizeDBAuto = (candN > 0)

End Function


Private Function HAMNREC_LoadOneSizeDB( _
    ByVal sizeDbPath As String, _
    ByVal dbName As String, _
    ByVal orderSizes As Object, _
    ByRef candN As Long, _
    ByRef candSize() As String, _
    ByRef candW() As Double, _
    ByRef candH() As Double, _
    ByRef candDB() As String, _
    ByRef detectLog As String _
) As Boolean

    On Error GoTo Fail

    Dim f As Integer
    Dim lineText As String
    Dim parts() As String
    Dim nums() As Double
    Dim numN As Long
    Dim sz As String
    Dim i As Long
    Dim v As Double
    Dim beforeN As Long

    beforeN = candN

    f = FreeFile
    Open sizeDbPath For Input As #f

    Do While Not EOF(f)
        Line Input #f, lineText
        lineText = Trim$(lineText)

        If lineText = "" Then GoTo NextLine
        If Left$(lineText, 1) = "@" Then GoTo NextLine
        If Left$(lineText, 1) = "#" Then GoTo NextLine
        If InStr(1, lineText, "|", vbBinaryCompare) = 0 Then GoTo NextLine

        parts = Split(lineText, "|")
        If UBound(parts) < 2 Then GoTo NextLine

        sz = HAMNREC_NormalizeSize(parts(0))
        If sz = "" Then GoTo NextLine

        'Kunci dengan Order.txt agar tidak salah ambil size yang tidak dipesan.
        If orderSizes.Count > 0 Then
            If Not orderSizes.Exists(sz) Then GoTo NextLine
        End If

        numN = 0
        Erase nums

        For i = 1 To UBound(parts)
            If HAMNREC_TryDbl(parts(i), v) Then
                numN = numN + 1
                If numN = 1 Then
                    ReDim nums(1 To numN)
                Else
                    ReDim Preserve nums(1 To numN)
                End If
                nums(numN) = v
            End If
        Next i

        If numN = 3 Then
            'Jersey: SIZE|LEBAR|TINGGI_DEPAN|TINGGI_BELAKANG
            HAMNREC_AddSizeCandidate candN, candSize, candW, candH, candDB, sz, nums(1), nums(2), dbName
            HAMNREC_AddSizeCandidate candN, candSize, candW, candH, candDB, sz, nums(1), nums(3), dbName
        ElseIf numN >= 4 Then
            'Jaket umum: SIZE|L_BELAKANG|L_DEPAN|T_DEPAN|T_BELAKANG
            HAMNREC_AddSizeCandidate candN, candSize, candW, candH, candDB, sz, nums(1), nums(4), dbName
            HAMNREC_AddSizeCandidate candN, candSize, candW, candH, candDB, sz, nums(2), nums(3), dbName

            'Fallback untuk DB custom.
            HAMNREC_AddSizeCandidate candN, candSize, candW, candH, candDB, sz, nums(1), nums(3), dbName
            HAMNREC_AddSizeCandidate candN, candSize, candW, candH, candDB, sz, nums(1), nums(2), dbName
        ElseIf numN = 2 Then
            'Celana / custom sederhana.
            HAMNREC_AddSizeCandidate candN, candSize, candW, candH, candDB, sz, nums(1), nums(2), dbName
        End If

NextLine:
    Loop

    Close #f

    detectLog = detectLog & dbName & " => candidates added: " & CStr(candN - beforeN) & vbCrLf
    HAMNREC_LoadOneSizeDB = True
    Exit Function

Fail:
    On Error Resume Next
    Close #f
    detectLog = detectLog & "Gagal membaca: " & sizeDbPath & " | " & Err.Description & vbCrLf
    HAMNREC_LoadOneSizeDB = False

End Function


Private Sub HAMNREC_AddSizeCandidate( _
    ByRef candN As Long, _
    ByRef candSize() As String, _
    ByRef candW() As Double, _
    ByRef candH() As Double, _
    ByRef candDB() As String, _
    ByVal sz As String, _
    ByVal w As Double, _
    ByVal h As Double, _
    ByVal dbName As String _
)

    If w <= 0# Or h <= 0# Then Exit Sub

    candN = candN + 1

    If candN = 1 Then
        ReDim candSize(1 To candN)
        ReDim candW(1 To candN)
        ReDim candH(1 To candN)
        ReDim candDB(1 To candN)
    Else
        ReDim Preserve candSize(1 To candN)
        ReDim Preserve candW(1 To candN)
        ReDim Preserve candH(1 To candN)
        ReDim Preserve candDB(1 To candN)
    End If

    candSize(candN) = sz
    candW(candN) = w
    candH(candN) = h
    candDB(candN) = dbName

End Sub


'==========================================================
' GROUP SIZE DETECTION
'==========================================================

Private Function HAMNREC_DetectOneGroupSizeStrong( _
    ByVal grp As Shape, _
    ByVal candN As Long, _
    ByRef candSize() As String, _
    ByRef candW() As Double, _
    ByRef candH() As Double, _
    ByRef candDB() As String, _
    ByRef outSize As String, _
    ByRef outDB As String, _
    ByRef detectLog As String, _
    ByVal groupIndex As Long _
) As Boolean

    Dim bestSize As String
    Dim bestDB As String
    Dim bestArea As Double
    Dim bestScore As Double
    Dim bestPanelW As Double
    Dim bestPanelH As Double

    bestSize = ""
    bestDB = ""
    bestArea = -1#
    bestScore = 999999#
    bestPanelW = 0#
    bestPanelH = 0#

    If grp.Type = cdrGroupShape Then
        Dim j As Long
        Dim ch As Shape

        For j = 1 To grp.Shapes.Count
            Set ch = grp.Shapes(j)
            HAMNREC_TestShapeAgainstCandidates ch, candN, candSize, candW, candH, candDB, _
                                                bestSize, bestDB, bestArea, bestScore, bestPanelW, bestPanelH
        Next j
    Else
        HAMNREC_TestShapeAgainstCandidates grp, candN, candSize, candW, candH, candDB, _
                                            bestSize, bestDB, bestArea, bestScore, bestPanelW, bestPanelH
    End If

    outSize = bestSize
    outDB = bestDB

    If bestSize <> "" Then
        detectLog = detectLog & "Group " & CStr(groupIndex) & " => " & bestSize & _
                    " | DB=" & bestDB & _
                    " | panel=" & HAMNREC_DblToStr(bestPanelW) & "x" & HAMNREC_DblToStr(bestPanelH) & _
                    " | score=" & HAMNREC_DblToStr(bestScore) & vbCrLf
        HAMNREC_DetectOneGroupSizeStrong = True
    Else
        detectLog = detectLog & "Group " & CStr(groupIndex) & " => FAIL. Tidak ada body panel yang cocok SizeDB." & vbCrLf
        HAMNREC_DetectOneGroupSizeStrong = False
    End If

End Function


Private Sub HAMNREC_TestShapeAgainstCandidates( _
    ByVal shp As Shape, _
    ByVal candN As Long, _
    ByRef candSize() As String, _
    ByRef candW() As Double, _
    ByRef candH() As Double, _
    ByRef candDB() As String, _
    ByRef bestSize As String, _
    ByRef bestDB As String, _
    ByRef bestArea As Double, _
    ByRef bestScore As Double, _
    ByRef bestPanelW As Double, _
    ByRef bestPanelH As Double _
)

    Dim i As Long
    Dim w As Double
    Dim h As Double
    Dim a As Double
    Dim score As Double

    w = shp.SizeWidth
    h = shp.SizeHeight

    If w < HAMNREC_MIN_PANEL_W Or h < HAMNREC_MIN_PANEL_H Then Exit Sub

    a = w * h

    For i = 1 To candN
        If HAMNREC_DimMatch(w, h, candW(i), candH(i), score) Then

            'Pilih panel terbesar yang match SizeDB. Ini biasanya badan depan/belakang.
            If a > bestArea + 1# Then
                bestArea = a
                bestScore = score
                bestSize = candSize(i)
                bestDB = candDB(i)
                bestPanelW = w
                bestPanelH = h
            ElseIf Abs(a - bestArea) <= 1# Then
                If score < bestScore Then
                    bestArea = a
                    bestScore = score
                    bestSize = candSize(i)
                    bestDB = candDB(i)
                    bestPanelW = w
                    bestPanelH = h
                End If
            End If

        End If
    Next i

End Sub


Private Function HAMNREC_DimMatch( _
    ByVal w As Double, _
    ByVal h As Double, _
    ByVal dbW As Double, _
    ByVal dbH As Double, _
    ByRef score As Double _
) As Boolean

    Dim s1 As Double
    Dim s2 As Double

    s1 = Abs(w - dbW) + Abs(h - dbH)
    s2 = Abs(w - dbH) + Abs(h - dbW)

    If Abs(w - dbW) <= HAMNREC_SIZE_MATCH_TOL And Abs(h - dbH) <= HAMNREC_SIZE_MATCH_TOL Then
        score = s1
        HAMNREC_DimMatch = True
        Exit Function
    End If

    If Abs(w - dbH) <= HAMNREC_SIZE_MATCH_TOL And Abs(h - dbW) <= HAMNREC_SIZE_MATCH_TOL Then
        score = s2 + 0.2
        HAMNREC_DimMatch = True
        Exit Function
    End If

    HAMNREC_DimMatch = False

End Function


Private Function HAMNREC_CheckDuplicateResolvedSizes( _
    ByVal n As Long, _
    ByRef resolvedSizes() As String, _
    ByRef msg As String _
) As Boolean

    Dim d As Object
    Dim i As Long
    Dim sz As String

    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1

    msg = ""

    For i = 1 To n
        sz = HAMNREC_NormalizeSize(resolvedSizes(i))
        If sz = "" Then
            msg = msg & "Group " & CStr(i) & " tidak memiliki size valid." & vbCrLf
        ElseIf d.Exists(sz) Then
            msg = msg & "Size terdeteksi dobel pada selection: " & sz & vbCrLf
        Else
            d.Add sz, True
        End If
    Next i

    HAMNREC_CheckDuplicateResolvedSizes = (msg = "")

End Function


Private Function HAMNREC_FindMissingOrderSizes( _
    ByVal orderSizes As Object, _
    ByVal n As Long, _
    ByRef resolvedSizes() As String _
) As String

    Dim resolved As Object
    Dim i As Long
    Dim key As Variant
    Dim msg As String

    Set resolved = CreateObject("Scripting.Dictionary")
    resolved.CompareMode = 1

    For i = 1 To n
        If HAMNREC_NormalizeSize(resolvedSizes(i)) <> "" Then
            resolved(HAMNREC_NormalizeSize(resolvedSizes(i))) = True
        End If
    Next i

    msg = ""

    For Each key In orderSizes.Keys
        If Not resolved.Exists(CStr(key)) Then
            msg = msg & "- " & CStr(key) & " | qty=" & CStr(orderSizes(key)) & vbCrLf
        End If
    Next key

    HAMNREC_FindMissingOrderSizes = msg

End Function


'==========================================================
' PANEL COLLECTOR + BUCKET CLASSIFIER
'==========================================================

Private Sub HAMNREC_AddPanel( _
    ByRef pN As Long, _
    ByRef pW() As Double, _
    ByRef pH() As Double, _
    ByRef pType() As String, _
    ByRef pName() As String, _
    ByRef pBucket() As String, _
    ByRef pLeft() As Double, _
    ByRef pTop() As Double, _
    ByRef pCX() As Double, _
    ByRef pCY() As Double, _
    ByRef pShape() As Shape, _
    ByRef pUID() As String, _
    ByRef pOrigName() As String, _
    ByVal w As Double, _
    ByVal h As Double, _
    ByVal typ As String, _
    ByVal nm As String, _
    ByVal bucket As String, _
    ByVal leftX As Double, _
    ByVal topY As Double, _
    ByVal centerX As Double, _
    ByVal centerY As Double, _
    ByVal panelShape As Shape _
)

    pN = pN + 1

    If pN = 1 Then
        ReDim pW(1 To pN)
        ReDim pH(1 To pN)
        ReDim pType(1 To pN)
        ReDim pName(1 To pN)
        ReDim pBucket(1 To pN)
        ReDim pLeft(1 To pN)
        ReDim pTop(1 To pN)
        ReDim pCX(1 To pN)
        ReDim pCY(1 To pN)
        ReDim pShape(1 To pN)
        ReDim pUID(1 To pN)
        ReDim pOrigName(1 To pN)
    Else
        ReDim Preserve pW(1 To pN)
        ReDim Preserve pH(1 To pN)
        ReDim Preserve pType(1 To pN)
        ReDim Preserve pName(1 To pN)
        ReDim Preserve pBucket(1 To pN)
        ReDim Preserve pLeft(1 To pN)
        ReDim Preserve pTop(1 To pN)
        ReDim Preserve pCX(1 To pN)
        ReDim Preserve pCY(1 To pN)
        ReDim Preserve pShape(1 To pN)
        ReDim Preserve pUID(1 To pN)
        ReDim Preserve pOrigName(1 To pN)
    End If

    pW(pN) = w
    pH(pN) = h
    pType(pN) = typ
    pName(pN) = HAMNREC_OriginalNameFromMaybeMarker(nm)
    pBucket(pN) = bucket
    pLeft(pN) = leftX
    pTop(pN) = topY
    pCX(pN) = centerX
    pCY(pN) = centerY
    Set pShape(pN) = panelShape
    pUID(pN) = ""
    pOrigName(pN) = pName(pN)

End Sub


Private Sub HAMNREC_ClassifyBuckets( _
    ByVal pN As Long, _
    ByRef pW() As Double, _
    ByRef pH() As Double, _
    ByRef pName() As String, _
    ByRef pBucket() As String _
)

    Dim i As Long
    Dim pass As Long
    Dim maxArea As Double
    Dim a As Double
    Dim bestIdx As Long
    Dim bestArea As Double
    Dim minDim As Double

    For i = 1 To pN
        pBucket(i) = HAMNREC_BucketFromName(pName(i))
        a = pW(i) * pH(i)
        If a > maxArea Then maxArea = a
    Next i

    If maxArea <= 0# Then Exit Sub

    'Dua panel terbesar per size dianggap BODY bila nama shape tidak memberi petunjuk.
    For pass = 1 To 2
        bestIdx = 0
        bestArea = -1#

        For i = 1 To pN
            If pBucket(i) = "" Then
                a = pW(i) * pH(i)
                If a > bestArea Then
                    bestArea = a
                    bestIdx = i
                End If
            End If
        Next i

        If bestIdx > 0 Then pBucket(bestIdx) = HAMNREC_BUCKET_BODY
    Next pass

    For i = 1 To pN
        If pBucket(i) = "" Then
            a = pW(i) * pH(i)
            minDim = HAMNREC_MinD(pW(i), pH(i))

            If minDim < 8# Or a <= (maxArea * 0.12) Then
                pBucket(i) = HAMNREC_BUCKET_SMALL
            ElseIf a <= (maxArea * 0.55) Then
                pBucket(i) = HAMNREC_BUCKET_SLEEVE
            Else
                pBucket(i) = HAMNREC_BUCKET_BODY
            End If
        End If
    Next i

End Sub


Private Function HAMNREC_BucketFromName(ByVal nm As String) As String

    Dim u As String
    u = UCase$(Trim$(nm))

    If u = "" Then
        HAMNREC_BucketFromName = ""
        Exit Function
    End If

    If InStr(1, u, "KERAH", vbTextCompare) > 0 Or _
       InStr(1, u, "COLLAR", vbTextCompare) > 0 Or _
       InStr(1, u, "NECK", vbTextCompare) > 0 Or _
       InStr(1, u, "TULANG", vbTextCompare) > 0 Or _
       InStr(1, u, "RIB", vbTextCompare) > 0 Or _
       InStr(1, u, "MANSET", vbTextCompare) > 0 Then
        HAMNREC_BucketFromName = HAMNREC_BUCKET_SMALL
        Exit Function
    End If

    If InStr(1, u, "LENGAN", vbTextCompare) > 0 Or _
       InStr(1, u, "SLEEVE", vbTextCompare) > 0 Or _
       InStr(1, u, "ARM", vbTextCompare) > 0 Then
        HAMNREC_BucketFromName = HAMNREC_BUCKET_SLEEVE
        Exit Function
    End If

    If InStr(1, u, "DEPAN", vbTextCompare) > 0 Or _
       InStr(1, u, "BELAKANG", vbTextCompare) > 0 Or _
       InStr(1, u, "FRONT", vbTextCompare) > 0 Or _
       InStr(1, u, "BACK", vbTextCompare) > 0 Or _
       InStr(1, u, "BODY", vbTextCompare) > 0 Or _
       InStr(1, u, "BADAN", vbTextCompare) > 0 Then
        HAMNREC_BucketFromName = HAMNREC_BUCKET_BODY
        Exit Function
    End If

    HAMNREC_BucketFromName = ""

End Function


Private Function HAMNREC_NormalizeBucket(ByVal s As String) As String

    s = UCase$(Trim$(s))

    Select Case s
        Case HAMNREC_BUCKET_BODY, HAMNREC_BUCKET_SLEEVE, HAMNREC_BUCKET_SMALL
            HAMNREC_NormalizeBucket = s
        Case Else
            HAMNREC_NormalizeBucket = HAMNREC_BUCKET_SMALL
    End Select

End Function



'==========================================================
' UID MARKER HELPERS
'==========================================================

Private Function HAMNREC_BuildPanelUID(ByVal sz As String, ByVal bucket As String, ByVal panelNo As Long) As String
    HAMNREC_BuildPanelUID = HAMNREC_NormalizeSize(sz) & "_" & HAMNREC_NormalizeBucket(bucket) & "_P" & Format$(panelNo, "000")
End Function


Private Function HAMNREC_SetPanelMarker(ByVal shp As Shape, ByVal uid As String, ByVal origName As String) As Boolean

    On Error GoTo Fail

    If shp Is Nothing Then GoTo Fail
    If Trim$(uid) = "" Then GoTo Fail

    shp.Name = HAMNREC_MARK_PREFIX & HAMNREC_CleanField(uid) & "|ORIG|" & HAMNREC_CleanField(origName)
    HAMNREC_SetPanelMarker = True
    Exit Function

Fail:
    HAMNREC_SetPanelMarker = False

End Function


Private Function HAMNREC_OriginalNameFromMaybeMarker(ByVal nm As String) As String

    Dim p() As String

    If UCase$(Left$(CStr(nm), Len(HAMNREC_MARK_PREFIX))) <> UCase$(HAMNREC_MARK_PREFIX) Then
        HAMNREC_OriginalNameFromMaybeMarker = CStr(nm)
        Exit Function
    End If

    p = Split(CStr(nm), "|")
    If UBound(p) >= 3 Then
        HAMNREC_OriginalNameFromMaybeMarker = CStr(p(3))
    Else
        HAMNREC_OriginalNameFromMaybeMarker = ""
    End If

End Function


'==========================================================
' HELPERS
'==========================================================

Private Function HAMNREC_ShouldSwap(ByVal a As Shape, ByVal b As Shape) As Boolean

    'Return True jika b harus berada sebelum a.
    'Sort visual: atas ke bawah, lalu kiri ke kanan.

    If Abs(a.TopY - b.TopY) > HAMNREC_VISUAL_Y_TOL Then
        HAMNREC_ShouldSwap = (a.TopY < b.TopY)
    Else
        HAMNREC_ShouldSwap = (a.LeftX > b.LeftX)
    End If

End Function


Private Function HAMNREC_GetDocumentsPath() As String
    HAMNREC_GetDocumentsPath = CreateObject("WScript.Shell").SpecialFolders("MyDocuments")
End Function


Private Function HAMNREC_NormalizeSize(ByVal s As String) As String

    s = UCase$(Trim$(s))

    Select Case s
        Case "XXL"
            s = "2XL"
        Case "XXXL"
            s = "3XL"
        Case "XXXXL"
            s = "4XL"
        Case "XXXXXL"
            s = "5XL"
        Case "XXXXXXL"
            s = "6XL"
    End Select

    Select Case s
        Case "XXS", "XS", "S", "M", "L", "XL", "2XL", "3XL", "4XL", "5XL", "6XL"
            HAMNREC_NormalizeSize = s
        Case Else
            HAMNREC_NormalizeSize = ""
    End Select

End Function


Private Function HAMNREC_TryDbl(ByVal s As String, ByRef outVal As Double) As Boolean

    On Error GoTo Fail

    s = Trim$(s)
    s = Replace(s, ",", ".")

    If s = "" Then GoTo Fail

    'Val memakai titik sebagai decimal separator dan lebih aman lintas locale VBA.
    outVal = Val(s)

    If outVal = 0# And Left$(s, 1) <> "0" Then GoTo Fail

    HAMNREC_TryDbl = True
    Exit Function

Fail:
    HAMNREC_TryDbl = False

End Function


Private Function HAMNREC_DblToStr(ByVal v As Double) As String
    HAMNREC_DblToStr = Replace(Format$(v, "0.000"), ",", ".")
End Function


Private Function HAMNREC_CleanField(ByVal s As String) As String

    s = Replace(s, "|", " ")
    s = Replace(s, vbCr, " ")
    s = Replace(s, vbLf, " ")
    s = Trim$(s)

    HAMNREC_CleanField = s

End Function


Private Function HAMNREC_MinD(ByVal a As Double, ByVal b As Double) As Double
    If a < b Then
        HAMNREC_MinD = a
    Else
        HAMNREC_MinD = b
    End If
End Function


Private Function HAMNREC_CenterX(ByVal shp As Shape) As Double
    HAMNREC_CenterX = shp.LeftX + (shp.SizeWidth / 2#)
End Function


Private Function HAMNREC_CenterY(ByVal shp As Shape) As Double
    HAMNREC_CenterY = shp.TopY - (shp.SizeHeight / 2#)
End Function


Private Function HAMNREC_DictKeysCsv(ByVal d As Object) As String

    Dim key As Variant
    Dim s As String

    s = ""

    If d Is Nothing Then
        HAMNREC_DictKeysCsv = ""
        Exit Function
    End If

    For Each key In d.Keys
        If s <> "" Then s = s & ","
        s = s & CStr(key)
    Next key

    HAMNREC_DictKeysCsv = s

End Function


Private Sub HAMNREC_WriteReport( _
    ByVal reportPath As String, _
    ByVal status As String, _
    ByVal orderInfo As String, _
    ByVal dbSource As String, _
    ByVal detectLog As String, _
    ByVal catalogPath As String, _
    ByVal hardError As String _
)

    On Error Resume Next

    Dim f As Integer

    f = FreeFile
    Open reportPath For Output As #f

    Print #f, "PROJECT HADES — RECORD PATTERN CATALOG V3.3 AUTO-ONLY + UID MARKER REPORT"
    Print #f, "CREATED=" & Format$(Now, "yyyy-mm-dd hh:nn:ss")
    Print #f, "STATUS=" & status
    Print #f, ""
    Print #f, "[ORDER]"
    Print #f, orderInfo
    Print #f, ""
    Print #f, "[SIZEDB SOURCE]"
    Print #f, dbSource
    Print #f, ""
    Print #f, "[DETECT LOG]"
    Print #f, detectLog
    Print #f, ""
    Print #f, "[CATALOG]"
    Print #f, catalogPath
    Print #f, ""
    Print #f, "[ERROR]"
    If hardError <> "" Then
        Print #f, hardError
    Else
        Print #f, "NONE"
    End If

    Close #f

End Sub
