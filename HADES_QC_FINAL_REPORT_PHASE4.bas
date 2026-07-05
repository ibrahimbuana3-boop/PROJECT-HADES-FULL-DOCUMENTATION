Option Explicit

'=========================================================
' PROJECT HADES — QC FINAL REPORT PHASE 4
' CorelDRAW 2021 VBA
'
' MAIN MACRO:
' HADES_QC_FINAL_REPORT_V4
'
' DEPENDENCY:
' - HADES_CORE_REPORT_PHASE2.bas
' - HADES_QC_FINAL_REPORT_PHASE3C.bas
' - QC_SIZE_CHECK V8.2 REPORT MODE
' - QC_TYPO_CHECK V12.1 REPORT MODE
'
' TUJUAN PHASE 4:
' - Menambahkan FINAL QC LOCK setelah Final QC Report V3C selesai.
' - Lock menyimpan fingerprint dokumen + selection layout.
' - HADES_FINALIZE_CONVERT_V4 nanti hanya boleh convert jika:
'   1) report terakhir PASS,
'   2) lock terakhir cocok dengan dokumen aktif,
'   3) selection saat convert sama dengan selection saat QC,
'   4) lock/report belum basi.
'
' CATATAN:
' - Macro ini tidak mengganti V3C.
' - Macro ini membungkus V3C lalu menulis lock file.
'=========================================================

Private Const H4_LOCK_VERSION As String = "4"
Private Const H4_LOCK_LATEST As String = "HADES_FINAL_QC_LOCK_LATEST.txt"

'=========================================================
' MAIN
'=========================================================
Public Sub HADES_QC_FINAL_REPORT_V4()

    Dim sigBefore As String
    Dim detailBefore As String
    Dim reportPath As String
    Dim finalStatus As String
    Dim permission As String
    Dim generatedText As String
    Dim lockPath As String

    On Error GoTo ERR_HANDLER

    If Not H4_HasSelection() Then
        MsgBox "Pilih hasil layout final yang ingin di-QC.", vbExclamation, "HADES QC FINAL V4"
        Exit Sub
    End If

    sigBefore = H4_BuildCurrentSelectionSignature(detailBefore)

    'Panggil report final otomatis Phase 3C.
    'V3C tetap menjadi validator utama; V4 hanya menambahkan lock produksi.
    Call HADES_QC_FINAL_REPORT_V3C

    reportPath = HADESR_LatestFinalReportPath()
    finalStatus = UCase$(Trim$(HADESR_ReadLatestFinalStatus()))
    permission = UCase$(Trim$(HADESR_ReadLatestConvertPermission()))
    generatedText = H4_ReadMachineValueFromFile(reportPath, "GENERATED")

    If finalStatus = "" Or permission = "" Or Dir$(reportPath) = "" Then
        MsgBox _
            "Final QC Report V3C belum menghasilkan report valid." & vbCrLf & vbCrLf & _
            "QC Lock V4 tidak dibuat." & vbCrLf & _
            "Perbaiki error pada Final QC Report lalu jalankan ulang.", _
            vbCritical, _
            "HADES QC FINAL V4"
        Exit Sub
    End If

    lockPath = H4_WriteFinalQCLock(sigBefore, detailBefore, reportPath, finalStatus, permission, generatedText)

    MsgBox _
        "HADES FINAL QC LOCK V4 SELESAI" & vbCrLf & vbCrLf & _
        "Final Status       : " & finalStatus & vbCrLf & _
        "Convert Permission : " & permission & vbCrLf & _
        "Report Generated   : " & generatedText & vbCrLf & vbCrLf & _
        "Lock:" & vbCrLf & lockPath & vbCrLf & vbCrLf & _
        "Langkah berikutnya:" & vbCrLf & _
        "Run HADES_FINALIZE_CONVERT_V4 pada selection layout yang sama.", _
        vbInformation, _
        "HADES QC FINAL V4"

    Exit Sub

ERR_HANDLER:

    MsgBox _
        "SYSTEM ERROR - HADES QC FINAL REPORT V4" & vbCrLf & vbCrLf & _
        "No : " & Err.Number & vbCrLf & _
        Err.Description, _
        vbCritical, _
        "HADES QC FINAL V4"

End Sub

'=========================================================
' LOCK WRITER
'=========================================================
Private Function H4_WriteFinalQCLock( _
    ByVal selectionSignature As String, _
    ByVal selectionDetail As String, _
    ByVal reportPath As String, _
    ByVal finalStatus As String, _
    ByVal permission As String, _
    ByVal generatedText As String) As String

    Dim folderPath As String
    Dim latestPath As String
    Dim timePath As String
    Dim ts As String
    Dim txt As String

    folderPath = HADESR_ReportFolderPath()
    H4_EnsureFolder folderPath

    ts = Format(Now, "yyyymmdd_hhnnss")

    latestPath = folderPath & "\" & H4_LOCK_LATEST
    timePath = folderPath & "\HADES_FINAL_QC_LOCK_" & ts & ".txt"

    txt = "# PROJECT_HADES_QC_LOCK" & vbCrLf
    txt = txt & "LOCK_VERSION=" & H4_LOCK_VERSION & vbCrLf
    txt = txt & "LOCK_TYPE=FINAL_QC" & vbCrLf
    txt = txt & "LOCK_CREATED=" & Format(Now, "yyyy-mm-dd hh:nn:ss") & vbCrLf
    txt = txt & "DOCUMENT_NAME=" & H4_DocumentName() & vbCrLf
    txt = txt & "DOCUMENT_FULLNAME=" & H4_DocumentFullName() & vbCrLf
    txt = txt & "REPORT_PATH=" & reportPath & vbCrLf
    txt = txt & "REPORT_GENERATED=" & generatedText & vbCrLf
    txt = txt & "FINAL_STATUS=" & finalStatus & vbCrLf
    txt = txt & "CONVERT_PERMISSION=" & permission & vbCrLf
    txt = txt & "SELECTION_SIGNATURE=" & selectionSignature & vbCrLf
    txt = txt & "# END_PROJECT_HADES_QC_LOCK" & vbCrLf & vbCrLf

    txt = txt & "PROJECT HADES — FINAL QC LOCK V4" & vbCrLf
    txt = txt & String(70, "=") & vbCrLf
    txt = txt & "Lock Created       : " & Format(Now, "yyyy-mm-dd hh:nn:ss") & vbCrLf
    txt = txt & "Document           : " & H4_DocumentName() & vbCrLf
    txt = txt & "Document FullName  : " & H4_DocumentFullName() & vbCrLf
    txt = txt & "Final Status       : " & finalStatus & vbCrLf
    txt = txt & "Convert Permission : " & permission & vbCrLf
    txt = txt & "Report Generated   : " & generatedText & vbCrLf
    txt = txt & "Report Path        : " & reportPath & vbCrLf
    txt = txt & String(70, "-") & vbCrLf
    txt = txt & "SELECTION SNAPSHOT" & vbCrLf
    txt = txt & String(70, "-") & vbCrLf
    txt = txt & selectionDetail & vbCrLf
    txt = txt & "Signature:" & vbCrLf & selectionSignature & vbCrLf

    H4_WriteText latestPath, txt
    H4_WriteText timePath, txt

    H4_WriteFinalQCLock = latestPath

End Function

'=========================================================
' PUBLIC HELPER
'=========================================================
Public Sub HADES4_OpenLatestQCLock()

    Dim p As String
    p = HADESR_ReportFolderPath() & "\" & H4_LOCK_LATEST

    If Dir$(p) = "" Then
        MsgBox "QC Lock latest belum ditemukan:" & vbCrLf & p, vbExclamation, "HADES QC LOCK"
        Exit Sub
    End If

    Shell "notepad.exe " & Chr$(34) & p & Chr$(34), vbNormalFocus

End Sub

'=========================================================
' SELECTION SNAPSHOT / SIGNATURE
'=========================================================
Private Function H4_BuildCurrentSelectionSignature(ByRef detailText As String) As String

    Dim sr As ShapeRange
    Dim s As Shape

    Dim topCount As Long
    Dim shapeCount As Long
    Dim groupCount As Long
    Dim textCount As Long
    Dim panelCount As Long
    Dim powerClipCount As Long
    Dim textHash As Double
    Dim panelHash As Double
    Dim geoHash As Double

    Dim x As Double
    Dim y As Double
    Dim w As Double
    Dim h As Double
    Dim bboxText As String

    On Error GoTo SAFE_FAIL

    Set sr = ActiveSelectionRange

    topCount = 0
    shapeCount = 0
    groupCount = 0
    textCount = 0
    panelCount = 0
    powerClipCount = 0
    textHash = 0
    panelHash = 0
    geoHash = 0

    For Each s In sr
        topCount = topCount + 1
        H4_ScanSignatureShape s, shapeCount, groupCount, textCount, panelCount, powerClipCount, textHash, panelHash, geoHash
    Next s

    On Error Resume Next
    sr.GetBoundingBox x, y, w, h, True
    On Error GoTo 0

    bboxText = _
        FormatNumber(x, 3, False, False, False) & "," & _
        FormatNumber(y, 3, False, False, False) & "," & _
        FormatNumber(w, 3, False, False, False) & "," & _
        FormatNumber(h, 3, False, False, False)

    detailText = "Top-level selected : " & topCount & vbCrLf & _
                 "Total shapes       : " & shapeCount & vbCrLf & _
                 "Groups             : " & groupCount & vbCrLf & _
                 "Active texts       : " & textCount & vbCrLf & _
                 "Panel outlines     : " & panelCount & vbCrLf & _
                 "PowerClip objects  : " & powerClipCount & vbCrLf & _
                 "BoundingBox x,y,w,h: " & bboxText & vbCrLf & _
                 "Text hash          : " & FormatNumber(textHash, 0, False, False, False) & vbCrLf & _
                 "Panel hash         : " & FormatNumber(panelHash, 0, False, False, False) & vbCrLf & _
                 "Geometry hash      : " & FormatNumber(geoHash, 0, False, False, False) & vbCrLf

    H4_BuildCurrentSelectionSignature = _
        "DOC=" & H4_DocumentKey() & _
        "|TOP=" & CStr(topCount) & _
        "|SHAPES=" & CStr(shapeCount) & _
        "|GROUPS=" & CStr(groupCount) & _
        "|TEXTS=" & CStr(textCount) & _
        "|PANELS=" & CStr(panelCount) & _
        "|PCLIPS=" & CStr(powerClipCount) & _
        "|BBOX=" & bboxText & _
        "|THASH=" & FormatNumber(textHash, 0, False, False, False) & _
        "|PHASH=" & FormatNumber(panelHash, 0, False, False, False) & _
        "|GHASH=" & FormatNumber(geoHash, 0, False, False, False)

    Exit Function

SAFE_FAIL:

    detailText = "Gagal membuat selection signature. Error " & Err.Number & ": " & Err.Description
    H4_BuildCurrentSelectionSignature = "ERROR_SIGNATURE"

End Function

Private Sub H4_ScanSignatureShape( _
    ByVal s As Shape, _
    ByRef shapeCount As Long, _
    ByRef groupCount As Long, _
    ByRef textCount As Long, _
    ByRef panelCount As Long, _
    ByRef powerClipCount As Long, _
    ByRef textHash As Double, _
    ByRef panelHash As Double, _
    ByRef geoHash As Double)

    Dim ch As Shape
    Dim pcShapes As Shapes
    Dim txt As String

    On Error Resume Next

    If s Is Nothing Then Exit Sub

    shapeCount = shapeCount + 1
    geoHash = geoHash + H4_ShapeGeometryCode(s)

    If s.Type = cdrGroupShape Then
        groupCount = groupCount + 1
        For Each ch In s.Shapes
            H4_ScanSignatureShape ch, shapeCount, groupCount, textCount, panelCount, powerClipCount, textHash, panelHash, geoHash
        Next ch
    End If

    If s.Type = cdrTextShape Then
        textCount = textCount + 1
        txt = ""
        txt = s.Text.Story.Text
        textHash = textHash + H4_TextChecksum(txt)
    End If

    If s.Type = cdrCurveShape Then
        If H4_IsPanelOutline(s) Then
            panelCount = panelCount + 1
            panelHash = panelHash + H4_PanelGeometryCode(s)
        End If
    End If

    Set pcShapes = Nothing
    Set pcShapes = s.PowerClip.Shapes

    If Not pcShapes Is Nothing Then
        powerClipCount = powerClipCount + 1
        For Each ch In pcShapes
            H4_ScanSignatureShape ch, shapeCount, groupCount, textCount, panelCount, powerClipCount, textHash, panelHash, geoHash
        Next ch
    End If

    On Error GoTo 0

End Sub

Private Function H4_TextChecksum(ByVal txt As String) As Double

    Dim i As Long
    Dim n As Long
    Dim code As Double

    txt = UCase$(Trim$(Replace(Replace(Replace(txt, vbCr, ""), vbLf, ""), vbTab, " ")))

    code = 0

    For i = 1 To Len(txt)
        n = AscW(Mid$(txt, i, 1))
        If n < 0 Then n = n + 65536
        code = code + (CDbl(n) * (i + 17))
    Next i

    H4_TextChecksum = code + Len(txt) * 100003

End Function

Private Function H4_ShapeGeometryCode(ByVal s As Shape) As Double

    On Error GoTo SAFE_FAIL

    H4_ShapeGeometryCode = _
        Round(Abs(CDbl(s.SizeWidth)) * 1000, 0) * 3 + _
        Round(Abs(CDbl(s.SizeHeight)) * 1000, 0) * 5

    Exit Function

SAFE_FAIL:

    H4_ShapeGeometryCode = 0

End Function

Private Function H4_PanelGeometryCode(ByVal s As Shape) As Double

    On Error GoTo SAFE_FAIL

    H4_PanelGeometryCode = _
        Round(Abs(CDbl(s.SizeWidth)) * 1000, 0) * 11 + _
        Round(Abs(CDbl(s.SizeHeight)) * 1000, 0) * 13

    Exit Function

SAFE_FAIL:

    H4_PanelGeometryCode = 0

End Function

'=========================================================
' PANEL OUTLINE DETECTION
'=========================================================
Private Function H4_IsPanelOutline(ByVal s As Shape) As Boolean

    If H4_IsRedOutline(s) Then
        H4_IsPanelOutline = True
        Exit Function
    End If

    If H4_IsGreenOutline(s) Then
        H4_IsPanelOutline = True
        Exit Function
    End If

    H4_IsPanelOutline = False

End Function

Private Function H4_IsRedOutline(ByVal s As Shape) As Boolean

    Dim r As Long
    Dim g As Long
    Dim b As Long

    On Error GoTo SAFE_EXIT

    H4_IsRedOutline = False

    If s.Outline.Width <= 0 Then Exit Function

    r = s.Outline.Color.RGBRed
    g = s.Outline.Color.RGBGreen
    b = s.Outline.Color.RGBBlue

    If r >= 230 And g <= 80 And b <= 80 Then
        H4_IsRedOutline = True
    End If

SAFE_EXIT:

End Function

Private Function H4_IsGreenOutline(ByVal s As Shape) As Boolean

    Dim r As Long
    Dim g As Long
    Dim b As Long

    On Error GoTo SAFE_EXIT

    H4_IsGreenOutline = False

    If s.Outline.Width <= 0 Then Exit Function

    r = s.Outline.Color.RGBRed
    g = s.Outline.Color.RGBGreen
    b = s.Outline.Color.RGBBlue

    If r <= 80 And g >= 180 And b <= 80 Then
        H4_IsGreenOutline = True
        Exit Function
    End If

    If Abs(r - 97) <= 25 And Abs(g - 186) <= 25 And Abs(b - 12) <= 25 Then
        H4_IsGreenOutline = True
        Exit Function
    End If

SAFE_EXIT:

End Function

'=========================================================
' FILE / MACHINE VALUE HELPERS
'=========================================================
Private Function H4_ReadMachineValueFromFile(ByVal path As String, ByVal keyName As String) As String

    Dim txt As String
    Dim lines() As String
    Dim i As Long
    Dim lineText As String
    Dim prefix As String

    txt = H4_ReadText(path)
    txt = Replace(txt, vbCrLf, vbLf)
    txt = Replace(txt, vbCr, vbLf)
    lines = Split(txt, vbLf)

    prefix = UCase$(Trim$(keyName)) & "="

    For i = LBound(lines) To UBound(lines)
        lineText = Trim$(CStr(lines(i)))
        If UCase$(Left$(lineText, Len(prefix))) = prefix Then
            H4_ReadMachineValueFromFile = Trim$(Mid$(lineText, Len(prefix) + 1))
            Exit Function
        End If
    Next i

    H4_ReadMachineValueFromFile = ""

End Function

Private Function H4_HasSelection() As Boolean

    On Error GoTo NO_SELECTION

    H4_HasSelection = False

    If ActiveSelection Is Nothing Then Exit Function
    If ActiveSelection.Shapes.Count <= 0 Then Exit Function

    H4_HasSelection = True
    Exit Function

NO_SELECTION:

    H4_HasSelection = False

End Function

Private Function H4_DocumentName() As String

    On Error GoTo SAFE_FAIL
    H4_DocumentName = ActiveDocument.Name
    Exit Function

SAFE_FAIL:
    H4_DocumentName = ""

End Function

Private Function H4_DocumentFullName() As String

    On Error GoTo SAFE_FAIL
    H4_DocumentFullName = ActiveDocument.FullFileName
    Exit Function

SAFE_FAIL:
    H4_DocumentFullName = H4_DocumentName()

End Function

Private Function H4_DocumentKey() As String

    Dim f As String
    f = Trim$(H4_DocumentFullName())

    If f <> "" Then
        H4_DocumentKey = f
    Else
        H4_DocumentKey = H4_DocumentName()
    End If

End Function

Private Sub H4_EnsureFolder(ByVal folderPath As String)

    On Error Resume Next
    If Dir$(folderPath, vbDirectory) = "" Then MkDir folderPath
    On Error GoTo 0

End Sub

Private Function H4_ReadText(ByVal path As String) As String

    On Error GoTo FAIL_UTF8

    Dim stm As Object
    Set stm = CreateObject("ADODB.Stream")

    With stm
        .Type = 2
        .Charset = "utf-8"
        .Open
        .LoadFromFile path
        H4_ReadText = .ReadText
        .Close
    End With

    Exit Function

FAIL_UTF8:

    On Error GoTo FAIL_ANSI

    Dim f As Integer
    f = FreeFile

    Open path For Input As #f
    H4_ReadText = Input$(LOF(f), #f)
    Close #f

    Exit Function

FAIL_ANSI:

    On Error Resume Next
    Close #f
    H4_ReadText = ""

End Function

Private Sub H4_WriteText(ByVal path As String, ByVal txt As String)

    Dim f As Integer
    f = FreeFile

    Open path For Output As #f
    Print #f, txt
    Close #f

End Sub
