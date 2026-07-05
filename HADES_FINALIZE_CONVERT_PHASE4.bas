Option Explicit

'=========================================================
' PROJECT HADES — FINALIZE CONVERT PHASE 4
' CorelDRAW 2021 VBA
'
' MAIN MACRO:
' HADES_FINALIZE_CONVERT_V4
'
' DEPENDENCY:
' - HADES_CORE_REPORT_PHASE2.bas
' - HADES_QC_FINAL_REPORT_PHASE4.bas sudah pernah dijalankan
'
' TUJUAN PHASE 4:
' - Convert Gate final/hardening.
' - Convert hanya boleh jika Final QC terakhir PASS dan ALLOWED.
' - Final QC Lock harus cocok dengan dokumen aktif.
' - Selection saat convert harus sama dengan selection saat Final QC.
' - Lock tidak boleh basi.
' - Active text dihitung sebelum convert.
' - Semua tindakan ditulis ke log gate/convert.
'=========================================================

Private Const H4C_LOCK_LATEST As String = "HADES_FINAL_QC_LOCK_LATEST.txt"
Private Const H4C_MAX_LOCK_AGE_MINUTES As Long = 480 '8 jam kerja. Ubah jika ingin lebih ketat/longgar.
Private Const H4C_CONFIRM_WORD As String = "FINAL"

'=========================================================
' MAIN
'=========================================================
Public Sub HADES_FINALIZE_CONVERT_V4()

    Dim gateOk As Boolean
    Dim gateReason As String
    Dim gateDetail As String

    Dim finalStatus As String
    Dim permission As String
    Dim latestReportPath As String
    Dim latestGenerated As String

    Dim lockPath As String
    Dim lockStatus As String
    Dim lockPermission As String
    Dim lockCreated As String
    Dim lockReportPath As String
    Dim lockReportGenerated As String
    Dim lockDocumentFullName As String
    Dim lockSignature As String
    Dim currentSignature As String
    Dim currentDetail As String
    Dim ageMinutes As Long

    Dim textShapes As Collection
    Dim convertedCount As Long
    Dim failedCount As Long
    Dim ans As String

    Dim oldUnit As cdrUnit
    Dim oldOptimization As Boolean
    Dim cmdStarted As Boolean

    On Error GoTo ERR_HANDLER

    latestReportPath = HADESR_LatestFinalReportPath()
    lockPath = HADESR_ReportFolderPath() & "\" & H4C_LOCK_LATEST

    finalStatus = UCase$(Trim$(HADESR_ReadLatestFinalStatus()))
    permission = UCase$(Trim$(HADESR_ReadLatestConvertPermission()))
    latestGenerated = H4C_ReadMachineValueFromFile(latestReportPath, "GENERATED")

    lockStatus = UCase$(Trim$(H4C_ReadLockValue("FINAL_STATUS")))
    lockPermission = UCase$(Trim$(H4C_ReadLockValue("CONVERT_PERMISSION")))
    lockCreated = Trim$(H4C_ReadLockValue("LOCK_CREATED"))
    lockReportPath = Trim$(H4C_ReadLockValue("REPORT_PATH"))
    lockReportGenerated = Trim$(H4C_ReadLockValue("REPORT_GENERATED"))
    lockDocumentFullName = Trim$(H4C_ReadLockValue("DOCUMENT_FULLNAME"))
    lockSignature = Trim$(H4C_ReadLockValue("SELECTION_SIGNATURE"))

    gateOk = True
    gateReason = ""
    gateDetail = ""

    '=====================================================
    ' GATE 1 — report dan lock harus ada
    '=====================================================
    If Dir$(latestReportPath) = "" Then
        gateOk = False
        gateReason = "Latest Final QC Report tidak ditemukan."
        gateDetail = gateDetail & "Missing report: " & latestReportPath & vbCrLf
    End If

    If Dir$(lockPath) = "" Then
        gateOk = False
        gateReason = "Final QC Lock V4 tidak ditemukan."
        gateDetail = gateDetail & "Missing lock: " & lockPath & vbCrLf
    End If

    '=====================================================
    ' GATE 2 — report harus PASS murni
    '=====================================================
    If finalStatus <> "PASS" Or permission <> "ALLOWED" Then
        gateOk = False
        gateReason = "Final QC terakhir belum PASS / ALLOWED."
        gateDetail = gateDetail & "Report FINAL_STATUS=" & finalStatus & vbCrLf
        gateDetail = gateDetail & "Report CONVERT_PERMISSION=" & permission & vbCrLf
    End If

    If lockStatus <> "PASS" Or lockPermission <> "ALLOWED" Then
        gateOk = False
        gateReason = "QC Lock terakhir bukan PASS / ALLOWED."
        gateDetail = gateDetail & "Lock FINAL_STATUS=" & lockStatus & vbCrLf
        gateDetail = gateDetail & "Lock CONVERT_PERMISSION=" & lockPermission & vbCrLf
    End If

    '=====================================================
    ' GATE 3 — lock/report harus saling cocok
    '=====================================================
    If Trim$(lockReportPath) = "" Then
        gateOk = False
        gateReason = "QC Lock tidak menyimpan REPORT_PATH."
    ElseIf LCase$(Trim$(lockReportPath)) <> LCase$(Trim$(latestReportPath)) Then
        gateOk = False
        gateReason = "QC Lock tidak mengarah ke latest report yang sedang dibaca."
        gateDetail = gateDetail & "Lock Report : " & lockReportPath & vbCrLf
        gateDetail = gateDetail & "Latest      : " & latestReportPath & vbCrLf
    End If

    If Trim$(lockReportGenerated) <> "" And Trim$(latestGenerated) <> "" Then
        If Trim$(lockReportGenerated) <> Trim$(latestGenerated) Then
            gateOk = False
            gateReason = "Report generated time berbeda dengan QC Lock."
            gateDetail = gateDetail & "Lock report generated  : " & lockReportGenerated & vbCrLf
            gateDetail = gateDetail & "Latest report generated: " & latestGenerated & vbCrLf
        End If
    End If

    '=====================================================
    ' GATE 4 — lock tidak boleh basi
    '=====================================================
    ageMinutes = H4C_MinutesSince(lockCreated)

    If ageMinutes < 0 Then
        gateOk = False
        gateReason = "LOCK_CREATED tidak valid."
        gateDetail = gateDetail & "LOCK_CREATED=" & lockCreated & vbCrLf
    ElseIf ageMinutes > H4C_MAX_LOCK_AGE_MINUTES Then
        gateOk = False
        gateReason = "Final QC Lock sudah basi. Jalankan Final QC V4 ulang."
        gateDetail = gateDetail & "Lock age minutes : " & ageMinutes & vbCrLf
        gateDetail = gateDetail & "Max allowed      : " & H4C_MAX_LOCK_AGE_MINUTES & vbCrLf
    End If

    '=====================================================
    ' GATE 5 — selection dan dokumen harus valid/cocok
    '=====================================================
    If Not H4C_HasSelection() Then
        gateOk = False
        gateReason = "Tidak ada selection layout untuk convert."
    Else
        currentSignature = H4C_BuildCurrentSelectionSignature(currentDetail)

        If Trim$(lockSignature) = "" Then
            gateOk = False
            gateReason = "QC Lock tidak menyimpan SELECTION_SIGNATURE."
        ElseIf Trim$(currentSignature) <> Trim$(lockSignature) Then
            gateOk = False
            gateReason = "Selection saat convert berbeda dari selection saat Final QC."
            gateDetail = gateDetail & vbCrLf & "CURRENT SELECTION:" & vbCrLf & currentDetail & vbCrLf
            gateDetail = gateDetail & "CURRENT SIGNATURE:" & vbCrLf & currentSignature & vbCrLf & vbCrLf
            gateDetail = gateDetail & "LOCK SIGNATURE:" & vbCrLf & lockSignature & vbCrLf
        End If

        If Trim$(lockDocumentFullName) <> "" Then
            If LCase$(Trim$(lockDocumentFullName)) <> LCase$(Trim$(H4C_DocumentFullName())) Then
                gateOk = False
                gateReason = "Dokumen aktif berbeda dari dokumen saat Final QC."
                gateDetail = gateDetail & "Lock Document  : " & lockDocumentFullName & vbCrLf
                gateDetail = gateDetail & "Active Document: " & H4C_DocumentFullName() & vbCrLf
            End If
        End If
    End If

    '=====================================================
    ' Gate failed: tulis log dan blokir convert.
    '=====================================================
    If Not gateOk Then
        H4C_WriteGateLog "BLOCKED", gateReason, gateDetail, 0, 0

        MsgBox _
            "FINAL CONVERT V4 DIBLOKIR" & vbCrLf & vbCrLf & _
            "Alasan:" & vbCrLf & gateReason & vbCrLf & vbCrLf & _
            "Yang harus dilakukan:" & vbCrLf & _
            "1. Select ulang layout final yang benar." & vbCrLf & _
            "2. Run HADES_QC_FINAL_REPORT_V4." & vbCrLf & _
            "3. Setelah PASS, run HADES_FINALIZE_CONVERT_V4 lagi.", _
            vbCritical, _
            "HADES CONVERT GATE V4"

        HADES4_OpenLatestQCLock
        Exit Sub
    End If

    '=====================================================
    ' GATE 6 — hitung active text sebelum convert
    '=====================================================
    Set textShapes = New Collection
    H4C_CollectTextShapes ActiveSelectionRange, textShapes

    If textShapes.Count = 0 Then
        H4C_WriteGateLog "NO_TEXT", "Tidak ada active text di selection.", currentDetail, 0, 0

        MsgBox _
            "Tidak ada active text di selection." & vbCrLf & vbCrLf & _
            "Tidak ada objek yang perlu di-convert.", _
            vbInformation, _
            "HADES FINALIZE CONVERT V4"
        Exit Sub
    End If

    ans = InputBox( _
        "FINAL CONVERT GATE V4 — PASS" & vbCrLf & vbCrLf & _
        "Report dan QC Lock cocok." & vbCrLf & _
        "Final Status       : " & finalStatus & vbCrLf & _
        "Convert Permission : " & permission & vbCrLf & _
        "Lock Created       : " & lockCreated & vbCrLf & _
        "Lock Age           : " & ageMinutes & " menit" & vbCrLf & _
        "Active Text        : " & textShapes.Count & vbCrLf & vbCrLf & _
        "Macro ini akan mengubah semua ACTIVE TEXT di selection menjadi curve." & vbCrLf & _
        "Setelah ini QC_TYPO_CHECK dan IDPO_CHECK tidak bisa membaca teks lagi." & vbCrLf & vbCrLf & _
        "Ketik FINAL untuk lanjut.", _
        "HADES FINALIZE CONVERT V4")

    ans = UCase$(Trim$(ans))

    If ans <> H4C_CONFIRM_WORD Then
        H4C_WriteGateLog "CANCELLED", "Operator membatalkan final convert.", currentDetail, 0, 0
        MsgBox "Final convert dibatalkan. Tidak ada objek yang diubah.", vbInformation, "HADES FINALIZE CONVERT V4"
        Exit Sub
    End If

    '=====================================================
    ' EXECUTE CONVERT
    '=====================================================
    oldUnit = ActiveDocument.Unit
    oldOptimization = Application.Optimization
    cmdStarted = False

    ActiveDocument.Unit = cdrCentimeter
    Application.Optimization = True
    ActiveDocument.BeginCommandGroup "HADES FINALIZE CONVERT V4 ACTIVE TEXT"
    cmdStarted = True

    H4C_ConvertTexts textShapes, convertedCount, failedCount

    ActiveDocument.EndCommandGroup
    cmdStarted = False

    Application.Optimization = oldOptimization
    ActiveDocument.Unit = oldUnit
    ActiveWindow.Refresh

    H4C_WriteGateLog "CONVERTED", "Final convert selesai.", currentDetail, convertedCount, failedCount

    MsgBox _
        "HADES FINAL CONVERT V4 SELESAI" & vbCrLf & vbCrLf & _
        "Converted : " & convertedCount & vbCrLf & _
        "Failed    : " & failedCount & vbCrLf & vbCrLf & _
        "Log convert ditulis ke Documents\HADES_REPORTS.", _
        vbInformation, _
        "HADES FINALIZE CONVERT V4"

    Exit Sub

ERR_HANDLER:

    On Error Resume Next

    If cmdStarted Then ActiveDocument.EndCommandGroup

    Application.Optimization = oldOptimization
    ActiveDocument.Unit = oldUnit
    ActiveWindow.Refresh

    H4C_WriteGateLog "ERROR", "System error saat final convert.", "Error " & Err.Number & ": " & Err.Description, convertedCount, failedCount

    MsgBox _
        "SYSTEM ERROR - HADES FINALIZE CONVERT V4" & vbCrLf & vbCrLf & _
        "No : " & Err.Number & vbCrLf & _
        Err.Description, _
        vbCritical, _
        "HADES FINALIZE CONVERT V4"

End Sub

'=========================================================
' TEXT COLLECTION / CONVERT
'=========================================================
Private Sub H4C_CollectTextShapes(ByVal sr As ShapeRange, ByVal textShapes As Collection)

    Dim s As Shape

    For Each s In sr
        H4C_CollectTextShapeRecursive s, textShapes
    Next s

End Sub

Private Sub H4C_CollectTextShapeRecursive(ByVal s As Shape, ByVal textShapes As Collection)

    Dim ch As Shape
    Dim pcShapes As Shapes

    On Error Resume Next

    If s Is Nothing Then Exit Sub

    If s.Type = cdrTextShape Then
        textShapes.Add s
        Exit Sub
    End If

    If s.Type = cdrGroupShape Then
        For Each ch In s.Shapes
            H4C_CollectTextShapeRecursive ch, textShapes
        Next ch
    End If

    Set pcShapes = Nothing
    Set pcShapes = s.PowerClip.Shapes

    If Not pcShapes Is Nothing Then
        For Each ch In pcShapes
            H4C_CollectTextShapeRecursive ch, textShapes
        Next ch
    End If

    On Error GoTo 0

End Sub

Private Sub H4C_ConvertTexts(ByVal textShapes As Collection, ByRef convertedCount As Long, ByRef failedCount As Long)

    Dim v As Variant
    Dim t As Shape

    convertedCount = 0
    failedCount = 0

    For Each v In textShapes

        Set t = v

        If H4C_IsShapeAlive(t) Then
            If H4C_ConvertOneText(t) Then
                convertedCount = convertedCount + 1
            Else
                failedCount = failedCount + 1
            End If
        Else
            failedCount = failedCount + 1
        End If

    Next v

End Sub

Private Function H4C_ConvertOneText(ByVal t As Shape) As Boolean

    On Error GoTo FAIL

    H4C_ConvertOneText = False

    If t Is Nothing Then Exit Function
    If t.Type <> cdrTextShape Then Exit Function

    On Error Resume Next
    t.Locked = False
    Err.Clear
    On Error GoTo FAIL

    t.ConvertToCurves

    H4C_ConvertOneText = True
    Exit Function

FAIL:

    Err.Clear
    H4C_ConvertOneText = False

End Function

Private Function H4C_IsShapeAlive(ByVal s As Shape) As Boolean

    Dim t As Long

    On Error GoTo DEAD

    If s Is Nothing Then GoTo DEAD

    t = s.Type
    H4C_IsShapeAlive = True
    Exit Function

DEAD:

    H4C_IsShapeAlive = False

End Function

'=========================================================
' SELECTION SIGNATURE — HARUS SAMA DENGAN PHASE 4 QC LOCK
'=========================================================
Private Function H4C_BuildCurrentSelectionSignature(ByRef detailText As String) As String

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
        H4C_ScanSignatureShape s, shapeCount, groupCount, textCount, panelCount, powerClipCount, textHash, panelHash, geoHash
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

    H4C_BuildCurrentSelectionSignature = _
        "DOC=" & H4C_DocumentKey() & _
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
    H4C_BuildCurrentSelectionSignature = "ERROR_SIGNATURE"

End Function

Private Sub H4C_ScanSignatureShape( _
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
    geoHash = geoHash + H4C_ShapeGeometryCode(s)

    If s.Type = cdrGroupShape Then
        groupCount = groupCount + 1
        For Each ch In s.Shapes
            H4C_ScanSignatureShape ch, shapeCount, groupCount, textCount, panelCount, powerClipCount, textHash, panelHash, geoHash
        Next ch
    End If

    If s.Type = cdrTextShape Then
        textCount = textCount + 1
        txt = ""
        txt = s.Text.Story.Text
        textHash = textHash + H4C_TextChecksum(txt)
    End If

    If s.Type = cdrCurveShape Then
        If H4C_IsPanelOutline(s) Then
            panelCount = panelCount + 1
            panelHash = panelHash + H4C_PanelGeometryCode(s)
        End If
    End If

    Set pcShapes = Nothing
    Set pcShapes = s.PowerClip.Shapes

    If Not pcShapes Is Nothing Then
        powerClipCount = powerClipCount + 1
        For Each ch In pcShapes
            H4C_ScanSignatureShape ch, shapeCount, groupCount, textCount, panelCount, powerClipCount, textHash, panelHash, geoHash
        Next ch
    End If

    On Error GoTo 0

End Sub

Private Function H4C_TextChecksum(ByVal txt As String) As Double

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

    H4C_TextChecksum = code + Len(txt) * 100003

End Function

Private Function H4C_ShapeGeometryCode(ByVal s As Shape) As Double

    On Error GoTo SAFE_FAIL

    H4C_ShapeGeometryCode = _
        Round(Abs(CDbl(s.SizeWidth)) * 1000, 0) * 3 + _
        Round(Abs(CDbl(s.SizeHeight)) * 1000, 0) * 5

    Exit Function

SAFE_FAIL:

    H4C_ShapeGeometryCode = 0

End Function

Private Function H4C_PanelGeometryCode(ByVal s As Shape) As Double

    On Error GoTo SAFE_FAIL

    H4C_PanelGeometryCode = _
        Round(Abs(CDbl(s.SizeWidth)) * 1000, 0) * 11 + _
        Round(Abs(CDbl(s.SizeHeight)) * 1000, 0) * 13

    Exit Function

SAFE_FAIL:

    H4C_PanelGeometryCode = 0

End Function

'=========================================================
' PANEL OUTLINE DETECTION
'=========================================================
Private Function H4C_IsPanelOutline(ByVal s As Shape) As Boolean

    If H4C_IsRedOutline(s) Then
        H4C_IsPanelOutline = True
        Exit Function
    End If

    If H4C_IsGreenOutline(s) Then
        H4C_IsPanelOutline = True
        Exit Function
    End If

    H4C_IsPanelOutline = False

End Function

Private Function H4C_IsRedOutline(ByVal s As Shape) As Boolean

    Dim r As Long
    Dim g As Long
    Dim b As Long

    On Error GoTo SAFE_EXIT

    H4C_IsRedOutline = False

    If s.Outline.Width <= 0 Then Exit Function

    r = s.Outline.Color.RGBRed
    g = s.Outline.Color.RGBGreen
    b = s.Outline.Color.RGBBlue

    If r >= 230 And g <= 80 And b <= 80 Then
        H4C_IsRedOutline = True
    End If

SAFE_EXIT:

End Function

Private Function H4C_IsGreenOutline(ByVal s As Shape) As Boolean

    Dim r As Long
    Dim g As Long
    Dim b As Long

    On Error GoTo SAFE_EXIT

    H4C_IsGreenOutline = False

    If s.Outline.Width <= 0 Then Exit Function

    r = s.Outline.Color.RGBRed
    g = s.Outline.Color.RGBGreen
    b = s.Outline.Color.RGBBlue

    If r <= 80 And g >= 180 And b <= 80 Then
        H4C_IsGreenOutline = True
        Exit Function
    End If

    If Abs(r - 97) <= 25 And Abs(g - 186) <= 25 And Abs(b - 12) <= 25 Then
        H4C_IsGreenOutline = True
        Exit Function
    End If

SAFE_EXIT:

End Function

'=========================================================
' LOGGING
'=========================================================
Private Sub H4C_WriteGateLog( _
    ByVal gateStatus As String, _
    ByVal reasonText As String, _
    ByVal detailText As String, _
    ByVal convertedCount As Long, _
    ByVal failedCount As Long)

    Dim folderPath As String
    Dim latestPath As String
    Dim timePath As String
    Dim ts As String
    Dim txt As String

    folderPath = HADESR_ReportFolderPath()
    H4C_EnsureFolder folderPath

    ts = Format(Now, "yyyymmdd_hhnnss")
    latestPath = folderPath & "\HADES_FINAL_CONVERT_GATE_LOG_LATEST.txt"
    timePath = folderPath & "\HADES_FINAL_CONVERT_GATE_LOG_" & ts & ".txt"

    txt = "PROJECT HADES — FINAL CONVERT GATE LOG V4" & vbCrLf
    txt = txt & String(70, "=") & vbCrLf
    txt = txt & "Logged At          : " & Format(Now, "yyyy-mm-dd hh:nn:ss") & vbCrLf
    txt = txt & "Gate Status        : " & gateStatus & vbCrLf
    txt = txt & "Reason             : " & reasonText & vbCrLf
    txt = txt & "Document           : " & H4C_DocumentName() & vbCrLf
    txt = txt & "Document FullName  : " & H4C_DocumentFullName() & vbCrLf
    txt = txt & "Latest Report      : " & HADESR_LatestFinalReportPath() & vbCrLf
    txt = txt & "Latest Lock        : " & HADESR_ReportFolderPath() & "\" & H4C_LOCK_LATEST & vbCrLf
    txt = txt & "Converted Text     : " & convertedCount & vbCrLf
    txt = txt & "Failed Convert     : " & failedCount & vbCrLf
    txt = txt & String(70, "-") & vbCrLf
    txt = txt & "DETAIL" & vbCrLf
    txt = txt & String(70, "-") & vbCrLf
    txt = txt & detailText & vbCrLf

    H4C_WriteText latestPath, txt
    H4C_WriteText timePath, txt

End Sub

'=========================================================
' FILE / VALUE HELPERS
'=========================================================
Private Function H4C_ReadLockValue(ByVal keyName As String) As String

    H4C_ReadLockValue = H4C_ReadMachineValueFromFile( _
        HADESR_ReportFolderPath() & "\" & H4C_LOCK_LATEST, _
        keyName)

End Function

Private Function H4C_ReadMachineValueFromFile(ByVal path As String, ByVal keyName As String) As String

    Dim txt As String
    Dim lines() As String
    Dim i As Long
    Dim lineText As String
    Dim prefix As String

    txt = H4C_ReadText(path)
    txt = Replace(txt, vbCrLf, vbLf)
    txt = Replace(txt, vbCr, vbLf)
    lines = Split(txt, vbLf)

    prefix = UCase$(Trim$(keyName)) & "="

    For i = LBound(lines) To UBound(lines)
        lineText = Trim$(CStr(lines(i)))
        If UCase$(Left$(lineText, Len(prefix))) = prefix Then
            H4C_ReadMachineValueFromFile = Trim$(Mid$(lineText, Len(prefix) + 1))
            Exit Function
        End If
    Next i

    H4C_ReadMachineValueFromFile = ""

End Function

Private Function H4C_MinutesSince(ByVal dateText As String) As Long

    On Error GoTo FAIL

    If Not IsDate(dateText) Then GoTo FAIL

    H4C_MinutesSince = CLng(DateDiff("n", CDate(dateText), Now))

    If H4C_MinutesSince < 0 Then GoTo FAIL

    Exit Function

FAIL:

    H4C_MinutesSince = -1

End Function

Private Function H4C_HasSelection() As Boolean

    On Error GoTo NO_SELECTION

    H4C_HasSelection = False

    If ActiveSelection Is Nothing Then Exit Function
    If ActiveSelection.Shapes.Count <= 0 Then Exit Function

    H4C_HasSelection = True
    Exit Function

NO_SELECTION:

    H4C_HasSelection = False

End Function

Private Function H4C_DocumentName() As String

    On Error GoTo SAFE_FAIL
    H4C_DocumentName = ActiveDocument.Name
    Exit Function

SAFE_FAIL:
    H4C_DocumentName = ""

End Function

Private Function H4C_DocumentFullName() As String

    On Error GoTo SAFE_FAIL
    H4C_DocumentFullName = ActiveDocument.FullFileName
    Exit Function

SAFE_FAIL:
    H4C_DocumentFullName = H4C_DocumentName()

End Function

Private Function H4C_DocumentKey() As String

    Dim f As String
    f = Trim$(H4C_DocumentFullName())

    If f <> "" Then
        H4C_DocumentKey = f
    Else
        H4C_DocumentKey = H4C_DocumentName()
    End If

End Function

Private Sub H4C_EnsureFolder(ByVal folderPath As String)

    On Error Resume Next
    If Dir$(folderPath, vbDirectory) = "" Then MkDir folderPath
    On Error GoTo 0

End Sub

Private Function H4C_ReadText(ByVal path As String) As String

    On Error GoTo FAIL_UTF8

    Dim stm As Object
    Set stm = CreateObject("ADODB.Stream")

    With stm
        .Type = 2
        .Charset = "utf-8"
        .Open
        .LoadFromFile path
        H4C_ReadText = .ReadText
        .Close
    End With

    Exit Function

FAIL_UTF8:

    On Error GoTo FAIL_ANSI

    Dim f As Integer
    f = FreeFile

    Open path For Input As #f
    H4C_ReadText = Input$(LOF(f), #f)
    Close #f

    Exit Function

FAIL_ANSI:

    On Error Resume Next
    Close #f
    H4C_ReadText = ""

End Function

Private Sub H4C_WriteText(ByVal path As String, ByVal txt As String)

    Dim f As Integer
    f = FreeFile

    Open path For Output As #f
    Print #f, txt
    Close #f

End Sub
