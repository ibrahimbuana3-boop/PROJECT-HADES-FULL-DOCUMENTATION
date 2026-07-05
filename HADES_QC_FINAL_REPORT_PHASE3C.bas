Option Explicit

'=========================================================
' PROJECT HADES — QC FINAL REPORT PHASE 3C
' CorelDRAW 2021 VBA
'
' MAIN MACRO:
' HADES_QC_FINAL_REPORT_V3C
'
' DEPENDENCY:
' - HADES_CORE_REPORT_PHASE2.bas sudah harus diimport.
' - QC_SIZE_CHECK V8.2 REPORT MODE dan QC_TYPO_CHECK V12.1 REPORT MODE harus ada.
'
' TUJUAN PHASE 3C:
' - Production polish untuk Final QC Report otomatis.
' - Menambahkan PREFLIGHT DATA CHECK sebelum QC berat dijalankan.
' - Menjadikan QC_SIZE_CHECK dan QC_TYPO_CHECK native report-mode.
' - QC_SIZE report-mode tidak mengubah red outline menjadi green.
' - QC_TYPO_CHECK tetap dijalankan SEBELUM QC_SIZE report-mode.
' - Operator tidak perlu lagi memilih manual P / W / F untuk Typo dan Size.
' - IDPO, GROUP STRUCTURE, dan POWERCLIP/TRANSPARENCY tetap otomatis.
'
' CATATAN:
' - Macro ini tidak mengganti VBA lama.
' - Macro ini menjadi shortcut baru untuk QC Final.
' - HADES_FINALIZE_CONVERT_V3 direkomendasikan untuk final convert setelah V3C.
'=========================================================

Private Const H3C_ORDER_FILE As String = "\Documents\Order.txt"
Private Const H3C_ID_MIN_H As Double = 0.28
Private Const H3C_ID_MAX_H As Double = 0.65

Private Const H3C_TR_FAIL_OVER As Double = 75#
Private Const H3C_TR_WARN_FROM As Double = 71#
Private Const H3C_TR_WARN_TO As Double = 75#
Private Const H3C_UNKNOWN_TR_IS_FAIL As Boolean = False
Private Const H3C_DETAIL_LIMIT As Long = 80

Private Const H3C_MIN_CHILD_GROUP As Long = 2

'=========================================================
' MAIN
'=========================================================
Public Sub HADES_QC_FINAL_REPORT_V3C()

    Dim ans As Long
    Dim reportPath As String
    Dim msg As String

    On Error GoTo ERR_HANDLER

    If Not H3C_HasSelection() Then
        MsgBox "Pilih hasil layout yang ingin divalidasi.", vbExclamation, "HADES QC FINAL V3C"
        Exit Sub
    End If

    ans = MsgBox( _
        "HADES QC FINAL REPORT V3C" & vbCrLf & vbCrLf & _
        "Macro ini akan membuat report gabungan." & vbCrLf & vbCrLf & _
        "Status otomatis:" & vbCrLf & _
        "- QC Typo Check report-mode" & vbCrLf & _
        "- QC Size Check report-mode" & vbCrLf & _
        "- IDPO Check" & vbCrLf & _
        "- Group Structure Check" & vbCrLf & _
        "- Transparency + PowerClip Check" & vbCrLf & vbCrLf & _
        "Lanjut?", _
        vbQuestion + vbYesNo, _
        "HADES QC FINAL V3C")

    If ans <> vbYes Then Exit Sub

    HADESR_Reset "PROJECT HADES — FINAL QC REPORT V3C"

    HADESR_AddNote _
        "Phase 3C: QC Typo dan QC Size sudah memakai report-mode otomatis. " & _
        "QC Typo tetap dijalankan sebelum QC Size untuk menjaga red outline reference. " & _
        "Preflight data check dijalankan sebelum QC berat agar workflow full otomatis tidak berhenti karena file dasar hilang. " & _
        "Final convert direkomendasikan memakai HADES_FINALIZE_CONVERT_V3."

    'PHASE 3C PREFLIGHT:
    'Cek Order.txt, TypoTemplate_Current.txt, dan SizeDB sebelum QC berat.
    'Jika preflight gagal, validator berat dilewati agar tidak memunculkan popup fallback.
    If H3C_RunPreflightAutoReport() Then

        'IMPORTANT PHASE 3C:
        'Run TYPO before SIZE. Then run SIZE using native report-mode
        'without changing red outline into green.
        H3C_RunTypoCheckAutoReport
        H3C_RunSizeCheckAutoReport
        H3C_RunIDPOAutoReport
        H3C_RunGroupStructureAutoReport
        H3C_RunTransparencyPowerClipAutoReport

    Else

        HADESR_AddNote _
            "QC berat dilewati karena PREFLIGHT DATA CHECK gagal. " & _
            "Perbaiki Order.txt / TypoTemplate_Current.txt / SizeDB, lalu jalankan ulang Final QC."

    End If

    reportPath = HADESR_WriteFinalReport()

    msg = "HADES FINAL QC REPORT V3C SELESAI" & vbCrLf & vbCrLf & _
          "Final Status       : " & HADESR_FinalStatus() & vbCrLf & _
          "Convert Permission : " & HADESR_ConvertPermission() & vbCrLf & vbCrLf & _
          "Report:" & vbCrLf & reportPath & vbCrLf & vbCrLf & _
          "Langkah berikutnya:" & vbCrLf & _
          "Run HADES_FINALIZE_CONVERT_V3 jika ingin final convert."

    If HADESR_FinalStatus() = "FAIL" Then
        MsgBox msg, vbCritical, "HADES FINAL QC FAIL"
    ElseIf HADESR_FinalStatus() = "WARNING" Then
        MsgBox msg, vbExclamation, "HADES FINAL QC WARNING"
    Else
        MsgBox msg, vbInformation, "HADES FINAL QC PASS"
    End If

    Exit Sub

ERR_HANDLER:

    MsgBox _
        "SYSTEM ERROR - HADES QC FINAL REPORT V3C" & vbCrLf & vbCrLf & _
        "No : " & Err.Number & vbCrLf & _
        Err.Description, _
        vbCritical, _
        "HADES QC FINAL V3C"

End Sub


'=========================================================
' PREFLIGHT DATA CHECK - PHASE 3C
'=========================================================
Private Function H3C_RunPreflightAutoReport() As Boolean

    Dim orderPath As String
    Dim tplPath As String
    Dim tplAltPath As String
    Dim dbName As String
    Dim dbPath As String
    Dim detail As String
    Dim ok As Boolean

    ok = True
    detail = ""

    orderPath = Environ$("USERPROFILE") & "\Documents\Order.txt"
    tplPath = Environ$("USERPROFILE") & "\Documents\TypoTemplate_Current.txt"
    tplAltPath = Environ$("USERPROFILE") & "\Documents\TypoTemplate_currents.txt"

    If Dir$(orderPath) = "" Then
        ok = False
        detail = detail & "FAIL : Order.txt tidak ditemukan: " & orderPath & vbCrLf
    Else
        detail = detail & "OK   : Order.txt ditemukan." & vbCrLf
        dbName = H3C_ReadDBNameFromKeyValueFile(orderPath, True)
    End If

    If Dir$(tplPath) = "" Then
        If Dir$(tplAltPath) <> "" Then
            tplPath = tplAltPath
        End If
    End If

    If Dir$(tplPath) = "" Then
        ok = False
        detail = detail & "FAIL : TypoTemplate_Current.txt tidak ditemukan." & vbCrLf
        detail = detail & "       Jalankan BUILD_TYPO_TEMPLATE sebelum Auto Duplicate / Auto Rename." & vbCrLf
    Else
        detail = detail & "OK   : TypoTemplate ditemukan." & vbCrLf
        If Trim$(dbName) = "" Then
            dbName = H3C_ReadDBNameFromKeyValueFile(tplPath, False)
        End If
    End If

    dbName = H3C_NormalizeDBFileName(dbName)

    If Trim$(dbName) = "" Then
        ok = False
        detail = detail & "FAIL : Nama SizeDB tidak ditemukan dari @SIZEDB Order.txt maupun template." & vbCrLf
        detail = detail & "       Tambahkan @SIZEDB=SizeDB_....txt pada Order.txt agar report berjalan otomatis." & vbCrLf
    Else
        dbPath = Environ$("USERPROFILE") & "\Documents\" & dbName

        If Dir$(dbPath) = "" Then
            ok = False
            detail = detail & "FAIL : SizeDB tidak ditemukan: " & dbPath & vbCrLf
        Else
            detail = detail & "OK   : SizeDB ditemukan: " & dbName & vbCrLf
        End If
    End If

    If ok Then
        H3C_RunPreflightAutoReport = True
        HADESR_AddResult _
            "PREFLIGHT DATA CHECK", _
            "PASS", _
            "Order.txt, TypoTemplate, dan SizeDB siap.", _
            detail
    Else
        H3C_RunPreflightAutoReport = False
        HADESR_AddResult _
            "PREFLIGHT DATA CHECK", _
            "FAIL", _
            "File dasar Project Hades belum lengkap.", _
            detail
    End If

End Function

Private Function H3C_ReadDBNameFromKeyValueFile(ByVal path As String, ByVal isOrderFile As Boolean) As String

    Dim txt As String
    Dim lines() As String
    Dim i As Long
    Dim ln As String
    Dim p As Long
    Dim k As String
    Dim v As String

    txt = H3C_ReadTextFileUTF8(path)
    txt = Replace(txt, vbCrLf, vbLf)
    txt = Replace(txt, vbCr, vbLf)
    lines = Split(txt, vbLf)

    For i = LBound(lines) To UBound(lines)

        ln = Trim$(CStr(lines(i)))
        If ln = "" Then GoTo NextLine

        If isOrderFile Then
            If Left$(ln, 1) <> "@" Then GoTo NextLine
            ln = Mid$(ln, 2)
        End If

        p = InStr(1, ln, "=", vbTextCompare)
        If p <= 0 Then GoTo NextLine

        k = UCase$(Trim$(Left$(ln, p - 1)))
        v = Trim$(Mid$(ln, p + 1))

        Select Case k
            Case "SIZEDB", "DB"
                H3C_ReadDBNameFromKeyValueFile = v
                Exit Function
        End Select

NextLine:

    Next i

End Function

Private Function H3C_NormalizeDBFileName(ByVal dbName As String) As String

    dbName = Trim$(dbName)

    If dbName = "" Then
        H3C_NormalizeDBFileName = ""
        Exit Function
    End If

    If InStr(1, UCase$(dbName), ".TXT", vbTextCompare) = 0 Then
        dbName = dbName & ".txt"
    End If

    H3C_NormalizeDBFileName = dbName

End Function

'=========================================================
' AUTO BRIDGE: TYPO REPORT MODE - PHASE 3C
'=========================================================
Private Sub H3C_RunTypoCheckAutoReport()

    On Error GoTo FAIL_RUN

    Call HADES_QC_TYPO_REPORT

    Exit Sub

FAIL_RUN:

    HADESR_AddResult _
        "TYPO CHECK", _
        "FAIL", _
        "QC Typo report-mode error saat dipanggil.", _
        "Pastikan module VBA_QC_TYPO_CHECK_V12_1_REPORT_MODE.bas sudah menggantikan QC_TYPO_CHECK lama." & vbCrLf & _
        "Error " & Err.Number & ": " & Err.Description

End Sub


'=========================================================
' AUTO BRIDGE: SIZE REPORT MODE - PHASE 3C
'=========================================================
Private Sub H3C_RunSizeCheckAutoReport()

    On Error GoTo FAIL_RUN

    Call HADES_QC_SIZE_REPORT

    Exit Sub

FAIL_RUN:

    HADESR_AddResult _
        "SIZE & QUANTITY CHECK", _
        "FAIL", _
        "QC Size report-mode error saat dipanggil.", _
        "Pastikan module VBA_QC_SIZE_CHECK_V8_2_REPORT_MODE.bas sudah menggantikan QC_SIZE_CHECK lama." & vbCrLf & _
        "Error " & Err.Number & ": " & Err.Description

End Sub


'=========================================================
' MANUAL BRIDGE: TYPO ONLY; SIZE MANUAL KEPT AS FALLBACK
'=========================================================
Private Sub H3C_RunSizeCheckManual()

    Dim st As String

    On Error GoTo FAIL_RUN

    Call QC_SIZE_CHECK

    st = H3C_AskStatus( _
        "SIZE & QUANTITY CHECK", _
        "Lihat hasil MsgBox QC_SIZE_CHECK tadi." & vbCrLf & vbCrLf & _
        "Ketik:" & vbCrLf & _
        "P = PASS" & vbCrLf & _
        "W = WARNING" & vbCrLf & _
        "F = FAIL")

    HADESR_AddResult _
        "SIZE & QUANTITY CHECK", _
        st, _
        "Status dikonfirmasi operator setelah QC_SIZE_CHECK.", _
        "Phase 2C belum mengambil status QC_SIZE_CHECK secara native. " & _
        "Phase berikutnya perlu menambahkan QC_SIZE_CHECK_REPORT_MODE."

    Exit Sub

FAIL_RUN:

    HADESR_AddResult _
        "SIZE & QUANTITY CHECK", _
        "FAIL", _
        "QC_SIZE_CHECK error saat dipanggil.", _
        "Error " & Err.Number & ": " & Err.Description

End Sub

Private Sub H3C_RunTypoCheckManual()

    Dim st As String

    On Error GoTo FAIL_RUN

    Call QC_TYPO_CHECK

    st = H3C_AskStatus( _
        "TYPO CHECK", _
        "Lihat hasil MsgBox QC_TYPO_CHECK tadi." & vbCrLf & vbCrLf & _
        "Ketik:" & vbCrLf & _
        "P = PASS" & vbCrLf & _
        "W = WARNING" & vbCrLf & _
        "F = FAIL")

    HADESR_AddResult _
        "TYPO CHECK", _
        st, _
        "Status dikonfirmasi operator setelah QC_TYPO_CHECK.", _
        "Phase 2C belum mengambil status QC_TYPO_CHECK secara native. " & _
        "Phase berikutnya perlu menambahkan QC_TYPO_CHECK_REPORT_MODE."

    Exit Sub

FAIL_RUN:

    HADESR_AddResult _
        "TYPO CHECK", _
        "FAIL", _
        "QC_TYPO_CHECK error saat dipanggil.", _
        "Error " & Err.Number & ": " & Err.Description

End Sub

Private Function H3C_AskStatus(ByVal titleText As String, ByVal promptText As String) As String

    Dim v As String

    Do
        v = UCase$(Trim$(InputBox(promptText, "HADES STATUS — " & titleText, "P")))

        If v = "P" Or v = "PASS" Then
            H3C_AskStatus = "PASS"
            Exit Function
        End If

        If v = "W" Or v = "WARN" Or v = "WARNING" Then
            H3C_AskStatus = "WARNING"
            Exit Function
        End If

        If v = "F" Or v = "FAIL" Then
            H3C_AskStatus = "FAIL"
            Exit Function
        End If

        If v = "" Then
            H3C_AskStatus = "FAIL"
            Exit Function
        End If

        MsgBox "Input tidak valid. Gunakan P, W, atau F.", vbExclamation, "HADES STATUS"

    Loop

End Function

'=========================================================
' IDPO AUTO REPORT MODE
'=========================================================
Private Sub H3C_RunIDPOAutoReport()

    Dim targetPO As String
    Dim foundTarget As Long
    Dim foundPlaceholder As Long
    Dim otherDict As Object
    Dim detail As String
    Dim statusText As String
    Dim summaryText As String
    Dim oldUnit As cdrUnit

    On Error GoTo FAIL_RUN

    targetPO = H3C_LoadIDPOFromOrderTxt()

    If Not H3C_IsSixDigit(targetPO) Then
        HADESR_AddResult _
            "IDPO CHECK", _
            "FAIL", _
            "@IDPO tidak terbaca dari Order.txt.", _
            "Pastikan Python ROBOT-PO menghasilkan metadata @IDPO=xxxxxx di Documents\Order.txt."
        Exit Sub
    End If

    Set otherDict = CreateObject("Scripting.Dictionary")

    oldUnit = ActiveDocument.Unit
    ActiveDocument.Unit = cdrCentimeter

    H3C_ScanIDPOSelection ActiveSelectionRange, targetPO, foundTarget, foundPlaceholder, otherDict

    ActiveDocument.Unit = oldUnit

    detail = "Target IDPO : " & targetPO & vbCrLf & _
             "Target found : " & foundTarget & vbCrLf & _
             "Placeholder  : " & foundPlaceholder & vbCrLf & _
             "Other IDPO   : " & otherDict.Count & vbCrLf

    If otherDict.Count > 0 Then
        detail = detail & "Other list   : " & H3C_JoinDictKeys(otherDict) & vbCrLf
    End If

    If foundTarget > 0 And foundPlaceholder = 0 And otherDict.Count = 0 Then
        statusText = "PASS"
        summaryText = "Target IDPO ditemukan dan tidak ada placeholder/IDPO lama."
    Else
        statusText = "FAIL"
        summaryText = "IDPO target hilang, placeholder masih ada, atau ditemukan IDPO lain."
    End If

    HADESR_AddResult "IDPO CHECK", statusText, summaryText, detail

    Exit Sub

FAIL_RUN:

    On Error Resume Next
    ActiveDocument.Unit = oldUnit

    HADESR_AddResult _
        "IDPO CHECK", _
        "FAIL", _
        "IDPO auto report error.", _
        "Error " & Err.Number & ": " & Err.Description

End Sub

Private Sub H3C_ScanIDPOSelection( _
    ByVal sr As ShapeRange, _
    ByVal targetPO As String, _
    ByRef foundTarget As Long, _
    ByRef foundPlaceholder As Long, _
    ByVal otherDict As Object)

    Dim s As Shape

    foundTarget = 0
    foundPlaceholder = 0

    For Each s In sr
        H3C_ScanIDPOShape s, targetPO, foundTarget, foundPlaceholder, otherDict
    Next s

End Sub

Private Sub H3C_ScanIDPOShape( _
    ByVal s As Shape, _
    ByVal targetPO As String, _
    ByRef foundTarget As Long, _
    ByRef foundPlaceholder As Long, _
    ByVal otherDict As Object)

    Dim ch As Shape
    Dim pcShapes As Shapes
    Dim txt As String

    On Error Resume Next

    If s Is Nothing Then Exit Sub

    If s.Type = cdrTextShape Then

        txt = H3C_NormalizeText(s.Text.Story.Text)

        If H3C_IsIDPOHeight(s) Then

            If txt = "IDPO" Then
                foundPlaceholder = foundPlaceholder + 1
            ElseIf H3C_IsSixDigit(txt) Then
                If txt = targetPO Then
                    foundTarget = foundTarget + 1
                Else
                    If Not otherDict.Exists(txt) Then
                        otherDict.Add txt, 1
                    Else
                        otherDict(txt) = CLng(otherDict(txt)) + 1
                    End If
                End If
            End If

        End If

    End If

    If s.Type = cdrGroupShape Then
        For Each ch In s.Shapes
            H3C_ScanIDPOShape ch, targetPO, foundTarget, foundPlaceholder, otherDict
        Next ch
    End If

    Set pcShapes = Nothing
    Set pcShapes = s.PowerClip.Shapes

    If Not pcShapes Is Nothing Then
        For Each ch In pcShapes
            H3C_ScanIDPOShape ch, targetPO, foundTarget, foundPlaceholder, otherDict
        Next ch
    End If

    On Error GoTo 0

End Sub

Private Function H3C_IsIDPOHeight(ByVal shp As Shape) As Boolean

    Dim h As Double
    Dim w As Double
    Dim smallSide As Double

    On Error GoTo FAIL

    h = Abs(shp.SizeHeight)
    w = Abs(shp.SizeWidth)

    If h < w Then
        smallSide = h
    Else
        smallSide = w
    End If

    H3C_IsIDPOHeight = (smallSide >= H3C_ID_MIN_H And smallSide <= H3C_ID_MAX_H)
    Exit Function

FAIL:
    H3C_IsIDPOHeight = False

End Function

Private Function H3C_LoadIDPOFromOrderTxt() As String

    Dim path As String
    Dim allText As String
    Dim lines As Variant
    Dim i As Long
    Dim line As String
    Dim p As Long
    Dim key As String
    Dim val As String

    On Error GoTo FAIL

    path = Environ$("USERPROFILE") & H3C_ORDER_FILE
    If Dir$(path) = "" Then Exit Function

    allText = H3C_ReadTextFileUTF8(path)
    allText = Replace(allText, vbCrLf, vbLf)
    allText = Replace(allText, vbCr, vbLf)

    lines = Split(allText, vbLf)

    For i = LBound(lines) To UBound(lines)

        line = H3C_RemoveBOM(CStr(lines(i)))
        line = Trim$(line)

        If Len(line) = 0 Then GoTo NextLine
        If Left$(line, 1) <> "@" Then GoTo NextLine

        p = InStr(1, line, "=", vbTextCompare)
        If p <= 2 Then GoTo NextLine

        key = UCase$(Trim$(Mid$(line, 2, p - 2)))
        val = H3C_NormalizeText(Mid$(line, p + 1))

        If key = "IDPO" Or key = "KODE_PRODUK" Or key = "KODE PRODUK" Then
            If H3C_IsSixDigit(val) Then
                H3C_LoadIDPOFromOrderTxt = val
                Exit Function
            End If
        End If

NextLine:
    Next i

    Exit Function

FAIL:
    H3C_LoadIDPOFromOrderTxt = ""

End Function

'=========================================================
' GROUP STRUCTURE AUTO REPORT MODE
'=========================================================
Private Sub H3C_RunGroupStructureAutoReport()

    Dim topTotal As Long
    Dim topGroup As Long
    Dim topLoose As Long
    Dim setChecked As Long
    Dim setFailed As Long
    Dim setWarning As Long
    Dim looseInside As Long
    Dim detail As String
    Dim statusText As String
    Dim summaryText As String

    On Error GoTo FAIL_RUN

    H3C_GroupScanSelection _
        ActiveSelectionRange, _
        topTotal, _
        topGroup, _
        topLoose, _
        setChecked, _
        setFailed, _
        setWarning, _
        looseInside, _
        detail

    If topLoose > 0 Or setFailed > 0 Then
        statusText = "FAIL"
        summaryText = "Ada top object bukan group atau child langsung yang belum digroup."
    ElseIf setWarning > 0 Then
        statusText = "WARNING"
        summaryText = "Struktur group lolos fatal check, tetapi ada warning child group terlalu sedikit."
    Else
        statusText = "PASS"
        summaryText = "Struktur group valid."
    End If

    detail = "Top total       : " & topTotal & vbCrLf & _
             "Top group       : " & topGroup & vbCrLf & _
             "Top loose       : " & topLoose & vbCrLf & _
             "Set checked     : " & setChecked & vbCrLf & _
             "Set failed      : " & setFailed & vbCrLf & _
             "Set warning     : " & setWarning & vbCrLf & _
             "Loose inside    : " & looseInside & vbCrLf & vbCrLf & _
             detail

    HADESR_AddResult "GROUP STRUCTURE CHECK", statusText, summaryText, detail

    Exit Sub

FAIL_RUN:

    HADESR_AddResult _
        "GROUP STRUCTURE CHECK", _
        "FAIL", _
        "Group structure auto report error.", _
        "Error " & Err.Number & ": " & Err.Description

End Sub

Private Sub H3C_GroupScanSelection( _
    ByVal sr As ShapeRange, _
    ByRef topTotal As Long, _
    ByRef topGroup As Long, _
    ByRef topLoose As Long, _
    ByRef setChecked As Long, _
    ByRef setFailed As Long, _
    ByRef setWarning As Long, _
    ByRef looseInside As Long, _
    ByRef detail As String)

    Dim s As Shape
    Dim topIndex As Long
    Dim groupIndex As Long

    topTotal = 0
    topGroup = 0
    topLoose = 0
    setChecked = 0
    setFailed = 0
    setWarning = 0
    looseInside = 0
    detail = ""

    For Each s In sr

        topIndex = topIndex + 1
        topTotal = topTotal + 1

        If s.Type = cdrGroupShape Then
            topGroup = topGroup + 1
            groupIndex = groupIndex + 1
            H3C_GroupCheckMain s, groupIndex, setChecked, setFailed, setWarning, looseInside, detail
        Else
            topLoose = topLoose + 1
            detail = detail & "- TOP OBJECT #" & topIndex & " bukan GROUP | Type=" & H3C_ShapeTypeName(s) & " | " & H3C_SizeInfo(s) & vbCrLf
        End If

    Next s

End Sub

Private Sub H3C_GroupCheckMain( _
    ByVal g As Shape, _
    ByVal groupIndex As Long, _
    ByRef setChecked As Long, _
    ByRef setFailed As Long, _
    ByRef setWarning As Long, _
    ByRef looseInside As Long, _
    ByRef detail As String)

    Dim child As Shape
    Dim childTotal As Long
    Dim childGroup As Long
    Dim childLoose As Long
    Dim localDetail As String

    setChecked = setChecked + 1

    For Each child In g.Shapes
        childTotal = childTotal + 1
        If child.Type = cdrGroupShape Then
            childGroup = childGroup + 1
        Else
            childLoose = childLoose + 1
            localDetail = localDetail & "    - Child #" & childTotal & " | Type=" & H3C_ShapeTypeName(child) & " | " & H3C_SizeInfo(child) & vbCrLf
        End If
    Next child

    If childLoose > 0 Then
        setFailed = setFailed + 1
        looseInside = looseInside + childLoose
        detail = detail & "GROUP SET #" & groupIndex & " punya " & childLoose & " objek langsung yang bukan group." & vbCrLf & _
                 "Child Group : " & childGroup & vbCrLf & _
                 "Child Total : " & childTotal & vbCrLf & _
                 localDetail & String(45, "-") & vbCrLf
    ElseIf childGroup < H3C_MIN_CHILD_GROUP Then
        setWarning = setWarning + 1
        detail = detail & "GROUP SET #" & groupIndex & " hanya punya " & childGroup & " child group." & vbCrLf & _
                 "Kemungkinan yang dipilih bukan group 1 set, atau struktur group terlalu sederhana." & vbCrLf & _
                 String(45, "-") & vbCrLf
    End If

End Sub

'=========================================================
' TRANSPARENCY + POWERCLIP AUTO REPORT MODE
'=========================================================
Private Sub H3C_RunTransparencyPowerClipAutoReport()

    Dim scanned As Long
    Dim powerClipCount As Long
    Dim trSafe As Long
    Dim trWarn As Long
    Dim trFail As Long
    Dim trUnknown As Long
    Dim detailCount As Long
    Dim detailSkipped As Long
    Dim detail As String
    Dim statusText As String
    Dim summaryText As String

    On Error GoTo FAIL_RUN

    H3C_QTPScanSelection _
        ActiveSelectionRange, _
        scanned, _
        powerClipCount, _
        trSafe, _
        trWarn, _
        trFail, _
        trUnknown, _
        detailCount, _
        detailSkipped, _
        detail

    If powerClipCount > 0 Or trFail > 0 Or (H3C_UNKNOWN_TR_IS_FAIL And trUnknown > 0) Then
        statusText = "FAIL"
        summaryText = "PowerClip ditemukan atau Transparency >75%."
    ElseIf trWarn > 0 Or trUnknown > 0 Then
        statusText = "WARNING"
        summaryText = "Ada transparency warning/unknown yang perlu cek manual."
    Else
        statusText = "PASS"
        summaryText = "Tidak ditemukan PowerClip dan tidak ada transparency berisiko."
    End If

    detail = "Scanned shape              : " & scanned & vbCrLf & _
             "PowerClip found            : " & powerClipCount & vbCrLf & _
             "Transparency safe 0-70     : " & trSafe & vbCrLf & _
             "Transparency warning 71-75 : " & trWarn & vbCrLf & _
             "Transparency FAIL >75      : " & trFail & vbCrLf & _
             "Transparency unknown       : " & trUnknown & vbCrLf & _
             "Detail skipped             : " & detailSkipped & vbCrLf & vbCrLf & _
             detail

    HADESR_AddResult "POWERCLIP / TRANSPARENCY", statusText, summaryText, detail

    Exit Sub

FAIL_RUN:

    HADESR_AddResult _
        "POWERCLIP / TRANSPARENCY", _
        "FAIL", _
        "Transparency/PowerClip auto report error.", _
        "Error " & Err.Number & ": " & Err.Description

End Sub

Private Sub H3C_QTPScanSelection( _
    ByVal sr As ShapeRange, _
    ByRef scanned As Long, _
    ByRef powerClipCount As Long, _
    ByRef trSafe As Long, _
    ByRef trWarn As Long, _
    ByRef trFail As Long, _
    ByRef trUnknown As Long, _
    ByRef detailCount As Long, _
    ByRef detailSkipped As Long, _
    ByRef detail As String)

    Dim s As Shape
    Dim i As Long

    scanned = 0
    powerClipCount = 0
    trSafe = 0
    trWarn = 0
    trFail = 0
    trUnknown = 0
    detailCount = 0
    detailSkipped = 0
    detail = ""

    i = 0
    For Each s In sr
        i = i + 1
        H3C_QTPScanShape s, "ROOT#" & CStr(i), scanned, powerClipCount, trSafe, trWarn, trFail, trUnknown, detailCount, detailSkipped, detail
    Next s

End Sub

Private Sub H3C_QTPScanShape( _
    ByVal s As Shape, _
    ByVal pathLabel As String, _
    ByRef scanned As Long, _
    ByRef powerClipCount As Long, _
    ByRef trSafe As Long, _
    ByRef trWarn As Long, _
    ByRef trFail As Long, _
    ByRef trUnknown As Long, _
    ByRef detailCount As Long, _
    ByRef detailSkipped As Long, _
    ByRef detail As String)

    Dim ch As Shape
    Dim pcShapes As Shapes

    On Error Resume Next

    If s Is Nothing Then Exit Sub

    scanned = scanned + 1

    H3C_QTPCheckTransparency s, pathLabel, trSafe, trWarn, trFail, trUnknown, detailCount, detailSkipped, detail

    If H3C_GetPowerClipShapes(s, pcShapes) Then
        powerClipCount = powerClipCount + 1
        H3C_AddLimitedDetail detail, detailCount, detailSkipped, "[POWERCLIP] " & pathLabel & " | " & H3C_ShapeInfo(s)

        For Each ch In pcShapes
            H3C_QTPScanShape ch, pathLabel & " > POWERCLIP", scanned, powerClipCount, trSafe, trWarn, trFail, trUnknown, detailCount, detailSkipped, detail
        Next ch
    End If

    If s.Type = cdrGroupShape Then
        For Each ch In s.Shapes
            H3C_QTPScanShape ch, pathLabel & " > GROUP", scanned, powerClipCount, trSafe, trWarn, trFail, trUnknown, detailCount, detailSkipped, detail
        Next ch
    End If

    On Error GoTo 0

End Sub

Private Sub H3C_QTPCheckTransparency( _
    ByVal s As Shape, _
    ByVal pathLabel As String, _
    ByRef trSafe As Long, _
    ByRef trWarn As Long, _
    ByRef trFail As Long, _
    ByRef trUnknown As Long, _
    ByRef detailCount As Long, _
    ByRef detailSkipped As Long, _
    ByRef detail As String)

    Dim hasTrans As Boolean
    Dim valueKnown As Boolean
    Dim pct As Double
    Dim desc As String

    If Not H3C_GetTransparencyInfo(s, hasTrans, valueKnown, pct, desc) Then Exit Sub
    If Not hasTrans Then Exit Sub

    If valueKnown Then
        If pct > H3C_TR_FAIL_OVER Then
            trFail = trFail + 1
            H3C_AddLimitedDetail detail, detailCount, detailSkipped, "[TR FAIL] " & pathLabel & " | Transparency=" & FormatNumber(pct, 1) & "% | " & desc & " | " & H3C_ShapeInfo(s)
        ElseIf pct >= H3C_TR_WARN_FROM And pct <= H3C_TR_WARN_TO Then
            trWarn = trWarn + 1
            H3C_AddLimitedDetail detail, detailCount, detailSkipped, "[TR WARNING] " & pathLabel & " | Transparency=" & FormatNumber(pct, 1) & "% | " & desc & " | " & H3C_ShapeInfo(s)
        Else
            trSafe = trSafe + 1
        End If
    Else
        trUnknown = trUnknown + 1
        H3C_AddLimitedDetail detail, detailCount, detailSkipped, "[TR UNKNOWN] " & pathLabel & " | Transparency ada tetapi nilai tidak terbaca | " & desc & " | " & H3C_ShapeInfo(s)
    End If

End Sub

Private Function H3C_GetPowerClipShapes(ByVal s As Shape, ByRef pcShapes As Shapes) As Boolean

    On Error GoTo NO_POWERCLIP

    Set pcShapes = s.PowerClip.Shapes

    If Not pcShapes Is Nothing Then
        If pcShapes.Count > 0 Then
            H3C_GetPowerClipShapes = True
            Exit Function
        End If
    End If

NO_POWERCLIP:
    H3C_GetPowerClipShapes = False

End Function

Private Function H3C_GetTransparencyInfo( _
    ByVal s As Shape, _
    ByRef hasTrans As Boolean, _
    ByRef valueKnown As Boolean, _
    ByRef pct As Double, _
    ByRef desc As String) As Boolean

    Dim tr As Object
    Dim typeRaw As Variant
    Dim typeKnown As Boolean
    Dim maxPct As Double
    Dim propDesc As String
    Dim valueFound As Boolean

    On Error GoTo SAFE_EXIT

    hasTrans = False
    valueKnown = False
    pct = 0
    desc = ""

    Set tr = s.Transparency

    If tr Is Nothing Then
        H3C_GetTransparencyInfo = False
        Exit Function
    End If

    typeKnown = H3C_TryGetProperty(tr, "Type", typeRaw)

    If typeKnown Then
        desc = "Type=" & H3C_TransparencyTypeName(CLng(typeRaw))

        If CLng(typeRaw) = 0 Then
            hasTrans = False
            valueKnown = True
            pct = 0
            H3C_GetTransparencyInfo = True
            Exit Function
        End If

        hasTrans = True
    End If

    valueFound = H3C_TryGetMaxTransparencyValue(tr, maxPct, propDesc)

    If valueFound Then
        pct = maxPct
        valueKnown = True

        If pct > 0 Then hasTrans = True

        If desc <> "" And propDesc <> "" Then
            desc = desc & " | " & propDesc
        ElseIf propDesc <> "" Then
            desc = propDesc
        End If

        H3C_GetTransparencyInfo = True
        Exit Function
    End If

    If hasTrans Then
        valueKnown = False
        H3C_GetTransparencyInfo = True
        Exit Function
    End If

    H3C_GetTransparencyInfo = False
    Exit Function

SAFE_EXIT:
    H3C_GetTransparencyInfo = False

End Function

Private Function H3C_TryGetMaxTransparencyValue(ByVal tr As Object, ByRef maxPct As Double, ByRef propDesc As String) As Boolean

    Dim props As Variant
    Dim i As Long
    Dim raw As Variant
    Dim pct As Double
    Dim found As Boolean
    Dim detail As String

    On Error Resume Next

    props = Array( _
        "UniformTransparency", _
        "Uniform", _
        "Transparency", _
        "TransparencyValue", _
        "Amount", _
        "StartTransparency", _
        "EndTransparency", _
        "MidPointTransparency", _
        "FountainStartTransparency", _
        "FountainEndTransparency" _
    )

    maxPct = 0
    found = False
    detail = ""

    For i = LBound(props) To UBound(props)
        If H3C_TryGetProperty(tr, CStr(props(i)), raw) Then
            If H3C_NormalizeTransparencyPercent(raw, pct) Then
                found = True
                If pct > maxPct Then maxPct = pct
                If detail <> "" Then detail = detail & ", "
                detail = detail & CStr(props(i)) & "=" & FormatNumber(pct, 1) & "%"
            End If
        End If
    Next i

    propDesc = detail
    H3C_TryGetMaxTransparencyValue = found

End Function

Private Function H3C_TryGetProperty(ByVal obj As Object, ByVal propName As String, ByRef result As Variant) As Boolean

    On Error GoTo FAIL_PROP

    result = CallByName(obj, propName, VbGet)
    H3C_TryGetProperty = True
    Exit Function

FAIL_PROP:
    Err.Clear
    H3C_TryGetProperty = False

End Function

Private Function H3C_NormalizeTransparencyPercent(ByVal raw As Variant, ByRef pct As Double) As Boolean

    On Error GoTo FAIL_VALUE

    If Not IsNumeric(raw) Then
        H3C_NormalizeTransparencyPercent = False
        Exit Function
    End If

    pct = CDbl(raw)

    If pct > 0 And pct <= 1 Then pct = pct * 100
    If pct < 0 Then GoTo FAIL_VALUE
    If pct > 100 Then GoTo FAIL_VALUE

    H3C_NormalizeTransparencyPercent = True
    Exit Function

FAIL_VALUE:
    H3C_NormalizeTransparencyPercent = False

End Function

Private Function H3C_TransparencyTypeName(ByVal t As Long) As String

    Select Case t
        Case 0: H3C_TransparencyTypeName = "None"
        Case 1: H3C_TransparencyTypeName = "Uniform / Type#1"
        Case 2: H3C_TransparencyTypeName = "Fountain / Type#2"
        Case 3: H3C_TransparencyTypeName = "Pattern / Type#3"
        Case 4: H3C_TransparencyTypeName = "Texture / Type#4"
        Case Else: H3C_TransparencyTypeName = "Type#" & CStr(t)
    End Select

End Function

'=========================================================
' COMMON HELPERS
'=========================================================
Private Function H3C_HasSelection() As Boolean

    On Error GoTo FAIL

    H3C_HasSelection = False

    If ActiveSelection Is Nothing Then Exit Function
    If ActiveSelection.Shapes.Count = 0 Then Exit Function

    H3C_HasSelection = True
    Exit Function

FAIL:
    H3C_HasSelection = False

End Function

Private Function H3C_ReadTextFileUTF8(ByVal path As String) As String

    Dim stm As Object

    On Error GoTo FALLBACK

    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 2
    stm.Charset = "utf-8"
    stm.Open
    stm.LoadFromFile path
    H3C_ReadTextFileUTF8 = stm.ReadText
    stm.Close
    Exit Function

FALLBACK:
    H3C_ReadTextFileUTF8 = H3C_ReadTextFileANSI(path)

End Function

Private Function H3C_ReadTextFileANSI(ByVal path As String) As String

    Dim f As Integer
    Dim txt As String

    On Error GoTo FAIL

    f = FreeFile
    Open path For Input As #f
    txt = Input$(LOF(f), f)
    Close #f

    H3C_ReadTextFileANSI = txt
    Exit Function

FAIL:
    On Error Resume Next
    Close #f
    H3C_ReadTextFileANSI = ""

End Function

Private Function H3C_RemoveBOM(ByVal s As String) As String

    s = Replace(s, ChrW$(&HFEFF), "")
    s = Replace(s, "ï»¿", "")
    H3C_RemoveBOM = s

End Function

Private Function H3C_NormalizeText(ByVal s As String) As String

    s = Replace(s, vbCr, "")
    s = Replace(s, vbLf, "")
    s = Replace(s, Chr$(160), " ")

    Do While InStr(1, s, "  ", vbTextCompare) > 0
        s = Replace(s, "  ", " ")
    Loop

    H3C_NormalizeText = UCase$(Trim$(s))

End Function

Private Function H3C_IsSixDigit(ByVal s As String) As Boolean

    s = Trim$(s)
    H3C_IsSixDigit = (Len(s) = 6 And IsNumeric(s))

End Function

Private Function H3C_JoinDictKeys(ByVal d As Object) As String

    Dim k As Variant
    Dim r As String

    For Each k In d.Keys
        If r <> "" Then r = r & ", "
        r = r & CStr(k) & "(" & CStr(d(k)) & "x)"
    Next k

    H3C_JoinDictKeys = r

End Function

Private Sub H3C_AddLimitedDetail(ByRef detail As String, ByRef detailCount As Long, ByRef detailSkipped As Long, ByVal lineText As String)

    If detailCount < H3C_DETAIL_LIMIT Then
        detail = detail & "- " & lineText & vbCrLf
        detailCount = detailCount + 1
    Else
        detailSkipped = detailSkipped + 1
    End If

End Sub

Private Function H3C_ShapeInfo(ByVal s As Shape) As String

    Dim nm As String

    On Error Resume Next

    nm = s.Name
    If Trim$(nm) = "" Then nm = "-"

    H3C_ShapeInfo = _
        "Shape=" & H3C_ShapeTypeName(s) & _
        " | Name=" & nm & _
        " | X=" & FormatNumber(s.CenterX, 2) & _
        " | Y=" & FormatNumber(s.CenterY, 2) & _
        " | W=" & FormatNumber(Abs(s.SizeWidth), 2) & _
        " | H=" & FormatNumber(Abs(s.SizeHeight), 2)

End Function

Private Function H3C_ShapeTypeName(ByVal s As Shape) As String

    On Error GoTo UNKNOWN_TYPE

    Select Case s.Type
        Case cdrGroupShape: H3C_ShapeTypeName = "GROUP"
        Case cdrCurveShape: H3C_ShapeTypeName = "CURVE"
        Case cdrTextShape: H3C_ShapeTypeName = "TEXT"
        Case cdrBitmapShape: H3C_ShapeTypeName = "BITMAP"
        Case cdrRectangleShape: H3C_ShapeTypeName = "RECTANGLE"
        Case cdrEllipseShape: H3C_ShapeTypeName = "ELLIPSE"
        Case cdrPolygonShape: H3C_ShapeTypeName = "POLYGON"
        Case Else: H3C_ShapeTypeName = "TYPE#" & CStr(s.Type)
    End Select

    Exit Function

UNKNOWN_TYPE:
    H3C_ShapeTypeName = "UNKNOWN"

End Function

Private Function H3C_SizeInfo(ByVal s As Shape) As String

    On Error Resume Next

    H3C_SizeInfo = _
        "W=" & FormatNumber(Abs(s.SizeWidth), 2) & _
        " | H=" & FormatNumber(Abs(s.SizeHeight), 2) & _
        " | X=" & FormatNumber(s.CenterX, 2) & _
        " | Y=" & FormatNumber(s.CenterY, 2)

End Function
