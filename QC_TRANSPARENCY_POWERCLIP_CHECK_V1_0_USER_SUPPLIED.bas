Option Explicit

'=========================================================
' PROJECT HADES — QC TRANSPARENCY + POWERCLIP CHECK V1.0
' CorelDRAW 2021 VBA
'
' MAIN MACRO:
' QC_TRANSPARENCY_POWERCLIP_CHECK
'
' TUJUAN:
' - Preflight reject check sebelum desain masuk proses layout/QC.
' - Mendeteksi PowerClip dan Transparency berisiko.
'
' ATURAN:
' - PowerClip ditemukan        = FAIL / REJECT
' - Transparency > 75%         = FAIL / REJECT
' - Transparency 71% - 75%     = WARNING
' - Transparency 0% - 70%      = PASS
' - Transparency ada tapi nilai tidak terbaca = WARNING
'
' CATATAN PENTING:
' - Yang dicek adalah nilai TRANSPARENCY CorelDRAW.
' - BUKAN opacity.
'=========================================================

Private Const QTP_TRANSPARENCY_FAIL_OVER As Double = 75#
Private Const QTP_TRANSPARENCY_WARN_FROM As Double = 71#
Private Const QTP_TRANSPARENCY_WARN_TO As Double = 75#
Private Const QTP_UNKNOWN_TRANSPARENCY_IS_FAIL As Boolean = False
Private Const QTP_DETAIL_LIMIT As Long = 200

Private qtpScanned As Long
Private qtpPowerClip As Long
Private qtpTransparencySafe As Long
Private qtpTransparencyWarning As Long
Private qtpTransparencyFail As Long
Private qtpTransparencyUnknown As Long
Private qtpDetailCount As Long
Private qtpDetailSkipped As Long
Private qtpPowerClipDetail As String
Private qtpTransparencyFailDetail As String
Private qtpTransparencyWarningDetail As String
Private qtpTransparencyUnknownDetail As String
Private qtpReportPath As String

Sub QC_TRANSPARENCY_POWERCLIP_CHECK()
    Dim oldUnit As cdrUnit
    Dim oldOptimization As Boolean
    Dim sr As ShapeRange
    Dim i As Long

    On Error GoTo ERR_HANDLER

    If ActiveSelection Is Nothing Then
        MsgBox "Tidak ada objek dipilih.", vbExclamation, "QC TRANSPARENCY POWERCLIP"
        Exit Sub
    End If

    If ActiveSelection.Shapes.Count = 0 Then
        MsgBox "Tidak ada objek dipilih.", vbExclamation, "QC TRANSPARENCY POWERCLIP"
        Exit Sub
    End If

    QTP_Reset

    oldUnit = ActiveDocument.Unit
    oldOptimization = Application.Optimization

    ActiveDocument.Unit = cdrCentimeter
    Application.Optimization = True

    Set sr = ActiveSelectionRange

    For i = 1 To sr.Count
        QTP_ScanShape sr(i), "ROOT#" & CStr(i)
    Next i

    Application.Optimization = oldOptimization
    ActiveDocument.Unit = oldUnit
    ActiveWindow.Refresh

    qtpReportPath = QTP_WriteReportFile
    QTP_ShowFinalResult
    Exit Sub

ERR_HANDLER:
    On Error Resume Next
    Application.Optimization = oldOptimization
    ActiveDocument.Unit = oldUnit
    ActiveWindow.Refresh
    MsgBox "SYSTEM ERROR - QC TRANSPARENCY POWERCLIP" & vbCrLf & vbCrLf & _
        "No : " & Err.Number & vbCrLf & Err.Description, vbCritical, "QC TRANSPARENCY POWERCLIP"
End Sub

Private Sub QTP_Reset()
    qtpScanned = 0: qtpPowerClip = 0
    qtpTransparencySafe = 0: qtpTransparencyWarning = 0: qtpTransparencyFail = 0: qtpTransparencyUnknown = 0
    qtpDetailCount = 0: qtpDetailSkipped = 0
    qtpPowerClipDetail = "": qtpTransparencyFailDetail = "": qtpTransparencyWarningDetail = "": qtpTransparencyUnknownDetail = ""
    qtpReportPath = ""
End Sub

Private Sub QTP_ScanShape(ByVal s As Shape, ByVal pathLabel As String)
    Dim ch As Shape
    Dim pcShapes As Shapes
    On Error Resume Next
    If s Is Nothing Then Exit Sub
    qtpScanned = qtpScanned + 1
    QTP_CheckTransparency s, pathLabel
    If QTP_GetPowerClipShapes(s, pcShapes) Then
        qtpPowerClip = qtpPowerClip + 1
        QTP_AddDetail qtpPowerClipDetail, pathLabel & " | " & QTP_ShapeInfo(s)
        For Each ch In pcShapes
            QTP_ScanShape ch, pathLabel & " > POWERCLIP"
        Next ch
    End If
    If s.Type = cdrGroupShape Then
        For Each ch In s.Shapes
            QTP_ScanShape ch, pathLabel & " > GROUP"
        Next ch
    End If
End Sub

Private Function QTP_GetPowerClipShapes(ByVal s As Shape, ByRef pcShapes As Shapes) As Boolean
    On Error GoTo NO_POWERCLIP
    Set pcShapes = s.PowerClip.Shapes
    If Not pcShapes Is Nothing Then
        If pcShapes.Count > 0 Then QTP_GetPowerClipShapes = True: Exit Function
    End If
NO_POWERCLIP:
    QTP_GetPowerClipShapes = False
End Function

Private Sub QTP_CheckTransparency(ByVal s As Shape, ByVal pathLabel As String)
    Dim hasTrans As Boolean, valueKnown As Boolean, pct As Double, desc As String
    On Error Resume Next
    If Not QTP_GetTransparencyInfo(s, hasTrans, valueKnown, pct, desc) Then Exit Sub
    If Not hasTrans Then Exit Sub
    If valueKnown Then
        If pct > QTP_TRANSPARENCY_FAIL_OVER Then
            qtpTransparencyFail = qtpTransparencyFail + 1
            QTP_AddDetail qtpTransparencyFailDetail, pathLabel & " | Transparency=" & FormatNumber(pct, 1) & "% | " & desc & " | " & QTP_ShapeInfo(s)
        ElseIf pct >= QTP_TRANSPARENCY_WARN_FROM And pct <= QTP_TRANSPARENCY_WARN_TO Then
            qtpTransparencyWarning = qtpTransparencyWarning + 1
            QTP_AddDetail qtpTransparencyWarningDetail, pathLabel & " | Transparency=" & FormatNumber(pct, 1) & "% | " & desc & " | " & QTP_ShapeInfo(s)
        Else
            qtpTransparencySafe = qtpTransparencySafe + 1
        End If
    Else
        qtpTransparencyUnknown = qtpTransparencyUnknown + 1
        QTP_AddDetail qtpTransparencyUnknownDetail, pathLabel & " | Transparency ada tetapi nilainya tidak terbaca | " & desc & " | " & QTP_ShapeInfo(s)
    End If
End Sub

Private Function QTP_GetTransparencyInfo(ByVal s As Shape, ByRef hasTrans As Boolean, ByRef valueKnown As Boolean, ByRef pct As Double, ByRef desc As String) As Boolean
    Dim tr As Object, typeRaw As Variant, typeKnown As Boolean, maxPct As Double, propDesc As String, valueFound As Boolean
    On Error GoTo SAFE_EXIT
    hasTrans = False: valueKnown = False: pct = 0: desc = ""
    Set tr = s.Transparency
    If tr Is Nothing Then QTP_GetTransparencyInfo = False: Exit Function
    typeKnown = QTP_TryGetProperty(tr, "Type", typeRaw)
    If typeKnown Then
        desc = "Type=" & QTP_TransparencyTypeName(CLng(typeRaw))
        If CLng(typeRaw) = 0 Then
            hasTrans = False: valueKnown = True: pct = 0: QTP_GetTransparencyInfo = True: Exit Function
        End If
        hasTrans = True
    End If
    valueFound = QTP_TryGetMaxTransparencyValue(tr, maxPct, propDesc)
    If valueFound Then
        pct = maxPct: valueKnown = True
        If pct > 0 Then hasTrans = True
        If desc <> "" And propDesc <> "" Then desc = desc & " | " & propDesc ElseIf propDesc <> "" Then desc = propDesc
        QTP_GetTransparencyInfo = True: Exit Function
    End If
    If hasTrans Then valueKnown = False: QTP_GetTransparencyInfo = True: Exit Function
    QTP_GetTransparencyInfo = False
    Exit Function
SAFE_EXIT:
    QTP_GetTransparencyInfo = False
End Function

Private Function QTP_TryGetMaxTransparencyValue(ByVal tr As Object, ByRef maxPct As Double, ByRef propDesc As String) As Boolean
    Dim props As Variant, i As Long, raw As Variant, pct As Double, found As Boolean, detail As String
    On Error Resume Next
    props = Array("UniformTransparency", "Uniform", "Transparency", "TransparencyValue", "Amount", "StartTransparency", "EndTransparency", "MidPointTransparency", "FountainStartTransparency", "FountainEndTransparency")
    maxPct = 0: found = False: detail = ""
    For i = LBound(props) To UBound(props)
        If QTP_TryGetProperty(tr, CStr(props(i)), raw) Then
            If QTP_NormalizeTransparencyPercent(raw, pct) Then
                found = True
                If pct > maxPct Then maxPct = pct
                If detail <> "" Then detail = detail & ", "
                detail = detail & CStr(props(i)) & "=" & FormatNumber(pct, 1) & "%"
            End If
        End If
    Next i
    propDesc = detail
    QTP_TryGetMaxTransparencyValue = found
End Function

Private Function QTP_TryGetProperty(ByVal obj As Object, ByVal propName As String, ByRef result As Variant) As Boolean
    On Error GoTo FAIL_PROP
    result = CallByName(obj, propName, VbGet)
    QTP_TryGetProperty = True
    Exit Function
FAIL_PROP:
    Err.Clear: QTP_TryGetProperty = False
End Function

Private Function QTP_NormalizeTransparencyPercent(ByVal raw As Variant, ByRef pct As Double) As Boolean
    On Error GoTo FAIL_VALUE
    If Not IsNumeric(raw) Then QTP_NormalizeTransparencyPercent = False: Exit Function
    pct = CDbl(raw)
    If pct > 0 And pct <= 1 Then pct = pct * 100
    If pct < 0 Or pct > 100 Then QTP_NormalizeTransparencyPercent = False: Exit Function
    QTP_NormalizeTransparencyPercent = True
    Exit Function
FAIL_VALUE:
    QTP_NormalizeTransparencyPercent = False
End Function

Private Function QTP_TransparencyTypeName(ByVal t As Long) As String
    Select Case t
        Case 0: QTP_TransparencyTypeName = "None"
        Case 1: QTP_TransparencyTypeName = "Uniform / Type#1"
        Case 2: QTP_TransparencyTypeName = "Fountain / Type#2"
        Case 3: QTP_TransparencyTypeName = "Pattern / Type#3"
        Case 4: QTP_TransparencyTypeName = "Texture / Type#4"
        Case Else: QTP_TransparencyTypeName = "Type#" & CStr(t)
    End Select
End Function

Private Function QTP_ShapeInfo(ByVal s As Shape) As String
    Dim nm As String, info As String
    On Error Resume Next
    nm = s.Name: If Trim(nm) = "" Then nm = "-"
    info = "Shape=" & QTP_ShapeTypeName(s) & " | Name=" & nm & " | X=" & FormatNumber(s.CenterX, 2) & " | Y=" & FormatNumber(s.CenterY, 2) & " | W=" & FormatNumber(Abs(s.SizeWidth), 2) & " | H=" & FormatNumber(Abs(s.SizeHeight), 2)
    QTP_ShapeInfo = info
End Function

Private Function QTP_ShapeTypeName(ByVal s As Shape) As String
    On Error GoTo UNKNOWN_TYPE
    Select Case s.Type
        Case cdrGroupShape: QTP_ShapeTypeName = "GROUP"
        Case cdrCurveShape: QTP_ShapeTypeName = "CURVE"
        Case cdrTextShape: QTP_ShapeTypeName = "TEXT"
        Case cdrBitmapShape: QTP_ShapeTypeName = "BITMAP"
        Case cdrRectangleShape: QTP_ShapeTypeName = "RECTANGLE"
        Case cdrEllipseShape: QTP_ShapeTypeName = "ELLIPSE"
        Case cdrPolygonShape: QTP_ShapeTypeName = "POLYGON"
        Case Else: QTP_ShapeTypeName = "TYPE#" & CStr(s.Type)
    End Select
    Exit Function
UNKNOWN_TYPE:
    QTP_ShapeTypeName = "UNKNOWN"
End Function

Private Sub QTP_AddDetail(ByRef target As String, ByVal lineText As String)
    If qtpDetailCount < QTP_DETAIL_LIMIT Then
        target = target & "- " & lineText & vbCrLf
        qtpDetailCount = qtpDetailCount + 1
    Else
        qtpDetailSkipped = qtpDetailSkipped + 1
    End If
End Sub

Private Function QTP_WriteReportFile() As String
    Dim path As String, f As Integer
    On Error GoTo FAIL_WRITE
    path = Environ("USERPROFILE") & "\Documents\HADES_QC_PREPRESS_REPORT.txt"
    f = FreeFile
    Open path For Output As #f
    Print #f, QTP_BuildFullReport()
    Close #f
    QTP_WriteReportFile = path
    Exit Function
FAIL_WRITE:
    On Error Resume Next: Close #f: On Error GoTo 0
    QTP_WriteReportFile = ""
End Function

Private Function QTP_BuildFullReport() As String
    Dim r As String, isFail As Boolean, isWarn As Boolean
    isFail = QTP_IsFail(): isWarn = QTP_IsWarningOnly()
    r = "PROJECT HADES — QC TRANSPARENCY + POWERCLIP REPORT" & vbCrLf
    r = r & "Generated : " & Format(Now, "yyyy-mm-dd hh:nn:ss") & vbCrLf
    r = r & String(60, "=") & vbCrLf & vbCrLf
    If isFail Then r = r & "STATUS : FAIL / REJECT" & vbCrLf ElseIf isWarn Then r = r & "STATUS : WARNING" & vbCrLf Else r = r & "STATUS : PASS" & vbCrLf
    r = r & vbCrLf & "SUMMARY" & vbCrLf
    r = r & "Scanned shape              : " & qtpScanned & vbCrLf
    r = r & "PowerClip found            : " & qtpPowerClip & vbCrLf
    r = r & "Transparency safe 0-70     : " & qtpTransparencySafe & vbCrLf
    r = r & "Transparency warning 71-75 : " & qtpTransparencyWarning & vbCrLf
    r = r & "Transparency FAIL >75      : " & qtpTransparencyFail & vbCrLf
    r = r & "Transparency unknown       : " & qtpTransparencyUnknown & vbCrLf
    r = r & "Detail skipped             : " & qtpDetailSkipped & vbCrLf & vbCrLf
    r = r & String(60, "-") & vbCrLf & "POWERCLIP DETAIL" & vbCrLf & String(60, "-") & vbCrLf
    If qtpPowerClipDetail <> "" Then r = r & qtpPowerClipDetail Else r = r & "Tidak ada PowerClip." & vbCrLf
    r = r & vbCrLf & String(60, "-") & vbCrLf & "TRANSPARENCY FAIL DETAIL (>75%)" & vbCrLf & String(60, "-") & vbCrLf
    If qtpTransparencyFailDetail <> "" Then r = r & qtpTransparencyFailDetail Else r = r & "Tidak ada transparency >75%." & vbCrLf
    r = r & vbCrLf & String(60, "-") & vbCrLf & "TRANSPARENCY WARNING DETAIL (71% - 75%)" & vbCrLf & String(60, "-") & vbCrLf
    If qtpTransparencyWarningDetail <> "" Then r = r & qtpTransparencyWarningDetail Else r = r & "Tidak ada transparency 71% - 75%." & vbCrLf
    r = r & vbCrLf & String(60, "-") & vbCrLf & "TRANSPARENCY UNKNOWN DETAIL" & vbCrLf & String(60, "-") & vbCrLf
    If qtpTransparencyUnknownDetail <> "" Then r = r & qtpTransparencyUnknownDetail Else r = r & "Tidak ada transparency unknown." & vbCrLf
    r = r & vbCrLf & String(60, "=") & vbCrLf & "RULES" & vbCrLf
    r = r & "- PowerClip ditemukan = FAIL / REJECT" & vbCrLf
    r = r & "- Transparency >75% = FAIL / REJECT" & vbCrLf
    r = r & "- Transparency 71% - 75% = WARNING" & vbCrLf
    r = r & "- Transparency 0% - 70% = PASS" & vbCrLf
    r = r & "- Unknown transparency = WARNING by default" & vbCrLf
    QTP_BuildFullReport = r
End Function

Private Function QTP_IsFail() As Boolean
    QTP_IsFail = False
    If qtpPowerClip > 0 Then QTP_IsFail = True: Exit Function
    If qtpTransparencyFail > 0 Then QTP_IsFail = True: Exit Function
    If QTP_UNKNOWN_TRANSPARENCY_IS_FAIL And qtpTransparencyUnknown > 0 Then QTP_IsFail = True: Exit Function
End Function

Private Function QTP_IsWarningOnly() As Boolean
    QTP_IsWarningOnly = False
    If QTP_IsFail() Then Exit Function
    If qtpTransparencyWarning > 0 Or qtpTransparencyUnknown > 0 Then QTP_IsWarningOnly = True
End Function

Private Sub QTP_ShowFinalResult()
    Dim msg As String
    msg = "QC TRANSPARENCY + POWERCLIP CHECK" & vbCrLf & vbCrLf
    msg = msg & "Scanned shape              : " & qtpScanned & vbCrLf
    msg = msg & "PowerClip found            : " & qtpPowerClip & vbCrLf
    msg = msg & "Transparency safe 0-70     : " & qtpTransparencySafe & vbCrLf
    msg = msg & "Transparency warning 71-75 : " & qtpTransparencyWarning & vbCrLf
    msg = msg & "Transparency FAIL >75      : " & qtpTransparencyFail & vbCrLf
    msg = msg & "Transparency unknown       : " & qtpTransparencyUnknown & vbCrLf & vbCrLf
    If qtpReportPath <> "" Then msg = msg & "Detail report:" & vbCrLf & qtpReportPath & vbCrLf & vbCrLf Else msg = msg & "Detail report gagal ditulis." & vbCrLf & vbCrLf
    If QTP_IsFail() Then
        msg = msg & "STATUS: FAIL / REJECT" & vbCrLf & vbCrLf & "Desain sebaiknya dikembalikan ke desainer untuk revisi." & vbCrLf & "Penyebab FAIL: PowerClip atau Transparency >75%."
        MsgBox msg, vbCritical, "QC PREPRESS FAIL"
    ElseIf QTP_IsWarningOnly() Then
        msg = msg & "STATUS: WARNING" & vbCrLf & vbCrLf & "Tidak ada reject fatal, tetapi ada object yang perlu dicek manual."
        MsgBox msg, vbExclamation, "QC PREPRESS WARNING"
    Else
        msg = msg & "STATUS: PASS" & vbCrLf & vbCrLf & "Tidak ditemukan PowerClip dan tidak ada Transparency berisiko."
        MsgBox msg, vbInformation, "QC PREPRESS PASS"
    End If
End Sub
