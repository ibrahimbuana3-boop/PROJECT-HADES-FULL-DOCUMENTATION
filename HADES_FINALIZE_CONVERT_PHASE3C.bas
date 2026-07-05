Option Explicit

'=========================================================
' PROJECT HADES — FINALIZE CONVERT PHASE 3C
' CorelDRAW 2021 VBA
'
' MAIN MACRO:
' HADES_FINALIZE_CONVERT_V3
'
' DEPENDENCY:
' - HADES_CORE_REPORT_PHASE2.bas sudah diimport.
' - HADES_QC_FINAL_REPORT_V3C sudah dijalankan dan membuat LATEST report.
'
' TUJUAN PHASE 3C:
' - Convert gate produksi setelah report otomatis V3C.
' - Membaca FINAL_STATUS dan CONVERT_PERMISSION dari latest report.
' - Mengconvert hanya ACTIVE TEXT di selection menjadi curves.
' - Menulis log final convert ke Documents\HADES_REPORTS.
'=========================================================

Public Sub HADES_FINALIZE_CONVERT_V3()

    Dim finalStatus As String
    Dim permission As String
    Dim ans As String
    Dim textShapes As Collection
    Dim oldUnit As cdrUnit
    Dim oldOptimization As Boolean
    Dim convertedCount As Long
    Dim failedCount As Long
    Dim cmdStarted As Boolean
    Dim latestReportPath As String
    Dim reportGenerated As String

    On Error GoTo ERR_HANDLER

    latestReportPath = HADESR_LatestFinalReportPath()
    finalStatus = UCase$(Trim$(HADESR_ReadLatestFinalStatus()))
    permission = UCase$(Trim$(HADESR_ReadLatestConvertPermission()))
    reportGenerated = H3C_ReadMachineValueFromLatest("GENERATED")

    If finalStatus = "" Or permission = "" Then
        MsgBox _
            "Final QC Report belum ditemukan atau belum valid." & vbCrLf & vbCrLf & _
            "Jalankan HADES_QC_FINAL_REPORT_V3C terlebih dahulu.", _
            vbExclamation, _
            "HADES FINALIZE CONVERT V3"
        Exit Sub
    End If

    If permission = "BLOCKED" Or finalStatus = "FAIL" Then
        MsgBox _
            "FINAL CONVERT DIBLOKIR." & vbCrLf & vbCrLf & _
            "Final Status       : " & finalStatus & vbCrLf & _
            "Convert Permission : " & permission & vbCrLf & _
            "Report Generated   : " & reportGenerated & vbCrLf & vbCrLf & _
            "Ada QC yang FAIL. Perbaiki layout, lalu jalankan Final QC Report lagi.", _
            vbCritical, _
            "HADES CONVERT BLOCKED"
        HADESR_OpenLatestFinalReport
        Exit Sub
    End If

    If ActiveSelection Is Nothing Then
        MsgBox "Pilih layout yang akan difinal-convert.", vbExclamation, "HADES FINALIZE CONVERT V3"
        Exit Sub
    End If

    If ActiveSelection.Shapes.Count = 0 Then
        MsgBox "Pilih layout yang akan difinal-convert.", vbExclamation, "HADES FINALIZE CONVERT V3"
        Exit Sub
    End If

    Set textShapes = New Collection
    H3C_CollectTextShapes ActiveSelectionRange, textShapes

    If textShapes.Count = 0 Then
        MsgBox _
            "Tidak ada active text di selection." & vbCrLf & _
            "Tidak ada objek yang di-convert.", _
            vbInformation, _
            "HADES FINALIZE CONVERT V3"
        Exit Sub
    End If

    If permission = "MANUAL_CONFIRM" Or finalStatus = "WARNING" Then
        ans = InputBox( _
            "FINAL QC STATUS: WARNING" & vbCrLf & vbCrLf & _
            "Ada WARNING pada report. Pastikan sudah dicek manual." & vbCrLf & vbCrLf & _
            "Report Generated: " & reportGenerated & vbCrLf & _
            "Jumlah active text terdeteksi: " & textShapes.Count & vbCrLf & vbCrLf & _
            "Ketik FINAL untuk tetap melanjutkan convert.", _
            "HADES FINALIZE CONVERT V3")
    Else
        ans = InputBox( _
            "FINAL QC STATUS: PASS" & vbCrLf & vbCrLf & _
            "Semua QC PASS berdasarkan report terakhir." & vbCrLf & vbCrLf & _
            "Report Generated: " & reportGenerated & vbCrLf & _
            "Jumlah active text terdeteksi: " & textShapes.Count & vbCrLf & vbCrLf & _
            "Macro ini akan mengubah semua ACTIVE TEXT di selection menjadi curve." & vbCrLf & _
            "Setelah ini QC_TYPO_CHECK dan IDPO_CHECK tidak bisa membaca teks lagi." & vbCrLf & vbCrLf & _
            "Ketik FINAL untuk lanjut.", _
            "HADES FINALIZE CONVERT V3")
    End If

    ans = UCase$(Trim$(ans))

    If ans <> "FINAL" Then
        MsgBox "Final convert dibatalkan. Tidak ada objek yang diubah.", vbInformation, "HADES FINALIZE CONVERT V3"
        Exit Sub
    End If

    oldUnit = ActiveDocument.Unit
    oldOptimization = Application.Optimization
    cmdStarted = False

    ActiveDocument.Unit = cdrCentimeter
    Application.Optimization = True
    ActiveDocument.BeginCommandGroup "HADES FINALIZE CONVERT V3 ACTIVE TEXT"
    cmdStarted = True

    H3C_ConvertTexts textShapes, convertedCount, failedCount

    ActiveDocument.EndCommandGroup
    cmdStarted = False

    Application.Optimization = oldOptimization
    ActiveDocument.Unit = oldUnit
    ActiveWindow.Refresh

    H3C_WriteConvertLog finalStatus, permission, reportGenerated, latestReportPath, convertedCount, failedCount

    MsgBox _
        "HADES FINAL CONVERT V3 SELESAI" & vbCrLf & vbCrLf & _
        "Converted : " & convertedCount & vbCrLf & _
        "Failed    : " & failedCount & vbCrLf & vbCrLf & _
        "Log convert ditulis ke folder HADES_REPORTS.", _
        vbInformation, _
        "HADES FINALIZE CONVERT V3"

    Exit Sub

ERR_HANDLER:

    On Error Resume Next

    If cmdStarted Then ActiveDocument.EndCommandGroup

    Application.Optimization = oldOptimization
    ActiveDocument.Unit = oldUnit
    ActiveWindow.Refresh

    MsgBox _
        "SYSTEM ERROR - HADES FINALIZE CONVERT V3" & vbCrLf & vbCrLf & _
        "No : " & Err.Number & vbCrLf & _
        Err.Description, _
        vbCritical, _
        "HADES FINALIZE CONVERT V3"

End Sub

'=========================================================
' TEXT COLLECTION
'=========================================================
Private Sub H3C_CollectTextShapes(ByVal sr As ShapeRange, ByVal textShapes As Collection)

    Dim s As Shape

    For Each s In sr
        H3C_CollectTextShapeRecursive s, textShapes
    Next s

End Sub

Private Sub H3C_CollectTextShapeRecursive(ByVal s As Shape, ByVal textShapes As Collection)

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
            H3C_CollectTextShapeRecursive ch, textShapes
        Next ch
    End If

    Set pcShapes = Nothing
    Set pcShapes = s.PowerClip.Shapes

    If Not pcShapes Is Nothing Then
        For Each ch In pcShapes
            H3C_CollectTextShapeRecursive ch, textShapes
        Next ch
    End If

    On Error GoTo 0

End Sub

'=========================================================
' CONVERT
'=========================================================
Private Sub H3C_ConvertTexts(ByVal textShapes As Collection, ByRef convertedCount As Long, ByRef failedCount As Long)

    Dim v As Variant
    Dim t As Shape

    convertedCount = 0
    failedCount = 0

    For Each v In textShapes

        Set t = v

        If H3C_IsShapeAlive(t) Then
            If H3C_ConvertOneText(t) Then
                convertedCount = convertedCount + 1
            Else
                failedCount = failedCount + 1
            End If
        Else
            failedCount = failedCount + 1
        End If

    Next v

End Sub

Private Function H3C_ConvertOneText(ByVal t As Shape) As Boolean

    On Error GoTo FAIL

    H3C_ConvertOneText = False

    If t Is Nothing Then Exit Function
    If t.Type <> cdrTextShape Then Exit Function

    On Error Resume Next
    t.Locked = False
    Err.Clear
    On Error GoTo FAIL

    t.ConvertToCurves

    H3C_ConvertOneText = True
    Exit Function

FAIL:

    Err.Clear
    H3C_ConvertOneText = False

End Function

Private Function H3C_IsShapeAlive(ByVal s As Shape) As Boolean

    Dim t As Long

    On Error GoTo DEAD

    If s Is Nothing Then GoTo DEAD

    t = s.Type
    H3C_IsShapeAlive = True
    Exit Function

DEAD:

    H3C_IsShapeAlive = False

End Function

'=========================================================
' LOGGING
'=========================================================
Private Sub H3C_WriteConvertLog( _
    ByVal finalStatus As String, _
    ByVal permission As String, _
    ByVal reportGenerated As String, _
    ByVal latestReportPath As String, _
    ByVal convertedCount As Long, _
    ByVal failedCount As Long)

    Dim folderPath As String
    Dim latestPath As String
    Dim timePath As String
    Dim ts As String
    Dim txt As String

    folderPath = HADESR_ReportFolderPath()
    H3C_EnsureFolder folderPath

    ts = Format(Now, "yyyymmdd_hhnnss")
    latestPath = folderPath & "\HADES_FINAL_CONVERT_LOG_LATEST.txt"
    timePath = folderPath & "\HADES_FINAL_CONVERT_LOG_" & ts & ".txt"

    txt = "PROJECT HADES — FINAL CONVERT LOG" & vbCrLf
    txt = txt & String(70, "=") & vbCrLf
    txt = txt & "Converted At       : " & Format(Now, "yyyy-mm-dd hh:nn:ss") & vbCrLf
    txt = txt & "Document           : " & ActiveDocument.Name & vbCrLf
    txt = txt & "Final Status       : " & finalStatus & vbCrLf
    txt = txt & "Convert Permission : " & permission & vbCrLf
    txt = txt & "Report Generated   : " & reportGenerated & vbCrLf
    txt = txt & "Report Path        : " & latestReportPath & vbCrLf
    txt = txt & "Converted Text     : " & convertedCount & vbCrLf
    txt = txt & "Failed Convert     : " & failedCount & vbCrLf
    txt = txt & String(70, "=") & vbCrLf

    H3C_WriteText latestPath, txt
    H3C_WriteText timePath, txt

End Sub

Private Function H3C_ReadMachineValueFromLatest(ByVal keyName As String) As String

    Dim txt As String
    Dim lines() As String
    Dim i As Long
    Dim lineText As String
    Dim prefix As String

    txt = HADESR_ReadText(HADESR_LatestFinalReportPath())
    txt = Replace(txt, vbCrLf, vbLf)
    txt = Replace(txt, vbCr, vbLf)
    lines = Split(txt, vbLf)

    prefix = UCase$(Trim$(keyName)) & "="

    For i = LBound(lines) To UBound(lines)
        lineText = Trim$(CStr(lines(i)))
        If UCase$(Left$(lineText, Len(prefix))) = prefix Then
            H3C_ReadMachineValueFromLatest = Trim$(Mid$(lineText, Len(prefix) + 1))
            Exit Function
        End If
    Next i

    H3C_ReadMachineValueFromLatest = ""

End Function

Private Sub H3C_EnsureFolder(ByVal folderPath As String)

    On Error Resume Next

    If Dir$(folderPath, vbDirectory) = "" Then
        MkDir folderPath
    End If

    On Error GoTo 0

End Sub

Private Sub H3C_WriteText(ByVal path As String, ByVal txt As String)

    Dim f As Integer
    f = FreeFile

    Open path For Output As #f
    Print #f, txt
    Close #f

End Sub
