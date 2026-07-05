Option Explicit

'=========================================================
' PROJECT HADES — CORE REPORT ENGINE PHASE 2
' CorelDRAW 2021 VBA
'
' MODULE:
' HADES_CORE_REPORT_PHASE2.bas
'
' TUJUAN:
' - Menjadi mesin report gabungan untuk QC Final.
' - Menulis report timestamp dan LATEST ke Documents\HADES_REPORTS.
' - Menyimpan machine-readable status untuk convert gate.
'
' CATATAN:
' - Module ini tidak mengubah object CorelDRAW.
' - Dipakai oleh HADES_QC_FINAL_REPORT_V2 dan HADES_FINALIZE_CONVERT_V2.
'=========================================================

Private Const HR_STATUS_PASS As String = "PASS"
Private Const HR_STATUS_WARNING As String = "WARNING"
Private Const HR_STATUS_FAIL As String = "FAIL"
Private Const HR_STATUS_SKIPPED As String = "SKIPPED"

Private HR_Title As String
Private HR_StartedAt As Date
Private HR_ModuleLines As Collection
Private HR_DetailLines As Collection
Private HR_HasFail As Boolean
Private HR_HasWarning As Boolean

'=========================================================
' PUBLIC API
'=========================================================
Public Sub HADESR_Reset(ByVal reportTitle As String)

    HR_Title = reportTitle
    HR_StartedAt = Now
    Set HR_ModuleLines = New Collection
    Set HR_DetailLines = New Collection
    HR_HasFail = False
    HR_HasWarning = False

End Sub

Public Sub HADESR_AddResult( _
    ByVal moduleName As String, _
    ByVal statusText As String, _
    ByVal summaryText As String, _
    Optional ByVal detailText As String = "")

    Dim st As String
    Dim lineText As String

    st = HADESR_NormalizeStatus(statusText)

    If st = HR_STATUS_FAIL Then HR_HasFail = True
    If st = HR_STATUS_WARNING Then HR_HasWarning = True

    lineText = HADESR_PadRight(moduleName, 32) & " : " & st

    If Trim$(summaryText) <> "" Then
        lineText = lineText & " — " & summaryText
    End If

    HR_ModuleLines.Add lineText

    If Trim$(detailText) <> "" Then
        HR_DetailLines.Add "[" & moduleName & "]" & vbCrLf & detailText
    End If

End Sub

Public Sub HADESR_AddNote(ByVal noteText As String)

    If Trim$(noteText) <> "" Then
        HR_DetailLines.Add "[NOTE]" & vbCrLf & noteText
    End If

End Sub

Public Function HADESR_FinalStatus() As String

    If HR_HasFail Then
        HADESR_FinalStatus = HR_STATUS_FAIL
    ElseIf HR_HasWarning Then
        HADESR_FinalStatus = HR_STATUS_WARNING
    Else
        HADESR_FinalStatus = HR_STATUS_PASS
    End If

End Function

Public Function HADESR_ConvertPermission() As String

    Select Case HADESR_FinalStatus()

        Case HR_STATUS_PASS
            HADESR_ConvertPermission = "ALLOWED"

        Case HR_STATUS_WARNING
            HADESR_ConvertPermission = "MANUAL_CONFIRM"

        Case Else
            HADESR_ConvertPermission = "BLOCKED"

    End Select

End Function

Public Function HADESR_ReportFolderPath() As String

    HADESR_ReportFolderPath = Environ$("USERPROFILE") & "\Documents\HADES_REPORTS"

End Function

Public Function HADESR_LatestFinalReportPath() As String

    HADESR_LatestFinalReportPath = HADESR_ReportFolderPath() & "\HADES_FINAL_QC_REPORT_LATEST.txt"

End Function

Public Function HADESR_WriteFinalReport() As String

    Dim folderPath As String
    Dim latestPath As String
    Dim timePath As String
    Dim reportText As String
    Dim ts As String

    folderPath = HADESR_ReportFolderPath()
    HADESR_EnsureFolder folderPath

    ts = Format(Now, "yyyymmdd_hhnnss")

    latestPath = folderPath & "\HADES_FINAL_QC_REPORT_LATEST.txt"
    timePath = folderPath & "\HADES_FINAL_QC_REPORT_" & ts & ".txt"

    reportText = HADESR_BuildFinalReportText()

    HADESR_WriteText latestPath, reportText
    HADESR_WriteText timePath, reportText

    HADESR_WriteFinalReport = latestPath

End Function

Public Function HADESR_ReadLatestFinalStatus() As String

    Dim txt As String
    txt = HADESR_ReadText(HADESR_LatestFinalReportPath())

    HADESR_ReadLatestFinalStatus = HADESR_ReadMachineValue(txt, "FINAL_STATUS")

End Function

Public Function HADESR_ReadLatestConvertPermission() As String

    Dim txt As String
    txt = HADESR_ReadText(HADESR_LatestFinalReportPath())

    HADESR_ReadLatestConvertPermission = HADESR_ReadMachineValue(txt, "CONVERT_PERMISSION")

End Function

Public Sub HADESR_OpenLatestFinalReport()

    Dim p As String
    p = HADESR_LatestFinalReportPath()

    If Dir$(p) = "" Then
        MsgBox "Report latest belum ditemukan:" & vbCrLf & p, vbExclamation, "HADES REPORT"
        Exit Sub
    End If

    Shell "notepad.exe " & Chr$(34) & p & Chr$(34), vbNormalFocus

End Sub

'=========================================================
' BUILD REPORT
'=========================================================
Private Function HADESR_BuildFinalReportText() As String

    Dim r As String
    Dim i As Long
    Dim finalStatus As String
    Dim permission As String
    Dim metaText As String

    finalStatus = HADESR_FinalStatus()
    permission = HADESR_ConvertPermission()
    metaText = HADESR_BuildOrderMetadataText()

    r = "# PROJECT_HADES_MACHINE_STATUS" & vbCrLf
    r = r & "FINAL_STATUS=" & finalStatus & vbCrLf
    r = r & "CONVERT_PERMISSION=" & permission & vbCrLf
    r = r & "GENERATED=" & Format(Now, "yyyy-mm-dd hh:nn:ss") & vbCrLf
    r = r & "# END_PROJECT_HADES_MACHINE_STATUS" & vbCrLf & vbCrLf

    r = r & "PROJECT HADES — FINAL QC REPORT" & vbCrLf
    r = r & String(70, "=") & vbCrLf
    r = r & "Generated : " & Format(Now, "yyyy-mm-dd hh:nn:ss") & vbCrLf
    r = r & "Started   : " & Format(HR_StartedAt, "yyyy-mm-dd hh:nn:ss") & vbCrLf
    r = r & String(70, "=") & vbCrLf & vbCrLf

    r = r & "JOB METADATA" & vbCrLf
    r = r & String(70, "-") & vbCrLf

    If Trim$(metaText) <> "" Then
        r = r & metaText
    Else
        r = r & "Order.txt metadata tidak terbaca." & vbCrLf
    End If

    r = r & vbCrLf
    r = r & "QC SUMMARY" & vbCrLf
    r = r & String(70, "-") & vbCrLf

    If HR_ModuleLines Is Nothing Then
        r = r & "Tidak ada hasil QC." & vbCrLf
    ElseIf HR_ModuleLines.Count = 0 Then
        r = r & "Tidak ada hasil QC." & vbCrLf
    Else
        For i = 1 To HR_ModuleLines.Count
            r = r & CStr(HR_ModuleLines(i)) & vbCrLf
        Next i
    End If

    r = r & vbCrLf
    r = r & "FINAL STATUS" & vbCrLf
    r = r & String(70, "-") & vbCrLf
    r = r & "FINAL STATUS       : " & finalStatus & vbCrLf
    r = r & "CONVERT PERMISSION : " & permission & vbCrLf

    Select Case permission
        Case "ALLOWED"
            r = r & "ACTION             : Semua QC PASS. Convert boleh ditawarkan." & vbCrLf
        Case "MANUAL_CONFIRM"
            r = r & "ACTION             : Ada WARNING. Convert hanya boleh setelah cek manual." & vbCrLf
        Case Else
            r = r & "ACTION             : Convert diblokir karena ada FAIL." & vbCrLf
    End Select

    r = r & vbCrLf
    r = r & "DETAIL" & vbCrLf
    r = r & String(70, "-") & vbCrLf

    If Not HR_DetailLines Is Nothing Then
        If HR_DetailLines.Count > 0 Then
            For i = 1 To HR_DetailLines.Count
                r = r & CStr(HR_DetailLines(i)) & vbCrLf & vbCrLf
            Next i
        Else
            r = r & "Tidak ada detail tambahan." & vbCrLf
        End If
    Else
        r = r & "Tidak ada detail tambahan." & vbCrLf
    End If

    r = r & String(70, "=") & vbCrLf
    r = r & "Generated by Project Hades Report Engine Phase 2" & vbCrLf

    HADESR_BuildFinalReportText = r

End Function

Private Function HADESR_BuildOrderMetadataText() As String

    Dim path As String
    Dim txt As String
    Dim lines() As String
    Dim i As Long
    Dim lineText As String
    Dim eqPos As Long
    Dim k As String
    Dim v As String
    Dim r As String

    path = Environ$("USERPROFILE") & "\Documents\Order.txt"

    If Dir$(path) = "" Then
        HADESR_BuildOrderMetadataText = "Order.txt tidak ditemukan: " & path & vbCrLf
        Exit Function
    End If

    txt = HADESR_ReadText(path)
    txt = Replace(txt, vbCrLf, vbLf)
    txt = Replace(txt, vbCr, vbLf)
    lines = Split(txt, vbLf)

    r = ""

    For i = LBound(lines) To UBound(lines)

        lineText = Trim$(lines(i))

        If Left$(lineText, 1) = "@" Then

            eqPos = InStr(1, lineText, "=", vbTextCompare)

            If eqPos > 1 Then
                k = Mid$(lineText, 2, eqPos - 2)
                v = Mid$(lineText, eqPos + 1)
                r = r & HADESR_PadRight(k, 18) & " : " & v & vbCrLf
            End If

        End If

    Next i

    HADESR_BuildOrderMetadataText = r

End Function

'=========================================================
' FILE UTILITIES
'=========================================================
Private Sub HADESR_EnsureFolder(ByVal folderPath As String)

    On Error Resume Next

    If Dir$(folderPath, vbDirectory) = "" Then
        MkDir folderPath
    End If

    On Error GoTo 0

End Sub

Private Sub HADESR_WriteText(ByVal path As String, ByVal txt As String)

    Dim f As Integer
    f = FreeFile

    Open path For Output As #f
    Print #f, txt
    Close #f

End Sub

Public Function HADESR_ReadText(ByVal path As String) As String

    On Error GoTo FAIL_UTF8

    Dim stm As Object
    Set stm = CreateObject("ADODB.Stream")

    With stm
        .Type = 2
        .Charset = "utf-8"
        .Open
        .LoadFromFile path
        HADESR_ReadText = .ReadText
        .Close
    End With

    Exit Function

FAIL_UTF8:

    On Error GoTo FAIL_ANSI

    Dim f As Integer
    f = FreeFile

    Open path For Input As #f
    HADESR_ReadText = Input$(LOF(f), #f)
    Close #f

    Exit Function

FAIL_ANSI:

    On Error Resume Next
    Close #f
    HADESR_ReadText = ""

End Function

Private Function HADESR_ReadMachineValue(ByVal txt As String, ByVal key As String) As String

    Dim lines() As String
    Dim i As Long
    Dim lineText As String
    Dim prefix As String

    txt = Replace(txt, vbCrLf, vbLf)
    txt = Replace(txt, vbCr, vbLf)
    lines = Split(txt, vbLf)

    prefix = UCase$(key) & "="

    For i = LBound(lines) To UBound(lines)
        lineText = Trim$(lines(i))
        If UCase$(Left$(lineText, Len(prefix))) = prefix Then
            HADESR_ReadMachineValue = Trim$(Mid$(lineText, Len(prefix) + 1))
            Exit Function
        End If
    Next i

    HADESR_ReadMachineValue = ""

End Function

'=========================================================
' STRING UTILITIES
'=========================================================
Public Function HADESR_NormalizeStatus(ByVal statusText As String) As String

    Dim st As String
    st = UCase$(Trim$(statusText))

    Select Case st
        Case "PASS", "P", "OK", "LULUS"
            HADESR_NormalizeStatus = HR_STATUS_PASS
        Case "WARNING", "WARN", "W", "PERINGATAN"
            HADESR_NormalizeStatus = HR_STATUS_WARNING
        Case "FAIL", "FAILED", "F", "REJECT", "ERROR", "GAGAL"
            HADESR_NormalizeStatus = HR_STATUS_FAIL
        Case "SKIP", "SKIPPED", "S", "LEWATI"
            HADESR_NormalizeStatus = HR_STATUS_SKIPPED
        Case Else
            HADESR_NormalizeStatus = HR_STATUS_WARNING
    End Select

End Function

Private Function HADESR_PadRight(ByVal s As String, ByVal totalLen As Long) As String

    If Len(s) >= totalLen Then
        HADESR_PadRight = s
    Else
        HADESR_PadRight = s & Space$(totalLen - Len(s))
    End If

End Function
