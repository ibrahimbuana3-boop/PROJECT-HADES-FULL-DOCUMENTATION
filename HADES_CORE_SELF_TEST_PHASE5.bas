Option Explicit

'=========================================================
' PROJECT HADES — CORE SELF TEST PHASE 5
' CorelDRAW 2021 VBA
'
' MAIN MACRO:
' - HADES5_CORE_SELF_TEST
' - HADES5_OpenLatestCoreSelfTestReport
'=========================================================

Private Const H5_SELFTEST_LATEST As String = "HADES_CORE_SELF_TEST_LATEST.txt"

Public Sub HADES5_CORE_SELF_TEST()
    Dim reportText As String
    Dim failCount As Long
    Dim warnCount As Long
    Dim ok As Boolean
    Dim latestPath As String

    ok = H5_RunCoreSelfTest(reportText, failCount, warnCount)
    latestPath = H5_WriteCoreSelfTestReport(reportText)

    If ok Then
        MsgBox _
            "HADES CORE SELF TEST PASSED" & vbCrLf & vbCrLf & _
            "Warning : " & warnCount & vbCrLf & _
            "Report  : " & latestPath, _
            vbInformation, _
            "HADES PHASE 5"
    Else
        MsgBox _
            "HADES CORE SELF TEST FAILED" & vbCrLf & vbCrLf & _
            "Fail    : " & failCount & vbCrLf & _
            "Warning : " & warnCount & vbCrLf & _
            "Report  : " & latestPath, _
            vbCritical, _
            "HADES PHASE 5"
    End If
End Sub

Public Function H5_RunCoreSelfTest( _
    ByRef reportText As String, _
    ByRef failCount As Long, _
    ByRef warnCount As Long) As Boolean

    Dim orderPath As String
    Dim tplPath As String
    Dim dbName As String
    Dim dbPath As String
    Dim meta As Object
    Dim rows As Collection
    Dim expected As Object
    Dim tpl As Object
    Dim db As Object
    Dim isPants As Boolean
    Dim isSplitFront As Boolean
    Dim detail As String
    Dim sig As String

    failCount = 0
    warnCount = 0

    reportText = ""
    reportText = reportText & "HADES CORE SELF TEST PHASE 5" & vbCrLf
    reportText = reportText & "GENERATED=" & H5_NowHuman() & vbCrLf
    reportText = reportText & String(60, "-") & vbCrLf & vbCrLf

    orderPath = H5_OrderPath()
    tplPath = H5_TemplatePath()

    '=====================================================
    ' ORDER
    '=====================================================
    reportText = reportText & "[ORDER]" & vbCrLf

    If Dir$(orderPath) = "" Then
        failCount = failCount + 1
        reportText = reportText & "FAIL | Order.txt tidak ditemukan: " & orderPath & vbCrLf
    Else
        Set meta = H5_LoadOrderMeta()
        Set rows = H5_LoadOrderRows()
        Set expected = H5_LoadOrderExpectedCounts()

        reportText = reportText & "PASS | Order.txt ditemukan." & vbCrLf
        reportText = reportText & "Rows valid : " & rows.Count & vbCrLf
        reportText = reportText & "Size count : " & expected.Count & vbCrLf

        If rows.Count = 0 Then
            failCount = failCount + 1
            reportText = reportText & "FAIL | Tidak ada baris order valid SIZE|NAMA|NOMOR|NICKNAME." & vbCrLf
        End If

        If meta.Exists("SIZEDB") Then
            reportText = reportText & "@SIZEDB    : " & CStr(meta("SIZEDB")) & vbCrLf
        Else
            warnCount = warnCount + 1
            reportText = reportText & "WARN | @SIZEDB tidak ditemukan. Sistem akan coba infer dari metadata." & vbCrLf
        End If
    End If

    reportText = reportText & vbCrLf

    '=====================================================
    ' TEMPLATE
    '=====================================================
    reportText = reportText & "[TYPO TEMPLATE]" & vbCrLf

    If Dir$(tplPath) = "" Then
        failCount = failCount + 1
        reportText = reportText & "FAIL | TypoTemplate_Current.txt tidak ditemukan: " & tplPath & vbCrLf
    Else
        Set tpl = H5_LoadTypoTemplate()
        reportText = reportText & "PASS | TypoTemplate_Current.txt ditemukan." & vbCrLf
        reportText = reportText & "Key count : " & tpl.Count & vbCrLf

        If Not tpl.Exists("MASTER_PANEL") Then
            failCount = failCount + 1
            reportText = reportText & "FAIL | MASTER_PANEL tidak ada. Jalankan BUILD_TYPO_TEMPLATE ulang." & vbCrLf
        ElseIf Val(CStr(tpl("MASTER_PANEL"))) <= 0 Then
            failCount = failCount + 1
            reportText = reportText & "FAIL | MASTER_PANEL tidak valid: " & CStr(tpl("MASTER_PANEL")) & vbCrLf
        Else
            reportText = reportText & "MASTER_PANEL : " & CStr(tpl("MASTER_PANEL")) & vbCrLf
        End If
    End If

    reportText = reportText & vbCrLf

    '=====================================================
    ' DATABASE
    '=====================================================
    reportText = reportText & "[SIZE DATABASE]" & vbCrLf

    dbName = H5_DetectCurrentSizeDBFileName()

    If Len(Trim$(dbName)) = 0 Then
        failCount = failCount + 1
        reportText = reportText & "FAIL | SizeDB tidak bisa ditentukan dari Order.txt / Template." & vbCrLf
    Else
        dbPath = H5_DocumentsFile(dbName)
        reportText = reportText & "Detected DB : " & dbName & vbCrLf

        If Dir$(dbPath) = "" Then
            failCount = failCount + 1
            reportText = reportText & "FAIL | SizeDB tidak ditemukan: " & dbPath & vbCrLf
        Else
            H5_ProductModeFromDB dbName, isPants, isSplitFront
            Set db = H5_LoadSizeDB(dbName, isPants, isSplitFront)

            If db.Count = 0 Then
                failCount = failCount + 1
                reportText = reportText & "FAIL | SizeDB kosong / format tidak valid." & vbCrLf
            Else
                reportText = reportText & "PASS | SizeDB ditemukan dan terbaca." & vbCrLf
                reportText = reportText & "DB rows : " & db.Count & vbCrLf
                reportText = reportText & "Mode    : " & H5_ModeText(isPants, isSplitFront) & vbCrLf
            End If
        End If
    End If

    reportText = reportText & vbCrLf

    '=====================================================
    ' SELECTION
    '=====================================================
    reportText = reportText & "[SELECTION]" & vbCrLf

    If H5_HasSelection() Then
        sig = H5_BuildCurrentSelectionSignature(detail)
        reportText = reportText & "PASS | Selection aktif terdeteksi." & vbCrLf
        reportText = reportText & "Signature: " & sig & vbCrLf
        reportText = reportText & detail
        reportText = reportText & "Active text count: " & H5_CountActiveTextInSelection() & vbCrLf
    Else
        warnCount = warnCount + 1
        reportText = reportText & "WARN | Tidak ada selection aktif. Self-test data tetap bisa jalan." & vbCrLf
    End If

    reportText = reportText & vbCrLf

    '=====================================================
    ' REPORT FOLDER
    '=====================================================
    reportText = reportText & "[REPORT FOLDER]" & vbCrLf
    reportText = reportText & "Path: " & H5_ReportFolderPath() & vbCrLf

    If Dir$(H5_ReportFolderPath(), vbDirectory) = "" Then
        failCount = failCount + 1
        reportText = reportText & "FAIL | Report folder tidak bisa dibuat." & vbCrLf
    Else
        reportText = reportText & "PASS | Report folder siap." & vbCrLf
    End If

    reportText = reportText & vbCrLf & String(60, "-") & vbCrLf

    If failCount = 0 Then
        reportText = reportText & "CORE_SELF_TEST_STATUS=PASS" & vbCrLf
        H5_RunCoreSelfTest = True
    Else
        reportText = reportText & "CORE_SELF_TEST_STATUS=FAIL" & vbCrLf
        H5_RunCoreSelfTest = False
    End If

    reportText = reportText & "FAIL_COUNT=" & failCount & vbCrLf
    reportText = reportText & "WARNING_COUNT=" & warnCount & vbCrLf
End Function

Public Function H5_WriteCoreSelfTestReport(ByVal reportText As String) As String
    Dim folder As String
    Dim latestPath As String
    Dim stampPath As String

    folder = H5_ReportFolderPath()
    latestPath = folder & "\" & H5_SELFTEST_LATEST
    stampPath = folder & "\HADES_CORE_SELF_TEST_" & H5_NowStamp() & ".txt"

    H5_WriteTextUTF8 latestPath, reportText
    H5_WriteTextUTF8 stampPath, reportText

    H5_WriteCoreSelfTestReport = latestPath
End Function

Public Sub HADES5_OpenLatestCoreSelfTestReport()
    H5_OpenFile H5_ReportFolderPath() & "\" & H5_SELFTEST_LATEST
End Sub

Private Function H5_ModeText(ByVal isPants As Boolean, ByVal isSplitFront As Boolean) As String
    If isPants Then
        H5_ModeText = "CELANA"
    ElseIf isSplitFront Then
        H5_ModeText = "JAKET"
    Else
        H5_ModeText = "JERSEY"
    End If
End Function
