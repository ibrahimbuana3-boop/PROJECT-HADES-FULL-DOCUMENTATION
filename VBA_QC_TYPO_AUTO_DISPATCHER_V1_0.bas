Option Explicit

'=========================================================
' PROJECT H.A.D.E.S. — QC TYPO AUTO DISPATCHER V1.0
' CorelDRAW 2021 VBA
'
' PUBLIC SHORTCUTS KEPT:
'   QC_TYPO_CHECK
'   HADES_QC_TYPO_REPORT
'
' PURPOSE:
' - Satu pintu QC Typo untuk dua workflow:
'   1) Normal / LRP layout hasil Auto Duplicate + Auto Rename
'   2) Auto Mass Nesting layout berbasis Record Pattern + AMN metadata
'
' IMPORTANT:
' - Module ini TIDAK menyatukan engine normal dan nest menjadi satu kode raksasa.
' - Module ini hanya mendeteksi selection lalu memanggil engine yang tepat.
'
' REQUIRED ENGINE MODULES:
' - VBA_QC_TYPO_NORMAL_ENGINE_V13_2.bas   exposes QTC_NORMAL_CHECK / QTC_NORMAL_REPORT
' - VBA_QC_TYPO_NEST_ENGINE_V1_3.bas      exposes QTN_NEST_CHECK / QTN_NEST_REPORT
'=========================================================

Private Const QTA_AMN_META_PREFIX As String = "HADES_AMN|"
Private Const QTA_AMN_ROW_TOKEN As String = "ROW="

Public Sub QC_TYPO_CHECK()

    Dim mode As String
    mode = QTA_DetectTypoMode()

    Select Case mode

        Case "NEST"
            QTN_NEST_CHECK

        Case "NORMAL"
            QTC_NORMAL_CHECK

        Case Else
            MsgBox _
                "QC TYPO AUTO tidak bisa menentukan mode layout." & vbCrLf & vbCrLf & _
                "Tidak ditemukan selection valid, atau struktur selection ambigu." & vbCrLf & vbCrLf & _
                "Jika ini hasil Auto Mass Nesting, pastikan metadata HADES_AMN|ROW= masih ada pada panel hasil nesting." & vbCrLf & _
                "Jika ini layout normal/LRP, select group jersey hasil layout lalu ulangi.", _
                vbExclamation, "HADES QC TYPO AUTO"

    End Select

End Sub

Public Sub HADES_QC_TYPO_REPORT()

    Dim mode As String
    mode = QTA_DetectTypoMode()

    Select Case mode

        Case "NEST"
            QTN_NEST_REPORT

        Case "NORMAL"
            QTC_NORMAL_REPORT

        Case Else
            HADESR_AddResult _
                "TYPO CHECK", _
                "FAIL", _
                "QC Typo Auto tidak bisa menentukan mode layout.", _
                "Tidak ditemukan metadata HADES_AMN|ROW= dan selection tidak dapat dipastikan sebagai layout normal. Pisahkan selection atau cek metadata Auto Mass Nesting."

    End Select

End Sub

Public Sub HADES_QC_TYPO_AUTO_SMOKE_TEST()

    Dim mode As String
    Dim metaCount As Long
    Dim scannedCount As Long

    mode = QTA_DetectTypoMode(metaCount, scannedCount)

    MsgBox _
        "QC TYPO AUTO SMOKE TEST" & vbCrLf & vbCrLf & _
        "Detected mode : " & mode & vbCrLf & _
        "AMN metadata  : " & metaCount & vbCrLf & _
        "Shapes scanned: " & scannedCount & vbCrLf & vbCrLf & _
        "NEST   = akan memanggil QTN_NEST_CHECK" & vbCrLf & _
        "NORMAL = akan memanggil QTC_NORMAL_CHECK", _
        vbInformation, "HADES QC TYPO AUTO"

End Sub

Private Function QTA_DetectTypoMode( _
    Optional ByRef metaCount As Long = 0, _
    Optional ByRef scannedCount As Long = 0) As String

    On Error GoTo FAIL

    metaCount = 0
    scannedCount = 0

    If ActiveSelection Is Nothing Then
        QTA_DetectTypoMode = "UNKNOWN"
        Exit Function
    End If

    If ActiveSelection.Shapes.Count = 0 Then
        QTA_DetectTypoMode = "UNKNOWN"
        Exit Function
    End If

    Dim s As Shape

    For Each s In ActiveSelection.Shapes
        QTA_ScanForAMNMeta s, metaCount, scannedCount
    Next s

    If metaCount > 0 Then
        QTA_DetectTypoMode = "NEST"
    Else
        QTA_DetectTypoMode = "NORMAL"
    End If

    Exit Function

FAIL:
    QTA_DetectTypoMode = "UNKNOWN"

End Function

Private Sub QTA_ScanForAMNMeta( _
    ByVal s As Shape, _
    ByRef metaCount As Long, _
    ByRef scannedCount As Long)

    On Error Resume Next

    scannedCount = scannedCount + 1

    If QTA_HasAMNMeta(s) Then
        metaCount = metaCount + 1
    End If

    If s.Type = cdrGroupShape Then
        Dim c As Shape
        For Each c In s.Shapes
            QTA_ScanForAMNMeta c, metaCount, scannedCount
        Next c
    End If

    On Error GoTo 0

End Sub

Private Function QTA_HasAMNMeta(ByVal s As Shape) As Boolean

    On Error Resume Next

    Dim nm As String
    nm = CStr(s.Name)

    QTA_HasAMNMeta = _
        (InStr(1, nm, QTA_AMN_META_PREFIX, vbTextCompare) > 0 And _
         InStr(1, nm, QTA_AMN_ROW_TOKEN, vbTextCompare) > 0)

    On Error GoTo 0

End Function
