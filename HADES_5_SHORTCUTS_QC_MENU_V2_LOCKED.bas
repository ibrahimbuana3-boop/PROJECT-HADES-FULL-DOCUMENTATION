Attribute VB_Name = "HADES_5_SHORTCUTS_QC_MENU_V2_LOCKED"
Option Explicit

'=========================================================
' PROJECT HADES — 5 SHORTCUT WRAPPER + QC FINAL MENU V2 LOCKED
' CorelDRAW 2021 VBA
'
' FILE INI MENGGANTIKAN:
' - HADES_5_SHORTCUTS_PHASE5D_CURRENT.bas
' - HADES_5_SHORTCUTS_QC_MENU_V1.bas
'
' SHORTCUT UTAMA:
' 1. HADES_PRECHECK_MASTER
' 2. HADES_PREPARE_MASTER
' 3. HADES_EXECUTE_LAYOUT
' 4. HADES_QC_FINAL       <-- menu pilihan QC
' 5. HADES_FINALIZE_CONVERT
'
' PERUBAHAN V2:
' - Pilihan 1-5 menjalankan QC mandiri dan MENGUNCI / INVALIDATE Finalize.
' - Pilihan 6 menjalankan Global Final Report V5D.
' - Hanya Global Final Report PASS + ALLOWED yang boleh membuka Finalize Convert.
' - Jika salah satu modul FAIL di Global Final Report, HADES_FINALIZE_CONVERT tetap diblokir.
'
' CATATAN PENTING:
' - Standalone engine tetap harus sudah diimport.
' - Core/Phase Final Report tetap harus lengkap.
' - Jangan import berdampingan dengan module shortcut lama,
'   karena akan terjadi Duplicate Procedure untuk 5 shortcut utama.
'=========================================================

Private Const HQC_LOCK_LATEST As String = "HADES_FINAL_QC_LOCK_LATEST.txt"
Private Const HQC_REPORT_FOLDER As String = "HADES_REPORTS"

'=========================================================
' SHORTCUT 1 — PRECHECK MASTER
'=========================================================
Public Sub HADES_PRECHECK_MASTER()
    On Error GoTo ERR_HANDLER

    Call QC_TRANSPARENCY_POWERCLIP_CHECK

    Exit Sub

ERR_HANDLER:
    MsgBox "HADES_PRECHECK_MASTER gagal." & vbCrLf & vbCrLf & _
           "Pastikan QC_TRANSPARENCY_POWERCLIP_CHECK sudah di-import." & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, vbCritical, "PROJECT HADES"
End Sub

'=========================================================
' SHORTCUT 2 — PREPARE MASTER
'=========================================================
Public Sub HADES_PREPARE_MASTER()
    On Error GoTo ERR_HANDLER

    Call AUTO_ARRANGE_MASTER_POLA
    Call AUTO_RECONTOUR_PLACEHOLDER
    Call BUILD_TYPO_TEMPLATE

    Exit Sub

ERR_HANDLER:
    MsgBox "HADES_PREPARE_MASTER gagal." & vbCrLf & vbCrLf & _
           "Pastikan AUTO_ARRANGE, AUTO_RECONTOUR, dan BUILD_TYPO_TEMPLATE sudah di-import." & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, vbCritical, "PROJECT HADES"
End Sub

'=========================================================
' SHORTCUT 3 — EXECUTE LAYOUT
'=========================================================
Public Sub HADES_EXECUTE_LAYOUT()
    On Error GoTo ERR_HANDLER

    Call HADES_AUTO_DUPLICATE_V21
    Call QC_AUTO_RENAME
    Call AI_TEXT

    Exit Sub

ERR_HANDLER:
    MsgBox "HADES_EXECUTE_LAYOUT gagal." & vbCrLf & vbCrLf & _
           "Pastikan Auto Duplicate, Auto Rename, dan AI Text sudah di-import." & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, vbCritical, "PROJECT HADES"
End Sub

'=========================================================
' SHORTCUT 4 — QC FINAL MENU
'=========================================================
Public Sub HADES_QC_FINAL()
    HADES_QC_FINAL_MENU
End Sub

Public Sub HADES_QC_FINAL_MENU()

    Dim pilih As String

    On Error GoTo ERR_HANDLER

    pilih = InputBox( _
        "PILIH QC YANG INGIN DIJALANKAN:" & vbCrLf & vbCrLf & _
        "1 = Jalankan PowerClip & Transparency" & vbCrLf & _
        "2 = Jalankan IDPO Check" & vbCrLf & _
        "3 = Jalankan Size Check" & vbCrLf & _
        "4 = Jalankan Typo Check" & vbCrLf & _
        "5 = Jalankan Group Structure Check" & vbCrLf & _
        "6 = Jalankan semuanya + buat Final QC Lock" & vbCrLf & vbCrLf & _
        "CATATAN:" & vbCrLf & _
        "- Pilihan 1-5 hanya QC mandiri dan akan mengunci Finalize." & vbCrLf & _
        "- Pilihan 6 wajib PASS agar Finalize Convert bisa dibuka." & vbCrLf & vbCrLf & _
        "0 / kosong = batal", _
        "HADES QC FINAL MENU V2 LOCKED")

    pilih = Trim$(pilih)

    If pilih = "" Or pilih = "0" Then Exit Sub

    Select Case pilih

        Case "1"
            Call HQC_RunPowerClipTransparency

        Case "2"
            Call HQC_RunIDPO

        Case "3"
            Call HQC_RunSize

        Case "4"
            Call HQC_RunTypo

        Case "5"
            Call HQC_RunGroupStructure

        Case "6"
            Call HQC_RunGlobalFinalReport

        Case Else
            MsgBox "Pilihan tidak valid." & vbCrLf & vbCrLf & _
                   "Masukkan angka 1 sampai 6.", _
                   vbExclamation, "HADES QC FINAL MENU V2"
    End Select

    Exit Sub

ERR_HANDLER:
    MsgBox "HADES_QC_FINAL_MENU gagal." & vbCrLf & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, _
           vbCritical, "PROJECT HADES"
End Sub

'=========================================================
' QC MENU RUNNERS — INDIVIDUAL MODE
'=========================================================
Private Sub HQC_RunPowerClipTransparency()
    On Error GoTo ERR_HANDLER

    If Not HQC_HasSelection() Then Exit Sub
    Call HQC_InvalidateFinalQCLock("QC mandiri PowerClip & Transparency dijalankan. Finalize harus menunggu Global Final Report PASS.")
    Call QC_TRANSPARENCY_POWERCLIP_CHECK
    Call HQC_ShowIndividualDoneMessage("PowerClip & Transparency")

    Exit Sub

ERR_HANDLER:
    Call HQC_InvalidateFinalQCLock("PowerClip & Transparency Check error. Finalize diblokir.")
    MsgBox "PowerClip & Transparency Check gagal." & vbCrLf & vbCrLf & _
           "Pastikan module QC_TRANSPARENCY_POWERCLIP_CHECK sudah di-import." & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, _
           vbCritical, "HADES QC FINAL MENU V2"
End Sub

Private Sub HQC_RunIDPO()
    On Error GoTo ERR_HANDLER

    If Not HQC_HasSelection() Then Exit Sub
    Call HQC_InvalidateFinalQCLock("QC mandiri IDPO Check dijalankan. Finalize harus menunggu Global Final Report PASS.")
    Call IDPO_CHECK
    Call HQC_ShowIndividualDoneMessage("IDPO Check")

    Exit Sub

ERR_HANDLER:
    Call HQC_InvalidateFinalQCLock("IDPO Check error. Finalize diblokir.")
    MsgBox "IDPO Check gagal." & vbCrLf & vbCrLf & _
           "Pastikan module VBA IDPO CHECK sudah di-import." & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, _
           vbCritical, "HADES QC FINAL MENU V2"
End Sub

Private Sub HQC_RunSize()
    On Error GoTo ERR_HANDLER

    If Not HQC_HasSelection() Then Exit Sub
    Call HQC_InvalidateFinalQCLock("QC mandiri Size Check dijalankan. Finalize harus menunggu Global Final Report PASS.")
    Call QC_SIZE_CHECK
    Call HQC_ShowIndividualDoneMessage("Size Check")

    Exit Sub

ERR_HANDLER:
    Call HQC_InvalidateFinalQCLock("Size Check error. Finalize diblokir.")
    MsgBox "Size Check gagal." & vbCrLf & vbCrLf & _
           "Pastikan module VBA_QC_SIZE_CHECK sudah di-import." & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, _
           vbCritical, "HADES QC FINAL MENU V2"
End Sub

Private Sub HQC_RunTypo()
    On Error GoTo ERR_HANDLER

    If Not HQC_HasSelection() Then Exit Sub
    Call HQC_InvalidateFinalQCLock("QC mandiri Typo Check dijalankan. Finalize harus menunggu Global Final Report PASS.")
    Call QC_TYPO_CHECK
    Call HQC_ShowIndividualDoneMessage("Typo Check")

    Exit Sub

ERR_HANDLER:
    Call HQC_InvalidateFinalQCLock("Typo Check error. Finalize diblokir.")
    MsgBox "Typo Check gagal." & vbCrLf & vbCrLf & _
           "Pastikan module VBA_QC_TYPO_CHECK sudah di-import." & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, _
           vbCritical, "HADES QC FINAL MENU V2"
End Sub

Private Sub HQC_RunGroupStructure()
    On Error GoTo ERR_HANDLER

    If Not HQC_HasSelection() Then Exit Sub
    Call HQC_InvalidateFinalQCLock("QC mandiri Group Structure Check dijalankan. Finalize harus menunggu Global Final Report PASS.")
    Call GROUP_STRUCTURE_CHECK
    Call HQC_ShowIndividualDoneMessage("Group Structure Check")

    Exit Sub

ERR_HANDLER:
    Call HQC_InvalidateFinalQCLock("Group Structure Check error. Finalize diblokir.")
    MsgBox "Group Structure Check gagal." & vbCrLf & vbCrLf & _
           "Pastikan module GROUP_STRUCTURE_CHECK sudah di-import." & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, _
           vbCritical, "HADES QC FINAL MENU V2"
End Sub

'=========================================================
' QC RUN ALL — GLOBAL REPORT + LOCK
'=========================================================
Private Sub HQC_RunGlobalFinalReport()

    Dim ans As Long

    On Error GoTo ERR_HANDLER

    If Not HQC_HasSelection() Then Exit Sub

    ans = MsgBox( _
        "HADES QC FINAL - RUN ALL + LOCK" & vbCrLf & vbCrLf & _
        "Macro akan menjalankan Global Final Report:" & vbCrLf & _
        "1. Preflight data" & vbCrLf & _
        "2. PowerClip & Transparency" & vbCrLf & _
        "3. IDPO Check" & vbCrLf & _
        "4. Size Check" & vbCrLf & _
        "5. Typo Check" & vbCrLf & _
        "6. Group Structure Check" & vbCrLf & vbCrLf & _
        "Jika semua PASS, Final QC Lock dibuat ALLOWED." & vbCrLf & _
        "Jika ada satu saja FAIL, Finalize Convert tetap diblokir." & vbCrLf & vbCrLf & _
        "Lanjut?", _
        vbQuestion + vbYesNo, _
        "HADES QC FINAL MENU V2")

    If ans <> vbYes Then Exit Sub

    Call HADES_QC_FINAL_REPORT_V5D

    Exit Sub

ERR_HANDLER:
    Call HQC_InvalidateFinalQCLock("Global Final Report error. Finalize diblokir.")
    MsgBox "Global Final Report gagal / berhenti di tengah proses." & vbCrLf & vbCrLf & _
           "Pastikan Final QC Chain sudah lengkap:" & vbCrLf & _
           "- HADES_CORE_REPORT_PHASE2" & vbCrLf & _
           "- HADES_QC_FINAL_REPORT_PHASE3C" & vbCrLf & _
           "- HADES_QC_FINAL_REPORT_PHASE4" & vbCrLf & _
           "- HADES_REPORT_VIEWER_PHASE5D_FIX3" & vbCrLf & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, _
           vbCritical, "HADES QC FINAL MENU V2"
End Sub

'=========================================================
' OPTIONAL ALIAS — GLOBAL FINAL REPORT
'=========================================================
Public Sub HADES_QC_GLOBAL_FINAL_REPORT()
    On Error GoTo ERR_HANDLER

    If Not HQC_HasSelection() Then Exit Sub
    Call HADES_QC_FINAL_REPORT_V5D

    Exit Sub

ERR_HANDLER:
    Call HQC_InvalidateFinalQCLock("HADES_QC_GLOBAL_FINAL_REPORT error. Finalize diblokir.")
    MsgBox "HADES_QC_GLOBAL_FINAL_REPORT gagal." & vbCrLf & vbCrLf & _
           "Pastikan Final QC Chain sudah lengkap:" & vbCrLf & _
           "- HADES_CORE_REPORT_PHASE2" & vbCrLf & _
           "- HADES_QC_FINAL_REPORT_PHASE3C" & vbCrLf & _
           "- HADES_QC_FINAL_REPORT_PHASE4" & vbCrLf & _
           "- HADES_REPORT_VIEWER_PHASE5D_FIX3" & vbCrLf & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, _
           vbCritical, "PROJECT HADES"
End Sub

'=========================================================
' SHORTCUT 5 — FINALIZE CONVERT
'=========================================================
Public Sub HADES_FINALIZE_CONVERT()
    On Error GoTo ERR_HANDLER

    'HADES_FINALIZE_CONVERT_V5D akan memanggil gate Phase4.
    'Gate Phase4 hanya ALLOWED bila latest report PASS dan QC Lock PASS.
    Call HADES_FINALIZE_CONVERT_V5D

    Exit Sub

ERR_HANDLER:
    MsgBox "HADES_FINALIZE_CONVERT gagal." & vbCrLf & vbCrLf & _
           "Finalize hanya boleh berjalan setelah HADES_QC_FINAL pilihan 6 menghasilkan PASS / ALLOWED." & vbCrLf & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, vbCritical, "PROJECT HADES"
End Sub

'=========================================================
' LOCK INVALIDATOR
'=========================================================
Private Sub HQC_InvalidateFinalQCLock(ByVal reasonText As String)

    On Error Resume Next

    Dim folderPath As String
    Dim lockPath As String
    Dim f As Integer

    folderPath = Environ$("USERPROFILE") & "\Documents\" & HQC_REPORT_FOLDER

    If Dir$(folderPath, vbDirectory) = "" Then
        MkDir folderPath
    End If

    lockPath = folderPath & "\" & HQC_LOCK_LATEST

    f = FreeFile
    Open lockPath For Output As #f

    Print #f, "# PROJECT_HADES_QC_LOCK"
    Print #f, "LOCK_VERSION=MENU_V2_INVALID"
    Print #f, "LOCK_TYPE=FINAL_QC"
    Print #f, "LOCK_CREATED=" & Format$(Now, "yyyy-mm-dd hh:nn:ss")
    Print #f, "FINAL_STATUS=INVALID"
    Print #f, "CONVERT_PERMISSION=BLOCKED"
    Print #f, "REPORT_PATH="
    Print #f, "REPORT_GENERATED="
    Print #f, "DOCUMENT_FULLNAME=" & HQC_CurrentDocumentFullName()
    Print #f, "SELECTION_SIGNATURE=INVALIDATED_BY_QC_MENU"
    Print #f, "REASON=" & Replace(Replace(reasonText, vbCr, " "), vbLf, " ")
    Print #f, "# END_PROJECT_HADES_QC_LOCK"
    Print #f, ""
    Print #f, "PROJECT HADES — FINAL QC LOCK INVALIDATED BY QC MENU V2"
    Print #f, "Reason: " & reasonText
    Print #f, ""
    Print #f, "Untuk membuka Finalize Convert:"
    Print #f, "1. Select layout final yang sama."
    Print #f, "2. Jalankan HADES_QC_FINAL."
    Print #f, "3. Pilih 6 = Jalankan semuanya + buat Final QC Lock."
    Print #f, "4. Pastikan status PASS / ALLOWED."

    Close #f

    On Error GoTo 0

End Sub

Private Function HQC_CurrentDocumentFullName() As String
    On Error GoTo FAIL

    If ActiveDocument Is Nothing Then GoTo FAIL
    HQC_CurrentDocumentFullName = ActiveDocument.FullFileName
    Exit Function

FAIL:
    HQC_CurrentDocumentFullName = ""
End Function

Private Sub HQC_ShowIndividualDoneMessage(ByVal qcName As String)
    MsgBox qcName & " selesai dijalankan sebagai QC mandiri." & vbCrLf & vbCrLf & _
           "FINALIZE CONVERT sekarang tetap DIBLOKIR sampai:" & vbCrLf & _
           "HADES_QC_FINAL -> pilih 6 -> Global Final Report PASS / ALLOWED.", _
           vbInformation, "HADES QC FINAL MENU V2"
End Sub

'=========================================================
' HELPER
'=========================================================
Private Function HQC_HasSelection() As Boolean
    On Error GoTo NO_SELECTION

    If ActiveDocument Is Nothing Then GoTo NO_SELECTION
    If ActiveSelection Is Nothing Then GoTo NO_SELECTION
    If ActiveSelection.Shapes.Count = 0 Then GoTo NO_SELECTION

    HQC_HasSelection = True
    Exit Function

NO_SELECTION:
    HQC_HasSelection = False
    MsgBox "Pilih / block hasil layout yang ingin di-QC dulu.", _
           vbExclamation, "HADES QC FINAL MENU V2"
End Function
