Attribute VB_Name = "HADES_REPORT_VIEWER_PHASE5D_FIX3"
Option Explicit

'=========================================================
' PROJECT HADES — PHASE 5D FIX3
' IN-COREL GLOBAL FINAL REPORT VIEWER WITHOUT .FRM — NO V5C DIRECT DEPENDENCY
' CorelDRAW 2021 VBA
'
' PURPOSE:
' - Replaces the broken UserForm/.frm viewer.
' - Does NOT require importing any .frm file.
' - Uses a native Windows EDIT window from pure .bas VBA.
' - Avoids compile error caused by VERSION 5.00 in .frm.
' - Avoids direct compile dependency on HADES_QC_FINAL_REPORT_V5C.
'
' REQUIRED EXISTING MACROS:
' - HADES_QC_FINAL_REPORT_V4
' - HADES_FINALIZE_CONVERT_V4
'
' NOTE:
' - This still uses your latest imported HADES_QC_SIZE_REPORT and
'   HADES_QC_TYPO_REPORT because V4/V3C calls those public validators.
' - Phase 5C modules may remain installed, but FIX3 does not require the
'   wrapper name HADES_QC_FINAL_REPORT_V5C to exist at compile time.
'
' SHORTCUTS:
' - HADES_QC_FINAL_REPORT_V5D
' - HADES_FINALIZE_CONVERT_V5D
' - HADES5D_ShowGlobalFinalReport
'=========================================================

'=========================================================
' WINDOWS API DECLARATIONS
'=========================================================

#If VBA7 Then

    Private Declare PtrSafe Function CreateWindowExW Lib "user32" ( _
        ByVal dwExStyle As Long, _
        ByVal lpClassName As LongPtr, _
        ByVal lpWindowName As LongPtr, _
        ByVal dwStyle As Long, _
        ByVal x As Long, _
        ByVal y As Long, _
        ByVal nWidth As Long, _
        ByVal nHeight As Long, _
        ByVal hWndParent As LongPtr, _
        ByVal hMenu As LongPtr, _
        ByVal hInstance As LongPtr, _
        ByVal lpParam As LongPtr) As LongPtr

    Private Declare PtrSafe Function SetWindowTextW Lib "user32" ( _
        ByVal hwnd As LongPtr, _
        ByVal lpString As LongPtr) As Long

    Private Declare PtrSafe Function ShowWindow Lib "user32" ( _
        ByVal hwnd As LongPtr, _
        ByVal nCmdShow As Long) As Long

    Private Declare PtrSafe Function SetForegroundWindow Lib "user32" ( _
        ByVal hwnd As LongPtr) As Long

    Private Declare PtrSafe Function SendMessageW Lib "user32" ( _
        ByVal hwnd As LongPtr, _
        ByVal wMsg As Long, _
        ByVal wParam As LongPtr, _
        ByVal lParam As LongPtr) As LongPtr

    Private Declare PtrSafe Function GetStockObject Lib "gdi32" ( _
        ByVal nIndex As Long) As LongPtr

#Else

    Private Declare Function CreateWindowExW Lib "user32" ( _
        ByVal dwExStyle As Long, _
        ByVal lpClassName As Long, _
        ByVal lpWindowName As Long, _
        ByVal dwStyle As Long, _
        ByVal x As Long, _
        ByVal y As Long, _
        ByVal nWidth As Long, _
        ByVal nHeight As Long, _
        ByVal hWndParent As Long, _
        ByVal hMenu As Long, _
        ByVal hInstance As Long, _
        ByVal lpParam As Long) As Long

    Private Declare Function SetWindowTextW Lib "user32" ( _
        ByVal hwnd As Long, _
        ByVal lpString As Long) As Long

    Private Declare Function ShowWindow Lib "user32" ( _
        ByVal hwnd As Long, _
        ByVal nCmdShow As Long) As Long

    Private Declare Function SetForegroundWindow Lib "user32" ( _
        ByVal hwnd As Long) As Long

    Private Declare Function SendMessageW Lib "user32" ( _
        ByVal hwnd As Long, _
        ByVal wMsg As Long, _
        ByVal wParam As Long, _
        ByVal lParam As Long) As Long

    Private Declare Function GetStockObject Lib "gdi32" ( _
        ByVal nIndex As Long) As Long

#End If

'=========================================================
' WINDOWS CONSTANTS
'=========================================================

Private Const WS_OVERLAPPED As Long = &H0&
Private Const WS_CAPTION As Long = &HC00000
Private Const WS_SYSMENU As Long = &H80000
Private Const WS_THICKFRAME As Long = &H40000
Private Const WS_MINIMIZEBOX As Long = &H20000
Private Const WS_MAXIMIZEBOX As Long = &H10000
Private Const WS_VISIBLE As Long = &H10000000
Private Const WS_VSCROLL As Long = &H200000
Private Const WS_HSCROLL As Long = &H100000

Private Const ES_MULTILINE As Long = &H4&
Private Const ES_AUTOVSCROLL As Long = &H40&
Private Const ES_AUTOHSCROLL As Long = &H80&
Private Const ES_READONLY As Long = &H800&
Private Const ES_WANTRETURN As Long = &H1000&

Private Const WS_EX_APPWINDOW As Long = &H40000
Private Const WS_EX_CLIENTEDGE As Long = &H200

Private Const SW_SHOW As Long = 5
Private Const WM_SETFONT As Long = &H30
Private Const DEFAULT_GUI_FONT As Long = 17

Private Const REPORT_FOLDER As String = "\Documents\HADES_REPORTS"
Private Const REPORT_LATEST As String = "HADES_FINAL_QC_REPORT_LATEST.txt"

#If VBA7 Then
    Private gViewerHwnd As LongPtr
#Else
    Private gViewerHwnd As Long
#End If

Private gLastReportPath As String

'=========================================================
' MAIN WRAPPERS
'=========================================================

Public Sub HADES_QC_FINAL_REPORT_V5D()

    On Error GoTo ErrHandler

    'Run stable hard-gate final QC.
    'Important: we call V4 directly to avoid compile error when V5C wrapper
    'is not present. V4 calls V3C, and V3C calls the currently imported
    'HADES_QC_TYPO_REPORT / HADES_QC_SIZE_REPORT validators.
    HADES_QC_FINAL_REPORT_V4

    'Then show global final report inside a lightweight viewer.
    HADES5D_ShowGlobalFinalReport

    Exit Sub

ErrHandler:

    MsgBox _
        "HADES_QC_FINAL_REPORT_V5D gagal." & vbCrLf & vbCrLf & _
        "Error " & Err.Number & ":" & vbCrLf & Err.Description, _
        vbCritical, _
        "HADES PHASE 5D FIX3"

End Sub


Public Sub HADES_FINALIZE_CONVERT_V5D()

    On Error GoTo ErrHandler

    'Alias aman ke convert gate hardening Phase 4.
    HADES_FINALIZE_CONVERT_V4

    Exit Sub

ErrHandler:

    MsgBox _
        "HADES_FINALIZE_CONVERT_V5D gagal." & vbCrLf & vbCrLf & _
        "Error " & Err.Number & ":" & vbCrLf & Err.Description, _
        vbCritical, _
        "HADES PHASE 5D FIX3"

End Sub


Public Sub HADES5D_ShowGlobalFinalReport()

    Dim reportPath As String
    Dim body As String
    Dim displayText As String

    On Error GoTo ErrHandler

    reportPath = H5D_FindLatestFinalReportPath()

    If Len(reportPath) = 0 Then
        MsgBox _
            "Global final report belum ditemukan." & vbCrLf & vbCrLf & _
            "Jalankan HADES_QC_FINAL_REPORT_V5C / V5D terlebih dahulu.", _
            vbExclamation, _
            "HADES REPORT VIEWER"
        Exit Sub
    End If

    gLastReportPath = reportPath
    body = H5D_ReadTextFileUTF8(reportPath)

    displayText = "PROJECT HADES - GLOBAL FINAL QC REPORT" & vbCrLf & _
                  "File: " & reportPath & vbCrLf & _
                  String(72, "=") & vbCrLf & _
                  "Tips: Klik di area teks, tekan CTRL+A lalu CTRL+C untuk copy report." & vbCrLf & _
                  String(72, "=") & vbCrLf & vbCrLf & _
                  body

    H5D_ShowTextWindow displayText

    Exit Sub

ErrHandler:

    MsgBox _
        "Gagal menampilkan global final report." & vbCrLf & vbCrLf & _
        "Error " & Err.Number & ":" & vbCrLf & Err.Description, _
        vbCritical, _
        "HADES REPORT VIEWER"

End Sub


Public Sub HADES5D_OpenLatestReportFolder()

    Dim folderPath As String

    folderPath = Environ$("USERPROFILE") & REPORT_FOLDER

    If Dir(folderPath, vbDirectory) = "" Then
        MsgBox "Folder report belum ada:" & vbCrLf & folderPath, vbExclamation, "HADES REPORTS"
        Exit Sub
    End If

    Shell "explorer.exe """ & folderPath & """", vbNormalFocus

End Sub


Public Sub HADES5D_OpenLatestReportInNotepad()

    Dim reportPath As String

    reportPath = H5D_FindLatestFinalReportPath()

    If Len(reportPath) = 0 Then
        MsgBox "Report belum ditemukan.", vbExclamation, "HADES REPORTS"
        Exit Sub
    End If

    Shell "notepad.exe """ & reportPath & """", vbNormalFocus

End Sub

'=========================================================
' PURE .BAS TEXT VIEWER
'=========================================================

Private Sub H5D_ShowTextWindow(ByVal textValue As String)

    Dim className As String
    Dim titleText As String
    Dim style As Long
    Dim exStyle As Long

#If VBA7 Then
    Dim hFont As LongPtr
#Else
    Dim hFont As Long
#End If

    className = "EDIT"
    titleText = "PROJECT HADES - FINAL QC REPORT"

    style = WS_OVERLAPPED Or WS_CAPTION Or WS_SYSMENU Or WS_THICKFRAME Or _
            WS_MINIMIZEBOX Or WS_MAXIMIZEBOX Or WS_VISIBLE Or _
            WS_VSCROLL Or WS_HSCROLL Or _
            ES_MULTILINE Or ES_AUTOVSCROLL Or ES_AUTOHSCROLL Or _
            ES_READONLY Or ES_WANTRETURN

    exStyle = WS_EX_APPWINDOW Or WS_EX_CLIENTEDGE

    gViewerHwnd = CreateWindowExW( _
                    exStyle, _
                    StrPtr(className), _
                    StrPtr(titleText), _
                    style, _
                    120, _
                    80, _
                    920, _
                    680, _
                    0, _
                    0, _
                    0, _
                    0)

    If gViewerHwnd = 0 Then
        MsgBox _
            "Viewer Windows gagal dibuat." & vbCrLf & vbCrLf & _
            "Fallback: report tetap tersimpan di Documents\HADES_REPORTS.", _
            vbExclamation, _
            "HADES REPORT VIEWER"
        Exit Sub
    End If

    hFont = GetStockObject(DEFAULT_GUI_FONT)
    If hFont <> 0 Then
        SendMessageW gViewerHwnd, WM_SETFONT, hFont, 1
    End If

    SetWindowTextW gViewerHwnd, StrPtr(textValue)
    ShowWindow gViewerHwnd, SW_SHOW
    SetForegroundWindow gViewerHwnd

End Sub

'=========================================================
' REPORT FILE DISCOVERY
'=========================================================

Private Function H5D_FindLatestFinalReportPath() As String

    Dim folderPath As String
    Dim p As String

    folderPath = Environ$("USERPROFILE") & REPORT_FOLDER

    p = folderPath & "\" & REPORT_LATEST

    If Dir(p) <> "" Then
        H5D_FindLatestFinalReportPath = p
        Exit Function
    End If

    'Fallback: find newest text report containing FINAL_QC_REPORT.
    H5D_FindLatestFinalReportPath = H5D_FindNewestMatchingTextFile(folderPath, "FINAL_QC_REPORT")

    If Len(H5D_FindLatestFinalReportPath) > 0 Then Exit Function

    'Last fallback: newest .txt in HADES_REPORTS.
    H5D_FindLatestFinalReportPath = H5D_FindNewestMatchingTextFile(folderPath, "")

End Function


Private Function H5D_FindNewestMatchingTextFile( _
    ByVal folderPath As String, _
    ByVal nameMustContain As String) As String

    Dim f As String
    Dim fullPath As String
    Dim bestPath As String
    Dim bestDate As Date
    Dim d As Date

    If Dir(folderPath, vbDirectory) = "" Then Exit Function

    f = Dir(folderPath & "\*.txt")

    Do While Len(f) > 0

        If Len(nameMustContain) = 0 Or _
           InStr(1, UCase$(f), UCase$(nameMustContain), vbTextCompare) > 0 Then

            fullPath = folderPath & "\" & f

            On Error Resume Next
            d = FileDateTime(fullPath)
            On Error GoTo 0

            If Len(bestPath) = 0 Or d > bestDate Then
                bestDate = d
                bestPath = fullPath
            End If

        End If

        f = Dir()

    Loop

    H5D_FindNewestMatchingTextFile = bestPath

End Function

'=========================================================
' UTF-8 READER
'=========================================================

Private Function H5D_ReadTextFileUTF8(ByVal path As String) As String

    Dim stm As Object

    On Error GoTo FallbackRead

    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 2
    stm.Charset = "utf-8"
    stm.Open
    stm.LoadFromFile path

    H5D_ReadTextFileUTF8 = stm.ReadText

    stm.Close
    Set stm = Nothing

    Exit Function

FallbackRead:

    On Error Resume Next
    If Not stm Is Nothing Then stm.Close
    Set stm = Nothing
    On Error GoTo 0

    Dim f As Integer
    Dim ln As String
    Dim result As String

    f = FreeFile

    Open path For Input As #f

    Do Until EOF(f)
        Line Input #f, ln
        result = result & ln & vbCrLf
    Loop

    Close #f

    H5D_ReadTextFileUTF8 = result

End Function
