Option Explicit

'=========================================================
' PROJECT HADES — GROUP STRUCTURE CHECK V1
'
' PURPOSE:
' Mengecek apakah hasil layout sudah digroup dengan benar.
'
' Struktur ideal:
'
' GROUP 1 SET / 1 SIZE
' ¦
' +-- GROUP Baju Depan
' +-- GROUP Baju Belakang
' +-- GROUP Lengan Kanan
' +-- GROUP Lengan Kiri
' +-- GROUP Kerah / komponen lain
'
' Macro ini mengecek:
' 1. Semua objek level atas harus GROUP.
' 2. Isi langsung dalam group utama harus GROUP.
' 3. Jika ada curve/text/bitmap langsung di dalam group utama,
'    berarti kemungkinan ada komponen yang belum digroup.
'
' TIDAK memakai:
' - Order.txt
' - SizeDB
' - AttributeDB
'
' MAIN MACRO:
' GROUP_STRUCTURE_CHECK
'
' ALIAS:
' UNGROUPED_CHECK
'=========================================================

Private GSC_TopTotal As Long
Private GSC_TopGroup As Long
Private GSC_TopLoose As Long

Private GSC_SetGroupChecked As Long
Private GSC_SetGroupFailed As Long
Private GSC_SetGroupWarning As Long

Private GSC_TotalLooseInside As Long

Private GSC_ReportTopLoose As String
Private GSC_ReportInsideLoose As String
Private GSC_ReportWarning As String

Private Const GSC_MIN_CHILD_GROUP As Long = 2

'=========================================================
' PUBLIC ENTRY
'=========================================================

Sub GROUP_STRUCTURE_CHECK()

On Error GoTo ERR_HANDLER

Dim sr As ShapeRange

On Error Resume Next
Set sr = ActiveSelectionRange
On Error GoTo ERR_HANDLER

If sr Is Nothing Then

    MsgBox _
        "Pilih HASIL LAYOUT terlebih dahulu.", _
        vbExclamation, _
        "Group Structure Check"

    Exit Sub

End If

If sr.count = 0 Then

    MsgBox _
        "Pilih HASIL LAYOUT terlebih dahulu.", _
        vbExclamation, _
        "Group Structure Check"

    Exit Sub

End If

GSC_Reset

Dim s As Shape
Dim topIndex As Long

topIndex = 0

For Each s In sr

    topIndex = topIndex + 1
    GSC_TopTotal = GSC_TopTotal + 1

    If s.Type = cdrGroupShape Then

        GSC_TopGroup = GSC_TopGroup + 1
        GSC_CheckMainGroup s, GSC_TopGroup, topIndex

    Else

        GSC_TopLoose = GSC_TopLoose + 1

        GSC_ReportTopLoose = GSC_ReportTopLoose & _
            "- TOP OBJECT #" & topIndex & _
            " bukan GROUP" & _
            " | Type: " & GSC_ShapeTypeName(s) & _
            " | Size: " & GSC_SizeInfo(s) & vbCrLf

    End If

Next s

GSC_ShowReport

Exit Sub

ERR_HANDLER:

MsgBox _
    "SYSTEM ERROR - GROUP STRUCTURE CHECK" & vbCrLf & vbCrLf & _
    "No : " & Err.Number & vbCrLf & _
    Err.Description, _
    vbCritical, _
    "Group Structure Check"

End Sub

Sub UNGROUPED_CHECK()

Call GROUP_STRUCTURE_CHECK

End Sub

'=========================================================
' RESET
'=========================================================

Private Sub GSC_Reset()

GSC_TopTotal = 0
GSC_TopGroup = 0
GSC_TopLoose = 0

GSC_SetGroupChecked = 0
GSC_SetGroupFailed = 0
GSC_SetGroupWarning = 0

GSC_TotalLooseInside = 0

GSC_ReportTopLoose = ""
GSC_ReportInsideLoose = ""
GSC_ReportWarning = ""

End Sub

'=========================================================
' CHECK MAIN GROUP
'=========================================================

Private Sub GSC_CheckMainGroup(ByVal g As Shape, ByVal groupIndex As Long, ByVal topIndex As Long)

Dim child As Shape

Dim childTotal As Long
Dim childGroup As Long
Dim childLoose As Long

Dim localLooseReport As String

childTotal = 0
childGroup = 0
childLoose = 0
localLooseReport = ""

GSC_SetGroupChecked = GSC_SetGroupChecked + 1

For Each child In g.Shapes

    childTotal = childTotal + 1

    If child.Type = cdrGroupShape Then

        childGroup = childGroup + 1

    Else

        childLoose = childLoose + 1

        localLooseReport = localLooseReport & _
            "    - Child #" & childTotal & _
            " | Type: " & GSC_ShapeTypeName(child) & _
            " | Size: " & GSC_SizeInfo(child) & vbCrLf

    End If

Next child

If childLoose > 0 Then

    GSC_SetGroupFailed = GSC_SetGroupFailed + 1
    GSC_TotalLooseInside = GSC_TotalLooseInside + childLoose

    GSC_ReportInsideLoose = GSC_ReportInsideLoose & _
        "GROUP SET #" & groupIndex & _
        " punya " & childLoose & _
        " objek langsung yang bukan group." & vbCrLf & _
        "Child Group : " & childGroup & vbCrLf & _
        "Child Total : " & childTotal & vbCrLf & _
        localLooseReport & _
        String(45, "-") & vbCrLf

Else

    If childGroup < GSC_MIN_CHILD_GROUP Then

        GSC_SetGroupWarning = GSC_SetGroupWarning + 1

        GSC_ReportWarning = GSC_ReportWarning & _
            "GROUP SET #" & groupIndex & _
            " hanya punya " & childGroup & _
            " child group." & vbCrLf & _
            "Kemungkinan yang dipilih bukan group 1 set, atau struktur group terlalu sederhana." & vbCrLf & _
            String(45, "-") & vbCrLf

    End If

End If

End Sub

'=========================================================
' REPORT
'=========================================================

Private Sub GSC_ShowReport()

Dim report As String
Dim isPass As Boolean

isPass = False

If GSC_TopLoose = 0 _
   And GSC_SetGroupFailed = 0 Then

    isPass = True

End If

If isPass Then

    report = _
        "GROUP STRUCTURE CHECK PASSED" & vbCrLf & vbCrLf & _
        "Top Object       : " & GSC_TopTotal & vbCrLf & _
        "Top Group        : " & GSC_TopGroup & vbCrLf & _
        "Objek lepas atas : " & GSC_TopLoose & vbCrLf & _
        "Group dicek      : " & GSC_SetGroupChecked & vbCrLf & vbCrLf & _
        "Tidak ada objek lepas di level atas maupun langsung di dalam group set."

    If GSC_SetGroupWarning > 0 Then

        report = report & vbCrLf & vbCrLf & _
            "WARNING:" & vbCrLf & _
            GSC_ReportWarning

        MsgBox _
            report, _
            vbExclamation, _
            "Group Structure Check"

    Else

        MsgBox _
            report, _
            vbInformation, _
            "Group Structure Check"

    End If

Else

    report = _
        "GROUP STRUCTURE CHECK FAILED" & vbCrLf & vbCrLf & _
        "Top Object             : " & GSC_TopTotal & vbCrLf & _
        "Top Group              : " & GSC_TopGroup & vbCrLf & _
        "Objek lepas level atas : " & GSC_TopLoose & vbCrLf & _
        "Group dicek            : " & GSC_SetGroupChecked & vbCrLf & _
        "Group bermasalah       : " & GSC_SetGroupFailed & vbCrLf & _
        "Objek lepas di dalam   : " & GSC_TotalLooseInside & vbCrLf & vbCrLf

    If GSC_ReportTopLoose <> "" Then

        report = report & _
            "OBJEK LEVEL ATAS YANG BELUM MASUK GROUP:" & vbCrLf & _
            GSC_ReportTopLoose & vbCrLf

    End If

    If GSC_ReportInsideLoose <> "" Then

        report = report & _
            "OBJEK LANGSUNG DI DALAM GROUP SET YANG BELUM DIGROUP:" & vbCrLf & _
            GSC_ReportInsideLoose & vbCrLf

    End If

    If GSC_ReportWarning <> "" Then

        report = report & _
            "WARNING:" & vbCrLf & _
            GSC_ReportWarning & vbCrLf

    End If

    report = report & _
        "Silakan cek group yang FAILED. Kemungkinan ada komponen hasil intersect yang belum digroup."

    MsgBox _
        report, _
        vbCritical, _
        "Group Structure Check"

End If

End Sub

'=========================================================
' HELPERS
'=========================================================

Private Function GSC_ShapeTypeName(ByVal s As Shape) As String

Select Case s.Type

    Case cdrGroupShape
        GSC_ShapeTypeName = "GROUP"

    Case cdrCurveShape
        GSC_ShapeTypeName = "CURVE"

    Case cdrTextShape
        GSC_ShapeTypeName = "TEXT"

    Case cdrBitmapShape
        GSC_ShapeTypeName = "BITMAP"

    Case cdrRectangleShape
        GSC_ShapeTypeName = "RECTANGLE"

    Case cdrEllipseShape
        GSC_ShapeTypeName = "ELLIPSE"

    Case Else
        GSC_ShapeTypeName = "TYPE " & CStr(s.Type)

End Select

End Function

Private Function GSC_SizeInfo(ByVal s As Shape) As String

On Error Resume Next

GSC_SizeInfo = _
    FormatNumber(Abs(s.SizeWidth), 2) & _
    " x " & _
    FormatNumber(Abs(s.SizeHeight), 2) & _
    " cm"

On Error GoTo 0

End Function
