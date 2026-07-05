Option Explicit

'=========================================================
' HADES - AUTO RE-CONTOUR PLACEHOLDER V2.2
' CorelDRAW 2021 VBA
'
' PERBAIKAN V2:
' - Contour baru TIDAK di-Separate / Break Apart.
' - Contour baru dibiarkan sebagai LIVE CONTOUR EFFECT.
' - Saat placeholder diganti / diedit, contour ikut berubah.
'
' TAMBAHAN V2.1:
' - Support placeholder nomor celana berupa digit 1-2 angka
' - Interval tinggi nomor celana: 6 cm - 8 cm
'
' REVISI V2.2:
' - Default bentuk contour: ROUNDED CORNERS
' - Menggunakan cdrContourCornerRound
'
' FUNGSI:
' - Scan selection / master desain
' - Cari placeholder active text
' - Abaikan IDPO
' - Ambil warna contour lama dari curve dalam group yang sama
' - Hapus contour lama yang sudah berupa curve / shape biasa
' - Buat contour baru sesuai interval tinggi text
' - Contour baru tetap live effect
'
' SYARAT:
' 1. Placeholder masih ACTIVE TEXT
' 2. Contour lama berada satu group dengan placeholder
' 3. Group placeholder tidak bercampur logo/motif
' 4. Unit akan dipaksa sementara ke centimeter
'
' CARA PAKAI:
' Select all master desain -> Run AUTO_RECONTOUR_PLACEHOLDER
'=========================================================


'=========================
' INTERVAL TINGGI TEXT
'=========================

Private Const ARC_BACK_NAME_MIN As Double = 3#
Private Const ARC_BACK_NAME_MAX As Double = 8#
Private Const ARC_BACK_NAME_CONTOUR As Double = 0.2

Private Const ARC_BACK_NUMBER_MIN As Double = 16#
Private Const ARC_BACK_NUMBER_MAX As Double = 30#
Private Const ARC_BACK_NUMBER_CONTOUR As Double = 0.4

Private Const ARC_CHEST_NAME_MIN As Double = 0.8
Private Const ARC_CHEST_NAME_MAX As Double = 3#
Private Const ARC_CHEST_NAME_CONTOUR As Double = 0.1

Private Const ARC_CHEST_NUMBER_MIN As Double = 9#
Private Const ARC_CHEST_NUMBER_MAX As Double = 13#
Private Const ARC_CHEST_NUMBER_CONTOUR As Double = 0.3

Private Const ARC_PANTS_NUMBER_MIN As Double = 6#
Private Const ARC_PANTS_NUMBER_MAX As Double = 8#
Private Const ARC_PANTS_NUMBER_CONTOUR As Double = 0.2


'=========================
' CATEGORY
'=========================

Private Const ARC_CAT_NONE As Long = 0
Private Const ARC_CAT_BACK_NAME As Long = 1
Private Const ARC_CAT_BACK_NUMBER As Long = 2
Private Const ARC_CAT_CHEST_NAME As Long = 3
Private Const ARC_CAT_CHEST_NUMBER As Long = 4
Private Const ARC_CAT_PANTS_NUMBER As Long = 5


'=========================
' COUNTERS
'=========================

Private arcProcessed As Long
Private arcBackName As Long
Private arcBackNumber As Long
Private arcChestName As Long
Private arcChestNumber As Long
Private arcPantsNumber As Long

Private arcIDPOSkipped As Long
Private arcUnclassified As Long
Private arcNoColor As Long
Private arcMultiPlaceholder As Long
Private arcFailed As Long
Private arcJobs As Long

Private arcWarnings As String
Private Const ARC_WARN_LIMIT As Long = 25


'=========================================================
' MAIN MACRO
'=========================================================

Sub AUTO_RECONTOUR_PLACEHOLDER()

    Dim oldUnit As Long
    Dim sr As ShapeRange
    Dim s As Shape
    Dim jobs As Collection
    Dim i As Long
    Dim jobShape As Shape

    On Error GoTo ErrHandler

    If ActiveSelection.Shapes.Count = 0 Then
        MsgBox "Tidak ada objek dipilih." & vbCrLf & _
               "Select all master desain terlebih dahulu.", _
               vbExclamation, "AUTO RE-CONTOUR"
        Exit Sub
    End If

    Set sr = ActiveSelectionRange
    Set jobs = New Collection

    ARC_Reset

    oldUnit = ActiveDocument.Unit
    ActiveDocument.Unit = cdrCentimeter

    Application.Optimization = True
    ActiveDocument.BeginCommandGroup "Hades Auto Re-Contour Placeholder V2.2"

    'Pre-scan untuk report IDPO dan text yang mencurigakan
    For Each s In sr.Shapes
        ARC_PreScanShape s
    Next s

    'Kumpulkan group placeholder
    For Each s In sr.Shapes
        ARC_CollectJobs s, jobs
    Next s

    arcJobs = jobs.Count

    'Eksekusi re-contour
    For i = 1 To jobs.Count
        Set jobShape = jobs(i)
        ARC_ProcessGroupPlaceholder jobShape
    Next i

    ActiveDocument.EndCommandGroup
    Application.Optimization = False
    ActiveDocument.Unit = oldUnit
    ActiveWindow.Refresh

    ARC_ShowReport

    Exit Sub

ErrHandler:
    On Error Resume Next
    ActiveDocument.EndCommandGroup
    Application.Optimization = False
    ActiveDocument.Unit = oldUnit
    ActiveWindow.Refresh

    MsgBox "AUTO RE-CONTOUR ERROR" & vbCrLf & vbCrLf & _
           "Error " & Err.Number & ":" & vbCrLf & Err.Description, _
           vbCritical, "AUTO RE-CONTOUR"

End Sub


'=========================================================
' RESET
'=========================================================

Private Sub ARC_Reset()

    arcProcessed = 0
    arcBackName = 0
    arcBackNumber = 0
    arcChestName = 0
    arcChestNumber = 0
    arcPantsNumber = 0

    arcIDPOSkipped = 0
    arcUnclassified = 0
    arcNoColor = 0
    arcMultiPlaceholder = 0
    arcFailed = 0
    arcJobs = 0

    arcWarnings = ""

End Sub


'=========================================================
' PRE-SCAN TEXT
'=========================================================

Private Sub ARC_PreScanShape(ByVal s As Shape)

    Dim ch As Shape
    Dim t As String
    Dim cat As Long
    Dim offset As Double

    On Error Resume Next

    If s.Type = cdrTextShape Then

        t = ARC_GetText(s)

        If ARC_IsIDPO(t) Then
            arcIDPOSkipped = arcIDPOSkipped + 1
        ElseIf ARC_IsPlaceholderCandidate(t) Then
            If Not ARC_ClassifyText(s, cat, offset) Then
                arcUnclassified = arcUnclassified + 1
                ARC_AddWarning "Unclassified text: """ & ARC_ShortText(t) & _
                               """ | H=" & FormatNumber(s.SizeHeight, 2) & " cm"
            End If
        End If

    ElseIf s.Type = cdrGroupShape Then

        For Each ch In s.Shapes
            ARC_PreScanShape ch
        Next ch

    Else

        ARC_PreScanPowerClip s

    End If

End Sub


Private Sub ARC_PreScanPowerClip(ByVal s As Shape)

    Dim pcShapes As Shapes
    Dim ch As Shape

    On Error Resume Next

    Set pcShapes = s.PowerClip.Shapes

    If Not pcShapes Is Nothing Then
        For Each ch In pcShapes
            ARC_PreScanShape ch
        Next ch
    End If

End Sub


'=========================================================
' COLLECT JOBS
'=========================================================

Private Sub ARC_CollectJobs(ByVal s As Shape, ByRef jobs As Collection)

    Dim ch As Shape

    On Error Resume Next

    If s.Type = cdrGroupShape Then

        If ARC_GroupHasDirectPlaceholder(s) Then
            jobs.Add s
        Else
            For Each ch In s.Shapes
                ARC_CollectJobs ch, jobs
            Next ch
        End If

    Else

        ARC_CollectJobsPowerClip s, jobs

    End If

End Sub


Private Sub ARC_CollectJobsPowerClip(ByVal s As Shape, ByRef jobs As Collection)

    Dim pcShapes As Shapes
    Dim ch As Shape

    On Error Resume Next

    Set pcShapes = s.PowerClip.Shapes

    If Not pcShapes Is Nothing Then
        For Each ch In pcShapes
            ARC_CollectJobs ch, jobs
        Next ch
    End If

End Sub


Private Function ARC_GroupHasDirectPlaceholder(ByVal grp As Shape) As Boolean

    Dim ch As Shape
    Dim cat As Long
    Dim offset As Double

    On Error Resume Next

    ARC_GroupHasDirectPlaceholder = False

    For Each ch In grp.Shapes
        If ch.Type = cdrTextShape Then
            If ARC_ClassifyText(ch, cat, offset) Then
                ARC_GroupHasDirectPlaceholder = True
                Exit Function
            End If
        End If
    Next ch

End Function


'=========================================================
' PROCESS GROUP
'=========================================================

Private Sub ARC_ProcessGroupPlaceholder(ByVal grp As Shape)

    Dim txt As Shape
    Dim ch As Shape
    Dim placeholderCount As Long
    Dim cat As Long
    Dim offset As Double
    Dim oldColor As Color
    Dim eff As Effect

    On Error GoTo FailedProcess

    placeholderCount = 0
    Set txt = Nothing

    'Cari direct child berupa placeholder active text
    For Each ch In grp.Shapes
        If ch.Type = cdrTextShape Then
            If ARC_ClassifyText(ch, cat, offset) Then
                placeholderCount = placeholderCount + 1
                Set txt = ch
            End If
        End If
    Next ch

    If placeholderCount = 0 Then
        Exit Sub
    End If

    If placeholderCount > 1 Then
        arcMultiPlaceholder = arcMultiPlaceholder + 1
        ARC_AddWarning "Skip group: ada lebih dari 1 placeholder dalam 1 group | " & _
                       ARC_ShapeInfo(grp)
        Exit Sub
    End If

    'Ambil warna contour lama dari shape non-text dalam group yang sama
    Set oldColor = ARC_GetOldContourColor(grp, txt)

    If oldColor Is Nothing Then
        arcNoColor = arcNoColor + 1
        ARC_AddWarning "Skip: warna contour lama tidak ditemukan | Text=""" & _
                       ARC_ShortText(ARC_GetText(txt)) & """ | " & ARC_ShapeInfo(txt)
        Exit Sub
    End If

    'Hapus contour lama yang sudah menjadi curve / shape biasa.
    'Text placeholder tetap dipertahankan.
    ARC_DeleteOldContourShapes grp, txt

    'Buat contour baru sebagai LIVE CONTOUR EFFECT.
    'JANGAN di-Separate, agar ketika text diganti, contour ikut berubah.
    'CornerType default: cdrContourCornerRound.
    Set eff = txt.CreateContour( _
                cdrContourOutside, _
                offset, _
                1, _
                cdrDirectFountainFillBlend, _
                oldColor, _
                oldColor, _
                oldColor, _
                0, _
                0, _
                cdrContourSquareCap, _
                cdrContourCornerRound, _
                15#)

    'Text tetap di depan
    On Error Resume Next
    txt.OrderToFront
    On Error GoTo FailedProcess

    arcProcessed = arcProcessed + 1
    ARC_CountCategory cat

    Exit Sub

FailedProcess:
    arcFailed = arcFailed + 1
    ARC_AddWarning "FAILED group/text: " & ARC_ShapeInfo(grp) & _
                   " | Error " & Err.Number & ": " & Err.Description
    Err.Clear

End Sub


Private Sub ARC_DeleteOldContourShapes(ByVal grp As Shape, ByVal txt As Shape)

    Dim i As Long
    Dim ch As Shape

    On Error Resume Next

    If grp.Type <> cdrGroupShape Then Exit Sub

    For i = grp.Shapes.Count To 1 Step -1

        Set ch = grp.Shapes(i)

        If ch Is txt Then

            'Jangan hapus placeholder active text.

        ElseIf ch.Type = cdrTextShape Then

            'Amankan text lain jika ada.
            'Idealnya group placeholder hanya berisi 1 active text.

        ElseIf ch.Type = cdrGroupShape Then

            'Jika contour lama berada dalam nested group, bersihkan juga.
            ARC_DeleteOldContourShapes ch, txt

            If ch.Shapes.Count = 0 Then
                ch.Locked = False
                ch.Delete
            End If

        Else

            'Semua non-text dalam group placeholder dianggap contour lama.
            ch.Locked = False
            ch.Delete

        End If

    Next i

End Sub


'=========================================================
' CLASSIFICATION
'=========================================================

Private Function ARC_ClassifyText( _
    ByVal s As Shape, _
    ByRef cat As Long, _
    ByRef offset As Double) As Boolean

    Dim rawText As String
    Dim t As String
    Dim h As Double

    On Error GoTo SafeExit

    rawText = ARC_GetText(s)
    t = ARC_NormalizeText(rawText)
    h = s.SizeHeight

    cat = ARC_CAT_NONE
    offset = 0#
    ARC_ClassifyText = False

    If ARC_IsIDPO(t) Then Exit Function

    'Nama / nickname
    If ARC_IsNamePlaceholder(t) Then

        If h >= ARC_BACK_NAME_MIN And h <= ARC_BACK_NAME_MAX Then
            cat = ARC_CAT_BACK_NAME
            offset = ARC_BACK_NAME_CONTOUR
            ARC_ClassifyText = True
            Exit Function
        End If

        If h >= ARC_CHEST_NAME_MIN And h < ARC_CHEST_NAME_MAX Then
            cat = ARC_CAT_CHEST_NAME
            offset = ARC_CHEST_NAME_CONTOUR
            ARC_ClassifyText = True
            Exit Function
        End If

    End If

    'Nomor
    If ARC_IsNumberPlaceholder(t) Then

        If h >= ARC_BACK_NUMBER_MIN And h <= ARC_BACK_NUMBER_MAX Then
            cat = ARC_CAT_BACK_NUMBER
            offset = ARC_BACK_NUMBER_CONTOUR
            ARC_ClassifyText = True
            Exit Function
        End If

        If h >= ARC_CHEST_NUMBER_MIN And h <= ARC_CHEST_NUMBER_MAX Then
            cat = ARC_CAT_CHEST_NUMBER
            offset = ARC_CHEST_NUMBER_CONTOUR
            ARC_ClassifyText = True
            Exit Function
        End If

        If h >= ARC_PANTS_NUMBER_MIN And h <= ARC_PANTS_NUMBER_MAX Then
            cat = ARC_CAT_PANTS_NUMBER
            offset = ARC_PANTS_NUMBER_CONTOUR
            ARC_ClassifyText = True
            Exit Function
        End If

    End If

SafeExit:

End Function


Private Function ARC_IsPlaceholderCandidate(ByVal txt As String) As Boolean

    Dim t As String

    t = ARC_NormalizeText(txt)

    If ARC_IsIDPO(t) Then
        ARC_IsPlaceholderCandidate = True
    ElseIf ARC_IsNamePlaceholder(t) Then
        ARC_IsPlaceholderCandidate = True
    ElseIf ARC_IsNumberPlaceholder(t) Then
        ARC_IsPlaceholderCandidate = True
    Else
        ARC_IsPlaceholderCandidate = False
    End If

End Function


Private Function ARC_IsIDPO(ByVal txt As String) As Boolean

    Dim t As String

    t = ARC_NormalizeText(txt)

    ARC_IsIDPO = (t = "IDPO" Or InStr(1, t, "IDPO", vbTextCompare) > 0)

End Function


Private Function ARC_IsNamePlaceholder(ByVal txt As String) As Boolean

    Dim t As String

    t = ARC_NormalizeText(txt)

    ARC_IsNamePlaceholder = False

    If InStr(1, t, "NAMA ATLIT", vbTextCompare) > 0 Then ARC_IsNamePlaceholder = True
    If InStr(1, t, "NAMA ATLET", vbTextCompare) > 0 Then ARC_IsNamePlaceholder = True
    If InStr(1, t, "NAMA", vbTextCompare) > 0 Then ARC_IsNamePlaceholder = True
    If InStr(1, t, "PLAYER", vbTextCompare) > 0 Then ARC_IsNamePlaceholder = True
    If InStr(1, t, "PLAYERS", vbTextCompare) > 0 Then ARC_IsNamePlaceholder = True
    If InStr(1, t, "NICKNAME", vbTextCompare) > 0 Then ARC_IsNamePlaceholder = True
    If InStr(1, t, "NICK", vbTextCompare) > 0 Then ARC_IsNamePlaceholder = True

End Function


Private Function ARC_IsNumberPlaceholder(ByVal txt As String) As Boolean

    Dim t As String
    Dim i As Long
    Dim c As String

    t = ARC_NormalizeText(txt)

    'Support placeholder berupa kata
    If t = "NO" Or t = "NOMOR" Or t = "NUMBER" Then
        ARC_IsNumberPlaceholder = True
        Exit Function
    End If

    'Support placeholder nomor berupa 1-2 digit
    If Len(t) < 1 Or Len(t) > 2 Then
        ARC_IsNumberPlaceholder = False
        Exit Function
    End If

    For i = 1 To Len(t)
        c = Mid$(t, i, 1)
        If c < "0" Or c > "9" Then
            ARC_IsNumberPlaceholder = False
            Exit Function
        End If
    Next i

    ARC_IsNumberPlaceholder = True

End Function


'=========================================================
' OLD CONTOUR COLOR DETECTION
'=========================================================

Private Function ARC_GetOldContourColor( _
    ByVal grp As Shape, _
    ByVal txt As Shape) As Color

    Dim ch As Shape
    Dim bestArea As Double
    Dim bestColor As Color

    On Error Resume Next

    bestArea = -1#
    Set bestColor = Nothing

    For Each ch In grp.Shapes
        If Not ch Is txt Then
            ARC_ProbeBestColor ch, bestArea, bestColor
        End If
    Next ch

    Set ARC_GetOldContourColor = bestColor

End Function


Private Sub ARC_ProbeBestColor( _
    ByVal s As Shape, _
    ByRef bestArea As Double, _
    ByRef bestColor As Color)

    Dim ch As Shape
    Dim area As Double
    Dim c As Color

    On Error Resume Next

    If s.Type = cdrGroupShape Then
        For Each ch In s.Shapes
            ARC_ProbeBestColor ch, bestArea, bestColor
        Next ch
        Exit Sub
    End If

    If s.Type = cdrTextShape Then Exit Sub

    area = s.SizeWidth * s.SizeHeight

    Set c = ARC_TryGetFillColor(s)

    If c Is Nothing Then
        Set c = ARC_TryGetOutlineColor(s)
    End If

    If Not c Is Nothing Then
        If area > bestArea Then
            bestArea = area
            Set bestColor = c
        End If
    End If

End Sub


Private Function ARC_TryGetFillColor(ByVal s As Shape) As Color

    Dim c As New Color

    On Error GoTo NoColor

    If s.Fill.Type = cdrUniformFill Then
        c.CopyAssign s.Fill.UniformColor
        Set ARC_TryGetFillColor = c
        Exit Function
    End If

NoColor:
    Set ARC_TryGetFillColor = Nothing

End Function


Private Function ARC_TryGetOutlineColor(ByVal s As Shape) As Color

    Dim c As New Color

    On Error GoTo NoColor

    If s.Outline.Width > 0 Then
        c.CopyAssign s.Outline.Color
        Set ARC_TryGetOutlineColor = c
        Exit Function
    End If

NoColor:
    Set ARC_TryGetOutlineColor = Nothing

End Function


'=========================================================
' TEXT HELPERS
'=========================================================

Private Function ARC_GetText(ByVal s As Shape) As String

    Dim t As String

    On Error Resume Next

    t = s.Text.Story.Text

    If Err.Number <> 0 Then
        Err.Clear
        t = CStr(s.Text.Story)
    End If

    ARC_GetText = t

End Function


Private Function ARC_NormalizeText(ByVal txt As String) As String

    Dim t As String

    t = CStr(txt)

    t = Replace(t, vbCr, " ")
    t = Replace(t, vbLf, " ")
    t = Replace(t, vbTab, " ")
    t = Trim$(t)

    Do While InStr(t, "  ") > 0
        t = Replace(t, "  ", " ")
    Loop

    ARC_NormalizeText = UCase$(t)

End Function


Private Function ARC_ShortText(ByVal txt As String) As String

    Dim t As String

    t = ARC_NormalizeText(txt)

    If Len(t) > 30 Then
        ARC_ShortText = Left$(t, 30) & "..."
    Else
        ARC_ShortText = t
    End If

End Function


'=========================================================
' REPORT HELPERS
'=========================================================

Private Sub ARC_CountCategory(ByVal cat As Long)

    Select Case cat
        Case ARC_CAT_BACK_NAME
            arcBackName = arcBackName + 1

        Case ARC_CAT_BACK_NUMBER
            arcBackNumber = arcBackNumber + 1

        Case ARC_CAT_CHEST_NAME
            arcChestName = arcChestName + 1

        Case ARC_CAT_CHEST_NUMBER
            arcChestNumber = arcChestNumber + 1

        Case ARC_CAT_PANTS_NUMBER
            arcPantsNumber = arcPantsNumber + 1
    End Select

End Sub


Private Function ARC_ShapeInfo(ByVal s As Shape) As String

    On Error Resume Next

    ARC_ShapeInfo = _
        "X=" & FormatNumber(s.CenterX, 2) & _
        " Y=" & FormatNumber(s.CenterY, 2) & _
        " W=" & FormatNumber(s.SizeWidth, 2) & _
        " H=" & FormatNumber(s.SizeHeight, 2)

End Function


Private Sub ARC_AddWarning(ByVal msg As String)

    Dim currentCount As Long

    currentCount = UBound(Split(arcWarnings & vbCrLf, vbCrLf))

    If currentCount <= ARC_WARN_LIMIT Then
        arcWarnings = arcWarnings & "- " & msg & vbCrLf
    End If

End Sub


Private Sub ARC_ShowReport()

    Dim msg As String

    msg = "AUTO RE-CONTOUR FINISHED" & vbCrLf & vbCrLf

    msg = msg & "Group kandidat ditemukan : " & arcJobs & vbCrLf
    msg = msg & "Berhasil re-contour     : " & arcProcessed & vbCrLf & vbCrLf

    msg = msg & "Rincian:" & vbCrLf
    msg = msg & "Nama/Nickname Punggung : " & arcBackName & vbCrLf
    msg = msg & "Nomor Punggung         : " & arcBackNumber & vbCrLf
    msg = msg & "Nama/Nickname Dada     : " & arcChestName & vbCrLf
    msg = msg & "Nomor Dada Tengah      : " & arcChestNumber & vbCrLf
    msg = msg & "Nomor Celana           : " & arcPantsNumber & vbCrLf
    msg = msg & "IDPO Diabaikan         : " & arcIDPOSkipped & vbCrLf & vbCrLf

    If arcUnclassified > 0 Or arcNoColor > 0 Or arcMultiPlaceholder > 0 Or arcFailed > 0 Then

        msg = msg & "WARNING:" & vbCrLf
        msg = msg & "Text tidak terklasifikasi : " & arcUnclassified & vbCrLf
        msg = msg & "Warna contour tidak ketemu: " & arcNoColor & vbCrLf
        msg = msg & "Group multi-placeholder   : " & arcMultiPlaceholder & vbCrLf
        msg = msg & "Failed process            : " & arcFailed & vbCrLf & vbCrLf

        If Len(arcWarnings) > 0 Then
            msg = msg & "Detail:" & vbCrLf & arcWarnings
        End If

        MsgBox msg, vbExclamation, "AUTO RE-CONTOUR"

    Else

        msg = msg & "Status: OK" & vbCrLf
        msg = msg & "Contour baru dibuat sebagai LIVE EFFECT." & vbCrLf
        msg = msg & "Corner contour default: ROUNDED." & vbCrLf
        msg = msg & "Jika placeholder diedit, contour seharusnya ikut berubah."

        MsgBox msg, vbInformation, "AUTO RE-CONTOUR"

    End If

End Sub