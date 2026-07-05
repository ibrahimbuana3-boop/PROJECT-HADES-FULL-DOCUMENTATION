Option Explicit

'=========================================
' AI_TEXT_V6.2
' HORIZONTAL STRETCH ONLY
' + OPTIONAL ARIAL FALLBACK
' + FINAL CONVERT ACTIVE TEXT ONLY
'
' MAIN MACRO:
' AI_TEXT_V1
'
' ALIAS:
' AI_TEXT
'
' Mode:
' 1 = Resize saja
' 2 = Resize + Arial hanya karakter khusus
' 3 = Resize + Arial seluruh teks jika mengandung karakter khusus
' 4 = FINAL CONVERT ACTIVE TEXT ONLY
'
' CATATAN MODE 4:
' - Hanya convert cdrTextShape.
' - Tidak convert bitmap.
' - Tidak convert curve.
' - Tidak convert group sebagai group.
' - Tidak resize.
' - Tidak font fallback.
' - Setelah convert, QC_TYPO_CHECK dan IDPO_CHECK
'   tidak bisa membaca text lagi.
'=========================================


'=========================================
' CONFIG
'=========================================

Private Const FRONT_MIN As Double = 1#
Private Const FRONT_MAX As Double = 3.5

Private Const BACK_MIN As Double = 3.6
Private Const BACK_MAX As Double = 7.8

Private Const LIMIT_FRONT As Double = 10#
Private Const LIMIT_BACK As Double = 30#

Private Const ID_MIN_H As Double = 0.3
Private Const ID_MAX_H As Double = 0.6


'=========================================
' GLOBAL COUNTERS
'=========================================

Private ArialMode As Long

Private FixedCount As Long
Private ResizedFront As Long
Private ResizedBack As Long

Private TextScanned As Long
Private TextConverted As Long
Private TextConvertFailed As Long


'=========================================
' ALIAS
'=========================================

Sub AI_TEXT()
    AI_TEXT_V1
End Sub


'=========================================
' MAIN
'=========================================

Sub AI_TEXT_V1()

    Dim oldUnit As cdrUnit
    Dim oldOptimization As Boolean
    Dim cmdStarted As Boolean

    On Error GoTo ERR_HANDLER

    If ActiveSelection Is Nothing Then
        MsgBox "Tidak ada objek dipilih.", vbExclamation, "AI TEXT"
        Exit Sub
    End If

    If ActiveSelection.Shapes.Count = 0 Then
        MsgBox "Tidak ada objek dipilih.", vbExclamation, "AI TEXT"
        Exit Sub
    End If

    ArialMode = AskArialMode()

    If ArialMode = 0 Then Exit Sub

    ResetCounters

    oldUnit = ActiveDocument.Unit
    oldOptimization = Application.Optimization
    cmdStarted = False

    ActiveDocument.Unit = cdrCentimeter

    '=====================================
    ' MODE 4: FINAL CONVERT ACTIVE TEXT
    '=====================================
    If ArialMode = 4 Then

        FinalConvertActiveTextOnly oldUnit, oldOptimization
        Exit Sub

    End If

    '=====================================
    ' MODE 1 - 3: AI TEXT NORMAL
    '=====================================
    Application.Optimization = True

    ActiveDocument.BeginCommandGroup "AI TEXT V6.2"
    cmdStarted = True

    Dim s As Shape

    For Each s In ActiveSelection.Shapes
        ScanShape s
    Next s

    ActiveDocument.EndCommandGroup
    cmdStarted = False

    Application.Optimization = oldOptimization
    ActiveDocument.Unit = oldUnit
    ActiveWindow.Refresh

    MsgBox _
        "AI TEXT selesai" & vbCrLf & vbCrLf & _
        "Mode           : " & ArialMode & vbCrLf & _
        "Text scanned   : " & TextScanned & vbCrLf & _
        "Fallback Arial : " & FixedCount & vbCrLf & _
        "Resize depan   : " & ResizedFront & vbCrLf & _
        "Resize belakang: " & ResizedBack, _
        vbInformation, _
        "AI TEXT"

    Exit Sub

ERR_HANDLER:

    On Error Resume Next

    If cmdStarted Then ActiveDocument.EndCommandGroup

    Application.Optimization = oldOptimization
    ActiveDocument.Unit = oldUnit
    ActiveWindow.Refresh

    MsgBox _
        "SYSTEM ERROR - AI TEXT" & vbCrLf & vbCrLf & _
        "No : " & err.Number & vbCrLf & _
        err.Description, _
        vbCritical, _
        "AI TEXT"

End Sub


'=========================================
' RESET COUNTERS
'=========================================

Private Sub ResetCounters()

    FixedCount = 0
    ResizedFront = 0
    ResizedBack = 0

    TextScanned = 0
    TextConverted = 0
    TextConvertFailed = 0

End Sub


'=========================================
' POPUP MODE
'=========================================

Private Function AskArialMode() As Long

    Dim ans As String

    ans = InputBox( _
        "PILIH MODE AI TEXT" & vbCrLf & vbCrLf & _
        "1 = Resize saja" & vbCrLf & _
        "2 = Resize + Arial hanya karakter khusus" & vbCrLf & _
        "3 = Resize + Arial seluruh teks jika ada karakter khusus" & vbCrLf & _
        "4 = FINAL CONVERT ACTIVE TEXT ONLY" & vbCrLf & vbCrLf & _
        "Mode 4 dipakai paling akhir setelah:" & vbCrLf & _
        "- QC_SIZE_CHECK" & vbCrLf & _
        "- QC_TYPO_CHECK" & vbCrLf & _
        "- IDPO_CHECK" & vbCrLf & vbCrLf & _
        "Rekomendasi harian:" & vbCrLf & _
        "Pakai 2 untuk resize + karakter khusus." & vbCrLf & _
        "Pakai 4 hanya saat final sebelum save produksi.", _
        "AI TEXT", _
        "2")

    ans = Trim$(ans)

    If ans = "" Then
        AskArialMode = 0
        Exit Function
    End If

    Select Case ans

        Case "1"
            AskArialMode = 1

        Case "2"
            AskArialMode = 2

        Case "3"
            AskArialMode = 3

        Case "4"
            AskArialMode = 4

        Case Else
            MsgBox "Pilihan tidak valid.", vbExclamation, "AI TEXT"
            AskArialMode = 0

    End Select

End Function


'=========================================
' SCAN MODE 1 - 3
'=========================================

Private Sub ScanShape(ByVal s As Shape)

    Dim c As Shape

    On Error Resume Next

    If s.Type = cdrGroupShape Then

        For Each c In s.Shapes
            ScanShape c
        Next c

        Exit Sub

    End If

    If s.Type <> cdrTextShape Then Exit Sub

    TextScanned = TextScanned + 1

    ProcessText s

End Sub


'=========================================
' PROCESS MODE 1 - 3
'=========================================

Private Sub ProcessText(ByVal t As Shape)

    On Error GoTo SAFE_EXIT

    Dim txt As String

    txt = Normalize(t.Text.Story.Text)

    If txt = "" Then Exit Sub

    If IsSmallIDPO(t, txt) Then Exit Sub

    If IsAthleteNumber(txt) Then Exit Sub

    FixLigature t

    If ArialMode = 2 Then

        If HasSpecialChar(txt) Then
            If ApplyArialToSpecialChars(t) Then
                FixedCount = FixedCount + 1
            End If
        End If

    ElseIf ArialMode = 3 Then

        If HasSpecialChar(txt) Then
            t.Text.Font = "Arial"
            FixedCount = FixedCount + 1
        End If

    End If

    ResizeTextIfNeeded t

SAFE_EXIT:

    On Error GoTo 0

End Sub


'=========================================
' RESIZE
'=========================================

Private Sub ResizeTextIfNeeded(ByVal t As Shape)

    On Error Resume Next

    Dim h As Double
    h = Round(t.SizeHeight, 2)

    If h >= FRONT_MIN And h <= FRONT_MAX Then

        If t.SizeWidth > LIMIT_FRONT Then
            StretchWidth t, LIMIT_FRONT
            ResizedFront = ResizedFront + 1
        End If

        Exit Sub

    End If

    If h >= BACK_MIN And h <= BACK_MAX Then

        If t.SizeWidth > LIMIT_BACK Then
            StretchWidth t, LIMIT_BACK
            ResizedBack = ResizedBack + 1
        End If

    End If

    On Error GoTo 0

End Sub


'=========================================
' STRETCH WIDTH ONLY
'=========================================

Private Sub StretchWidth(ByVal t As Shape, ByVal targetWidth As Double)

    On Error Resume Next

    Dim oldW As Double
    oldW = t.SizeWidth

    If oldW <= targetWidth Then Exit Sub
    If oldW <= 0 Then Exit Sub

    Dim factor As Double
    factor = targetWidth / oldW

    Dim lx As Double
    Dim rx As Double
    Dim cx As Double

    lx = t.LeftX
    rx = t.RightX
    cx = t.CenterX

    Dim al As Long
    al = t.Text.AlignProperties.Alignment

    t.Stretch factor, 1

    Select Case al

        Case cdrLeftAlignment
            t.LeftX = lx

        Case cdrRightAlignment
            t.RightX = rx

        Case Else
            t.CenterX = cx

    End Select

    On Error GoTo 0

End Sub


'=========================================
' ARIAL FALLBACK SPECIAL ONLY
'=========================================

Private Function ApplyArialToSpecialChars(ByVal t As Shape) As Boolean

    On Error GoTo FAIL

    Dim raw As String
    raw = t.Text.Story.Text

    Dim i As Long
    Dim ch As String
    Dim changed As Boolean

    For i = 1 To Len(raw)

        ch = Mid$(raw, i, 1)

        If IsSpecialChar(ch) Then

            On Error Resume Next

            err.Clear

            'CorelDRAW biasanya mendukung Characters(start, length)
            t.Text.Story.Characters(i, 1).Font = "Arial"

            If err.Number = 0 Then
                changed = True
            End If

            err.Clear
            On Error GoTo FAIL

        End If

    Next i

    ApplyArialToSpecialChars = changed
    Exit Function

FAIL:

    ApplyArialToSpecialChars = False

End Function


'=========================================
' FIX LIGATURE
'=========================================

Private Sub FixLigature(ByVal t As Shape)

    On Error Resume Next

    t.Text.Story.OpenTypeFeatures.Ligatures = False

    On Error GoTo 0

End Sub


'=========================================
' MODE 4 — FINAL CONVERT ACTIVE TEXT ONLY
'=========================================

Private Sub FinalConvertActiveTextOnly( _
    ByVal oldUnit As cdrUnit, _
    ByVal oldOptimization As Boolean)

    Dim textShapes As Collection
    Dim s As Shape
    Dim ans As String
    Dim cmdStarted As Boolean

    On Error GoTo ERR_HANDLER

    Set textShapes = New Collection

    For Each s In ActiveSelection.Shapes
        CollectTextShapesForConvert s, textShapes
    Next s

    If textShapes.Count = 0 Then

        ActiveDocument.Unit = oldUnit
        Application.Optimization = oldOptimization
        ActiveWindow.Refresh

        MsgBox _
            "Tidak ada active text di dalam selection." & vbCrLf & vbCrLf & _
            "Tidak ada objek yang diconvert.", _
            vbInformation, _
            "AI TEXT FINAL CONVERT"

        Exit Sub

    End If

    ans = InputBox( _
        "FINAL CONVERT ACTIVE TEXT ONLY" & vbCrLf & vbCrLf & _
        "Jumlah active text terdeteksi: " & textShapes.Count & vbCrLf & vbCrLf & _
        "Macro ini akan mengubah semua ACTIVE TEXT" & vbCrLf & _
        "di dalam selection menjadi curve." & vbCrLf & vbCrLf & _
        "Yang TIDAK disentuh:" & vbCrLf & _
        "- bitmap" & vbCrLf & _
        "- motif" & vbCrLf & _
        "- curve" & vbCrLf & _
        "- pattern outline" & vbCrLf & _
        "- group non-text" & vbCrLf & vbCrLf & _
        "PERINGATAN:" & vbCrLf & _
        "Setelah ini QC_TYPO_CHECK dan IDPO_CHECK" & vbCrLf & _
        "tidak bisa membaca text lagi." & vbCrLf & vbCrLf & _
        "Ketik FINAL untuk lanjut.", _
        "AI TEXT FINAL CONVERT")

    ans = UCase$(Trim$(ans))

    If ans <> "FINAL" Then

        ActiveDocument.Unit = oldUnit
        Application.Optimization = oldOptimization
        ActiveWindow.Refresh

        MsgBox _
            "Final convert dibatalkan." & vbCrLf & _
            "Tidak ada objek yang diubah.", _
            vbInformation, _
            "AI TEXT FINAL CONVERT"

        Exit Sub

    End If

    Application.Optimization = True

    ActiveDocument.BeginCommandGroup "AI TEXT FINAL CONVERT ACTIVE TEXT ONLY"
    cmdStarted = True

    ConvertCollectedTexts textShapes

    ActiveDocument.EndCommandGroup
    cmdStarted = False

    Application.Optimization = oldOptimization
    ActiveDocument.Unit = oldUnit
    ActiveWindow.Refresh

    MsgBox _
        "FINAL CONVERT ACTIVE TEXT ONLY selesai." & vbCrLf & vbCrLf & _
        "Text terdeteksi : " & textShapes.Count & vbCrLf & _
        "Text converted  : " & TextConverted & vbCrLf & _
        "Gagal convert   : " & TextConvertFailed & vbCrLf & vbCrLf & _
        "Hanya active text yang diproses.", _
        vbInformation, _
        "AI TEXT FINAL CONVERT"

    Exit Sub

ERR_HANDLER:

    On Error Resume Next

    If cmdStarted Then ActiveDocument.EndCommandGroup

    Application.Optimization = oldOptimization
    ActiveDocument.Unit = oldUnit
    ActiveWindow.Refresh

    MsgBox _
        "SYSTEM ERROR - FINAL CONVERT ACTIVE TEXT ONLY" & vbCrLf & vbCrLf & _
        "No : " & err.Number & vbCrLf & _
        err.Description & vbCrLf & vbCrLf & _
        "Jika sebagian text sudah berubah, tekan Ctrl+Z satu kali.", _
        vbCritical, _
        "AI TEXT FINAL CONVERT"

End Sub


Private Sub CollectTextShapesForConvert( _
    ByVal s As Shape, _
    ByRef textShapes As Collection)

    Dim c As Shape

    On Error Resume Next

    If Not IsShapeAlive(s) Then Exit Sub

    If s.Type = cdrGroupShape Then

        For Each c In s.Shapes
            CollectTextShapesForConvert c, textShapes
        Next c

        Exit Sub

    End If

    If s.Type = cdrTextShape Then
        textShapes.Add s
    End If

End Sub


Private Sub ConvertCollectedTexts(ByVal textShapes As Collection)

    Dim v As Variant
    Dim t As Shape

    TextConverted = 0
    TextConvertFailed = 0

    For Each v In textShapes

        Set t = v

        If IsShapeAlive(t) Then

            If ConvertOneTextToCurve(t) Then
                TextConverted = TextConverted + 1
            Else
                TextConvertFailed = TextConvertFailed + 1
            End If

        Else

            TextConvertFailed = TextConvertFailed + 1

        End If

    Next v

End Sub


Private Function ConvertOneTextToCurve(ByVal t As Shape) As Boolean

    On Error GoTo FAIL

    ConvertOneTextToCurve = False

    If t Is Nothing Then Exit Function
    If t.Type <> cdrTextShape Then Exit Function

    On Error Resume Next
    t.Locked = False
    err.Clear
    On Error GoTo FAIL

    t.ConvertToCurves

    ConvertOneTextToCurve = True
    Exit Function

FAIL:

    err.Clear
    ConvertOneTextToCurve = False

End Function


Private Function IsShapeAlive(ByVal s As Shape) As Boolean

    Dim t As Long

    On Error GoTo DEAD

    If s Is Nothing Then GoTo DEAD

    t = s.Type

    IsShapeAlive = True
    Exit Function

DEAD:

    IsShapeAlive = False

End Function


'=========================================
' SPECIAL CHARACTER RULES
'=========================================

Private Function HasSpecialChar(ByVal s As String) As Boolean

    Dim i As Long
    Dim ch As String

    For i = 1 To Len(s)

        ch = Mid$(s, i, 1)

        If IsSpecialChar(ch) Then
            HasSpecialChar = True
            Exit Function
        End If

    Next i

End Function


Private Function IsSpecialChar(ByVal ch As String) As Boolean

    Select Case ch

        Case "_", "-", """", "'", "*", "@", "+", "&", ".", ",", "/", "\", "(", ")", "[", "]"
            IsSpecialChar = True

        Case Else
            IsSpecialChar = False

    End Select

End Function


'=========================================
' UTIL
'=========================================

Private Function Normalize(ByVal s As String) As String

    s = Replace(s, vbCr, "")
    s = Replace(s, vbLf, "")
    s = Replace(s, Chr$(160), " ")

    Do While InStr(1, s, "  ", vbTextCompare) > 0
        s = Replace(s, "  ", " ")
    Loop

    Normalize = UCase$(Trim$(s))

End Function


Private Function IsAthleteNumber(ByVal s As String) As Boolean

    s = Trim$(s)

    If Len(s) = 0 Then Exit Function
    If Len(s) > 3 Then Exit Function

    If IsNumeric(s) Then
        IsAthleteNumber = True
    End If

End Function


Private Function IsSmallIDPO(ByVal t As Shape, ByVal s As String) As Boolean

    s = Normalize(s)

    If s = "IDPO" Then

        If t.SizeHeight >= ID_MIN_H And t.SizeHeight <= ID_MAX_H Then
            IsSmallIDPO = True
        End If

        Exit Function

    End If

    If Len(s) = 6 And IsNumeric(s) Then

        If t.SizeHeight >= ID_MIN_H And t.SizeHeight <= ID_MAX_H Then
            IsSmallIDPO = True
        End If

    End If

End Function

