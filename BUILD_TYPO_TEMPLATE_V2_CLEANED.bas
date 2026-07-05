Option Explicit

'=========================================================
' PROJECT HADES — BUILD TYPO TEMPLATE V3.1 SMART SENSOR
' CorelDRAW 2021 VBA
'
' MAIN MACRO TETAP:
'   BUILD_TYPO_TEMPLATE
'
' TUJUAN:
' - Helper untuk QC_TYPO_CHECK.
' - Merekam titik nol dari 1 master layout sample.
' - Menghindari aturan statis tinggi teks.
' - Output tetap Documents\TypoTemplate_Current.txt.
'
' FITUR V3.1:
' - Fix bug variable shadowing dictionary vs Shape.
' - Recursive group + PowerClip scan.
' - Deteksi panel merah/hijau terbesar sebagai MASTER_PANEL.
' - Simpan MASTER_PANEL_WIDTH, MASTER_PANEL_HEIGHT, MASTER_PANEL_AREA.
' - Simpan ROLE_COUNT + slot metrics per placeholder.
' - Simpan height, width, align, placeholder, relative X/Y, label ringan.
' - Skip IDPO kecil, empty text, dan marker attribute @A: / @ATTR:.
' - Membaca @SIZEDB dari Documents\Order.txt jika ada.
' - Backward compatible: key lama tetap dibuat.
'=========================================================

Private BT_Tpl As Object

Private BT_MasterPanelHeight As Double
Private BT_MasterPanelWidth As Double
Private BT_MasterPanelArea As Double
Private BT_MasterPanelCenterX As Double
Private BT_MasterPanelCenterY As Double
Private BT_BestPanelArea As Double
Private BT_FallbackPanelHeight As Double
Private BT_FallbackPanelWidth As Double

Private Const BT_ID_MIN_H As Double = 0.28
Private Const BT_ID_MAX_H As Double = 0.65
Private Const BT_MIN_TEXT_H As Double = 0.1

'=========================================================
' PUBLIC ENTRY — JANGAN GANTI NAMA SHORTCUT
'=========================================================

Public Sub BUILD_TYPO_TEMPLATE()

    Dim oldUnit As cdrUnit
    Dim cmdOpened As Boolean
    Dim s As Shape

    If ActiveSelection Is Nothing Then
        MsgBox "Pilih 1 MASTER LAYOUT SAMPLE terlebih dahulu.", vbExclamation, "BUILD TYPO TEMPLATE"
        Exit Sub
    End If

    If ActiveSelection.Shapes.Count = 0 Then
        MsgBox "Pilih 1 MASTER LAYOUT SAMPLE terlebih dahulu.", vbExclamation, "BUILD TYPO TEMPLATE"
        Exit Sub
    End If

    oldUnit = ActiveDocument.Unit
    cmdOpened = False

    On Error GoTo ERR_HANDLER

    Set BT_Tpl = CreateObject("Scripting.Dictionary")

    BT_MasterPanelHeight = 0
    BT_MasterPanelWidth = 0
    BT_MasterPanelArea = 0
    BT_MasterPanelCenterX = 0
    BT_MasterPanelCenterY = 0
    BT_BestPanelArea = 0
    BT_FallbackPanelHeight = 0
    BT_FallbackPanelWidth = 0

    ActiveDocument.Unit = cdrCentimeter
    ActiveDocument.BeginCommandGroup "BUILD TYPO TEMPLATE V3.1"
    cmdOpened = True

    For Each s In ActiveSelection.Shapes
        BT_ScanShape s
    Next s

    If BT_MasterPanelHeight <= 0 Then
        BT_MasterPanelHeight = BT_FallbackPanelHeight
        BT_MasterPanelWidth = BT_FallbackPanelWidth
        BT_MasterPanelArea = BT_MasterPanelHeight * BT_MasterPanelWidth
    End If

    If BT_MasterPanelHeight <= 0 Then
        MsgBox _
            "BUILD TEMPLATE GAGAL" & vbCrLf & vbCrLf & _
            "MASTER_PANEL tidak bisa dideteksi." & vbCrLf & _
            "Pastikan sample layout berisi pola merah/hijau atau object panel terbesar.", _
            vbCritical, _
            "BUILD TYPO TEMPLATE"
        GoTo EXIT_CLEAN
    End If

    BT_SaveTemplate

    ActiveDocument.EndCommandGroup
    cmdOpened = False
    ActiveDocument.Unit = oldUnit

    MsgBox _
        "BUILD TEMPLATE BERHASIL" & vbCrLf & vbCrLf & _
        "File:" & vbCrLf & _
        Environ$("USERPROFILE") & "\Documents\TypoTemplate_Current.txt" & vbCrLf & vbCrLf & _
        "MASTER_PANEL : " & Format(BT_MasterPanelHeight, "0.000") & " cm" & vbCrLf & _
        "NAMA_ATLIT   : " & BT_CountOf("NAMA_ATLIT") & vbCrLf & _
        "NAMA         : " & BT_CountOf("NAMA") & vbCrLf & _
        "NICKNAME     : " & BT_CountOf("NICKNAME") & vbCrLf & _
        "NUMBER       : " & BT_CountOf("NUMBER"), _
        vbInformation, _
        "BUILD TYPO TEMPLATE"

    Exit Sub

EXIT_CLEAN:
    On Error Resume Next
    If cmdOpened Then ActiveDocument.EndCommandGroup
    ActiveDocument.Unit = oldUnit
    On Error GoTo 0
    Exit Sub

ERR_HANDLER:
    On Error Resume Next
    If cmdOpened Then ActiveDocument.EndCommandGroup
    ActiveDocument.Unit = oldUnit
    On Error GoTo 0

    MsgBox _
        "SYSTEM ERROR - BUILD TYPO TEMPLATE V3.1" & vbCrLf & vbCrLf & _
        "No : " & Err.Number & vbCrLf & _
        Err.Description, _
        vbCritical, _
        "BUILD TYPO TEMPLATE"

End Sub

Public Sub BUILD_TYPO_TEMPLATE_V3()
    BUILD_TYPO_TEMPLATE
End Sub

Public Sub BUILD_TYPO_TEMPLATE_V31()
    BUILD_TYPO_TEMPLATE
End Sub

'=========================================================
' SCAN ENGINE
'=========================================================

Private Sub BT_ScanShape(ByVal shp As Shape)

    Dim c As Shape
    Dim pcShapes As Shapes

    On Error Resume Next

    If shp.SizeHeight > BT_FallbackPanelHeight Then
        BT_FallbackPanelHeight = Round(shp.SizeHeight, 3)
        BT_FallbackPanelWidth = Round(shp.SizeWidth, 3)
    End If

    If shp.Type = cdrGroupShape Then
        For Each c In shp.Shapes
            BT_ScanShape c
        Next c
        Exit Sub
    End If

    Set pcShapes = shp.PowerClip.Shapes
    If Not pcShapes Is Nothing Then
        For Each c In pcShapes
            BT_ScanShape c
        Next c
    End If

    On Error GoTo 0

    If shp.Type = cdrCurveShape Then
        BT_MinePanel shp
        Exit Sub
    End If

    If shp.Type = cdrTextShape Then
        BT_MineText shp
        Exit Sub
    End If

End Sub

'=========================================================
' PANEL SENSOR
'=========================================================

Private Sub BT_MinePanel(ByVal shp As Shape)

    If Not BT_IsPanelOutline(shp) Then Exit Sub

    Dim w As Double
    Dim h As Double
    Dim area As Double
    Dim mx As Double
    Dim mn As Double

    w = Round(shp.SizeWidth, 3)
    h = Round(shp.SizeHeight, 3)

    If w <= 0 Or h <= 0 Then Exit Sub

    area = w * h

    If w > h Then
        mx = w
        mn = h
    Else
        mx = h
        mn = w
    End If

    If area > BT_BestPanelArea Then
        BT_BestPanelArea = area
        BT_MasterPanelHeight = mx
        BT_MasterPanelWidth = mn
        BT_MasterPanelArea = area

        On Error Resume Next
        BT_MasterPanelCenterX = shp.PositionX
        BT_MasterPanelCenterY = shp.PositionY
        On Error GoTo 0
    End If

End Sub

Private Function BT_IsPanelOutline(ByVal shp As Shape) As Boolean

    If BT_IsRedOutline(shp) Then
        BT_IsPanelOutline = True
        Exit Function
    End If

    If BT_IsGreenOutline(shp) Then
        BT_IsPanelOutline = True
        Exit Function
    End If

End Function

Private Function BT_IsRedOutline(ByVal shp As Shape) As Boolean

    On Error Resume Next

    If shp.Outline Is Nothing Then Exit Function
    If shp.Outline.Type = cdrNoOutline Then Exit Function

    BT_IsRedOutline = _
        shp.Outline.Color.RGBRed > 200 And _
        shp.Outline.Color.RGBGreen < 80 And _
        shp.Outline.Color.RGBBlue < 80

    On Error GoTo 0

End Function

Private Function BT_IsGreenOutline(ByVal shp As Shape) As Boolean

    On Error Resume Next

    If shp.Outline Is Nothing Then Exit Function
    If shp.Outline.Type = cdrNoOutline Then Exit Function

    BT_IsGreenOutline = _
        shp.Outline.Color.RGBRed <= 100 And _
        shp.Outline.Color.RGBGreen >= 160 And _
        shp.Outline.Color.RGBBlue <= 100

    On Error GoTo 0

End Function

'=========================================================
' TEXT SENSOR
'=========================================================

Private Sub BT_MineText(ByVal shpText As Shape)

    On Error GoTo FAIL

    Dim raw As String
    Dim txt As String
    Dim role As String

    raw = shpText.Text.Story.Text
    txt = BT_Normalize(raw)

    If txt = "" Then Exit Sub
    If shpText.SizeHeight < BT_MIN_TEXT_H Then Exit Sub
    If BT_IsAttributeMarker(txt) Then Exit Sub
    If BT_IgnoreSmallID(shpText, txt) Then Exit Sub

    role = BT_DetectPlaceholderRole(txt)

    If role = "" Then Exit Sub

    BT_AddRoleMetric role, shpText, txt

    Exit Sub

FAIL:
    Debug.Print "BT_MineText gagal | " & Err.Description

End Sub

Private Function BT_DetectPlaceholderRole(ByVal txt As String) As String

    Select Case txt

        Case "NAMA ATLIT", "NAMA ATLET", "PLAYER", "PLAYERS", "PLAYER NAME", "NAMA PEMAIN"
            BT_DetectPlaceholderRole = "NAMA_ATLIT"

        Case "NAMA"
            BT_DetectPlaceholderRole = "NAMA"

        Case "NICKNAME", "NICK NAME", "NICK", "NAMA PANGGILAN"
            BT_DetectPlaceholderRole = "NICKNAME"

        Case Else
            If BT_IsNumberPlaceholder(txt) Then
                BT_DetectPlaceholderRole = "NUMBER"
            End If

    End Select

End Function

Private Sub BT_AddRoleMetric( _
    ByVal role As String, _
    ByVal shpText As Shape, _
    ByVal originalText As String)

    Dim h As Double
    Dim w As Double
    Dim al As String
    Dim countKey As String
    Dim idx As Long
    Dim relX As Double
    Dim relY As Double
    Dim label As String

    h = Round(shpText.SizeHeight, 3)
    w = Round(shpText.SizeWidth, 3)
    al = BT_GetAlign(shpText)

    countKey = role & "_COUNT"

    If BT_Tpl.Exists(countKey) Then
        idx = CLng(BT_Tpl(countKey)) + 1
        BT_Tpl(countKey) = idx
    Else
        idx = 1
        BT_Tpl.Add countKey, idx
    End If

    relX = BT_RelX(shpText)
    relY = BT_RelY(shpText)
    label = BT_GuessSlotLabel(role, h, relX, relY)

    BT_SetValue role & "_" & idx & "_H", h
    BT_SetValue role & "_" & idx & "_W", w
    BT_SetValue role & "_" & idx & "_ALIGN", al
    BT_SetValue role & "_" & idx & "_PLACEHOLDER", originalText
    BT_SetValue role & "_" & idx & "_REL_X", Format(relX, "0.000")
    BT_SetValue role & "_" & idx & "_REL_Y", Format(relY, "0.000")
    BT_SetValue role & "_" & idx & "_LABEL", label

    ' Backward compatibility untuk QC lama.
    If idx = 1 Then
        BT_AddValue role, h
        BT_AddValue role & "_H", h
        BT_AddValue role & "_W", w
        BT_AddValue role & "_ALIGN", al
    End If

End Sub

'=========================================================
' TEMPLATE SAVE
'=========================================================

Private Sub BT_SaveTemplate()

    Dim path As String
    Dim f As Integer
    Dim k As Variant
    Dim dbName As String

    path = Environ$("USERPROFILE") & "\Documents\TypoTemplate_Current.txt"
    dbName = BT_ReadOrderMetaValue("SIZEDB")

    f = FreeFile

    Open path For Output As #f

    Print #f, "TEMPLATE_VERSION=3.1"
    Print #f, "ENGINE=BUILD_TYPO_TEMPLATE_V3_1"

    If BT_BestPanelArea > 0 Then
        Print #f, "PANEL_SOURCE=RED_OR_GREEN_PANEL"
    Else
        Print #f, "PANEL_SOURCE=FALLBACK_MAX_SHAPE"
    End If

    Print #f, "MASTER_PANEL=" & Format(BT_MasterPanelHeight, "0.000")
    Print #f, "MASTER_PANEL_HEIGHT=" & Format(BT_MasterPanelHeight, "0.000")
    Print #f, "MASTER_PANEL_WIDTH=" & Format(BT_MasterPanelWidth, "0.000")
    Print #f, "MASTER_PANEL_AREA=" & Format(BT_MasterPanelArea, "0.000")

    If Len(Trim$(dbName)) > 0 Then
        Print #f, "SIZEDB=" & dbName
    End If

    Print #f, "PAIRING_MODE=STRICT_NAME_NUMBER"
    Print #f, "TEXT_SCALE_MODE=PANEL_RATIO"

    For Each k In BT_Tpl.Keys
        Print #f, CStr(k) & "=" & BT_Tpl(k)
    Next k

    Close #f

End Sub

'=========================================================
' HELPERS
'=========================================================

Private Function BT_CountOf(ByVal role As String) As Long

    If BT_Tpl Is Nothing Then Exit Function

    role = UCase$(Trim$(role))

    If BT_Tpl.Exists(role & "_COUNT") Then
        BT_CountOf = CLng(BT_Tpl(role & "_COUNT"))
    End If

End Function

Private Sub BT_AddValue(ByVal k As String, ByVal v As Variant)

    If Not BT_Tpl.Exists(k) Then
        BT_Tpl.Add k, v
    End If

End Sub

Private Sub BT_SetValue(ByVal k As String, ByVal v As Variant)

    If BT_Tpl.Exists(k) Then
        BT_Tpl(k) = v
    Else
        BT_Tpl.Add k, v
    End If

End Sub

Private Function BT_GetAlign(ByVal shpText As Shape) As String

    On Error Resume Next

    Dim a As Long
    a = shpText.Text.AlignProperties.Alignment

    Select Case a
        Case cdrLeftAlignment
            BT_GetAlign = "LEFT"
        Case cdrRightAlignment
            BT_GetAlign = "RIGHT"
        Case Else
            BT_GetAlign = "CENTER"
    End Select

    On Error GoTo 0

End Function

Private Function BT_RelX(ByVal shp As Shape) As Double

    On Error Resume Next

    If BT_MasterPanelWidth <= 0 Then
        BT_RelX = 0
    Else
        BT_RelX = (shp.PositionX - BT_MasterPanelCenterX) / BT_MasterPanelWidth
    End If

    On Error GoTo 0

End Function

Private Function BT_RelY(ByVal shp As Shape) As Double

    On Error Resume Next

    If BT_MasterPanelHeight <= 0 Then
        BT_RelY = 0
    Else
        BT_RelY = (BT_MasterPanelCenterY - shp.PositionY) / BT_MasterPanelHeight
    End If

    On Error GoTo 0

End Function

Private Function BT_GuessSlotLabel( _
    ByVal role As String, _
    ByVal h As Double, _
    ByVal relX As Double, _
    ByVal relY As Double) As String

    role = UCase$(Trim$(role))

    If role = "NUMBER" Then
        If h >= 15 Then
            BT_GuessSlotLabel = "NOMOR_PUNGGUNG"
        ElseIf h >= 7 Then
            BT_GuessSlotLabel = "NOMOR_DADA_TENGAH"
        Else
            BT_GuessSlotLabel = "NOMOR_SLOT"
        End If
        Exit Function
    End If

    If role = "NAMA_ATLIT" Or role = "NAMA" Or role = "NICKNAME" Then
        If h >= 3 Then
            BT_GuessSlotLabel = role & "_BESAR"
        Else
            BT_GuessSlotLabel = role & "_KECIL"
        End If
        Exit Function
    End If

    BT_GuessSlotLabel = role & "_SLOT"

End Function

Private Function BT_Normalize(ByVal s As String) As String

    s = Replace(s, vbCr, "")
    s = Replace(s, vbLf, "")
    s = Replace(s, Chr(160), " ")

    On Error Resume Next
    s = Replace(s, ChrW(&HFB01), "FI")
    s = Replace(s, ChrW(&HFB02), "FL")
    On Error GoTo 0

    Do While InStr(1, s, "  ", vbTextCompare) > 0
        s = Replace(s, "  ", " ")
    Loop

    BT_Normalize = UCase$(Trim$(s))

End Function

Private Function BT_IsNumberPlaceholder(ByVal s As String) As Boolean

    s = Trim$(s)

    If Len(s) = 0 Then Exit Function
    If Len(s) > 3 Then Exit Function

    If IsNumeric(s) Then
        BT_IsNumberPlaceholder = True
    End If

End Function

Private Function BT_IsAttributeMarker(ByVal s As String) As Boolean

    s = UCase$(Trim$(s))

    If Left$(s, 3) = "@A:" Then
        BT_IsAttributeMarker = True
        Exit Function
    End If

    If Left$(s, 6) = "@ATTR:" Then
        BT_IsAttributeMarker = True
        Exit Function
    End If

End Function

Private Function BT_IgnoreSmallID(ByVal shpText As Shape, ByVal txt As String) As Boolean

    txt = Trim$(txt)

    If BT_Normalize(txt) = "IDPO" Then
        If shpText.SizeHeight >= BT_ID_MIN_H And shpText.SizeHeight <= BT_ID_MAX_H Then
            BT_IgnoreSmallID = True
        End If
        Exit Function
    End If

    If Len(txt) = 6 And IsNumeric(txt) Then
        If shpText.SizeHeight >= BT_ID_MIN_H And shpText.SizeHeight <= BT_ID_MAX_H Then
            BT_IgnoreSmallID = True
        End If
    End If

End Function

Private Function BT_ReadOrderMetaValue(ByVal keyName As String) As String

    On Error GoTo FAIL

    Dim path As String
    Dim f As Integer
    Dim ln As String
    Dim p As Long
    Dim k As String
    Dim v As String

    path = Environ$("USERPROFILE") & "\Documents\Order.txt"

    If Dir(path) = "" Then Exit Function

    keyName = UCase$(Trim$(keyName))

    f = FreeFile
    Open path For Input As #f

    Do Until EOF(f)
        Line Input #f, ln
        ln = Trim$(ln)

        If Left$(ln, 1) = "@" Then
            p = InStr(1, ln, "=", vbTextCompare)
            If p > 2 Then
                k = UCase$(Trim$(Mid$(ln, 2, p - 2)))
                v = Trim$(Mid$(ln, p + 1))
                If k = keyName Then
                    BT_ReadOrderMetaValue = v
                    Close #f
                    Exit Function
                End If
            End If
        End If
    Loop

    Close #f
    Exit Function

FAIL:
    On Error Resume Next
    Close #f
    On Error GoTo 0

End Function
