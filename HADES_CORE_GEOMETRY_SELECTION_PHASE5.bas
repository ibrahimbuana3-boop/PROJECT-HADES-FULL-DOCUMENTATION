Option Explicit

'=========================================================
' PROJECT HADES — CORE GEOMETRY / SELECTION PHASE 5
'=========================================================

Private Const H5_SIZE_TOL As Double = 1#
Private Const H5_PANTS_TOL As Double = 0.35

Public Function H5_HasSelection() As Boolean
    On Error Resume Next
    H5_HasSelection = False
    If ActiveSelection Is Nothing Then Exit Function
    If ActiveSelection.Shapes.Count <= 0 Then Exit Function
    H5_HasSelection = True
    On Error GoTo 0
End Function

Public Function H5_IsRedRGB(ByVal r As Long, ByVal g As Long, ByVal b As Long) As Boolean
    If r >= 150 And g <= 150 And b <= 150 Then
        If r > g + 35 And r > b + 35 Then H5_IsRedRGB = True
    End If
End Function

Public Function H5_IsGreenRGB(ByVal r As Long, ByVal g As Long, ByVal b As Long) As Boolean
    If r <= 80 And g >= 180 And b <= 80 Then
        H5_IsGreenRGB = True
        Exit Function
    End If

    If Abs(r - 97) <= 25 And Abs(g - 186) <= 25 And Abs(b - 12) <= 25 Then
        H5_IsGreenRGB = True
    End If
End Function

Public Function H5_IsRedShape(ByVal s As Shape) As Boolean
    Dim r As Long
    Dim g As Long
    Dim b As Long

    On Error Resume Next
    H5_IsRedShape = False

    If Not s.Outline Is Nothing Then
        If s.Outline.Type <> cdrNoOutline Then
            r = s.Outline.Color.RGBRed
            g = s.Outline.Color.RGBGreen
            b = s.Outline.Color.RGBBlue
            If H5_IsRedRGB(r, g, b) Then
                H5_IsRedShape = True
                Exit Function
            End If
        End If
    End If

    If s.Fill.Type = cdrUniformFill Then
        r = s.Fill.UniformColor.RGBRed
        g = s.Fill.UniformColor.RGBGreen
        b = s.Fill.UniformColor.RGBBlue
        If H5_IsRedRGB(r, g, b) Then
            H5_IsRedShape = True
            Exit Function
        End If
    End If

    On Error GoTo 0
End Function

Public Function H5_IsGreenOutline(ByVal s As Shape) As Boolean
    Dim r As Long
    Dim g As Long
    Dim b As Long

    On Error Resume Next
    H5_IsGreenOutline = False

    If s.Outline.Width <= 0 Then Exit Function

    r = s.Outline.Color.RGBRed
    g = s.Outline.Color.RGBGreen
    b = s.Outline.Color.RGBBlue

    H5_IsGreenOutline = H5_IsGreenRGB(r, g, b)
    On Error GoTo 0
End Function

Public Function H5_IsPanelOutline(ByVal s As Shape) As Boolean
    H5_IsPanelOutline = False

    On Error Resume Next
    If s.Type <> cdrCurveShape Then Exit Function

    If H5_IsRedShape(s) Then
        H5_IsPanelOutline = True
        Exit Function
    End If

    If H5_IsGreenOutline(s) Then
        H5_IsPanelOutline = True
        Exit Function
    End If

    On Error GoTo 0
End Function

Public Function H5_DetectSizeFromDimensions( _
    ByVal w As Double, _
    ByVal h As Double, _
    ByVal SizeDB As Object, _
    ByVal isPants As Boolean, _
    ByVal isSplitFront As Boolean) As String

    Dim mn As Double
    Dim mx As Double
    Dim key As Variant
    Dim db As Variant
    Dim errA As Double
    Dim errB As Double
    Dim err As Double
    Dim bestSize As String
    Dim bestErr As Double

    mn = H5_MinD(Abs(w), Abs(h))
    mx = H5_MaxD(Abs(w), Abs(h))

    bestSize = ""
    bestErr = 999999#

    If SizeDB Is Nothing Then Exit Function

    For Each key In SizeDB.keys
        db = SizeDB(key)

        If isPants Then
            If UBound(db) >= 2 Then
                errA = H5_MinD(Abs(mn - H5_ToDbl(db(1))), Abs(mx - H5_ToDbl(db(1))))
                errB = H5_MinD(Abs(mn - H5_ToDbl(db(2))), Abs(mx - H5_ToDbl(db(2))))
                err = H5_MinD(errA, errB)

                If err <= H5_PANTS_TOL Then
                    If err < bestErr Then
                        bestErr = err
                        bestSize = CStr(key)
                    End If
                End If
            End If

        ElseIf isSplitFront Then
            If UBound(db) >= 4 Then
                errA = Abs(mn - H5_ToDbl(db(1))) + Abs(mx - H5_ToDbl(db(4)))
                If errA <= (H5_SIZE_TOL * 2) Then
                    If errA < bestErr Then
                        bestErr = errA
                        bestSize = CStr(key)
                    End If
                End If

                errB = Abs(mn - H5_ToDbl(db(2))) + Abs(mx - H5_ToDbl(db(3)))
                If errB <= (H5_SIZE_TOL * 2) Then
                    If errB < bestErr Then
                        bestErr = errB
                        bestSize = CStr(key)
                    End If
                End If
            End If

        Else
            If UBound(db) >= 3 Then
                If Abs(mn - H5_ToDbl(db(1))) <= H5_SIZE_TOL Then
                    errA = Abs(mx - H5_ToDbl(db(2)))
                    errB = Abs(mx - H5_ToDbl(db(3)))
                    err = Abs(mn - H5_ToDbl(db(1))) + H5_MinD(errA, errB)

                    If errA <= H5_SIZE_TOL Or errB <= H5_SIZE_TOL Then
                        If err < bestErr Then
                            bestErr = err
                            bestSize = CStr(key)
                        End If
                    End If
                End If
            End If
        End If
    Next key

    H5_DetectSizeFromDimensions = bestSize
End Function

Public Function H5_BuildCurrentSelectionSignature(ByRef detailText As String) As String
    Dim sr As ShapeRange
    Dim s As Shape
    Dim raw As String
    Dim countTop As Long
    Dim minX As Double
    Dim minY As Double
    Dim maxX As Double
    Dim maxY As Double
    Dim initialized As Boolean

    detailText = ""
    H5_BuildCurrentSelectionSignature = "NO_SELECTION"

    If Not H5_HasSelection() Then Exit Function

    Set sr = ActiveSelectionRange

    For Each s In sr.Shapes
        countTop = countTop + 1
        H5_ExpandBBox s, minX, minY, maxX, maxY, initialized
        raw = raw & H5_ShapeSignatureLine(s) & vbLf
    Next s

    detailText = _
        "TOP_LEVEL_SHAPES=" & CStr(countTop) & vbCrLf & _
        "BBOX_MIN_X=" & Format$(minX, "0.000") & vbCrLf & _
        "BBOX_MIN_Y=" & Format$(minY, "0.000") & vbCrLf & _
        "BBOX_MAX_X=" & Format$(maxX, "0.000") & vbCrLf & _
        "BBOX_MAX_Y=" & Format$(maxY, "0.000") & vbCrLf

    H5_BuildCurrentSelectionSignature = _
        "COUNT=" & CStr(countTop) & _
        "|MINX=" & Format$(minX, "0.000") & _
        "|MINY=" & Format$(minY, "0.000") & _
        "|MAXX=" & Format$(maxX, "0.000") & _
        "|MAXY=" & Format$(maxY, "0.000") & _
        "|HASH=" & CStr(H5_SimpleHash(raw))
End Function

Private Sub H5_ExpandBBox( _
    ByVal s As Shape, _
    ByRef minX As Double, _
    ByRef minY As Double, _
    ByRef maxX As Double, _
    ByRef maxY As Double, _
    ByRef initialized As Boolean)

    Dim x1 As Double
    Dim y1 As Double
    Dim x2 As Double
    Dim y2 As Double

    On Error Resume Next

    x1 = s.LeftX
    y1 = s.BottomY
    x2 = s.RightX
    y2 = s.TopY

    If Not initialized Then
        minX = x1
        minY = y1
        maxX = x2
        maxY = y2
        initialized = True
    Else
        If x1 < minX Then minX = x1
        If y1 < minY Then minY = y1
        If x2 > maxX Then maxX = x2
        If y2 > maxY Then maxY = y2
    End If

    On Error GoTo 0
End Sub

Private Function H5_ShapeSignatureLine(ByVal s As Shape) As String
    On Error Resume Next
    H5_ShapeSignatureLine = _
        CStr(s.Type) & "|" & _
        Format$(s.LeftX, "0.000") & "|" & _
        Format$(s.BottomY, "0.000") & "|" & _
        Format$(s.SizeWidth, "0.000") & "|" & _
        Format$(s.SizeHeight, "0.000")
    On Error GoTo 0
End Function

Public Function H5_SimpleHash(ByVal s As String) As Long
    Dim i As Long
    Dim h As Double
    Dim codePoint As Long

    h = 5381

    For i = 1 To Len(s)
        codePoint = AscW(Mid$(s, i, 1))
        If codePoint < 0 Then codePoint = codePoint + 65536
        h = (h * 33 + codePoint)
        Do While h > 2147483647#
            h = h - 2147483647#
        Loop
    Next i

    H5_SimpleHash = CLng(h)
End Function

Public Function H5_CountActiveTextInSelection() As Long
    Dim sr As ShapeRange
    Dim s As Shape

    If Not H5_HasSelection() Then Exit Function

    Set sr = ActiveSelectionRange

    For Each s In sr.Shapes
        H5_CountActiveTextInSelection = H5_CountActiveTextInSelection + H5_CountActiveTextRecursive(s)
    Next s
End Function

Private Function H5_CountActiveTextRecursive(ByVal s As Shape) As Long
    Dim ch As Shape
    Dim pcShapes As Shapes

    On Error Resume Next

    If s.Type = cdrTextShape Then
        H5_CountActiveTextRecursive = 1
        Exit Function
    End If

    If s.Type = cdrGroupShape Then
        For Each ch In s.Shapes
            H5_CountActiveTextRecursive = H5_CountActiveTextRecursive + H5_CountActiveTextRecursive(ch)
        Next ch
    End If

    Set pcShapes = s.PowerClip.Shapes
    If Not pcShapes Is Nothing Then
        For Each ch In pcShapes
            H5_CountActiveTextRecursive = H5_CountActiveTextRecursive + H5_CountActiveTextRecursive(ch)
        Next ch
    End If

    On Error GoTo 0
End Function
