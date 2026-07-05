Option Explicit

'=================================================================
' PROJECT HADES — DATABASE MINER SINGLE / JERSEY BODY
'
' PURPOSE:
' Menambang dimensi pola body depan + belakang dari size terkecil
' sampai terbesar dan menyimpannya sebagai SizeDB_*.txt.
'
' OUTPUT FORMAT:
' SIZE|LEBAR|TINGGI_DEPAN|TINGGI_BELAKANG
'
' CATATAN:
' - Digunakan sebagai tool database/admin, bukan workflow layout harian.
' - Select panel body depan + belakang dari semua size yang ingin ditambang.
' - Target panel adalah curve dengan outline merah.
'=================================================================

Private PanelsCount As Long
Private RawWidths() As Double
Private RawHeights() As Double

Private Const MINE_TOL As Double = 1#  ' cm

Public Sub MINE_SIZE_DATABASE()

    If ActiveSelection Is Nothing Or ActiveSelection.Shapes.Count = 0 Then
        MsgBox "MINING ERROR:" & vbCrLf & _
               "Harap blok pola body depan & belakang dari size terkecil sampai terbesar.", _
               vbExclamation, "Hades Miner"
        Exit Sub
    End If

    PanelsCount = 0
    ReDim RawWidths(1 To 1000)
    ReDim RawHeights(1 To 1000)

    Dim origUnit As cdrUnit
    origUnit = ActiveDocument.Unit
    ActiveDocument.Unit = cdrCentimeter

    Dim s As Shape
    For Each s In ActiveSelection.Shapes
        Mine_ExtractPanel s
    Next s

    ActiveDocument.Unit = origUnit

    If PanelsCount = 0 Then
        MsgBox "Tidak ada pola valid yang ditemukan." & vbCrLf & _
               "Pastikan outline merah dan objek berupa Curve.", _
               vbCritical, "Hades Miner"
        Exit Sub
    End If

    If PanelsCount Mod 2 <> 0 Then
        MsgBox "MINING GAGAL: jumlah panel ganjil (" & PanelsCount & ")." & vbCrLf & _
               "Setiap size harus punya sepasang body depan + belakang.", _
               vbCritical, "Hades Miner"
        Exit Sub
    End If

    ReDim Preserve RawWidths(1 To PanelsCount)
    ReDim Preserve RawHeights(1 To PanelsCount)

    Mine_SortByWidth PanelsCount, RawWidths, RawHeights

    Dim sizeLabels As Variant
    sizeLabels = Array("XS", "S", "M", "L", "XL", "2XL", "3XL", "4XL", "5XL", "6XL", "7XL", "8XL")

    Dim startSize As String
    startSize = InputBox( _
        "Sistem menemukan " & CStr(PanelsCount / 2) & " size." & vbCrLf & vbCrLf & _
        "Ketik nama size terkecil dari susunan pola ini:", _
        "Set Starting Size", _
        "XS")

    If Trim$(startSize) = "" Then Exit Sub

    Dim startIndex As Long
    startIndex = Mine_FindSizeIndex(sizeLabels, UCase$(Trim$(startSize)))

    If startIndex < 0 Then
        MsgBox "Size '" & startSize & "' tidak standar." & vbCrLf & _
               "Gunakan XS, S, M, L, XL, 2XL, dst.", _
               vbCritical, "Hades Miner"
        Exit Sub
    End If

    If startIndex + (PanelsCount / 2) - 1 > UBound(sizeLabels) Then
        MsgBox "Jumlah size melebihi daftar size label internal.", vbCritical, "Hades Miner"
        Exit Sub
    End If

    Dim outputText As String
    Dim pairIndex As Long
    Dim i As Long

    pairIndex = startIndex
    i = 1

    Do While i <= PanelsCount

        Dim W1 As Double, W2 As Double
        Dim H1 As Double, H2 As Double

        W1 = RawWidths(i)
        H1 = RawHeights(i)
        W2 = RawWidths(i + 1)
        H2 = RawHeights(i + 1)

        If Abs(W1 - W2) <= MINE_TOL Then

            Dim finalW As Double
            Dim frontH As Double
            Dim backH As Double

            finalW = Round((W1 + W2) / 2#, 3)

            If H1 > H2 Then
                frontH = Round(H1, 3)
                backH = Round(H2, 3)
            Else
                frontH = Round(H2, 3)
                backH = Round(H1, 3)
            End If

            outputText = outputText & _
                         CStr(sizeLabels(pairIndex)) & "|" & _
                         Mine_FormatNum(finalW) & "|" & _
                         Mine_FormatNum(frontH) & "|" & _
                         Mine_FormatNum(backH) & vbCrLf

            pairIndex = pairIndex + 1
            i = i + 2

        Else
            MsgBox "MINING GAGAL: pasangan panel tidak sinkron." & vbCrLf & _
                   "Width: " & Mine_FormatNum(W1) & " vs " & Mine_FormatNum(W2), _
                   vbCritical, "Hades Miner"
            Exit Sub
        End If

    Loop

    Dim fileNameInput As String
    fileNameInput = InputBox( _
        "Data berhasil ditambang." & vbCrLf & _
        "Simpan dengan nama file apa?", _
        "Save SizeDB", _
        "SizeDB_Baru.txt")

    If Trim$(fileNameInput) = "" Then Exit Sub
    If LCase$(Right$(fileNameInput, 4)) <> ".txt" Then fileNameInput = fileNameInput & ".txt"

    Dim savePath As String
    savePath = Environ$("USERPROFILE") & "\Documents\" & fileNameInput

    Dim f As Integer
    f = FreeFile
    Open savePath For Output As #f
    Print #f, Trim$(outputText)
    Close #f

    MsgBox "Penambangan database berhasil." & vbCrLf & vbCrLf & _
           "File tersimpan di:" & vbCrLf & savePath, _
           vbInformation, "Mining Complete"

End Sub

Private Sub Mine_ExtractPanel(ByVal shp As Shape)

    Dim c As Shape

    If shp.Type = cdrGroupShape Then
        For Each c In shp.Shapes
            Mine_ExtractPanel c
        Next c
        Exit Sub
    End If

    If shp.Type <> cdrCurveShape Then Exit Sub
    If Not Mine_IsRedOutline(shp) Then Exit Sub

    Dim w As Double, h As Double
    Dim maxD As Double, minD As Double

    w = shp.SizeWidth
    h = shp.SizeHeight

    If w > h Then
        maxD = w
        minD = h
    Else
        maxD = h
        minD = w
    End If

    PanelsCount = PanelsCount + 1
    RawWidths(PanelsCount) = minD
    RawHeights(PanelsCount) = maxD

End Sub

Private Function Mine_IsRedOutline(ByVal shp As Shape) As Boolean

    On Error Resume Next

    Mine_IsRedOutline = False

    If shp.Outline Is Nothing Then Exit Function
    If shp.Outline.Type = cdrNoOutline Then Exit Function

    Dim r As Long, g As Long, b As Long
    r = shp.Outline.Color.RGBRed
    g = shp.Outline.Color.RGBGreen
    b = shp.Outline.Color.RGBBlue

    If (r > 200) And (g < 60) And (b < 60) Then
        Mine_IsRedOutline = True
    End If

    On Error GoTo 0

End Function

Private Sub Mine_SortByWidth(ByVal n As Long, ByRef arrW() As Double, ByRef arrH() As Double)

    Dim i As Long, j As Long
    Dim tempW As Double, tempH As Double

    For i = 1 To n - 1
        For j = i + 1 To n
            If arrW(i) > arrW(j) Then
                tempW = arrW(i): arrW(i) = arrW(j): arrW(j) = tempW
                tempH = arrH(i): arrH(i) = arrH(j): arrH(j) = tempH
            End If
        Next j
    Next i

End Sub

Private Function Mine_FindSizeIndex(ByVal labels As Variant, ByVal sizeName As String) As Long

    Dim k As Long

    Mine_FindSizeIndex = -1

    For k = LBound(labels) To UBound(labels)
        If UCase$(Trim$(CStr(labels(k)))) = sizeName Then
            Mine_FindSizeIndex = k
            Exit Function
        End If
    Next k

End Function

Private Function Mine_FormatNum(ByVal v As Double) As String
    Mine_FormatNum = Trim$(Replace(CStr(Round(v, 3)), ",", "."))
End Function
