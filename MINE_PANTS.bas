Option Explicit

'=========================================================
' PROJECT HADES — MINING PANTS WIDTH ONLY V2
'
' Untuk celana 4 panel:
' - Depan Kanan
' - Depan Kiri
' - Belakang Kanan
' - Belakang Kiri
'
' Output SizeDB:
' SIZE|L_DEPAN|L_BELAKANG
'
' Tujuan:
' - QC Size Pants hanya mengecek SIZE
' - Tidak membedakan lis besar / lis kecil / tanpa lis
' - Mengurangi jumlah database
'
' Cara pakai:
' 1. Kumpulkan semua pola celana dari size terkecil ke terbesar
' 2. Tiap size harus terdiri dari 4 panel merah
' 3. Select semua panel
' 4. Run MINE_PANTS_DATABASE
' 5. Input size terkecil
' 6. Input nama file SizeDB
'=========================================================

Dim PantsCount As Long

Dim PantsW() As Double
Dim PantsH() As Double
Dim PantsCX() As Double
Dim PantsCY() As Double

Dim BlockCount As Long
Dim BlockFrontW() As Double
Dim BlockBackW() As Double
Dim BlockAvgW() As Double

Const PAIR_TOL As Double = 2#
Const ROW_TOL As Double = 20#

Const FRONT_IS_SMALLER As Boolean = True

'=========================================================
' MAIN
'=========================================================

Sub MINE_PANTS_DATABASE()

```
If ActiveSelection Is Nothing Then

    MsgBox _
        "MINING ERROR:" & vbCrLf & _
        "Blok semua pola celana terlebih dahulu.", _
        vbExclamation, _
        "Mining Pants"

    Exit Sub

End If

If ActiveSelection.Shapes.Count = 0 Then

    MsgBox _
        "MINING ERROR:" & vbCrLf & _
        "Blok semua pola celana terlebih dahulu.", _
        vbExclamation, _
        "Mining Pants"

    Exit Sub

End If

PantsCount = 0

ReDim PantsW(1 To 3000)
ReDim PantsH(1 To 3000)
ReDim PantsCX(1 To 3000)
ReDim PantsCY(1 To 3000)

Dim oldUnit As cdrUnit

oldUnit = ActiveDocument.Unit

On Error GoTo ERR_HANDLER

ActiveDocument.Unit = cdrCentimeter

Dim s As Shape

For Each s In ActiveSelection.Shapes
    ExtractPantsPanel s
Next

ActiveDocument.Unit = oldUnit

If PantsCount = 0 Then

    MsgBox _
        "MINING GAGAL:" & vbCrLf & _
        "Tidak ada panel merah valid yang ditemukan." & vbCrLf & _
        "Pastikan outline pola berwarna merah dan objek berupa Curve.", _
        vbCritical, _
        "Mining Pants"

    Exit Sub

End If

If PantsCount Mod 4 <> 0 Then

    MsgBox _
        "MINING GAGAL:" & vbCrLf & _
        "Jumlah panel harus kelipatan 4." & vbCrLf & vbCrLf & _
        "Panel ditemukan : " & PantsCount & vbCrLf & _
        "Setiap size celana wajib punya 4 panel.", _
        vbCritical, _
        "Mining Pants"

    Exit Sub

End If

ReDim Preserve PantsW(1 To PantsCount)
ReDim Preserve PantsH(1 To PantsCount)
ReDim Preserve PantsCX(1 To PantsCount)
ReDim Preserve PantsCY(1 To PantsCount)

'Urutkan panel secara visual agar 4 panel yang berdekatan
'masuk sebagai 1 block size.
SortPanelsVisual

BlockCount = PantsCount / 4

ReDim BlockFrontW(1 To BlockCount)
ReDim BlockBackW(1 To BlockCount)
ReDim BlockAvgW(1 To BlockCount)

If Not BuildBlocks Then Exit Sub

'Setelah setiap block punya front/back width,
'urutkan block dari size terkecil ke terbesar berdasarkan rata-rata lebar.
SortBlocksByWidth

Dim sizeLabels As Variant

sizeLabels = Array( _
    "XXS", "XS", "S", "M", "L", "XL", _
    "2XL", "3XL", "4XL", "5XL", "6XL", _
    "7XL", "8XL", "9XL", "10XL")

Dim startSize As String

startSize = InputBox( _
    "Sistem menemukan " & BlockCount & " size celana." & vbCrLf & vbCrLf & _
    "Ketik nama size TERKECIL dari susunan pola ini:", _
    "Set Starting Size", _
    "XS")

If startSize = "" Then Exit Sub

Dim startIndex As Long

startIndex = FindSizeIndex(sizeLabels, startSize)

If startIndex < 0 Then

    MsgBox _
        "Size '" & startSize & "' tidak dikenal." & vbCrLf & _
        "Gunakan: XXS, XS, S, M, L, XL, 2XL, dst.", _
        vbCritical, _
        "Mining Pants"

    Exit Sub

End If

If startIndex + BlockCount - 1 > UBound(sizeLabels) Then

    MsgBox _
        "Jumlah size melebihi daftar sizeLabels." & vbCrLf & _
        "Tambahkan label size di array sizeLabels.", _
        vbCritical, _
        "Mining Pants"

    Exit Sub

End If

Dim outputText As String
Dim i As Long

For i = 1 To BlockCount

    outputText = outputText & _
        sizeLabels(startIndex + i - 1) & "|" & _
        FormatNum(Round(BlockFrontW(i), 3)) & "|" & _
        FormatNum(Round(BlockBackW(i), 3)) & vbCrLf

Next i

Dim fileNameInput As String

fileNameInput = InputBox( _
    "Data celana berhasil ditambang!" & vbCrLf & vbCrLf & _
    "Format output:" & vbCrLf & _
    "SIZE|L_DEPAN|L_BELAKANG" & vbCrLf & vbCrLf & _
    "Simpan dengan nama file apa?", _
    "Save Pants Database", _
    "SizeDB_Celana.txt")

If fileNameInput = "" Then Exit Sub

If LCase(Right(fileNameInput, 4)) <> ".txt" Then
    fileNameInput = fileNameInput & ".txt"
End If

Dim savePath As String

savePath = Environ("USERPROFILE") & "\Documents\" & fileNameInput

Dim f As Integer

f = FreeFile

Open savePath For Output As #f
Print #f, Trim(outputText)
Close #f

MsgBox _
    "Mining Pants berhasil!" & vbCrLf & vbCrLf & _
    "Jumlah panel : " & PantsCount & vbCrLf & _
    "Jumlah size  : " & BlockCount & vbCrLf & vbCrLf & _
    "File tersimpan di:" & vbCrLf & _
    savePath & vbCrLf & vbCrLf & _
    "Format:" & vbCrLf & _
    "SIZE|L_DEPAN|L_BELAKANG", _
    vbInformation, _
    "Mining Pants Complete"

Exit Sub
```

ERR_HANDLER:

```
On Error Resume Next

ActiveDocument.Unit = oldUnit

On Error GoTo 0

MsgBox _
    "SYSTEM ERROR - MINING PANTS" & vbCrLf & vbCrLf & _
    "No : " & Err.Number & vbCrLf & _
    Err.Description, _
    vbCritical, _
    "Mining Pants"
```

End Sub

'=========================================================
' EXTRACT PANEL
'=========================================================

Sub ExtractPantsPanel(ByVal shp As Shape)

```
Dim c As Shape

If shp.Type = cdrGroupShape Then

    For Each c In shp.Shapes
        ExtractPantsPanel c
    Next

    Exit Sub

End If

If shp.Type <> cdrCurveShape Then Exit Sub
If Not IsRed(shp) Then Exit Sub

Dim w As Double
Dim h As Double

w = shp.SizeWidth
h = shp.SizeHeight

If w <= 0 Or h <= 0 Then Exit Sub

PantsCount = PantsCount + 1

'Lebar pola = dimensi pendek.
'Tinggi pola = dimensi panjang, hanya untuk pairing kanan/kiri.
If w > h Then

    PantsW(PantsCount) = Round(h, 3)
    PantsH(PantsCount) = Round(w, 3)

Else

    PantsW(PantsCount) = Round(w, 3)
    PantsH(PantsCount) = Round(h, 3)

End If

PantsCX(PantsCount) = shp.CenterX
PantsCY(PantsCount) = shp.CenterY
```

End Sub

'=========================================================
' SORT PANEL VISUAL
'=========================================================

Sub SortPanelsVisual()

```
Dim i As Long
Dim j As Long

For i = 1 To PantsCount - 1

    For j = i + 1 To PantsCount

        If NeedSwapVisual(i, j) Then
            SwapPanels i, j
        End If

    Next j

Next i
```

End Sub

Function NeedSwapVisual(ByVal i As Long, ByVal j As Long) As Boolean

```
'Corel umumnya Y lebih besar berarti lebih atas.
'Urutan: atas ke bawah, kiri ke kanan.
'Kalau susunanmu terbalik, bagian CY ini bisa dibalik.

If Abs(PantsCY(i) - PantsCY(j)) > ROW_TOL Then

    NeedSwapVisual = PantsCY(i) < PantsCY(j)

Else

    NeedSwapVisual = PantsCX(i) > PantsCX(j)

End If
```

End Function

Sub SwapPanels(ByVal i As Long, ByVal j As Long)

```
Dim tw As Double
Dim th As Double
Dim tx As Double
Dim ty As Double

tw = PantsW(i)
th = PantsH(i)
tx = PantsCX(i)
ty = PantsCY(i)

PantsW(i) = PantsW(j)
PantsH(i) = PantsH(j)
PantsCX(i) = PantsCX(j)
PantsCY(i) = PantsCY(j)

PantsW(j) = tw
PantsH(j) = th
PantsCX(j) = tx
PantsCY(j) = ty
```

End Sub

'=========================================================
' BUILD BLOCKS
'=========================================================

Function BuildBlocks() As Boolean

```
BuildBlocks = False

Dim b As Long
Dim idx As Long

idx = 1

For b = 1 To BlockCount

    Dim w1 As Double, h1 As Double
    Dim w2 As Double, h2 As Double
    Dim w3 As Double, h3 As Double
    Dim w4 As Double, h4 As Double

    w1 = PantsW(idx)
    h1 = PantsH(idx)

    w2 = PantsW(idx + 1)
    h2 = PantsH(idx + 1)

    w3 = PantsW(idx + 2)
    h3 = PantsH(idx + 2)

    w4 = PantsW(idx + 3)
    h4 = PantsH(idx + 3)

    Dim a1 As Long
    Dim a2 As Long
    Dim b1 As Long
    Dim b2 As Long
    Dim bestScore As Double

    If Not BestPairing4( _
        w1, h1, w2, h2, w3, h3, w4, h4, _
        a1, a2, b1, b2, bestScore) _
    Then

        MsgBox _
            "MINING GAGAL:" & vbCrLf & _
            "Gagal memasangkan 4 panel pada block size ke-" & b & "." & vbCrLf & vbCrLf & _
            "Kemungkinan:" & vbCrLf & _
            "- Ada panel kurang / lebih" & vbCrLf & _
            "- Panel beda size masuk dalam satu block" & vbCrLf & _
            "- Susunan visual terlalu campur" & vbCrLf & vbCrLf & _
            "Mismatch score: " & FormatNum(bestScore), _
            vbCritical, _
            "Mining Pants"

        Exit Function

    End If

    Dim pairAW As Double
    Dim pairAH As Double
    Dim pairBW As Double
    Dim pairBH As Double

    GetPairAverage a1, a2, _
                   w1, h1, w2, h2, w3, h3, w4, h4, _
                   pairAW, pairAH

    GetPairAverage b1, b2, _
                   w1, h1, w2, h2, w3, h3, w4, h4, _
                   pairBW, pairBH

    'Front/back dari lebar.
    'Default:
    'lebar lebih kecil = depan
    'lebar lebih besar = belakang.

    If FRONT_IS_SMALLER Then

        If pairAW <= pairBW Then

            BlockFrontW(b) = pairAW
            BlockBackW(b) = pairBW

        Else

            BlockFrontW(b) = pairBW
            BlockBackW(b) = pairAW

        End If

    Else

        If pairAW >= pairBW Then

            BlockFrontW(b) = pairAW
            BlockBackW(b) = pairBW

        Else

            BlockFrontW(b) = pairBW
            BlockBackW(b) = pairAW

        End If

    End If

    BlockAvgW(b) = Round((BlockFrontW(b) + BlockBackW(b)) / 2, 3)

    idx = idx + 4

Next b

BuildBlocks = True
```

End Function

'=========================================================
' BEST PAIRING 4 PANELS
'=========================================================

Function BestPairing4( _
ByVal w1 As Double, ByVal h1 As Double, _
ByVal w2 As Double, ByVal h2 As Double, _
ByVal w3 As Double, ByVal h3 As Double, _
ByVal w4 As Double, ByVal h4 As Double, _
ByRef a1 As Long, ByRef a2 As Long, _
ByRef b1 As Long, ByRef b2 As Long, _
ByRef bestScore As Double) As Boolean

```
Dim score1 As Double
Dim score2 As Double
Dim score3 As Double

'Kemungkinan pairing:
'1) (1,2) + (3,4)
'2) (1,3) + (2,4)
'3) (1,4) + (2,3)

score1 = PairDistance(w1, h1, w2, h2) + PairDistance(w3, h3, w4, h4)
score2 = PairDistance(w1, h1, w3, h3) + PairDistance(w2, h2, w4, h4)
score3 = PairDistance(w1, h1, w4, h4) + PairDistance(w2, h2, w3, h3)

bestScore = score1
a1 = 1
a2 = 2
b1 = 3
b2 = 4

If score2 < bestScore Then

    bestScore = score2
    a1 = 1
    a2 = 3
    b1 = 2
    b2 = 4

End If

If score3 < bestScore Then

    bestScore = score3
    a1 = 1
    a2 = 4
    b1 = 2
    b2 = 3

End If

If bestScore <= PAIR_TOL Then
    BestPairing4 = True
Else
    BestPairing4 = False
End If
```

End Function

Function PairDistance( _
ByVal wA As Double, ByVal hA As Double, _
ByVal wB As Double, ByVal hB As Double) As Double

```
'Output hanya memakai lebar.
'Tetapi pairing kanan/kiri tetap memakai width + height
'agar pasangan depan kanan/kiri dan belakang kanan/kiri lebih akurat.

PairDistance = Abs(wA - wB) + Abs(hA - hB)
```

End Function

Sub GetPairAverage( _
ByVal p1 As Long, _
ByVal p2 As Long, _
ByVal w1 As Double, ByVal h1 As Double, _
ByVal w2 As Double, ByVal h2 As Double, _
ByVal w3 As Double, ByVal h3 As Double, _
ByVal w4 As Double, ByVal h4 As Double, _
ByRef avgW As Double, _
ByRef avgH As Double)

```
Dim pw1 As Double
Dim ph1 As Double
Dim pw2 As Double
Dim ph2 As Double

GetPanelDim p1, w1, h1, w2, h2, w3, h3, w4, h4, pw1, ph1
GetPanelDim p2, w1, h1, w2, h2, w3, h3, w4, h4, pw2, ph2

avgW = Round((pw1 + pw2) / 2, 3)
avgH = Round((ph1 + ph2) / 2, 3)
```

End Sub

Sub GetPanelDim( _
ByVal p As Long, _
ByVal w1 As Double, ByVal h1 As Double, _
ByVal w2 As Double, ByVal h2 As Double, _
ByVal w3 As Double, ByVal h3 As Double, _
ByVal w4 As Double, ByVal h4 As Double, _
ByRef outW As Double, _
ByRef outH As Double)

```
Select Case p

    Case 1

        outW = w1
        outH = h1

    Case 2

        outW = w2
        outH = h2

    Case 3

        outW = w3
        outH = h3

    Case 4

        outW = w4
        outH = h4

End Select
```

End Sub

'=========================================================
' SORT BLOCKS BY WIDTH
'=========================================================

Sub SortBlocksByWidth()

```
Dim i As Long
Dim j As Long

For i = 1 To BlockCount - 1

    For j = i + 1 To BlockCount

        If BlockAvgW(i) > BlockAvgW(j) Then
            SwapBlocks i, j
        End If

    Next j

Next i
```

End Sub

Sub SwapBlocks(ByVal i As Long, ByVal j As Long)

```
Dim tf As Double
Dim tb As Double
Dim ta As Double

tf = BlockFrontW(i)
tb = BlockBackW(i)
ta = BlockAvgW(i)

BlockFrontW(i) = BlockFrontW(j)
BlockBackW(i) = BlockBackW(j)
BlockAvgW(i) = BlockAvgW(j)

BlockFrontW(j) = tf
BlockBackW(j) = tb
BlockAvgW(j) = ta
```

End Sub

'=========================================================
' SIZE LABEL
'=========================================================

Function FindSizeIndex(ByVal arr As Variant, ByVal startSize As String) As Long

```
Dim k As Long

FindSizeIndex = -1

startSize = UCase(Trim(startSize))

For k = LBound(arr) To UBound(arr)

    If startSize = UCase(arr(k)) Then
        FindSizeIndex = k
        Exit Function
    End If

Next k
```

End Function

'=========================================================
' RED OUTLINE
'=========================================================

Function IsRed(ByVal shp As Shape) As Boolean

```
On Error Resume Next

IsRed = False

If shp.Outline Is Nothing Then Exit Function
If shp.Outline.Type = cdrNoOutline Then Exit Function

Dim r As Long
Dim g As Long
Dim b As Long

r = shp.Outline.Color.RGBRed
g = shp.Outline.Color.RGBGreen
b = shp.Outline.Color.RGBBlue

If (r > 200) And (g < 60) And (b < 60) Then
    IsRed = True
End If

On Error GoTo 0
```

End Function

'=========================================================
' FORMAT NUMBER
'=========================================================

Function FormatNum(ByVal val As Double) As String

```
FormatNum = Trim(Replace(CStr(val), ",", "."))
```

End Function