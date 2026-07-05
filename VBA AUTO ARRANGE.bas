Option Explicit

'==========================================================
' PROJECT HADES
' AUTO ARRANGE MASTER POLA V1.0
'
' Fungsi:
' - Menyusun master pola yang terseleksi dari kiri ke kanan.
' - Urutan berdasarkan ukuran bounding box terkecil ke terbesar.
' - Cocok setelah drag beberapa file pola ke workspace CorelDRAW.
'
' Prinsip:
' - Tidak ungroup.
' - Tidak menembus group.
' - Tidak mengubah isi objek.
' - Hanya memindahkan posisi top-level selected objects.
'
' Cara pakai:
' 1. Drag beberapa pola ke workspace CorelDRAW.
' 2. Pastikan tiap pola / tiap size menjadi 1 group utama.
' 3. Select semua group master pola.
' 4. Run AUTO_ARRANGE_MASTER_POLA.
'==========================================================


Sub AUTO_ARRANGE_MASTER_POLA()

    On Error GoTo ErrHandler

    Dim oldUnit As cdrUnit
    oldUnit = ActiveDocument.Unit
    ActiveDocument.Unit = cdrCentimeter

    If ActiveSelection Is Nothing Then
        MsgBox "Tidak ada objek yang dipilih." & vbCrLf & _
               "Select dulu master pola yang ingin disusun.", _
               vbExclamation, "AUTO ARRANGE MASTER POLA"
        GoTo SafeExit
    End If

    If ActiveSelection.Shapes.Count < 2 Then
        MsgBox "Minimal pilih 2 objek / group master pola.", _
               vbExclamation, "AUTO ARRANGE MASTER POLA"
        GoTo SafeExit
    End If

    Dim gapText As String
    Dim gapX As Double

    gapText = InputBox( _
        "Masukkan jarak antar master pola dalam cm:" & vbCrLf & _
        "Contoh: 5", _
        "AUTO ARRANGE MASTER POLA", _
        "5" _
    )

    If Trim$(gapText) = "" Then GoTo SafeExit

    gapText = Replace(gapText, ",", ".")
    gapX = Val(gapText)

    If gapX < 0 Then gapX = 5

    ActiveDocument.BeginCommandGroup "AUTO ARRANGE MASTER POLA"

    Dim sr As ShapeRange
    Set sr = ActiveSelectionRange

    Dim n As Long
    n = sr.Shapes.Count

    Dim arrSh() As Shape
    Dim arrX() As Double
    Dim arrY() As Double
    Dim arrW() As Double
    Dim arrH() As Double
    Dim arrKey() As Double

    ReDim arrSh(1 To n)
    ReDim arrX(1 To n)
    ReDim arrY(1 To n)
    ReDim arrW(1 To n)
    ReDim arrH(1 To n)
    ReDim arrKey(1 To n)

    Dim i As Long
    Dim x As Double
    Dim y As Double
    Dim w As Double
    Dim h As Double

    Dim anchorLeft As Double
    Dim anchorTop As Double

    For i = 1 To n

        Set arrSh(i) = sr.Shapes(i)

        ' GetBoundingBox:
        ' x = kiri
        ' y = bawah
        ' w = lebar
        ' h = tinggi
        arrSh(i).GetBoundingBox x, y, w, h, True

        arrX(i) = x
        arrY(i) = y
        arrW(i) = w
        arrH(i) = h

        ' Sort key:
        ' area bounding box.
        ' Umumnya size kecil memiliki area lebih kecil.
        arrKey(i) = w * h

        If i = 1 Then
            anchorLeft = x
            anchorTop = y + h
        Else
            If x < anchorLeft Then anchorLeft = x
            If y + h > anchorTop Then anchorTop = y + h
        End If

    Next i

    ' Urutkan dari area terkecil ke terbesar
    SortShapeArraysByKey arrSh, arrX, arrY, arrW, arrH, arrKey, n

    ' Susun horizontal dari kiri ke kanan
    Dim currentX As Double
    Dim targetX As Double
    Dim targetY As Double
    Dim dx As Double
    Dim dy As Double

    currentX = anchorLeft

    For i = 1 To n

        targetX = currentX
        targetY = anchorTop - arrH(i)

        dx = targetX - arrX(i)
        dy = targetY - arrY(i)

        arrSh(i).Move dx, dy

        currentX = currentX + arrW(i) + gapX

    Next i

    ActiveDocument.EndCommandGroup

    MsgBox "AUTO ARRANGE MASTER POLA selesai." & vbCrLf & vbCrLf & _
           "Jumlah objek/group disusun : " & n & vbCrLf & _
           "Jarak antar pola           : " & gapX & " cm" & vbCrLf & vbCrLf & _
           "Urutan: kiri terkecil → kanan terbesar", _
           vbInformation, "PROJECT HADES"

SafeExit:
    On Error Resume Next
    ActiveDocument.Unit = oldUnit
    Exit Sub

ErrHandler:
    On Error Resume Next
    ActiveDocument.EndCommandGroup
    ActiveDocument.Unit = oldUnit

    MsgBox "ERROR AUTO ARRANGE MASTER POLA:" & vbCrLf & _
           Err.Description, _
           vbCritical, "PROJECT HADES"

End Sub


Private Sub SortShapeArraysByKey( _
    ByRef arrSh() As Shape, _
    ByRef arrX() As Double, _
    ByRef arrY() As Double, _
    ByRef arrW() As Double, _
    ByRef arrH() As Double, _
    ByRef arrKey() As Double, _
    ByVal n As Long)

    Dim i As Long
    Dim j As Long

    Dim tmpShape As Shape
    Dim tmpD As Double

    For i = 1 To n - 1
        For j = i + 1 To n

            If arrKey(j) < arrKey(i) Then

                Set tmpShape = arrSh(i)
                Set arrSh(i) = arrSh(j)
                Set arrSh(j) = tmpShape

                tmpD = arrX(i)
                arrX(i) = arrX(j)
                arrX(j) = tmpD

                tmpD = arrY(i)
                arrY(i) = arrY(j)
                arrY(j) = tmpD

                tmpD = arrW(i)
                arrW(i) = arrW(j)
                arrW(j) = tmpD

                tmpD = arrH(i)
                arrH(i) = arrH(j)
                arrH(j) = tmpD

                tmpD = arrKey(i)
                arrKey(i) = arrKey(j)
                arrKey(j) = tmpD

            End If

        Next j
    Next i

End Sub