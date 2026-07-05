Option Explicit

'============================================================
' PROJECT HADES — UNIVERSAL SPLIT FRONT MINER
'
' PURPOSE:
' Menambang database untuk produk split-front:
' - Jaket
' - Jaket Anak
' - Kemeja
' - Rompi
'
' SELECT:
' - 1 panel belakang
' - 1 panel depan kanan/kiri
'
' OUTPUT FORMAT:
' SIZE|LEBAR_BELAKANG|LEBAR_DEPAN|TINGGI_DEPAN|TINGGI_BELAKANG
'============================================================

Public Sub MINE_SPLIT_FRONT()

    If ActiveSelection.Shapes.Count <> 2 Then
        MsgBox "MINING ERROR:" & vbCrLf & _
               "Harap select tepat:" & vbCrLf & _
               "- 1 Panel Belakang" & vbCrLf & _
               "- 1 Panel Depan (kanan atau kiri)", _
               vbExclamation, "Split Front Miner"
        Exit Sub
    End If

    Dim oldUnit As cdrUnit
    oldUnit = ActiveDocument.Unit
    ActiveDocument.Unit = cdrCentimeter

    Dim W1 As Double, H1 As Double
    Dim W2 As Double, H2 As Double

    W1 = ActiveSelection.Shapes(1).SizeWidth
    H1 = ActiveSelection.Shapes(1).SizeHeight
    W2 = ActiveSelection.Shapes(2).SizeWidth
    H2 = ActiveSelection.Shapes(2).SizeHeight

    Dim Max1 As Double, Min1 As Double
    Dim Max2 As Double, Min2 As Double

    If W1 > H1 Then
        Max1 = W1
        Min1 = H1
    Else
        Max1 = H1
        Min1 = W1
    End If

    If W2 > H2 Then
        Max2 = W2
        Min2 = H2
    Else
        Max2 = H2
        Min2 = W2
    End If

    ActiveDocument.Unit = oldUnit

    Dim LebarBelakang As Double
    Dim LebarDepan As Double
    Dim TinggiBelakang As Double
    Dim TinggiDepan As Double

    If Min1 > Min2 Then
        LebarBelakang = Min1
        TinggiBelakang = Max1
        LebarDepan = Min2
        TinggiDepan = Max2
    Else
        LebarBelakang = Min2
        TinggiBelakang = Max2
        LebarDepan = Min1
        TinggiDepan = Max1
    End If

    Dim sizeName As String
    sizeName = InputBox( _
        "Dimensi terdeteksi:" & vbCrLf & vbCrLf & _
        "Lebar Belakang : " & SplitMine_FormatNum(LebarBelakang) & " cm" & vbCrLf & _
        "Lebar Depan    : " & SplitMine_FormatNum(LebarDepan) & " cm" & vbCrLf & _
        "Tinggi Depan   : " & SplitMine_FormatNum(TinggiDepan) & " cm" & vbCrLf & _
        "Tinggi Belakang: " & SplitMine_FormatNum(TinggiBelakang) & " cm" & vbCrLf & vbCrLf & _
        "Masukkan nama size:", _
        "Input Size")

    If Trim$(sizeName) = "" Then Exit Sub

    Dim fileName As String
    fileName = InputBox( _
        "Simpan ke database mana?" & vbCrLf & vbCrLf & _
        "Contoh:" & vbCrLf & _
        "SizeDB_Jaket.txt" & vbCrLf & _
        "SizeDB_JaketAnak.txt" & vbCrLf & _
        "SizeDB_Kemeja.txt", _
        "Target Database", _
        "SizeDB_Jaket.txt")

    If Trim$(fileName) = "" Then Exit Sub
    If LCase$(Right$(fileName, 4)) <> ".txt" Then fileName = fileName & ".txt"

    Dim output As String
    output = UCase$(Trim$(sizeName)) & "|" & _
             SplitMine_FormatNum(LebarBelakang) & "|" & _
             SplitMine_FormatNum(LebarDepan) & "|" & _
             SplitMine_FormatNum(TinggiDepan) & "|" & _
             SplitMine_FormatNum(TinggiBelakang)

    Dim savePath As String
    savePath = Environ$("USERPROFILE") & "\Documents\" & fileName

    Dim f As Integer
    f = FreeFile

    Open savePath For Append As #f
    Print #f, output
    Close #f

    MsgBox "MINING SUCCESS" & vbCrLf & vbCrLf & _
           output & vbCrLf & vbCrLf & _
           "Disimpan ke:" & vbCrLf & fileName, _
           vbInformation, "Split Front Miner"

End Sub

Private Function SplitMine_FormatNum(ByVal v As Double) As String
    SplitMine_FormatNum = Trim$(Replace(CStr(Round(v, 3)), ",", "."))
End Function
