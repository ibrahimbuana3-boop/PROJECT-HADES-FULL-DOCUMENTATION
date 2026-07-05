Option Explicit

'=========================================================
' PROJECT HADES — IDPO CHECK V3 AUTO METADATA
'
' PURPOSE:
' Mengecek apakah semua IDPO kecil sudah diganti
' menjadi nomor IDPO / Kode Produk target 6 digit.
'
' SUMBER TARGET:
' 1. AUTO dari Documents\Order.txt:
'    @IDPO=355863
'
' 2. Fallback manual popup jika @IDPO tidak ada / kosong / invalid.
'
' Target scan:
' - Active text
' - Isi text = "IDPO" atau angka 6 digit
' - Tinggi kecil sekitar 0.28 - 0.65 cm
'
' PASS jika:
' - IDPO target ditemukan minimal 1x
' - Tidak ada placeholder "IDPO"
' - Tidak ada nomor IDPO lain / IDPO lama
'
' FAIL jika:
' - IDPO target tidak ditemukan
' - Masih ada placeholder "IDPO"
' - Masih ada IDPO lain selain target
'
' MAIN MACRO:
' IDPO_CHECK
'=========================================================


'=========================================================
' GLOBAL
'=========================================================

Private targetPO As String
Private targetSource As String

Private foundTarget As Long
Private foundPlaceholder As Long

Private dictOtherPO As Object

Private Const ID_MIN_H As Double = 0.28
Private Const ID_MAX_H As Double = 0.65

Private Const ORDER_FILE As String = "\Documents\Order.txt"


'=========================================================
' MAIN
'=========================================================

Sub IDPO_CHECK()

    Dim oldUnit As cdrUnit
    Dim sr As ShapeRange
    Dim s As Shape

    oldUnit = ActiveDocument.Unit

    On Error GoTo ERR_HANDLER

    On Error Resume Next
    Set sr = ActiveSelectionRange
    On Error GoTo ERR_HANDLER

    If sr Is Nothing Then

        MsgBox _
            "Pilih HASIL LAYOUT terlebih dahulu.", _
            vbExclamation, _
            "IDPO Check"

        Exit Sub

    End If

    If sr.Count = 0 Then

        MsgBox _
            "Pilih HASIL LAYOUT terlebih dahulu.", _
            vbExclamation, _
            "IDPO Check"

        Exit Sub

    End If

    targetPO = ""
    targetSource = ""

    '=====================================================
    ' V3:
    ' Ambil IDPO otomatis dari Order.txt.
    '=====================================================
    targetPO = LoadIDPOFromOrderTxt()

    If IsSixDigit(targetPO) Then

        targetSource = "AUTO dari Order.txt @IDPO"

    Else

        targetPO = AskIDPOManual()

        If targetPO = "" Then Exit Sub

        targetSource = "MANUAL POPUP"

    End If

    foundTarget = 0
    foundPlaceholder = 0

    Set dictOtherPO = CreateObject("Scripting.Dictionary")

    ActiveDocument.Unit = cdrCentimeter

    For Each s In sr
        ScanIDPO s
    Next s

    ActiveDocument.Unit = oldUnit

    ShowIDPOReport

    Exit Sub

ERR_HANDLER:

    On Error Resume Next
    ActiveDocument.Unit = oldUnit
    On Error GoTo 0

    MsgBox _
        "SYSTEM ERROR - IDPO CHECK" & vbCrLf & vbCrLf & _
        "No : " & Err.Number & vbCrLf & _
        Err.Description, _
        vbCritical, _
        "IDPO Check"

End Sub


'=========================================================
' LOAD IDPO FROM ORDER.TXT
'=========================================================

Private Function LoadIDPOFromOrderTxt() As String

    Dim path As String
    Dim allText As String
    Dim lines As Variant
    Dim i As Long
    Dim line As String
    Dim p As Long
    Dim key As String
    Dim val As String

    On Error GoTo FAIL

    LoadIDPOFromOrderTxt = ""

    path = Environ$("USERPROFILE") & ORDER_FILE

    If Dir(path) = "" Then Exit Function

    allText = ReadTextFileUTF8(path)

    allText = Replace(allText, vbCrLf, vbLf)
    allText = Replace(allText, vbCr, vbLf)

    lines = Split(allText, vbLf)

    For i = LBound(lines) To UBound(lines)

        line = NormalizeIDPOText(CStr(lines(i)))

        If Len(line) = 0 Then GoTo NextLine

        line = RemoveBOM(line)

        If Left$(line, 1) <> "@" Then GoTo NextLine

        p = InStr(1, line, "=", vbTextCompare)

        If p <= 2 Then GoTo NextLine

        key = UCase$(Trim$(Mid$(line, 2, p - 2)))
        val = Trim$(Mid$(line, p + 1))

        If key = "IDPO" Or key = "KODE_PRODUK" Or key = "KODE PRODUK" Then

            val = NormalizeIDPOText(val)

            If IsSixDigit(val) Then
                LoadIDPOFromOrderTxt = val
                Exit Function
            End If

        End If

NextLine:

    Next i

    Exit Function

FAIL:

    LoadIDPOFromOrderTxt = ""

End Function


Private Function AskIDPOManual() As String

    Dim v As String

    AskIDPOManual = ""

    v = InputBox( _
        "IDPO / Kode Produk tidak ditemukan otomatis dari Order.txt." & vbCrLf & vbCrLf & _
        "Paste IDPO target 6 digit:", _
        "IDPO Check")

    v = Trim$(v)

    If v = "" Then Exit Function

    If Not IsSixDigit(v) Then

        MsgBox _
            "IDPO / Kode Produk harus tepat 6 digit angka." & vbCrLf & vbCrLf & _
            "Contoh: 355863", _
            vbExclamation, _
            "IDPO Check"

        Exit Function

    End If

    AskIDPOManual = v

End Function


Private Function ReadTextFileUTF8(ByVal path As String) As String

    Dim stm As Object

    On Error GoTo FALLBACK

    Set stm = CreateObject("ADODB.Stream")

    stm.Type = 2
    stm.CharSet = "utf-8"
    stm.Open
    stm.LoadFromFile path

    ReadTextFileUTF8 = stm.ReadText

    stm.Close

    Exit Function

FALLBACK:

    On Error Resume Next

    If Not stm Is Nothing Then stm.Close

    On Error GoTo ANSI_FAIL

    ReadTextFileUTF8 = ReadTextFileANSI(path)
    Exit Function

ANSI_FAIL:

    ReadTextFileUTF8 = ""

End Function


Private Function ReadTextFileANSI(ByVal path As String) As String

    Dim f As Integer
    Dim line As String
    Dim result As String

    On Error GoTo FAIL

    f = FreeFile

    Open path For Input As #f

    Do Until EOF(f)
        Line Input #f, line
        result = result & line & vbLf
    Loop

    Close #f

    ReadTextFileANSI = result
    Exit Function

FAIL:

    On Error Resume Next
    Close #f
    ReadTextFileANSI = ""

End Function


Private Function RemoveBOM(ByVal s As String) As String

    On Error Resume Next

    If Len(s) > 0 Then
        If AscW(Left$(s, 1)) = &HFEFF Then
            RemoveBOM = Mid$(s, 2)
            Exit Function
        End If
    End If

    RemoveBOM = s

End Function


'=========================================================
' SCAN RECURSIVE
'=========================================================

Private Sub ScanIDPO(ByVal shp As Shape)

    Dim c As Shape

    On Error Resume Next

    If shp.Type = cdrGroupShape Then

        For Each c In shp.Shapes
            ScanIDPO c
        Next c

        ScanPowerClipIfAny shp

        Exit Sub

    End If

    ScanPowerClipIfAny shp

    If shp.Type <> cdrTextShape Then Exit Sub

    If Not IsIDPOHeight(shp) Then Exit Sub

    Dim raw As String
    Dim txt As String

    raw = ""

    Err.Clear
    raw = shp.Text.Story.Text
    Err.Clear

    txt = NormalizeIDPOText(raw)

    If txt = "" Then Exit Sub

    If UCase$(txt) = "IDPO" Then

        foundPlaceholder = foundPlaceholder + 1
        Exit Sub

    End If

    If IsSixDigit(txt) Then

        If txt = targetPO Then

            foundTarget = foundTarget + 1

        Else

            AddCount dictOtherPO, txt

        End If

    End If

End Sub


Private Sub ScanPowerClipIfAny(ByVal shp As Shape)

    On Error Resume Next

    Dim pc As Object
    Dim c As Shape

    Set pc = shp.PowerClip

    If pc Is Nothing Then Exit Sub

    For Each c In pc.Shapes
        ScanIDPO c
    Next c

    On Error GoTo 0

End Sub


'=========================================================
' REPORT
'=========================================================

Private Sub ShowIDPOReport()

    Dim report As String
    Dim k As Variant
    Dim isPass As Boolean

    isPass = False

    If foundTarget > 0 _
       And foundPlaceholder = 0 _
       And dictOtherPO.Count = 0 Then

        isPass = True

    End If

    If isPass Then

        report = _
            "IDPO CHECK PASSED" & vbCrLf & vbCrLf & _
            "Target IDPO : " & targetPO & vbCrLf & _
            "Source      : " & targetSource & vbCrLf & _
            "Ditemukan   : " & foundTarget & "x" & vbCrLf & vbCrLf & _
            "Semua IDPO sudah sesuai."

        MsgBox _
            report, _
            vbInformation, _
            "IDPO Check"

    Else

        report = _
            "IDPO CHECK FAILED" & vbCrLf & vbCrLf & _
            "Target IDPO      : " & targetPO & vbCrLf & _
            "Source           : " & targetSource & vbCrLf & _
            "Target ditemukan : " & foundTarget & "x" & vbCrLf

        If foundPlaceholder > 0 Then

            report = report & _
                "Placeholder IDPO : " & foundPlaceholder & "x" & vbCrLf

        End If

        If dictOtherPO.Count > 0 Then

            report = report & vbCrLf & _
                "IDPO lain terdeteksi:" & vbCrLf

            For Each k In dictOtherPO.Keys

                report = report & _
                    CStr(k) & " : " & dictOtherPO(k) & "x" & vbCrLf

            Next k

        End If

        report = report & vbCrLf

        If foundTarget = 0 Then

            report = report & _
                "- IDPO target belum ditemukan." & vbCrLf

        End If

        If foundPlaceholder > 0 Then

            report = report & _
                "- Masih ada placeholder IDPO." & vbCrLf

        End If

        If dictOtherPO.Count > 0 Then

            report = report & _
                "- Masih ada IDPO lama / IDPO lain." & vbCrLf

        End If

        If foundTarget = 0 _
           And foundPlaceholder = 0 _
           And dictOtherPO.Count = 0 Then

            report = report & _
                "- Tidak ada IDPO aktif yang cocok aturan tinggi " & _
                FormatNumber(ID_MIN_H, 2) & " - " & _
                FormatNumber(ID_MAX_H, 2) & " cm." & vbCrLf

        End If

        MsgBox _
            report, _
            vbCritical, _
            "IDPO Check"

    End If

End Sub


'=========================================================
' HELPERS
'=========================================================

Private Function NormalizeIDPOText(ByVal s As String) As String

    s = Replace(s, vbCr, "")
    s = Replace(s, vbLf, "")
    s = Replace(s, Chr$(160), " ")

    On Error Resume Next
    s = Replace(s, ChrW(&H200B), "")
    s = Replace(s, ChrW(&H200C), "")
    s = Replace(s, ChrW(&H200D), "")
    s = Replace(s, ChrW(&HFEFF), "")
    On Error GoTo 0

    Do While InStr(1, s, "  ", vbTextCompare) > 0
        s = Replace(s, "  ", " ")
    Loop

    NormalizeIDPOText = Trim$(s)

End Function


Private Function IsSixDigit(ByVal s As String) As Boolean

    Dim i As Long
    Dim ch As String

    s = Trim$(s)

    If Len(s) <> 6 Then Exit Function

    For i = 1 To 6

        ch = Mid$(s, i, 1)

        If ch < "0" Or ch > "9" Then
            Exit Function
        End If

    Next i

    IsSixDigit = True

End Function


Private Function IsIDPOHeight(ByVal shp As Shape) As Boolean

    Dim w As Double
    Dim h As Double
    Dim mn As Double

    w = shp.SizeWidth
    h = shp.SizeHeight

    If w <= 0 Or h <= 0 Then Exit Function

    If w < h Then
        mn = w
    Else
        mn = h
    End If

    If mn >= ID_MIN_H And mn <= ID_MAX_H Then
        IsIDPOHeight = True
    End If

End Function


Private Sub AddCount(ByVal dict As Object, ByVal key As String)

    If dict.Exists(key) Then
        dict(key) = CLng(dict(key)) + 1
    Else
        dict.Add key, 1
    End If

End Sub