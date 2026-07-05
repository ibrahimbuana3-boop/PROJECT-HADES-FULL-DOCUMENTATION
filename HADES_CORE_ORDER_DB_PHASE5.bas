Option Explicit

'=========================================================
' PROJECT HADES — CORE ORDER / DATABASE PHASE 5
'=========================================================

Public Function H5_OrderPath() As String
    H5_OrderPath = H5_DocumentsFile("Order.txt")
End Function

Public Function H5_TemplatePath() As String
    H5_TemplatePath = H5_DocumentsFile("TypoTemplate_Current.txt")
End Function

Public Function H5_LoadKeyValueFile(ByVal path As String) As Object
    Dim d As Object
    Dim txt As String
    Dim lines As Variant
    Dim i As Long
    Dim ln As String
    Dim p As Long
    Dim k As String
    Dim v As String

    Set d = CreateObject("Scripting.Dictionary")

    If Dir$(path) = "" Then
        Set H5_LoadKeyValueFile = d
        Exit Function
    End If

    txt = H5_ReadTextUTF8(path)
    txt = Replace(txt, vbCrLf, vbLf)
    txt = Replace(txt, vbCr, vbLf)
    lines = Split(txt, vbLf)

    For i = LBound(lines) To UBound(lines)
        ln = Trim$(CStr(lines(i)))
        If Len(ln) = 0 Then GoTo NEXT_LINE

        p = InStr(1, ln, "=", vbTextCompare)
        If p <= 0 Then GoTo NEXT_LINE

        k = UCase$(Trim$(Left$(ln, p - 1)))
        v = Trim$(Mid$(ln, p + 1))

        If Left$(k, 1) = "@" Then k = Mid$(k, 2)

        If Len(k) > 0 Then
            If d.Exists(k) Then
                d(k) = v
            Else
                d.Add k, v
            End If
        End If

NEXT_LINE:
    Next i

    Set H5_LoadKeyValueFile = d
End Function

Public Function H5_LoadOrderMeta() As Object
    Dim d As Object
    Dim txt As String
    Dim lines As Variant
    Dim i As Long
    Dim ln As String
    Dim p As Long
    Dim k As String
    Dim v As String

    Set d = CreateObject("Scripting.Dictionary")

    If Dir$(H5_OrderPath()) = "" Then
        Set H5_LoadOrderMeta = d
        Exit Function
    End If

    txt = H5_ReadTextUTF8(H5_OrderPath())
    txt = Replace(txt, vbCrLf, vbLf)
    txt = Replace(txt, vbCr, vbLf)
    lines = Split(txt, vbLf)

    For i = LBound(lines) To UBound(lines)
        ln = Trim$(CStr(lines(i)))
        If Len(ln) = 0 Then GoTo NEXT_LINE
        If Left$(ln, 1) <> "@" Then GoTo NEXT_LINE

        p = InStr(1, ln, "=", vbTextCompare)
        If p <= 2 Then GoTo NEXT_LINE

        k = UCase$(Trim$(Mid$(ln, 2, p - 2)))
        v = Trim$(Mid$(ln, p + 1))

        If Len(k) > 0 Then
            If d.Exists(k) Then
                d(k) = v
            Else
                d.Add k, v
            End If
        End If

NEXT_LINE:
    Next i

    Set H5_LoadOrderMeta = d
End Function

Public Function H5_LoadOrderRows() As Collection
    Dim c As New Collection
    Dim txt As String
    Dim lines As Variant
    Dim i As Long
    Dim ln As String
    Dim arr As Variant

    If Dir$(H5_OrderPath()) = "" Then
        Set H5_LoadOrderRows = c
        Exit Function
    End If

    txt = H5_ReadTextUTF8(H5_OrderPath())
    txt = Replace(txt, vbCrLf, vbLf)
    txt = Replace(txt, vbCr, vbLf)
    lines = Split(txt, vbLf)

    For i = LBound(lines) To UBound(lines)
        ln = Trim$(CStr(lines(i)))
        If Len(ln) = 0 Then GoTo NEXT_LINE
        If Left$(ln, 1) = "@" Then GoTo NEXT_LINE

        arr = Split(ln, "|")
        If UBound(arr) >= 3 Then c.Add arr

NEXT_LINE:
    Next i

    Set H5_LoadOrderRows = c
End Function

Public Function H5_LoadOrderExpectedCounts() As Object
    Dim d As Object
    Dim rows As Collection
    Dim arr As Variant
    Dim sz As String

    Set d = CreateObject("Scripting.Dictionary")
    Set rows = H5_LoadOrderRows()

    For Each arr In rows
        sz = H5_NormalizeSizeKey(CStr(arr(0)))
        If H5_IsStandardSize(sz) Then
            If Not d.Exists(sz) Then d.Add sz, 0
            d(sz) = CLng(d(sz)) + 1
        End If
    Next arr

    Set H5_LoadOrderExpectedCounts = d
End Function

Public Function H5_LoadTypoTemplate() As Object
    Dim path As String
    Dim pathAlt As String

    path = H5_TemplatePath()
    pathAlt = H5_DocumentsFile("TypoTemplate_currents.txt")

    If Dir$(path) = "" Then
        If Dir$(pathAlt) <> "" Then path = pathAlt
    End If

    Set H5_LoadTypoTemplate = H5_LoadKeyValueFile(path)
End Function

Public Function H5_DetectCurrentSizeDBFileName() As String
    Dim meta As Object
    Dim tpl As Object
    Dim db As String

    Set meta = H5_LoadOrderMeta()

    If meta.Exists("SIZEDB") Then db = Trim$(CStr(meta("SIZEDB")))
    If Len(db) = 0 Then
        If meta.Exists("DB") Then db = Trim$(CStr(meta("DB")))
    End If

    If Len(db) = 0 Then
        db = H5_InferDBFromOrderMeta(meta)
    End If

    If Len(db) = 0 Then
        Set tpl = H5_LoadTypoTemplate()
        If tpl.Exists("SIZEDB") Then db = Trim$(CStr(tpl("SIZEDB")))
        If Len(db) = 0 Then
            If tpl.Exists("DB") Then db = Trim$(CStr(tpl("DB")))
        End If
    End If

    If Len(db) > 0 Then
        If InStr(1, UCase$(db), ".TXT", vbTextCompare) = 0 Then db = db & ".txt"
    End If

    H5_DetectCurrentSizeDBFileName = db
End Function

Public Function H5_InferDBFromOrderMeta(ByVal meta As Object) As String
    Dim jenis As String
    Dim pola As String
    Dim model As String
    Dim allText As String

    If meta Is Nothing Then Exit Function

    If meta.Exists("JENIS_PESANAN") Then jenis = UCase$(CStr(meta("JENIS_PESANAN")))
    If meta.Exists("JENIS_POLA") Then pola = UCase$(CStr(meta("JENIS_POLA")))
    If meta.Exists("MODEL_JAHIT") Then model = UCase$(CStr(meta("MODEL_JAHIT")))

    allText = jenis & " " & pola & " " & model

    If InStr(1, allText, "CELANA", vbTextCompare) > 0 Then
        If InStr(1, allText, "ANAK", vbTextCompare) > 0 Then
            H5_InferDBFromOrderMeta = "SizeDB_CelanaAnak.txt"
        ElseIf InStr(1, allText, "WANITA", vbTextCompare) > 0 Or _
               InStr(1, allText, "PEREMPUAN", vbTextCompare) > 0 Or _
               InStr(1, allText, "CEWEK", vbTextCompare) > 0 Then
            H5_InferDBFromOrderMeta = "SizeDB_CelanaWanita.txt"
        Else
            H5_InferDBFromOrderMeta = "SizeDB_CelanaPria.txt"
        End If
        Exit Function
    End If

    If InStr(1, allText, "JAKET", vbTextCompare) > 0 Then
        If InStr(1, allText, "ANAK", vbTextCompare) > 0 Then
            H5_InferDBFromOrderMeta = "SizeDB_JaketAnak.txt"
        Else
            H5_InferDBFromOrderMeta = "SizeDB_Jaket.txt"
        End If
        Exit Function
    End If

    If InStr(1, allText, "JERSEY", vbTextCompare) > 0 Then
        If InStr(1, allText, "ANAK", vbTextCompare) > 0 Then
            H5_InferDBFromOrderMeta = "SizeDB_Anak.txt"
            Exit Function
        End If

        If InStr(1, allText, "SLIM", vbTextCompare) > 0 Then
            If InStr(1, allText, "WANITA", vbTextCompare) > 0 Or _
               InStr(1, allText, "PEREMPUAN", vbTextCompare) > 0 Or _
               InStr(1, allText, "CEWEK", vbTextCompare) > 0 Then
                H5_InferDBFromOrderMeta = "SizeDB_WanitaSlimFit.txt"
            Else
                H5_InferDBFromOrderMeta = "SizeDB_PriaSlimFit.txt"
            End If
            Exit Function
        End If

        If InStr(1, allText, "WANITA", vbTextCompare) > 0 Or _
           InStr(1, allText, "PEREMPUAN", vbTextCompare) > 0 Or _
           InStr(1, allText, "CEWEK", vbTextCompare) > 0 Then
            H5_InferDBFromOrderMeta = "SizeDB_Wanita.txt"
        Else
            H5_InferDBFromOrderMeta = "SizeDB_Pria.txt"
        End If
        Exit Function
    End If
End Function

Public Sub H5_ProductModeFromDB(ByVal dbName As String, ByRef isPants As Boolean, ByRef isSplitFront As Boolean)
    dbName = UCase$(Trim$(dbName))

    isPants = False
    isSplitFront = False

    If InStr(1, dbName, "CELANA", vbTextCompare) > 0 Then
        isPants = True
        isSplitFront = False
        Exit Sub
    End If

    If InStr(1, dbName, "JAKET", vbTextCompare) > 0 Then
        isPants = False
        isSplitFront = True
        Exit Sub
    End If
End Sub

Public Function H5_LoadSizeDB(ByVal dbFileName As String, ByVal isPants As Boolean, ByVal isSplitFront As Boolean) As Object
    Dim d As Object
    Dim path As String
    Dim txt As String
    Dim lines As Variant
    Dim i As Long
    Dim ln As String
    Dim arr As Variant
    Dim sz As String

    Set d = CreateObject("Scripting.Dictionary")

    If Len(Trim$(dbFileName)) = 0 Then
        Set H5_LoadSizeDB = d
        Exit Function
    End If

    path = H5_DocumentsFile(dbFileName)
    If Dir$(path) = "" Then
        Set H5_LoadSizeDB = d
        Exit Function
    End If

    txt = H5_ReadTextUTF8(path)
    txt = Replace(txt, vbCrLf, vbLf)
    txt = Replace(txt, vbCr, vbLf)
    lines = Split(txt, vbLf)

    For i = LBound(lines) To UBound(lines)
        ln = Trim$(CStr(lines(i)))
        If Len(ln) = 0 Then GoTo NEXT_LINE
        If Left$(ln, 1) = "@" Then GoTo NEXT_LINE

        arr = Split(ln, "|")
        If UBound(arr) < 0 Then GoTo NEXT_LINE

        sz = H5_NormalizeSizeKey(CStr(arr(0)))
        If Not H5_IsStandardSize(sz) Then GoTo NEXT_LINE

        If isPants Then
            If UBound(arr) < 2 Then GoTo NEXT_LINE
        ElseIf isSplitFront Then
            If UBound(arr) < 4 Then GoTo NEXT_LINE
        Else
            If UBound(arr) < 3 Then GoTo NEXT_LINE
        End If

        If d.Exists(sz) Then
            d(sz) = arr
        Else
            d.Add sz, arr
        End If

NEXT_LINE:
    Next i

    Set H5_LoadSizeDB = d
End Function
