Option Explicit

'=========================================================
' PROJECT HADES — CORE IO / PATHS PHASE 5
' CorelDRAW 2021 VBA
'
' TUJUAN:
' - Satu pintu untuk path Documents, report folder, read/write UTF-8,
'   timestamp, dan pembacaan key=value report.
'
' CATATAN:
' - Semua nama memakai prefix H5_ agar tidak tabrakan dengan VBA lama.
' - Module ini pondasi refactor; belum memaksa macro lama berubah total.
'=========================================================

Private Const H5_REPORT_FOLDER_NAME As String = "HADES_REPORTS"

Public Function H5_DocumentsPath() As String
    H5_DocumentsPath = Environ$("USERPROFILE") & "\Documents"
End Function

Public Function H5_DocumentsFile(ByVal fileName As String) As String
    H5_DocumentsFile = H5_DocumentsPath() & "\" & fileName
End Function

Public Function H5_ReportFolderPath() As String
    Dim p As String
    p = H5_DocumentsPath() & "\" & H5_REPORT_FOLDER_NAME
    H5_EnsureFolder p
    H5_ReportFolderPath = p
End Function

Public Sub H5_EnsureFolder(ByVal folderPath As String)
    On Error Resume Next
    If Len(Trim$(folderPath)) = 0 Then Exit Sub
    If Dir$(folderPath, vbDirectory) = "" Then MkDir folderPath
    On Error GoTo 0
End Sub

Public Function H5_FileExists(ByVal path As String) As Boolean
    H5_FileExists = (Len(Dir$(path)) > 0)
End Function

Public Function H5_FolderExists(ByVal path As String) As Boolean
    H5_FolderExists = (Len(Dir$(path, vbDirectory)) > 0)
End Function

Public Function H5_NowStamp() As String
    H5_NowStamp = Format$(Now, "yyyymmdd_hhnnss")
End Function

Public Function H5_NowHuman() As String
    H5_NowHuman = Format$(Now, "yyyy-mm-dd hh:nn:ss")
End Function

Public Function H5_ReadTextUTF8(ByVal path As String) As String
    Dim stm As Object
    Dim f As Integer
    Dim line As String
    Dim result As String

    On Error GoTo FALLBACK

    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 2
    stm.Charset = "utf-8"
    stm.Open
    stm.LoadFromFile path
    H5_ReadTextUTF8 = stm.ReadText
    stm.Close
    Exit Function

FALLBACK:
    On Error Resume Next
    If Not stm Is Nothing Then stm.Close
    On Error GoTo 0

    If Dir$(path) = "" Then
        H5_ReadTextUTF8 = ""
        Exit Function
    End If

    f = FreeFile
    Open path For Input As #f
    Do Until EOF(f)
        Line Input #f, line
        result = result & line & vbLf
    Loop
    Close #f

    H5_ReadTextUTF8 = result
End Function

Public Sub H5_WriteTextUTF8(ByVal path As String, ByVal textData As String)
    Dim stm As Object
    Dim f As Integer

    On Error GoTo FALLBACK

    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 2
    stm.Charset = "utf-8"
    stm.Open
    stm.WriteText textData
    stm.SaveToFile path, 2
    stm.Close
    Exit Sub

FALLBACK:
    On Error Resume Next
    If Not stm Is Nothing Then stm.Close
    On Error GoTo 0

    f = FreeFile
    Open path For Output As #f
    Print #f, textData
    Close #f
End Sub

Public Function H5_ReadMachineValueFromFile(ByVal path As String, ByVal keyName As String) As String
    Dim txt As String

    If Len(Dir$(path)) = 0 Then
        H5_ReadMachineValueFromFile = ""
        Exit Function
    End If

    txt = H5_ReadTextUTF8(path)
    H5_ReadMachineValueFromFile = H5_ReadMachineValue(txt, keyName)
End Function

Public Function H5_ReadMachineValue(ByVal textData As String, ByVal keyName As String) As String
    Dim lines As Variant
    Dim i As Long
    Dim ln As String
    Dim p As Long
    Dim k As String
    Dim v As String

    keyName = UCase$(Trim$(keyName))
    textData = Replace(textData, vbCrLf, vbLf)
    textData = Replace(textData, vbCr, vbLf)
    lines = Split(textData, vbLf)

    For i = LBound(lines) To UBound(lines)
        ln = Trim$(CStr(lines(i)))
        If Len(ln) = 0 Then GoTo NEXT_LINE

        p = InStr(1, ln, "=", vbTextCompare)
        If p <= 0 Then GoTo NEXT_LINE

        k = UCase$(Trim$(Left$(ln, p - 1)))
        v = Trim$(Mid$(ln, p + 1))

        If k = keyName Then
            H5_ReadMachineValue = v
            Exit Function
        End If

NEXT_LINE:
    Next i
End Function

Public Sub H5_OpenFolder(ByVal folderPath As String)
    On Error Resume Next
    If Len(Trim$(folderPath)) = 0 Then Exit Sub
    Shell "explorer.exe " & Chr$(34) & folderPath & Chr$(34), vbNormalFocus
    On Error GoTo 0
End Sub

Public Sub H5_OpenFile(ByVal filePath As String)
    On Error Resume Next
    If Len(Dir$(filePath)) = 0 Then Exit Sub
    Shell "explorer.exe " & Chr$(34) & filePath & Chr$(34), vbNormalFocus
    On Error GoTo 0
End Sub
