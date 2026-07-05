Option Explicit

'=========================================================
' PROJECT HADES — CORE TEXT / NORMALIZE PHASE 5
'=========================================================

Public Function H5_NormalizeText(ByVal s As String) As String
    s = Replace(s, vbCr, "")
    s = Replace(s, vbLf, "")
    s = Replace(s, vbTab, " ")
    s = Replace(s, Chr$(160), " ")

    Do While InStr(1, s, "  ", vbTextCompare) > 0
        s = Replace(s, "  ", " ")
    Loop

    H5_NormalizeText = UCase$(Trim$(s))
End Function

Public Function H5_NormalizeSizeKey(ByVal sz As String) As String
    Dim s As String

    s = UCase$(Trim$(sz))

    Select Case s
        Case "XXL"
            H5_NormalizeSizeKey = "2XL"
        Case "XXXL"
            H5_NormalizeSizeKey = "3XL"
        Case "XXXXL"
            H5_NormalizeSizeKey = "4XL"
        Case "XXXXXL"
            H5_NormalizeSizeKey = "5XL"
        Case "XXXXXXL"
            H5_NormalizeSizeKey = "6XL"
        Case Else
            H5_NormalizeSizeKey = s
    End Select
End Function

Public Function H5_IsStandardSize(ByVal sz As String) As Boolean
    Select Case H5_NormalizeSizeKey(sz)
        Case "XXS", "XS", "S", "M", "L", "XL", "2XL", "3XL", "4XL", "5XL", "6XL"
            H5_IsStandardSize = True
        Case Else
            H5_IsStandardSize = False
    End Select
End Function

Public Function H5_ToDbl(ByVal v As Variant) As Double
    Dim s As String

    s = Trim$(CStr(v))
    s = Replace(s, ",", ".")

    H5_ToDbl = CDbl(Val(s))
End Function

Public Function H5_MinD(ByVal a As Double, ByVal b As Double) As Double
    If a < b Then
        H5_MinD = a
    Else
        H5_MinD = b
    End If
End Function

Public Function H5_MaxD(ByVal a As Double, ByVal b As Double) As Double
    If a > b Then
        H5_MaxD = a
    Else
        H5_MaxD = b
    End If
End Function

Public Function H5_IsNumericText(ByVal s As String, Optional ByVal maxLen As Long = 3) As Boolean
    s = Trim$(s)
    If Len(s) = 0 Then Exit Function
    If maxLen > 0 Then
        If Len(s) > maxLen Then Exit Function
    End If
    If IsNumeric(s) Then H5_IsNumericText = True
End Function
