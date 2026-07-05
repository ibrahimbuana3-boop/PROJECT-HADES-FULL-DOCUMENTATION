Option Explicit

'==========================================================
' PROJECT HADES
' HADES AUTO MASS NESTING V3.5 - ROW MAJOR 6x2 + AUTO RENAME
'
' SHORTCUT UTAMA:
' HADES_AUTO_MASS_NESTING
'
' ALIAS:
' HADES_AUTO_MASS_NESTING_V2
' HADES_AUTO_MASS_NESTING_V3
' HADES_AUTO_MASS_NESTING_V3_AUTO_RENAME
'
' BASIS:
' HADES AUTO MASS NESTING V2 SAFE BUCKET yang sudah working.
'
' FUNGSI:
' - Membaca Documents\Order.txt sebagai record lengkap:
'   SIZE|NAMA|NOMOR|NICKNAME
' - Membaca Documents\HADES_PATTERN_CATALOG_CURRENT.txt
' - Mencari source panel dari selection loose master
' - Duplicate panel sesuai setiap baris Order.txt
' - Auto rename placeholder active text pada duplicate sebelum nesting
' - Menyusun hasil duplicate ke area 178 x 255 cm
' - Output diletakkan di atas ActivePage, bukan di kanan selection
' - Gap antar panel 1 cm
'
' FITUR V3:
' - Auto Rename internal, bukan memanggil QC_AUTO_RENAME.
' - Rename dilakukan sebelum panel disebar ke area nesting.
' - 1 baris Order.txt = 1 identitas atlet.
' - Support remove ligature seperti Auto Rename terbaru.
' - Support Japanese/CJK font fallback.
' - Skip IDPO kecil / angka PO 6 digit kecil / marker @A: dan @ATTR:.
' - Output anchor mengikuti konsep Auto Duplicate: di atas ActivePage.
' - Tetap tidak mengubah rotasi.
' - Tetap tidak menembus group panel untuk pergerakan; recursive scan hanya untuk mengganti active text.
'
' CATATAN:
' - V3 tetap kompatibel dengan Pattern Catalog V1.
' - Lebih optimal memakai Pattern Catalog V2 karena sudah ada field BUCKET.
'==========================================================

Private Const HAMN_CATALOG_FILE As String = "HADES_PATTERN_CATALOG_CURRENT.txt"
Private Const HAMN_ORDER_FILE As String = "Order.txt"
Private Const HAMN_REPORT_FILE As String = "HADES_MASS_NESTING_REPORT_LATEST.txt"

Private Const HAMN_AREA_W As Double = 178#
Private Const HAMN_AREA_H As Double = 255#
Private Const HAMN_GAP As Double = 1#

Private Const HAMN_AREA_SPACING As Double = 10#
Private Const HAMN_GRID_COLS As Long = 6
Private Const HAMN_GRID_ROWS As Long = 2
Private Const HAMN_OUTPUT_OFFSET_RIGHT As Double = 10# 'Legacy; tidak dipakai di V3.1 page anchor.
Private Const HAMN_PAGE_TOP_OFFSET As Double = 210#
Private Const HAMN_PAGE_CENTER_X_CORRECTION As Double = 0#
Private Const HAMN_PAGE_TOP_Y_CORRECTION As Double = 0#

Private Const HAMN_MATCH_TOL As Double = 0.25

Private Const HAMN_BUCKET_BODY As String = "BODY"
Private Const HAMN_BUCKET_SLEEVE As String = "SLEEVE"
Private Const HAMN_BUCKET_SMALL As String = "SMALL"

Private Const HAMN_MIN_TEXT_H As Double = 1#
Private Const HAMN_ID_MIN_H As Double = 0.28
Private Const HAMN_ID_MAX_H As Double = 0.65

Private Const HAMN_CJK_FONT_1 As String = "Meiryo"
Private Const HAMN_CJK_FONT_2 As String = "Yu Gothic"
Private Const HAMN_CJK_FONT_3 As String = "Noto Sans CJK JP"
Private Const HAMN_CJK_FONT_4 As String = "Noto Sans JP"
Private Const HAMN_CJK_FONT_5 As String = "MS Gothic"

Private HAMN_DeleteQueue As Collection
Private HAMN_DeletedTextCount As Long
Private HAMN_CJKDetectedCount As Long
Private HAMN_CJKFontAppliedCount As Long
Private HAMN_CJKFontFailedCount As Long


Public Sub HADES_AUTO_MASS_NESTING()
    HADES_AUTO_MASS_NESTING_V3_AUTO_RENAME
End Sub

Public Sub HADES_AUTO_MASS_NESTING_V2()
    HADES_AUTO_MASS_NESTING_V3_AUTO_RENAME
End Sub

Public Sub HADES_AUTO_MASS_NESTING_V3()
    HADES_AUTO_MASS_NESTING_V3_AUTO_RENAME
End Sub


Public Sub HADES_AUTO_MASS_NESTING_V3_AUTO_RENAME()

    On Error GoTo ErrHandler

    Dim oldUnit As Long
    Dim sr As ShapeRange

    Dim catalogPath As String
    Dim orderPath As String
    Dim reportPath As String

    Dim catN As Long
    Dim catSize() As String
    Dim catPanelNo() As Long
    Dim catW() As Double
    Dim catH() As Double
    Dim catType() As String
    Dim catName() As String
    Dim catBucket() As String

    Dim orderN As Long
    Dim orderSize() As String
    Dim orderName() As String
    Dim orderNo() As String
    Dim orderNick() As String
    Dim orderQty As Object
    Dim orderSeqMap As Object

    Dim recNameR() As Long
    Dim recNumR() As Long
    Dim recNickR() As Long

    Dim catSizeMap As Object
    Dim catPanelCountBySize As Object

    Dim srcN As Long
    Dim srcShapes() As Shape
    Dim srcW() As Double
    Dim srcH() As Double
    Dim srcUsed() As Boolean

    Dim outN As Long
    Dim outShapes() As Shape
    Dim outW() As Double
    Dim outH() As Double
    Dim outArea() As Double
    Dim outLabel() As String
    Dim outBucket() As String

    Dim i As Long
    Dim q As Long
    Dim qty As Long
    Dim srcIdx As Long
    Dim recIdx As Long

    Dim dup As Shape
    Dim warn As String
    Dim hardError As String
    Dim logText As String
    Dim renameLog As String

    Dim baseLeft As Double
    Dim baseTop As Double
    Dim pageCenterX As Double
    Dim pageTopY As Double
    Dim anchorMethod As String

    Dim usedAreaCount As Long
    Dim expectedPanels As Long

    Dim nameR As Long
    Dim numR As Long
    Dim nickR As Long

    Dim totalNameR As Long
    Dim totalNumR As Long
    Dim totalNickR As Long

    Dim cmdStarted As Boolean

    Set HAMN_DeleteQueue = New Collection
    HAMN_DeletedTextCount = 0
    HAMN_CJKDetectedCount = 0
    HAMN_CJKFontAppliedCount = 0
    HAMN_CJKFontFailedCount = 0

    If ActiveDocument Is Nothing Then
        MsgBox "Tidak ada dokumen aktif.", vbExclamation, "Hades Auto Mass Nesting V3.5"
        Exit Sub
    End If

    If ActiveSelection.Shapes.Count = 0 Then
        MsgBox "Select semua panel master yang sudah loose dulu." & vbCrLf & vbCrLf & _
               "Urutan workflow:" & vbCrLf & _
               "1. Record catalog saat masih group per size." & vbCrLf & _
               "2. Ungroup sekali." & vbCrLf & _
               "3. Select semua panel master loose." & vbCrLf & _
               "4. Jalankan Auto Mass Nesting V3.", _
               vbExclamation, "Hades Auto Mass Nesting V3.5"
        Exit Sub
    End If

    oldUnit = ActiveDocument.Unit
    ActiveDocument.Unit = cdrCentimeter

    catalogPath = HAMN_GetDocumentsPath() & "\" & HAMN_CATALOG_FILE
    orderPath = HAMN_GetDocumentsPath() & "\" & HAMN_ORDER_FILE
    reportPath = HAMN_GetDocumentsPath() & "\" & HAMN_REPORT_FILE

    If Dir$(catalogPath) = "" Then
        MsgBox "Pattern Catalog belum ditemukan." & vbCrLf & vbCrLf & _
               "File yang dicari:" & vbCrLf & _
               catalogPath & vbCrLf & vbCrLf & _
               "Jalankan HADES_RECORD_PATTERN_CATALOG dulu.", _
               vbCritical, "Hades Auto Mass Nesting V3.5"
        ActiveDocument.Unit = oldUnit
        Exit Sub
    End If

    If Dir$(orderPath) = "" Then
        MsgBox "Order.txt belum ditemukan." & vbCrLf & vbCrLf & _
               "File yang dicari:" & vbCrLf & _
               orderPath, _
               vbCritical, "Hades Auto Mass Nesting V3.5"
        ActiveDocument.Unit = oldUnit
        Exit Sub
    End If

    If Not HAMN_LoadCatalog(catalogPath, catN, catSize, catPanelNo, catW, catH, catType, catName, catBucket) Then
        MsgBox "Gagal membaca Pattern Catalog.", vbCritical, "Hades Auto Mass Nesting V3.5"
        ActiveDocument.Unit = oldUnit
        Exit Sub
    End If

    If catN = 0 Then
        MsgBox "Pattern Catalog kosong / tidak ada panel.", vbCritical, "Hades Auto Mass Nesting V3.5"
        ActiveDocument.Unit = oldUnit
        Exit Sub
    End If

    HAMN_ClassifyCatalogBuckets catN, catSize, catW, catH, catName, catBucket

    hardError = HAMN_PreflightCatalogSize(catN, catSize, catPanelNo, catW, catH)
    If hardError <> "" Then
        MsgBox "AUTO MASS NESTING DIBATALKAN." & vbCrLf & vbCrLf & _
               "Ada panel yang lebih besar dari area 178 x 255 cm:" & vbCrLf & vbCrLf & _
               hardError, _
               vbCritical, "Hades Auto Mass Nesting V3.5"
        ActiveDocument.Unit = oldUnit
        Exit Sub
    End If

    Set orderQty = CreateObject("Scripting.Dictionary")
    orderQty.CompareMode = 1

    Set orderSeqMap = CreateObject("Scripting.Dictionary")
    orderSeqMap.CompareMode = 1

    If Not HAMN_LoadOrderRecords(orderPath, orderN, orderSize, orderName, orderNo, orderNick, orderQty, orderSeqMap) Then
        MsgBox "Gagal membaca Order.txt.", vbCritical, "Hades Auto Mass Nesting V3.5"
        ActiveDocument.Unit = oldUnit
        Exit Sub
    End If

    If orderN = 0 Or orderQty.Count = 0 Then
        MsgBox "Order.txt tidak berisi record order valid.", vbCritical, "Hades Auto Mass Nesting V3.5"
        ActiveDocument.Unit = oldUnit
        Exit Sub
    End If

    ReDim recNameR(1 To orderN)
    ReDim recNumR(1 To orderN)
    ReDim recNickR(1 To orderN)

    Set catSizeMap = CreateObject("Scripting.Dictionary")
    catSizeMap.CompareMode = 1

    Set catPanelCountBySize = CreateObject("Scripting.Dictionary")
    catPanelCountBySize.CompareMode = 1

    For i = 1 To catN
        If Not catSizeMap.Exists(catSize(i)) Then
            catSizeMap.Add catSize(i), True
        End If

        If catPanelCountBySize.Exists(catSize(i)) Then
            catPanelCountBySize(catSize(i)) = CLng(catPanelCountBySize(catSize(i))) + 1
        Else
            catPanelCountBySize.Add catSize(i), 1
        End If
    Next i

    warn = ""
    expectedPanels = 0

    Dim key As Variant
    For Each key In orderQty.keys
        If Not catSizeMap.Exists(CStr(key)) Then
            warn = warn & "- FAIL: Size " & CStr(key) & " ada di Order.txt tetapi tidak ada di Pattern Catalog." & vbCrLf
        Else
            expectedPanels = expectedPanels + (CLng(orderQty(key)) * CLng(catPanelCountBySize(CStr(key))))
        End If
    Next key

    Set sr = ActiveSelectionRange

    srcN = sr.Shapes.Count

    ReDim srcShapes(1 To srcN)
    ReDim srcW(1 To srcN)
    ReDim srcH(1 To srcN)
    ReDim srcUsed(1 To srcN)

    For i = 1 To srcN
        Set srcShapes(i) = sr.Shapes(i)
        srcW(i) = srcShapes(i).SizeWidth
        srcH(i) = srcShapes(i).SizeHeight
        srcUsed(i) = False
    Next i

    HAMN_GetAccuratePageAnchor pageCenterX, pageTopY, anchorMethod

    'Output diletakkan di atas ActivePage seperti Auto Duplicate.
    'Titik bawah area nesting pertama = pageTopY + HAMN_PAGE_TOP_OFFSET.
    'Karena layout shelf berjalan dari atas ke bawah, baseTop = bottom anchor + tinggi area.
    baseLeft = pageCenterX - (HAMN_AREA_W / 2#) + HAMN_PAGE_CENTER_X_CORRECTION
    baseTop = pageTopY + HAMN_PAGE_TOP_OFFSET + HAMN_AREA_H + HAMN_PAGE_TOP_Y_CORRECTION

    ActiveDocument.BeginCommandGroup "HADES AUTO MASS NESTING V3.5 ROW MAJOR 6x2"
    cmdStarted = True

    outN = 0

    '======================================================
    ' Generate duplicate berdasarkan Catalog x setiap record Order.txt
    ' Rename dilakukan pada duplicate sebelum masuk queue nesting.
    '======================================================
    For i = 1 To catN

        qty = 0

        If orderQty.Exists(catSize(i)) Then
            qty = CLng(orderQty(catSize(i)))
        End If

        If qty <= 0 Then
            GoTo NextCatalogItem
        End If

        srcIdx = HAMN_FindMatchingSource(catW(i), catH(i), srcN, srcW, srcH, srcUsed)

        If srcIdx = 0 Then
            warn = warn & "- FAIL: Source panel tidak ditemukan untuk " & _
                          catSize(i) & " panel " & CStr(catPanelNo(i)) & _
                          " (" & HAMN_DblToStr(catW(i)) & " x " & HAMN_DblToStr(catH(i)) & " cm)" & vbCrLf
            GoTo NextCatalogItem
        End If

        srcUsed(srcIdx) = True

        For q = 1 To qty

            recIdx = 0
            If orderSeqMap.Exists(catSize(i) & "|" & CStr(q)) Then
                recIdx = CLng(orderSeqMap(catSize(i) & "|" & CStr(q)))
            End If

            Set dup = srcShapes(srcIdx).Duplicate

            If recIdx > 0 Then
                nameR = 0
                numR = 0
                nickR = 0

                HAMN_RenameTextRecursive dup, orderName(recIdx), orderNo(recIdx), orderNick(recIdx), nameR, numR, nickR

                recNameR(recIdx) = recNameR(recIdx) + nameR
                recNumR(recIdx) = recNumR(recIdx) + numR
                recNickR(recIdx) = recNickR(recIdx) + nickR

                totalNameR = totalNameR + nameR
                totalNumR = totalNumR + numR
                totalNickR = totalNickR + nickR
            Else
                warn = warn & "- WARNING: Record order tidak ditemukan untuk " & catSize(i) & " urutan " & CStr(q) & vbCrLf
            End If

            HAMN_AddOutputShape _
                outN, outShapes, outW, outH, outArea, outLabel, outBucket, _
                dup, _
                dup.SizeWidth, _
                dup.SizeHeight, _
                catBucket(i), _
                catSize(i) & "_" & catBucket(i) & "_P" & CStr(catPanelNo(i)) & "_COPY" & CStr(q)

        Next q

NextCatalogItem:
    Next i

    HAMN_DeleteQueuedTextShapes

    renameLog = HAMN_BuildRenameCheck(orderN, orderSize, orderName, orderNo, orderNick, recNameR, recNumR, recNickR)
    If renameLog <> "" Then warn = warn & renameLog

    If outN = 0 Then
        If cmdStarted Then ActiveDocument.EndCommandGroup
        ActiveDocument.Unit = oldUnit
        MsgBox "Tidak ada panel yang berhasil dibuat." & vbCrLf & vbCrLf & _
               warn, vbCritical, "Hades Auto Mass Nesting V3.5"
        Exit Sub
    End If

    '======================================================
    ' Sort output: BODY dulu, SLEEVE, SMALL, lalu area terbesar
    '======================================================
    HAMN_SortOutputByBucketArea outN, outShapes, outW, outH, outArea, outLabel, outBucket

    '======================================================
    ' Nesting bounding box shelf algorithm
    '======================================================
    usedAreaCount = HAMN_LayoutShelf(outN, outShapes, outW, outH, outLabel, outBucket, baseLeft, baseTop, logText, warn)

    logText = "OUTPUT_ANCHOR|ABOVE_ACTIVE_PAGE|METHOD=" & anchorMethod & _
              "|PAGE_CENTER_X=" & HAMN_DblToStr(pageCenterX) & _
              "|PAGE_TOP_Y=" & HAMN_DblToStr(pageTopY) & _
              "|BASE_LEFT=" & HAMN_DblToStr(baseLeft) & _
              "|BASE_TOP=" & HAMN_DblToStr(baseTop) & vbCrLf & logText

    If outN <> expectedPanels Then
        warn = warn & "- FAIL: Total panel placed tidak sama dengan expected. Expected=" & _
                      CStr(expectedPanels) & ", Placed=" & CStr(outN) & vbCrLf
    End If

    If cmdStarted Then ActiveDocument.EndCommandGroup

    ActiveDocument.Unit = oldUnit

    HAMN_WriteReport reportPath, outN, expectedPanels, usedAreaCount, orderN, orderSize, orderName, orderNo, orderNick, _
                     recNameR, recNumR, recNickR, orderQty, catPanelCountBySize, _
                     totalNameR, totalNumR, totalNickR, warn, logText

    MsgBox "HADES AUTO MASS NESTING V3 selesai." & vbCrLf & vbCrLf & _
           "Order records       : " & orderN & vbCrLf & _
           "Expected panel      : " & expectedPanels & vbCrLf & _
           "Panel hasil duplicate: " & outN & vbCrLf & _
           "Blok terpakai       : " & usedAreaCount & vbCrLf & _
           "Area size           : 178 x 255 cm" & vbCrLf & _
           "Output              : di atas ActivePage" & vbCrLf & _
           "Jarak dari page     : 210 cm" & vbCrLf & _
           "Gap                 : 1 cm" & vbCrLf & _
           "Mode                : SAFE BUCKET + AUTO RENAME + ROW-MAJOR 6x2" & vbCrLf & _
           "Rename nama/nomor/nick: " & totalNameR & "/" & totalNumR & "/" & totalNickR & vbCrLf & vbCrLf & _
           "Report:" & vbCrLf & _
           reportPath & vbCrLf & vbCrLf & _
           IIf(warn <> "", "STATUS: WARNING / FAIL. Cek report.", "STATUS: PASS."), _
           vbInformation, "Hades Auto Mass Nesting V3.5"

    Exit Sub

ErrHandler:
    On Error Resume Next

    If cmdStarted Then ActiveDocument.EndCommandGroup
    ActiveDocument.Unit = oldUnit

    MsgBox "ERROR HADES_AUTO_MASS_NESTING_V3:" & vbCrLf & _
           err.Description, vbCritical, "Hades Auto Mass Nesting V3.5"

End Sub


'==========================================================
' LOAD CATALOG
'==========================================================

Private Function HAMN_LoadCatalog( _
    ByVal path As String, _
    ByRef catN As Long, _
    ByRef catSize() As String, _
    ByRef catPanelNo() As Long, _
    ByRef catW() As Double, _
    ByRef catH() As Double, _
    ByRef catType() As String, _
    ByRef catName() As String, _
    ByRef catBucket() As String _
) As Boolean

    On Error GoTo ErrHandler

    Dim f As Integer
    Dim ln As String
    Dim p() As String

    catN = 0

    f = FreeFile
    Open path For Input As #f

    Do While Not EOF(f)

        Line Input #f, ln
        ln = Trim$(ln)

        If ln <> "" Then

            If UCase$(Left$(ln, 6)) = "PANEL|" Then

                p = Split(ln, "|")

                If UBound(p) >= 4 Then

                    catN = catN + 1

                    If catN = 1 Then
                        ReDim catSize(1 To catN)
                        ReDim catPanelNo(1 To catN)
                        ReDim catW(1 To catN)
                        ReDim catH(1 To catN)
                        ReDim catType(1 To catN)
                        ReDim catName(1 To catN)
                        ReDim catBucket(1 To catN)
                    Else
                        ReDim Preserve catSize(1 To catN)
                        ReDim Preserve catPanelNo(1 To catN)
                        ReDim Preserve catW(1 To catN)
                        ReDim Preserve catH(1 To catN)
                        ReDim Preserve catType(1 To catN)
                        ReDim Preserve catName(1 To catN)
                        ReDim Preserve catBucket(1 To catN)
                    End If

                    catSize(catN) = HAMN_NormalizeSize(p(1))
                    catPanelNo(catN) = CLng(val(p(2)))
                    catW(catN) = HAMN_Val(p(3))
                    catH(catN) = HAMN_Val(p(4))

                    If UBound(p) >= 5 Then
                        catType(catN) = CStr(p(5))
                    Else
                        catType(catN) = ""
                    End If

                    If UBound(p) >= 6 Then
                        catName(catN) = CStr(p(6))
                    Else
                        catName(catN) = ""
                    End If

                    If UBound(p) >= 7 Then
                        catBucket(catN) = HAMN_NormalizeBucket(CStr(p(7)))
                    Else
                        catBucket(catN) = ""
                    End If

                End If

            End If

        End If

    Loop

    Close #f

    HAMN_LoadCatalog = True
    Exit Function

ErrHandler:
    On Error Resume Next
    Close #f
    HAMN_LoadCatalog = False

End Function


'==========================================================
' LOAD ORDER RECORDS
'==========================================================

Private Function HAMN_LoadOrderRecords( _
    ByVal path As String, _
    ByRef orderN As Long, _
    ByRef orderSize() As String, _
    ByRef orderName() As String, _
    ByRef orderNo() As String, _
    ByRef orderNick() As String, _
    ByRef orderQty As Object, _
    ByRef orderSeqMap As Object _
) As Boolean

    On Error GoTo ErrHandler

    Dim f As Integer
    Dim ln As String
    Dim p() As String
    Dim sz As String
    Dim nm As String
    Dim no As String
    Dim nick As String
    Dim seqDict As Object
    Dim seq As Long

    Set seqDict = CreateObject("Scripting.Dictionary")
    seqDict.CompareMode = 1

    orderN = 0

    f = FreeFile
    Open path For Input As #f

    Do While Not EOF(f)

        Line Input #f, ln
        ln = Trim$(ln)

        If ln <> "" Then

            If Left$(ln, 1) <> "@" Then

                If InStr(1, ln, "|") > 0 Then

                    p = Split(ln, "|")

                    sz = ""
                    nm = ""
                    no = ""
                    nick = ""

                    If UBound(p) >= 0 Then sz = HAMN_NormalizeSize(p(0))
                    If UBound(p) >= 1 Then nm = CStr(p(1))
                    If UBound(p) >= 2 Then no = CStr(p(2))
                    If UBound(p) >= 3 Then nick = CStr(p(3))

                    If sz <> "" Then

                        orderN = orderN + 1

                        If orderN = 1 Then
                            ReDim orderSize(1 To orderN)
                            ReDim orderName(1 To orderN)
                            ReDim orderNo(1 To orderN)
                            ReDim orderNick(1 To orderN)
                        Else
                            ReDim Preserve orderSize(1 To orderN)
                            ReDim Preserve orderName(1 To orderN)
                            ReDim Preserve orderNo(1 To orderN)
                            ReDim Preserve orderNick(1 To orderN)
                        End If

                        orderSize(orderN) = sz
                        orderName(orderN) = HAMN_RemoveLigatures(Trim$(nm))
                        orderNo(orderN) = HAMN_RemoveLigatures(Trim$(no))
                        orderNick(orderN) = HAMN_RemoveLigatures(Trim$(nick))

                        If orderQty.Exists(sz) Then
                            orderQty(sz) = CLng(orderQty(sz)) + 1
                        Else
                            orderQty.Add sz, 1
                        End If

                        If seqDict.Exists(sz) Then
                            seqDict(sz) = CLng(seqDict(sz)) + 1
                        Else
                            seqDict.Add sz, 1
                        End If

                        seq = CLng(seqDict(sz))
                        orderSeqMap.Add sz & "|" & CStr(seq), orderN

                    End If

                End If

            End If

        End If

    Loop

    Close #f

    HAMN_LoadOrderRecords = True
    Exit Function

ErrHandler:
    On Error Resume Next
    Close #f
    HAMN_LoadOrderRecords = False

End Function


'==========================================================
' CATALOG BUCKET CLASSIFIER
'==========================================================

Private Sub HAMN_ClassifyCatalogBuckets( _
    ByVal catN As Long, _
    ByRef catSize() As String, _
    ByRef catW() As Double, _
    ByRef catH() As Double, _
    ByRef catName() As String, _
    ByRef catBucket() As String _
)

    Dim i As Long
    Dim sz As Variant
    Dim sizes As Object

    Set sizes = CreateObject("Scripting.Dictionary")
    sizes.CompareMode = 1

    For i = 1 To catN
        If Not sizes.Exists(catSize(i)) Then sizes.Add catSize(i), True

        If catBucket(i) = "" Then
            catBucket(i) = HAMN_BucketFromName(catName(i))
        Else
            catBucket(i) = HAMN_NormalizeBucket(catBucket(i))
        End If
    Next i

    For Each sz In sizes.keys
        HAMN_ClassifyOneSize CStr(sz), catN, catSize, catW, catH, catBucket
    Next sz

End Sub


Private Sub HAMN_ClassifyOneSize( _
    ByVal sz As String, _
    ByVal catN As Long, _
    ByRef catSize() As String, _
    ByRef catW() As Double, _
    ByRef catH() As Double, _
    ByRef catBucket() As String _
)

    Dim i As Long
    Dim pass As Long
    Dim maxArea As Double
    Dim a As Double
    Dim bestIdx As Long
    Dim bestArea As Double
    Dim minDim As Double

    maxArea = 0#

    For i = 1 To catN
        If UCase$(catSize(i)) = UCase$(sz) Then
            a = catW(i) * catH(i)
            If a > maxArea Then maxArea = a
        End If
    Next i

    If maxArea <= 0# Then Exit Sub

    ' Ambil dua panel terbesar per size sebagai BODY bila belum diberi bucket khusus.
    For pass = 1 To 2
        bestIdx = 0
        bestArea = -1#

        For i = 1 To catN
            If UCase$(catSize(i)) = UCase$(sz) Then
                If catBucket(i) = "" Then
                    a = catW(i) * catH(i)
                    If a > bestArea Then
                        bestArea = a
                        bestIdx = i
                    End If
                End If
            End If
        Next i

        If bestIdx > 0 Then catBucket(bestIdx) = HAMN_BUCKET_BODY
    Next pass

    ' Sisanya: kecil sebagai SMALL, medium sebagai SLEEVE.
    For i = 1 To catN
        If UCase$(catSize(i)) = UCase$(sz) Then
            If catBucket(i) = "" Then
                a = catW(i) * catH(i)
                minDim = HAMN_MinD(catW(i), catH(i))

                If minDim < 8# Or a <= (maxArea * 0.12) Then
                    catBucket(i) = HAMN_BUCKET_SMALL
                ElseIf a <= (maxArea * 0.55) Then
                    catBucket(i) = HAMN_BUCKET_SLEEVE
                Else
                    catBucket(i) = HAMN_BUCKET_BODY
                End If
            End If
        End If
    Next i

End Sub


Private Function HAMN_BucketFromName(ByVal nm As String) As String

    Dim u As String
    u = UCase$(Trim$(nm))

    If u = "" Then
        HAMN_BucketFromName = ""
        Exit Function
    End If

    If InStr(1, u, "KERAH", vbTextCompare) > 0 Or _
       InStr(1, u, "COLLAR", vbTextCompare) > 0 Or _
       InStr(1, u, "NECK", vbTextCompare) > 0 Or _
       InStr(1, u, "TULANG", vbTextCompare) > 0 Or _
       InStr(1, u, "RIB", vbTextCompare) > 0 Then
        HAMN_BucketFromName = HAMN_BUCKET_SMALL
        Exit Function
    End If

    If InStr(1, u, "LENGAN", vbTextCompare) > 0 Or _
       InStr(1, u, "SLEEVE", vbTextCompare) > 0 Or _
       InStr(1, u, "ARM", vbTextCompare) > 0 Then
        HAMN_BucketFromName = HAMN_BUCKET_SLEEVE
        Exit Function
    End If

    If InStr(1, u, "DEPAN", vbTextCompare) > 0 Or _
       InStr(1, u, "BELAKANG", vbTextCompare) > 0 Or _
       InStr(1, u, "FRONT", vbTextCompare) > 0 Or _
       InStr(1, u, "BACK", vbTextCompare) > 0 Or _
       InStr(1, u, "BODY", vbTextCompare) > 0 Or _
       InStr(1, u, "BADAN", vbTextCompare) > 0 Then
        HAMN_BucketFromName = HAMN_BUCKET_BODY
        Exit Function
    End If

    HAMN_BucketFromName = ""

End Function


Private Function HAMN_NormalizeBucket(ByVal s As String) As String

    s = UCase$(Trim$(s))

    Select Case s
        Case HAMN_BUCKET_BODY, HAMN_BUCKET_SLEEVE, HAMN_BUCKET_SMALL
            HAMN_NormalizeBucket = s
        Case Else
            HAMN_NormalizeBucket = ""
    End Select

End Function


Private Function HAMN_PreflightCatalogSize( _
    ByVal catN As Long, _
    ByRef catSize() As String, _
    ByRef catPanelNo() As Long, _
    ByRef catW() As Double, _
    ByRef catH() As Double _
) As String

    Dim i As Long
    Dim msg As String

    msg = ""

    For i = 1 To catN
        If catW(i) > HAMN_AREA_W Or catH(i) > HAMN_AREA_H Then
            msg = msg & "- " & catSize(i) & " panel " & CStr(catPanelNo(i)) & _
                  " = " & HAMN_DblToStr(catW(i)) & " x " & HAMN_DblToStr(catH(i)) & " cm" & vbCrLf
        End If
    Next i

    HAMN_PreflightCatalogSize = msg

End Function


'==========================================================
' MATCH SOURCE PANEL
'==========================================================

Private Function HAMN_FindMatchingSource( _
    ByVal targetW As Double, _
    ByVal targetH As Double, _
    ByVal srcN As Long, _
    ByRef srcW() As Double, _
    ByRef srcH() As Double, _
    ByRef srcUsed() As Boolean _
) As Long

    Dim i As Long
    Dim dw As Double
    Dim dh As Double
    Dim bestIdx As Long
    Dim bestScore As Double
    Dim score As Double

    bestIdx = 0
    bestScore = 999999#

    For i = 1 To srcN

        If Not srcUsed(i) Then

            dw = Abs(srcW(i) - targetW)
            dh = Abs(srcH(i) - targetH)

            If dw <= HAMN_MATCH_TOL And dh <= HAMN_MATCH_TOL Then
                score = dw + dh
                If score < bestScore Then
                    bestScore = score
                    bestIdx = i
                End If
            End If

        End If

    Next i

    HAMN_FindMatchingSource = bestIdx

End Function


'==========================================================
' AUTO RENAME INTERNAL
'==========================================================

Private Sub HAMN_RenameTextRecursive( _
    ByVal s As Shape, _
    ByVal nm As String, _
    ByVal no As String, _
    ByVal nick As String, _
    ByRef nameR As Long, _
    ByRef numR As Long, _
    ByRef nickR As Long _
)

    Dim c As Shape
    Dim pcShapes As Shapes
    Dim raw As String

    On Error Resume Next

    If s.Type = cdrGroupShape Then
        For Each c In s.Shapes
            HAMN_RenameTextRecursive c, nm, no, nick, nameR, numR, nickR
        Next c
        Exit Sub
    End If

    Set pcShapes = s.PowerClip.Shapes

    If Not pcShapes Is Nothing Then
        For Each c In pcShapes
            HAMN_RenameTextRecursive c, nm, no, nick, nameR, numR, nickR
        Next c
    End If

    On Error GoTo 0

    If s.Type <> cdrTextShape Then Exit Sub

    raw = ""
    On Error Resume Next
    raw = s.Text.Story.Text
    On Error GoTo 0

    raw = Trim$(raw)

    If raw = "" Then
        HAMN_QueueDeleteTextShape s
        Exit Sub
    End If

    If HAMN_IgnoreTextShape(s, raw) Then Exit Sub

    If HAMN_IsNamePlaceholder(raw) Then
        If HAMN_SetTextSafe(s, nm) Then
            If Trim$(nm) <> "" Then nameR = nameR + 1
        End If
        Exit Sub
    End If

    If HAMN_IsNicknamePlaceholder(raw) Then
        If HAMN_SetTextSafe(s, nick) Then
            If Trim$(nick) <> "" Then nickR = nickR + 1
        End If
        Exit Sub
    End If

    If HAMN_IsNumberPlaceholder(s, raw) Then
        If HAMN_SetTextSafe(s, no) Then
            If Trim$(no) <> "" Then numR = numR + 1
        End If
        Exit Sub
    End If

End Sub


Private Function HAMN_SetTextSafe(ByVal t As Shape, ByVal newText As String) As Boolean

    On Error Resume Next

    Dim hasCJK As Boolean
    Dim fontReadyBefore As Boolean
    Dim fontReadyAfter As Boolean

    newText = CStr(newText)
    newText = HAMN_RemoveLigatures(newText)

    If Trim$(newText) = "" Then
        HAMN_QueueDeleteTextShape t
        HAMN_SetTextSafe = True
        Exit Function
    End If

    hasCJK = HAMN_ContainsJapaneseOrCJK(newText)

    If hasCJK Then
        HAMN_CJKDetectedCount = HAMN_CJKDetectedCount + 1
        fontReadyBefore = HAMN_ApplyJapaneseFont(t)
    End If

    err.Clear
    t.Text.Story.Text = newText

    If err.Number <> 0 Then
        HAMN_SetTextSafe = False
        err.Clear
        On Error GoTo 0
        Exit Function
    End If

    If hasCJK Then

        fontReadyAfter = HAMN_ApplyJapaneseFont(t)

        If Trim$(t.Text.Story.Text) = "" Then
            err.Clear
            t.Text.Story.Text = newText
            fontReadyAfter = HAMN_ApplyJapaneseFont(t)
        End If

        If fontReadyBefore Or fontReadyAfter Then
            HAMN_CJKFontAppliedCount = HAMN_CJKFontAppliedCount + 1
        Else
            HAMN_CJKFontFailedCount = HAMN_CJKFontFailedCount + 1
        End If

    End If

    HAMN_SetTextSafe = True

    On Error GoTo 0

End Function


Private Sub HAMN_QueueDeleteTextShape(ByVal t As Shape)
    On Error Resume Next
    If HAMN_DeleteQueue Is Nothing Then Set HAMN_DeleteQueue = New Collection
    HAMN_DeleteQueue.Add t
End Sub


Private Sub HAMN_DeleteQueuedTextShapes()

    On Error Resume Next

    Dim i As Long
    Dim shp As Shape

    If HAMN_DeleteQueue Is Nothing Then Exit Sub

    For i = HAMN_DeleteQueue.Count To 1 Step -1
        Set shp = HAMN_DeleteQueue(i)
        If Not shp Is Nothing Then
            shp.Delete
            HAMN_DeletedTextCount = HAMN_DeletedTextCount + 1
        End If
    Next i

End Sub


Private Function HAMN_RemoveLigatures(ByVal s As String) As String

    On Error Resume Next

    s = Replace(s, ChrW(&HFB00), "ff")
    s = Replace(s, ChrW(&HFB01), "fi")
    s = Replace(s, ChrW(&HFB02), "fl")
    s = Replace(s, ChrW(&HFB03), "ffi")
    s = Replace(s, ChrW(&HFB04), "ffl")
    s = Replace(s, ChrW(&HFB05), "st")
    s = Replace(s, ChrW(&HFB06), "st")

    On Error GoTo 0

    HAMN_RemoveLigatures = s

End Function


Private Function HAMN_IsNamePlaceholder(ByVal txt As String) As Boolean

    Dim s As String
    s = HAMN_NormalizeText(txt)

    Select Case s
        Case "NAMA ATLIT", _
             "NAMA ATLET", _
             "NAMA", _
             "PLAYER", _
             "PLAYERS", _
             "PLAYER NAME", _
             "NAMA PEMAIN"
            HAMN_IsNamePlaceholder = True
    End Select

End Function


Private Function HAMN_IsNicknamePlaceholder(ByVal txt As String) As Boolean

    Dim s As String
    s = HAMN_NormalizeText(txt)

    Select Case s
        Case "NICKNAME", _
             "NICK NAME", _
             "NICK", _
             "NAMA PANGGILAN"
            HAMN_IsNicknamePlaceholder = True
    End Select

End Function


Private Function HAMN_IsNumberPlaceholder(ByVal t As Shape, ByVal txt As String) As Boolean

    Dim s As String
    s = Trim$(txt)

    If s = "" Then Exit Function
    If Len(s) > 3 Then Exit Function
    If Not IsNumeric(s) Then Exit Function

    If t.SizeHeight < HAMN_MIN_TEXT_H Then Exit Function

    HAMN_IsNumberPlaceholder = True

End Function


Private Function HAMN_IgnoreTextShape(ByVal t As Shape, ByVal txt As String) As Boolean

    Dim s As String
    Dim u As String

    s = Trim$(txt)
    u = HAMN_NormalizeText(s)

    If Left$(u, 3) = "@A:" Then
        HAMN_IgnoreTextShape = True
        Exit Function
    End If

    If Left$(u, 6) = "@ATTR:" Then
        HAMN_IgnoreTextShape = True
        Exit Function
    End If

    If u = "IDPO" Then
        If t.SizeHeight >= HAMN_ID_MIN_H And t.SizeHeight <= HAMN_ID_MAX_H Then
            HAMN_IgnoreTextShape = True
        End If
        Exit Function
    End If

    If Len(s) = 6 And IsNumeric(s) Then
        If t.SizeHeight >= HAMN_ID_MIN_H And t.SizeHeight <= HAMN_ID_MAX_H Then
            HAMN_IgnoreTextShape = True
        End If
    End If

End Function


Private Function HAMN_BuildRenameCheck( _
    ByVal orderN As Long, _
    ByRef orderSize() As String, _
    ByRef orderName() As String, _
    ByRef orderNo() As String, _
    ByRef orderNick() As String, _
    ByRef recNameR() As Long, _
    ByRef recNumR() As Long, _
    ByRef recNickR() As Long _
) As String

    Dim i As Long
    Dim msg As String

    msg = ""

    For i = 1 To orderN

        If Trim$(orderName(i)) <> "" And recNameR(i) = 0 Then
            msg = msg & "- WARNING: Nama tidak menemukan placeholder. Record " & CStr(i) & " " & _
                  orderSize(i) & " | " & orderName(i) & " | " & orderNo(i) & vbCrLf
        End If

        If Trim$(orderNo(i)) <> "" And recNumR(i) = 0 Then
            msg = msg & "- WARNING: Nomor tidak menemukan placeholder. Record " & CStr(i) & " " & _
                  orderSize(i) & " | " & orderName(i) & " | " & orderNo(i) & vbCrLf
        End If

        If Trim$(orderNick(i)) <> "" And recNickR(i) = 0 Then
            msg = msg & "- WARNING: Nickname tidak menemukan placeholder. Record " & CStr(i) & " " & _
                  orderSize(i) & " | " & orderName(i) & " | " & orderNo(i) & " | " & orderNick(i) & vbCrLf
        End If

    Next i

    HAMN_BuildRenameCheck = msg

End Function


Private Function HAMN_ContainsJapaneseOrCJK(ByVal s As String) As Boolean

    Dim i As Long
    Dim ch As String
    Dim code As Long

    For i = 1 To Len(s)

        ch = Mid$(s, i, 1)
        code = AscW(ch)

        If code < 0 Then code = code + 65536

        If HAMN_IsJapaneseOrCJKCode(code) Then
            HAMN_ContainsJapaneseOrCJK = True
            Exit Function
        End If

    Next i

End Function


Private Function HAMN_IsJapaneseOrCJKCode(ByVal code As Long) As Boolean

    Select Case code
        Case &H3000 To &H303F, _
             &H3040 To &H309F, _
             &H30A0 To &H30FF, _
             &H31F0 To &H31FF, _
             &H3400 To &H4DBF, _
             &H4E00 To &H9FFF, _
             &HF900 To &HFAFF, _
             &HFF00 To &HFFEF
            HAMN_IsJapaneseOrCJKCode = True
    End Select

End Function


Private Function HAMN_ApplyJapaneseFont(ByVal t As Shape) As Boolean

    HAMN_ApplyJapaneseFont = False

    If t Is Nothing Then Exit Function

    If HAMN_TrySetFontName(t, HAMN_CJK_FONT_1, True) Then
        HAMN_ApplyJapaneseFont = True
        Exit Function
    End If

    If HAMN_TrySetFontName(t, HAMN_CJK_FONT_2, True) Then
        HAMN_ApplyJapaneseFont = True
        Exit Function
    End If

    If HAMN_TrySetFontName(t, HAMN_CJK_FONT_3, True) Then
        HAMN_ApplyJapaneseFont = True
        Exit Function
    End If

    If HAMN_TrySetFontName(t, HAMN_CJK_FONT_4, True) Then
        HAMN_ApplyJapaneseFont = True
        Exit Function
    End If

    If HAMN_TrySetFontName(t, HAMN_CJK_FONT_5, True) Then
        HAMN_ApplyJapaneseFont = True
        Exit Function
    End If

End Function


Private Function HAMN_TrySetFontName( _
    ByVal t As Shape, _
    ByVal fontName As String, _
    Optional ByVal makeBold As Boolean = True _
) As Boolean

    HAMN_TrySetFontName = False

    If t Is Nothing Then Exit Function
    If Trim$(fontName) = "" Then Exit Function

    On Error Resume Next

    err.Clear
    t.Text.Story.Font = fontName

    If err.Number <> 0 Then
        err.Clear
        On Error GoTo 0
        Exit Function
    End If

    If makeBold Then
        err.Clear
        t.Text.Story.Bold = True
        err.Clear
    End If

    HAMN_TrySetFontName = True

    On Error GoTo 0

End Function


'==========================================================
' OUTPUT ARRAY
'==========================================================

Private Sub HAMN_AddOutputShape( _
    ByRef outN As Long, _
    ByRef outShapes() As Shape, _
    ByRef outW() As Double, _
    ByRef outH() As Double, _
    ByRef outArea() As Double, _
    ByRef outLabel() As String, _
    ByRef outBucket() As String, _
    ByVal shp As Shape, _
    ByVal w As Double, _
    ByVal h As Double, _
    ByVal bucket As String, _
    ByVal label As String _
)

    outN = outN + 1

    If outN = 1 Then
        ReDim outShapes(1 To outN)
        ReDim outW(1 To outN)
        ReDim outH(1 To outN)
        ReDim outArea(1 To outN)
        ReDim outLabel(1 To outN)
        ReDim outBucket(1 To outN)
    Else
        ReDim Preserve outShapes(1 To outN)
        ReDim Preserve outW(1 To outN)
        ReDim Preserve outH(1 To outN)
        ReDim Preserve outArea(1 To outN)
        ReDim Preserve outLabel(1 To outN)
        ReDim Preserve outBucket(1 To outN)
    End If

    Set outShapes(outN) = shp
    outW(outN) = w
    outH(outN) = h
    outArea(outN) = w * h
    outLabel(outN) = label
    outBucket(outN) = bucket

End Sub


Private Sub HAMN_SortOutputByBucketArea( _
    ByVal outN As Long, _
    ByRef outShapes() As Shape, _
    ByRef outW() As Double, _
    ByRef outH() As Double, _
    ByRef outArea() As Double, _
    ByRef outLabel() As String, _
    ByRef outBucket() As String _
)

    Dim i As Long
    Dim j As Long

    Dim tmpShape As Shape
    Dim tmpD As Double
    Dim tmpS As String

    For i = 1 To outN - 1
        For j = i + 1 To outN

            If HAMN_ShouldSwapOutput(i, j, outArea, outBucket, outLabel) Then

                Set tmpShape = outShapes(i)
                Set outShapes(i) = outShapes(j)
                Set outShapes(j) = tmpShape

                tmpD = outW(i)
                outW(i) = outW(j)
                outW(j) = tmpD

                tmpD = outH(i)
                outH(i) = outH(j)
                outH(j) = tmpD

                tmpD = outArea(i)
                outArea(i) = outArea(j)
                outArea(j) = tmpD

                tmpS = outLabel(i)
                outLabel(i) = outLabel(j)
                outLabel(j) = tmpS

                tmpS = outBucket(i)
                outBucket(i) = outBucket(j)
                outBucket(j) = tmpS

            End If

        Next j
    Next i

End Sub


Private Function HAMN_ShouldSwapOutput( _
    ByVal i As Long, _
    ByVal j As Long, _
    ByRef outArea() As Double, _
    ByRef outBucket() As String, _
    ByRef outLabel() As String _
) As Boolean

    Dim pi As Long
    Dim pj As Long

    pi = HAMN_BucketPriority(outBucket(i))
    pj = HAMN_BucketPriority(outBucket(j))

    If pj < pi Then
        HAMN_ShouldSwapOutput = True
        Exit Function
    End If

    If pj = pi Then
        If outArea(j) > outArea(i) Then
            HAMN_ShouldSwapOutput = True
            Exit Function
        End If

        If outArea(j) = outArea(i) Then
            If outLabel(j) < outLabel(i) Then
                HAMN_ShouldSwapOutput = True
                Exit Function
            End If
        End If
    End If

    HAMN_ShouldSwapOutput = False

End Function


Private Function HAMN_BucketPriority(ByVal bucket As String) As Long

    Select Case UCase$(Trim$(bucket))
        Case HAMN_BUCKET_BODY
            HAMN_BucketPriority = 1
        Case HAMN_BUCKET_SLEEVE
            HAMN_BucketPriority = 2
        Case HAMN_BUCKET_SMALL
            HAMN_BucketPriority = 3
        Case Else
            HAMN_BucketPriority = 9
    End Select

End Function



'==========================================================
' ACCURATE PAGE ANCHOR
'==========================================================

Private Sub HAMN_GetAccuratePageAnchor( _
    ByRef outCenterX As Double, _
    ByRef outTopY As Double, _
    ByRef outMethod As String _
)

    'Fallback lama berbasis ukuran page.
    outCenterX = ActivePage.SizeWidth / 2#
    outTopY = ActivePage.SizeHeight
    outMethod = "Fallback Width/Height"

    Dim tmp As Shape
    Dim centerY As Double
    Dim failed As Boolean

    failed = False

    On Error Resume Next

    Set tmp = ActiveLayer.CreateRectangle2(0, 0, 0.2, 0.2)

    If tmp Is Nothing Then
        On Error GoTo 0
        Exit Sub
    End If

    err.Clear
    CallByName tmp, "AlignToPageCenter", VbMethod, cdrAlignHCenter
    If err.Number <> 0 Then failed = True

    err.Clear
    CallByName tmp, "AlignToPageCenter", VbMethod, cdrAlignVCenter
    If err.Number <> 0 Then failed = True

    If failed = False Then
        outCenterX = tmp.PositionX
        centerY = tmp.PositionY
        outTopY = centerY + (ActivePage.SizeHeight / 2#)
        outMethod = "Smart Temporary Shape Align"
    End If

    tmp.Delete

    On Error GoTo 0

End Sub


'==========================================================
' LAYOUT SHELF ALGORITHM
'==========================================================

Private Function HAMN_LayoutShelf( _
    ByVal outN As Long, _
    ByRef outShapes() As Shape, _
    ByRef outW() As Double, _
    ByRef outH() As Double, _
    ByRef outLabel() As String, _
    ByRef outBucket() As String, _
    ByVal baseLeft As Double, _
    ByVal baseTop As Double, _
    ByRef logText As String, _
    ByRef warn As String _
) As Long

    '======================================================
    ' V3.5 ROW-MAJOR 6 x 2
    '
    ' Yang dimaksud 2 row:
    '   ROW 1 : item 1, 2, 3, 4, 5, 6
    '   ROW 2 : item 7, 8, 9, 10, 11, 12
    '
    ' Bukan zig-zag:
    '   1 atas, 2 bawah, 3 atas, 4 bawah, dst.
    '
    ' Setelah 12 item, lanjut blok berikutnya secara horizontal.
    ' Movement tetap pakai bounding box top-left dan tidak menembus group.
    '======================================================

    Dim blockStart As Long
    Dim blockEnd As Long
    Dim slotCount As Long
    Dim localIdx As Long
    Dim rowIdx As Long
    Dim colIdx As Long

    Dim i As Long
    Dim c As Long
    Dim r As Long

    Dim colW(0 To 5) As Double
    Dim rowH(0 To 1) As Double
    Dim colX(0 To 5) As Double
    Dim rowY(0 To 1) As Double

    Dim usedCols As Long
    Dim usedRows As Long
    Dim blockW As Double
    Dim blockH As Double
    Dim blockLeft As Double
    Dim blockOffsetX As Double

    Dim w As Double
    Dim h As Double
    Dim targetLeft As Double
    Dim targetTop As Double

    Dim blockIndex As Long

    logText = ""

    If outN <= 0 Then
        HAMN_LayoutShelf = 0
        Exit Function
    End If

    blockStart = 1
    blockIndex = 1
    blockOffsetX = 0#

    Do While blockStart <= outN

        blockEnd = blockStart + (HAMN_GRID_COLS * HAMN_GRID_ROWS) - 1
        If blockEnd > outN Then blockEnd = outN

        slotCount = blockEnd - blockStart + 1

        If slotCount <= HAMN_GRID_COLS Then
            usedCols = slotCount
            usedRows = 1
        Else
            usedCols = HAMN_GRID_COLS
            usedRows = 2
        End If

        For c = 0 To 5
            colW(c) = 0#
            colX(c) = 0#
        Next c

        For r = 0 To 1
            rowH(r) = 0#
            rowY(r) = 0#
        Next r

        'Hitung lebar kolom dan tinggi row berdasarkan item dalam blok 6x2.
        For i = blockStart To blockEnd

            w = outW(i)
            h = outH(i)

            If w > HAMN_AREA_W Or h > HAMN_AREA_H Then
                warn = warn & "- FAIL: Panel terlalu besar untuk area 178 x 255 setelah rename: " & _
                              outLabel(i) & " (" & HAMN_DblToStr(w) & " x " & HAMN_DblToStr(h) & " cm)" & vbCrLf
                GoTo NextMeasureItem
            End If

            localIdx = i - blockStart
            rowIdx = localIdx \ HAMN_GRID_COLS
            colIdx = localIdx Mod HAMN_GRID_COLS

            If w > colW(colIdx) Then colW(colIdx) = w
            If h > rowH(rowIdx) Then rowH(rowIdx) = h

NextMeasureItem:
        Next i

        'Offset kolom.
        colX(0) = 0#
        For c = 1 To usedCols - 1
            colX(c) = colX(c - 1) + colW(c - 1) + HAMN_GAP
        Next c

        'Offset row.
        rowY(0) = 0#
        If usedRows > 1 Then
            rowY(1) = rowH(0) + HAMN_GAP
        End If

        blockW = 0#
        If usedCols > 0 Then
            For c = 0 To usedCols - 1
                blockW = blockW + colW(c)
            Next c
            blockW = blockW + ((usedCols - 1) * HAMN_GAP)
        End If

        blockH = rowH(0)
        If usedRows > 1 Then
            blockH = rowH(0) + HAMN_GAP + rowH(1)
        End If

        'Tetap catat jika blok 6 kolom secara fisik lebih besar dari area 178.
        'Tidak dipecah zig-zag supaya perilaku sesuai request: row-major 6x2.
        If blockW > HAMN_AREA_W Then
            warn = warn & "- WARNING: Blok " & CStr(blockIndex) & _
                          " row-major 6 kolom melebihi lebar 178 cm. Width=" & _
                          HAMN_DblToStr(blockW) & " cm." & vbCrLf
        End If

        If blockH > HAMN_AREA_H Then
            warn = warn & "- WARNING: Blok " & CStr(blockIndex) & _
                          " row-major 2 row melebihi tinggi 255 cm. Height=" & _
                          HAMN_DblToStr(blockH) & " cm." & vbCrLf
        End If

        blockLeft = baseLeft + blockOffsetX

        'Place item row-major:
        '1-6 di row atas, 7-12 di row bawah.
        For i = blockStart To blockEnd

            localIdx = i - blockStart
            rowIdx = localIdx \ HAMN_GRID_COLS
            colIdx = localIdx Mod HAMN_GRID_COLS

            targetLeft = blockLeft + colX(colIdx)
            targetTop = baseTop - rowY(rowIdx)

            HAMN_MoveShapeTopLeft outShapes(i), targetLeft, targetTop

            logText = logText & _
                      "BLOCK " & CStr(blockIndex) & " | " & _
                      "ROW " & CStr(rowIdx + 1) & " COL " & CStr(colIdx + 1) & " | " & _
                      outBucket(i) & " | " & _
                      outLabel(i) & " | " & _
                      HAMN_DblToStr(outW(i)) & " x " & HAMN_DblToStr(outH(i)) & _
                      " | X=" & HAMN_DblToStr(colX(colIdx)) & _
                      " Y=" & HAMN_DblToStr(rowY(rowIdx)) & vbCrLf

        Next i

        If blockW <= 0# Then
            blockOffsetX = blockOffsetX + HAMN_AREA_W + HAMN_AREA_SPACING
        Else
            blockOffsetX = blockOffsetX + blockW + HAMN_AREA_SPACING
        End If

        blockStart = blockEnd + 1
        blockIndex = blockIndex + 1

    Loop

    HAMN_LayoutShelf = blockIndex - 1

End Function


Private Sub HAMN_MoveShapeTopLeft(ByVal shp As Shape, ByVal targetLeft As Double, ByVal targetTop As Double)

    Dim dx As Double
    Dim dy As Double

    dx = targetLeft - shp.leftX
    dy = targetTop - shp.topY

    shp.Move dx, dy

End Sub


'==========================================================
' REPORT
'==========================================================

Private Sub HAMN_WriteReport( _
    ByVal path As String, _
    ByVal outN As Long, _
    ByVal expectedPanels As Long, _
    ByVal areaCount As Long, _
    ByVal orderN As Long, _
    ByRef orderSize() As String, _
    ByRef orderName() As String, _
    ByRef orderNo() As String, _
    ByRef orderNick() As String, _
    ByRef recNameR() As Long, _
    ByRef recNumR() As Long, _
    ByRef recNickR() As Long, _
    ByVal orderQty As Object, _
    ByVal catPanelCountBySize As Object, _
    ByVal totalNameR As Long, _
    ByVal totalNumR As Long, _
    ByVal totalNickR As Long, _
    ByVal warn As String, _
    ByVal logText As String _
)

    On Error Resume Next

    Dim f As Integer
    Dim key As Variant
    Dim statusText As String
    Dim i As Long

    If warn <> "" Or outN <> expectedPanels Then
        statusText = "WARNING_OR_FAIL"
    Else
        statusText = "PASS"
    End If

    f = FreeFile
    Open path For Output As #f

    Print #f, "PROJECT HADES - AUTO MASS NESTING REPORT V3.5 ROW MAJOR 6x2"
    Print #f, "CREATED=" & Format$(Now, "yyyy-mm-dd hh:nn:ss")
    Print #f, "STATUS=" & statusText
    Print #f, "AUTO_RENAME=ON"
    Print #f, ""
    Print #f, "AREA_WIDTH_CM=" & HAMN_DblToStr(HAMN_AREA_W)
    Print #f, "AREA_HEIGHT_CM=" & HAMN_DblToStr(HAMN_AREA_H)
    Print #f, "GAP_CM=" & HAMN_DblToStr(HAMN_GAP)
    Print #f, "METHOD=ROW_MAJOR_6X2_BLOCK"
    Print #f, "BUCKET_PRIORITY=BODY>SLEEVE>SMALL"
    Print #f, ""
    Print #f, "ORDER_RECORDS=" & CStr(orderN)
    Print #f, "TOTAL_EXPECTED_PANEL=" & CStr(expectedPanels)
    Print #f, "TOTAL_OUTPUT_PANEL=" & CStr(outN)
    Print #f, "TOTAL_BLOCK_USED=" & CStr(areaCount)
    Print #f, ""
    Print #f, "[AUTO RENAME SUMMARY]"
    Print #f, "NAME_REPLACED=" & CStr(totalNameR)
    Print #f, "NUMBER_REPLACED=" & CStr(totalNumR)
    Print #f, "NICKNAME_REPLACED=" & CStr(totalNickR)
    Print #f, "EMPTY_PLACEHOLDER_DELETED=" & CStr(HAMN_DeletedTextCount)
    Print #f, "CJK_DETECTED=" & CStr(HAMN_CJKDetectedCount)
    Print #f, "CJK_FONT_APPLIED=" & CStr(HAMN_CJKFontAppliedCount)
    Print #f, "CJK_FONT_FAILED=" & CStr(HAMN_CJKFontFailedCount)
    Print #f, ""

    Print #f, "[ORDER QTY]"
    For Each key In orderQty.keys
        Print #f, CStr(key) & "=" & CStr(orderQty(key))
    Next key

    Print #f, ""
    Print #f, "[CATALOG PANEL COUNT PER SIZE]"
    For Each key In catPanelCountBySize.keys
        Print #f, CStr(key) & "=" & CStr(catPanelCountBySize(key))
    Next key

    Print #f, ""
    Print #f, "[RENAME CHECK - ONLY NONZERO / WARNING RELEVANT]"
    For i = 1 To orderN
        If Trim$(orderName(i)) <> "" Or Trim$(orderNo(i)) <> "" Or Trim$(orderNick(i)) <> "" Then
            Print #f, "REC " & CStr(i) & " | " & orderSize(i) & " | " & _
                      orderName(i) & " | " & orderNo(i) & " | " & orderNick(i) & _
                      " | N=" & CStr(recNameR(i)) & _
                      " NO=" & CStr(recNumR(i)) & _
                      " NICK=" & CStr(recNickR(i))
        End If
    Next i

    Print #f, ""
    Print #f, "[WARNING / FAIL]"
    If warn <> "" Then
        Print #f, warn
    Else
        Print #f, "NONE"
    End If

    Print #f, ""
    Print #f, "[LAYOUT LOG]"
    Print #f, logText

    Close #f

End Sub


'==========================================================
' HELPERS
'==========================================================

Private Function HAMN_GetDocumentsPath() As String
    HAMN_GetDocumentsPath = CreateObject("WScript.Shell").SpecialFolders("MyDocuments")
End Function


Private Function HAMN_NormalizeSize(ByVal s As String) As String

    s = UCase$(Trim$(s))

    Select Case s
        Case "XXL"
            s = "2XL"
        Case "XXXL"
            s = "3XL"
        Case "XXXXL"
            s = "4XL"
        Case "XXXXXL"
            s = "5XL"
        Case "XXXXXXL"
            s = "6XL"
    End Select

    Select Case s
        Case "XXS", "XS", "S", "M", "L", "XL", "2XL", "3XL", "4XL", "5XL", "6XL"
            HAMN_NormalizeSize = s
        Case Else
            HAMN_NormalizeSize = ""
    End Select

End Function


Private Function HAMN_NormalizeText(ByVal s As String) As String

    s = Replace(s, vbCr, "")
    s = Replace(s, vbLf, "")
    s = Replace(s, Chr(160), " ")

    On Error Resume Next
    s = Replace(s, ChrW(&HFB00), "FF")
    s = Replace(s, ChrW(&HFB01), "FI")
    s = Replace(s, ChrW(&HFB02), "FL")
    s = Replace(s, ChrW(&HFB03), "FFI")
    s = Replace(s, ChrW(&HFB04), "FFL")
    s = Replace(s, ChrW(&HFB05), "ST")
    s = Replace(s, ChrW(&HFB06), "ST")
    On Error GoTo 0

    Do While InStr(1, s, "  ", vbTextCompare) > 0
        s = Replace(s, "  ", " ")
    Loop

    HAMN_NormalizeText = UCase$(Trim$(s))

End Function


Private Function HAMN_Val(ByVal s As String) As Double

    s = Trim$(s)
    s = Replace(s, ",", ".")

    HAMN_Val = val(s)

End Function


Private Function HAMN_DblToStr(ByVal v As Double) As String
    HAMN_DblToStr = Replace(Format$(v, "0.000"), ",", ".")
End Function


Private Function HAMN_MinD(ByVal a As Double, ByVal b As Double) As Double
    If a < b Then
        HAMN_MinD = a
    Else
        HAMN_MinD = b
    End If
End Function
