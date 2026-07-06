'=========================================================
' AUTO-GENERATED PATCH: NEST ENGINE RENAMED FOR SMART DISPATCHER
' Do not import together with old QC Typo Nest module that still exposes QC_TYPO_NEST_CHECK/HADES_QC_TYPO_NEST_REPORT.
'=========================================================
Option Explicit

'=========================================================
' PROJECT H.A.D.E.S.
' QC TYPO NEST ENGINE V1.3
' CorelDRAW 2021 VBA
'
' PUBLIC SHORTCUT:
'   QTN_NEST_CHECK
'
' REPORT MODE FOR HADES FINAL REPORT:
'   HADES_QC_TYPO_NEST_REPORT
'
' TUJUAN:
' - QC Typo khusus hasil HADES Auto Mass Nesting.
' - Satu atlet / satu baris Order.txt boleh tersebar ke banyak panel/box.
' - Validasi bukan hanya "teks ada di Order.txt", tetapi:
'     1) pasangan nama-nomor-nickname benar per row Order.txt
'     2) nomor depan/belakang seragam
'     3) jumlah kemunculan nomor sesuai Build Typo Template
'     4) mendeteksi kemungkinan pasangan nama-nomor tertukar
'     5) popup report langsung + marker hijau hanya pada panel yang memuat teks/nomor bermasalah
'
' SUMBER DATA:
' - Documents\Order.txt
' - Documents\TypoTemplate_Current.txt
' - Documents\SizeDB_*.txt sesuai @SIZEDB atau SIZEDB template
'
' SYARAT PENTING UNTUK MODE MASS NESTING:
' - Hasil duplicate Auto Mass Nesting harus menyimpan metadata pada Shape.Name:
'     HADES_AMN|ROW=7|SIZE=M|PANEL=BODY|UID=M_BODY_P001|PANELNO=1
'
'   Minimal wajib ada ROW dan SIZE.
'   PANEL boleh BODY/SLEEVE/SMALL/BODY_FRONT/BODY_BACK.
'
' CATATAN:
' - Module ini sengaja berdiri sendiri dan memakai prefix QTN_ agar tidak bentrok
'   dengan QC_TYPO_CHECK normal.
' - Tidak menggantikan QC_TYPO_CHECK biasa.
'=========================================================

'=========================================================
' CONFIG
'=========================================================
Private Const QTN_ORDER_FILE As String = "Order.txt"
Private Const QTN_TEMPLATE_FILE As String = "TypoTemplate_Current.txt"
Private Const QTN_REPORT_FILE As String = "HADES_QC_TYPO_NEST_LATEST.txt"
Private Const QTN_POPUP_MAX_CHARS As Long = 3800

Private Const QTN_META_PREFIX As String = "HADES_AMN|"
Private Const QTN_ID_MIN_H As Double = 0.28
Private Const QTN_ID_MAX_H As Double = 0.65
Private Const QTN_MIN_TEXT_H As Double = 0.5
Private Const QTN_SIDE_TOL As Double = 1.2

Private Const QTN_NAME_BIG_H As Double = 3#
Private Const QTN_NUMBER_BACK_H As Double = 15#
Private Const QTN_NUMBER_FRONT_H As Double = 7#

Private Const QTN_ENFORCE_ROW_METADATA As Boolean = True
Private Const QTN_CHECK_UNKNOWN_TEXT As Boolean = False

Private qtnReportMode As Boolean
Private qtnLastStatus As String
Private qtnLastSummary As String
Private qtnLastDetail As String

'=========================================================
' PUBLIC ENTRY
'=========================================================
Public Sub QTN_NEST_CHECK()

    On Error GoTo ERR_HANDLER

    Dim oldUnit As Long
    Dim orderPath As String
    Dim tplPath As String
    Dim reportPath As String
    Dim dbName As String
    Dim dbPath As String

    Dim ordN As Long
    Dim ordSize() As String
    Dim ordName() As String
    Dim ordNo() As String
    Dim ordNick() As String

    Dim tpl As Object
    Dim db As Object

    Dim foundName As Object
    Dim foundNo As Object
    Dim foundNick As Object
    Dim foundNoFront As Object
    Dim foundNoBack As Object
    Dim foundNoUnknownSide As Object
    Dim foundUnknown As Object
    Dim rowTouched As Object
    Dim rowPanelShapes As Object
    Dim rowNameShapes As Object
    Dim rowNoShapes As Object
    Dim rowNoFrontShapes As Object
    Dim rowNoBackShapes As Object
    Dim failRows As Object
    Dim failMarkShapes As Object

    Dim detail As String
    Dim summary As String
    Dim status As String
    Dim failN As Long
    Dim warnN As Long
    Dim scannedTextN As Long
    Dim metaPanelN As Long
    Dim markerN As Long

    qtnLastStatus = ""
    qtnLastSummary = ""
    qtnLastDetail = ""

    If ActiveDocument Is Nothing Then
        QTN_FailOut "Tidak ada dokumen aktif.", ""
        Exit Sub
    End If

    If ActiveSelection Is Nothing Then
        QTN_FailOut "Select semua hasil Auto Mass Nesting terlebih dahulu.", ""
        Exit Sub
    End If

    If ActiveSelection.Shapes.Count = 0 Then
        QTN_FailOut "Select semua hasil Auto Mass Nesting terlebih dahulu.", ""
        Exit Sub
    End If

    oldUnit = ActiveDocument.Unit
    ActiveDocument.Unit = cdrCentimeter

    orderPath = QTN_DocumentsPath() & "\" & QTN_ORDER_FILE
    tplPath = QTN_DocumentsPath() & "\" & QTN_TEMPLATE_FILE
    reportPath = QTN_DocumentsPath() & "\" & QTN_REPORT_FILE

    If Dir$(orderPath) = "" Then
        ActiveDocument.Unit = oldUnit
        QTN_FailOut "Order.txt tidak ditemukan.", orderPath
        Exit Sub
    End If

    If Dir$(tplPath) = "" Then
        ActiveDocument.Unit = oldUnit
        QTN_FailOut "TypoTemplate_Current.txt tidak ditemukan.", _
                    "Jalankan BUILD_TYPO_TEMPLATE terlebih dahulu pada sample master."
        Exit Sub
    End If

    If Not QTN_LoadOrder(orderPath, ordN, ordSize, ordName, ordNo, ordNick) Then
        ActiveDocument.Unit = oldUnit
        QTN_FailOut "Gagal membaca Order.txt.", orderPath
        Exit Sub
    End If

    If ordN = 0 Then
        ActiveDocument.Unit = oldUnit
        QTN_FailOut "Order.txt tidak memiliki baris order.", "Metadata @ otomatis diskip."
        Exit Sub
    End If

    Set tpl = QTN_LoadKeyValueFile(tplPath)
    If tpl Is Nothing Or tpl.Count = 0 Then
        ActiveDocument.Unit = oldUnit
        QTN_FailOut "TypoTemplate_Current.txt kosong / gagal dibaca.", tplPath
        Exit Sub
    End If

    dbName = QTN_ReadOrderMeta(orderPath, "SIZEDB")
    If Len(Trim$(dbName)) = 0 Then
        If tpl.Exists("SIZEDB") Then dbName = CStr(tpl("SIZEDB"))
    End If

    If Len(Trim$(dbName)) = 0 Then
        ActiveDocument.Unit = oldUnit
        QTN_FailOut "SIZEDB tidak ditemukan di Order.txt maupun TypoTemplate.", _
                    "Tambahkan @SIZEDB=SizeDB_xxx.txt di Order.txt atau rebuild TypoTemplate setelah Order.txt aktif."
        Exit Sub
    End If

    dbPath = QTN_DocumentsPath() & "\" & dbName
    If Dir$(dbPath) = "" Then
        ActiveDocument.Unit = oldUnit
        QTN_FailOut "File SizeDB tidak ditemukan.", dbPath
        Exit Sub
    End If

    Set db = QTN_LoadSizeDB(dbPath)
    If db Is Nothing Or db.Count = 0 Then
        ActiveDocument.Unit = oldUnit
        QTN_FailOut "Gagal membaca SizeDB.", dbPath
        Exit Sub
    End If

    Set foundName = CreateObject("Scripting.Dictionary")
    Set foundNo = CreateObject("Scripting.Dictionary")
    Set foundNick = CreateObject("Scripting.Dictionary")
    Set foundNoFront = CreateObject("Scripting.Dictionary")
    Set foundNoBack = CreateObject("Scripting.Dictionary")
    Set foundNoUnknownSide = CreateObject("Scripting.Dictionary")
    Set foundUnknown = CreateObject("Scripting.Dictionary")
    Set rowTouched = CreateObject("Scripting.Dictionary")
    Set rowPanelShapes = CreateObject("Scripting.Dictionary")
    Set rowNameShapes = CreateObject("Scripting.Dictionary")
    Set rowNoShapes = CreateObject("Scripting.Dictionary")
    Set rowNoFrontShapes = CreateObject("Scripting.Dictionary")
    Set rowNoBackShapes = CreateObject("Scripting.Dictionary")
    Set failRows = CreateObject("Scripting.Dictionary")
    Set failMarkShapes = CreateObject("Scripting.Dictionary")

    QTN_InitDict foundName
    QTN_InitDict foundNo
    QTN_InitDict foundNick
    QTN_InitDict foundNoFront
    QTN_InitDict foundNoBack
    QTN_InitDict foundNoUnknownSide
    QTN_InitDict foundUnknown
    QTN_InitDict rowTouched
    QTN_InitDict rowPanelShapes
    QTN_InitDict rowNameShapes
    QTN_InitDict rowNoShapes
    QTN_InitDict rowNoFrontShapes
    QTN_InitDict rowNoBackShapes
    QTN_InitDict failRows
    QTN_InitDict failMarkShapes

    scannedTextN = 0
    metaPanelN = 0

    Dim s As Shape
    For Each s In ActiveSelection.Shapes
        QTN_ScanShape s, 0, "", "", 0, 0, Nothing, db, foundName, foundNo, foundNick, _
                      foundNoFront, foundNoBack, foundNoUnknownSide, foundUnknown, _
                      rowTouched, rowPanelShapes, rowNameShapes, rowNoShapes, rowNoFrontShapes, rowNoBackShapes, _
                      scannedTextN, metaPanelN
    Next s

    detail = QTN_ValidateRows( _
                ordN, ordSize, ordName, ordNo, ordNick, _
                tpl, db, foundName, foundNo, foundNick, _
                foundNoFront, foundNoBack, foundNoUnknownSide, foundUnknown, _
                rowTouched, rowPanelShapes, rowNameShapes, rowNoShapes, rowNoFrontShapes, rowNoBackShapes, _
                failRows, failMarkShapes, failN, warnN)

    If QTN_ENFORCE_ROW_METADATA And metaPanelN = 0 Then
        failN = failN + 1
        detail = "FAIL - Metadata Auto Mass Nesting tidak ditemukan pada selection." & vbCrLf & _
                 "QC Typo Nest membutuhkan Shape.Name dengan prefix HADES_AMN|ROW=...|SIZE=..." & vbCrLf & vbCrLf & detail
    End If

    If failN = 0 Then
        status = "PASS"
    Else
        status = "FAIL"
    End If

    summary = "Rows=" & CStr(ordN) & _
              ", TextScanned=" & CStr(scannedTextN) & _
              ", AMNPanelMeta=" & CStr(metaPanelN) & _
              ", Fail=" & CStr(failN) & _
              ", Warn=" & CStr(warnN)

    QTN_WriteReport reportPath, status, summary, detail, orderPath, tplPath, dbPath

    markerN = 0
    If failN > 0 And Not qtnReportMode Then
        markerN = QTN_MarkFailedRowsGreen(failRows, failMarkShapes)
    End If

    qtnLastStatus = status
    qtnLastSummary = summary
    qtnLastDetail = detail

    ActiveDocument.Unit = oldUnit

    If Not qtnReportMode Then
        QTN_ShowFinalPopup status, summary, detail, reportPath, markerN
    End If

    Exit Sub

ERR_HANDLER:
    On Error Resume Next
    ActiveDocument.Unit = oldUnit
    qtnLastStatus = "FAIL"
    qtnLastSummary = "QC Typo Nest runtime error."
    qtnLastDetail = "Error " & Err.Number & ": " & Err.Description

    If Not qtnReportMode Then
        MsgBox "QC TYPO NEST ERROR" & vbCrLf & vbCrLf & _
               "Error " & Err.Number & ": " & Err.Description, _
               vbCritical, "HADES QC TYPO NEST"
    End If

End Sub

Public Sub QTN_NEST_REPORT()

    On Error GoTo ERR_HANDLER

    qtnReportMode = True
    qtnLastStatus = ""
    qtnLastSummary = ""
    qtnLastDetail = ""

    QTN_NEST_CHECK

    If Trim$(qtnLastStatus) = "" Then
        HADESR_AddResult "TYPO CHECK (NEST)", "FAIL", _
                         "QC Typo Nest tidak menghasilkan status.", _
                         "Kemungkinan Order.txt, TypoTemplate, SizeDB, atau selection gagal sebelum report terbentuk."
    Else
        HADESR_AddResult "TYPO CHECK (NEST)", qtnLastStatus, qtnLastSummary, qtnLastDetail
    End If

SAFE_EXIT:
    qtnReportMode = False
    Exit Sub

ERR_HANDLER:
    HADESR_AddResult "TYPO CHECK (NEST)", "FAIL", _
                     "QC Typo Nest report mode error.", _
                     "Error " & Err.Number & ": " & Err.Description
    Resume SAFE_EXIT

End Sub

'=========================================================
' SCAN ENGINE
'=========================================================
Private Sub QTN_ScanShape( _
    ByVal shp As Shape, _
    ByVal ctxRow As Long, _
    ByVal ctxSize As String, _
    ByVal ctxPanel As String, _
    ByVal ctxPanelW As Double, _
    ByVal ctxPanelH As Double, _
    ByVal ctxHostPanel As Shape, _
    ByVal db As Object, _
    ByVal foundName As Object, _
    ByVal foundNo As Object, _
    ByVal foundNick As Object, _
    ByVal foundNoFront As Object, _
    ByVal foundNoBack As Object, _
    ByVal foundNoUnknownSide As Object, _
    ByVal foundUnknown As Object, _
    ByVal rowTouched As Object, _
    ByVal rowPanelShapes As Object, _
    ByVal rowNameShapes As Object, _
    ByVal rowNoShapes As Object, _
    ByVal rowNoFrontShapes As Object, _
    ByVal rowNoBackShapes As Object, _
    ByRef scannedTextN As Long, _
    ByRef metaPanelN As Long)

    On Error Resume Next

    Dim row2 As Long
    Dim size2 As String
    Dim panel2 As String
    Dim panelW2 As Double
    Dim panelH2 As Double
    Dim uid2 As String
    Dim c As Shape
    Dim pcShapes As Shapes
    Dim host2 As Shape

    row2 = ctxRow
    size2 = ctxSize
    panel2 = ctxPanel
    panelW2 = ctxPanelW
    panelH2 = ctxPanelH
    Set host2 = ctxHostPanel

    If QTN_ParseAMNMeta(QTN_SafeName(shp), row2, size2, panel2, uid2) Then
        metaPanelN = metaPanelN + 1
        size2 = QTN_NormalizeSize(size2)
        panel2 = UCase$(Trim$(panel2))
        panelW2 = shp.SizeWidth
        panelH2 = shp.SizeHeight
        If row2 > 0 Then QTN_AddRowShape rowPanelShapes, CStr(row2), shp
        Set host2 = shp
    End If

    If shp.Type = cdrTextShape Then
        QTN_ProcessText shp, row2, size2, panel2, panelW2, panelH2, host2, db, foundName, foundNo, foundNick, _
                        foundNoFront, foundNoBack, foundNoUnknownSide, foundUnknown, _
                        rowTouched, rowNameShapes, rowNoShapes, rowNoFrontShapes, rowNoBackShapes, scannedTextN
        Exit Sub
    End If

    If shp.Type = cdrGroupShape Then
        For Each c In shp.Shapes
            QTN_ScanShape c, row2, size2, panel2, panelW2, panelH2, host2, db, foundName, foundNo, foundNick, _
                          foundNoFront, foundNoBack, foundNoUnknownSide, foundUnknown, _
                          rowTouched, rowPanelShapes, rowNameShapes, rowNoShapes, rowNoFrontShapes, rowNoBackShapes, _
                          scannedTextN, metaPanelN
        Next c
    End If

    Set pcShapes = shp.PowerClip.Shapes
    If Not pcShapes Is Nothing Then
        For Each c In pcShapes
            QTN_ScanShape c, row2, size2, panel2, panelW2, panelH2, host2, db, foundName, foundNo, foundNick, _
                          foundNoFront, foundNoBack, foundNoUnknownSide, foundUnknown, _
                          rowTouched, rowPanelShapes, rowNameShapes, rowNoShapes, rowNoFrontShapes, rowNoBackShapes, _
                          scannedTextN, metaPanelN
        Next c
    End If

End Sub

Private Sub QTN_ProcessText( _
    ByVal shpText As Shape, _
    ByVal rowNo As Long, _
    ByVal rowSize As String, _
    ByVal panelHint As String, _
    ByVal panelW As Double, _
    ByVal panelH As Double, _
    ByVal hostPanel As Shape, _
    ByVal db As Object, _
    ByVal foundName As Object, _
    ByVal foundNo As Object, _
    ByVal foundNick As Object, _
    ByVal foundNoFront As Object, _
    ByVal foundNoBack As Object, _
    ByVal foundNoUnknownSide As Object, _
    ByVal foundUnknown As Object, _
    ByVal rowTouched As Object, _
    ByVal rowNameShapes As Object, _
    ByVal rowNoShapes As Object, _
    ByVal rowNoFrontShapes As Object, _
    ByVal rowNoBackShapes As Object, _
    ByRef scannedTextN As Long)

    On Error GoTo SAFE_EXIT

    Dim raw As String
    Dim txt As String
    Dim h As Double
    Dim side As String

    If rowNo <= 0 Then Exit Sub

    raw = shpText.Text.Story.Text
    txt = QTN_NormalizeText(raw)
    h = shpText.SizeHeight

    If Len(txt) = 0 Then Exit Sub
    If h < QTN_MIN_TEXT_H Then Exit Sub
    If QTN_IsMarkerText(txt) Then Exit Sub
    If QTN_IsSmallIDPO(txt, h) Then Exit Sub

    scannedTextN = scannedTextN + 1
    QTN_AddString rowTouched, CStr(rowNo), "1"

    If QTN_IsNumberText(txt) Then
        QTN_AddString foundNo, CStr(rowNo), txt
        If Not hostPanel Is Nothing Then QTN_AddRowShape rowNoShapes, CStr(rowNo), hostPanel

        side = QTN_DetectSideFromPanelOrDims(panelHint, rowSize, panelW, panelH, db)
        If side = "FRONT" Then
            QTN_AddString foundNoFront, CStr(rowNo), txt
            If Not hostPanel Is Nothing Then QTN_AddRowShape rowNoFrontShapes, CStr(rowNo), hostPanel
        ElseIf side = "BACK" Then
            QTN_AddString foundNoBack, CStr(rowNo), txt
            If Not hostPanel Is Nothing Then QTN_AddRowShape rowNoBackShapes, CStr(rowNo), hostPanel
        Else
            QTN_AddString foundNoUnknownSide, CStr(rowNo), txt
        End If
        Exit Sub
    End If

    'Klasifikasi nama/nickname sengaja sederhana dan exact-content.
    'Validasi final akan membandingkan against Order.txt per ROW.
    QTN_AddString foundName, CStr(rowNo), txt
    If Not hostPanel Is Nothing Then QTN_AddRowShape rowNameShapes, CStr(rowNo), hostPanel

SAFE_EXIT:
End Sub

'=========================================================
' VALIDATION ENGINE
'=========================================================
Private Function QTN_ValidateRows( _
    ByVal ordN As Long, _
    ByRef ordSize() As String, _
    ByRef ordName() As String, _
    ByRef ordNo() As String, _
    ByRef ordNick() As String, _
    ByVal tpl As Object, _
    ByVal db As Object, _
    ByVal foundName As Object, _
    ByVal foundNo As Object, _
    ByVal foundNick As Object, _
    ByVal foundNoFront As Object, _
    ByVal foundNoBack As Object, _
    ByVal foundNoUnknownSide As Object, _
    ByVal foundUnknown As Object, _
    ByVal rowTouched As Object, _
    ByVal rowPanelShapes As Object, _
    ByVal rowNameShapes As Object, _
    ByVal rowNoShapes As Object, _
    ByVal rowNoFrontShapes As Object, _
    ByVal rowNoBackShapes As Object, _
    ByVal failRows As Object, _
    ByVal failMarkShapes As Object, _
    ByRef failN As Long, _
    ByRef warnN As Long) As String

    Dim i As Long
    Dim detail As String
    Dim expNameSlots As Long
    Dim expNickSlots As Long
    Dim expNumSlots As Long
    Dim expNumFront As Long
    Dim expNumBack As Long
    Dim expNumUnknown As Long

    expNameSlots = QTN_TemplateNameSlotCount(tpl)
    expNickSlots = QTN_TemplateRoleCount(tpl, "NICKNAME")
    expNumSlots = QTN_TemplateRoleCount(tpl, "NUMBER")
    expNumFront = QTN_TemplateNumberSideCount(tpl, "FRONT")
    expNumBack = QTN_TemplateNumberSideCount(tpl, "BACK")
    expNumUnknown = expNumSlots - expNumFront - expNumBack

    If expNumSlots = 0 Then expNumSlots = 1 'Fallback safety jika template lama tidak punya NUMBER_COUNT.

    detail = "QC TYPO NEST V1.3" & vbCrLf & _
             "Template expected slots: NAME=" & CStr(expNameSlots) & _
             ", NICK=" & CStr(expNickSlots) & _
             ", NUMBER=" & CStr(expNumSlots) & _
             " (FRONT=" & CStr(expNumFront) & ", BACK=" & CStr(expNumBack) & ")" & vbCrLf & vbCrLf

    For i = 1 To ordN

        Dim eName As String
        Dim eNo As String
        Dim eNick As String
        Dim fNames As String
        Dim fNos As String
        Dim fNicks As String
        Dim fNoFront As String
        Dim fNoBack As String
        Dim fNoUnk As String
        Dim rowFail As String
        Dim rowWarn As String
        Dim rowHasIdentity As Boolean

        eName = QTN_NormalizeText(ordName(i))
        eNo = QTN_NormalizeText(ordNo(i))
        eNick = QTN_NormalizeText(ordNick(i))

        rowHasIdentity = (eName <> "" Or eNo <> "" Or eNick <> "")
        If Not rowHasIdentity Then GoTo NextRow

        fNames = QTN_GetString(foundName, CStr(i))
        fNos = QTN_GetString(foundNo, CStr(i))
        fNicks = QTN_GetString(foundNick, CStr(i))
        fNoFront = QTN_GetString(foundNoFront, CStr(i))
        fNoBack = QTN_GetString(foundNoBack, CStr(i))
        fNoUnk = QTN_GetString(foundNoUnknownSide, CStr(i))

        rowFail = ""
        rowWarn = ""

        If Not rowTouched.Exists(CStr(i)) Then
            rowFail = rowFail & "- Row tidak ditemukan pada hasil nesting / tidak punya metadata HADES_AMN." & vbCrLf
            QTN_CopyRowShapes failMarkShapes, rowPanelShapes, CStr(i)
        End If

        If eName <> "" Then
            If QTN_CountExactInList(fNames, eName) < 1 Then
                rowFail = rowFail & "- Nama expected tidak ditemukan: " & eName & vbCrLf
                QTN_CopyRowShapesPrefer failMarkShapes, rowNameShapes, rowPanelShapes, CStr(i)
            End If
        End If

        If eNick <> "" And expNickSlots > 0 Then
            If QTN_CountExactInList(fNames, eNick) < 1 And QTN_CountExactInList(fNicks, eNick) < 1 Then
                rowFail = rowFail & "- Nickname expected tidak ditemukan: " & eNick & vbCrLf
                QTN_CopyRowShapesPrefer failMarkShapes, rowNameShapes, rowPanelShapes, CStr(i)
            End If
        End If

        If eNo <> "" Then
            Dim cntNo As Long
            Dim cntWrongNo As Long
            Dim wrongNos As String
            Dim cntFront As Long
            Dim cntBack As Long

            cntNo = QTN_CountExactInList(fNos, eNo)
            wrongNos = QTN_ListValuesNotEqual(fNos, eNo, cntWrongNo)

            If cntNo <> expNumSlots Then
                rowFail = rowFail & "- Jumlah nomor tidak sesuai. Expected " & eNo & " x" & CStr(expNumSlots) & _
                          ", Found exact x" & CStr(cntNo) & ". Semua nomor found: [" & QTN_DisplayList(fNos) & "]" & vbCrLf
                QTN_CopyRowShapesPrefer failMarkShapes, rowNoShapes, rowPanelShapes, CStr(i)
            End If

            If cntWrongNo > 0 Then
                rowFail = rowFail & "- Nomor tidak seragam / salah. Expected semua = " & eNo & _
                          ", tetapi ditemukan: [" & QTN_DisplayList(wrongNos) & "]" & vbCrLf & _
                          QTN_ExplainNumberOwners(wrongNos, ordN, ordSize, ordName, ordNo, i)
                QTN_CopyRowShapesPrefer failMarkShapes, rowNoShapes, rowPanelShapes, CStr(i)
            End If

            If expNumFront > 0 Then
                cntFront = QTN_CountExactInList(fNoFront, eNo)
                If cntFront <> expNumFront Then
                    rowFail = rowFail & "- Nomor depan/BODY_FRONT tidak sesuai. Expected x" & CStr(expNumFront) & _
                              ", Found exact x" & CStr(cntFront) & ". Found front: [" & QTN_DisplayList(fNoFront) & "]" & vbCrLf
                    QTN_CopyRowShapesPrefer failMarkShapes, rowNoFrontShapes, rowNoShapes, CStr(i)
                End If
            End If

            If expNumBack > 0 Then
                cntBack = QTN_CountExactInList(fNoBack, eNo)
                If cntBack <> expNumBack Then
                    rowFail = rowFail & "- Nomor belakang/BODY_BACK tidak sesuai. Expected x" & CStr(expNumBack) & _
                              ", Found exact x" & CStr(cntBack) & ". Found back: [" & QTN_DisplayList(fNoBack) & "]" & vbCrLf
                    QTN_CopyRowShapesPrefer failMarkShapes, rowNoBackShapes, rowNoShapes, CStr(i)
                End If
            End If

            If Len(Trim$(fNoUnk)) > 0 Then
                rowWarn = rowWarn & "- Ada nomor dengan sisi panel tidak diketahui: [" & QTN_DisplayList(fNoUnk) & "]" & vbCrLf
            End If
        End If

        If rowFail <> "" Then
            failN = failN + 1
            If Not failRows Is Nothing Then
                If Not failRows.Exists(CStr(i)) Then failRows.Add CStr(i), "1"
            End If
            detail = detail & "FAIL ROW " & CStr(i) & " | " & ordSize(i) & " | " & _
                     QTN_ShowField(ordName(i)) & " | " & QTN_ShowField(ordNo(i)) & " | " & QTN_ShowField(ordNick(i)) & vbCrLf & _
                     "Found names/text : [" & QTN_DisplayList(fNames) & "]" & vbCrLf & _
                     "Found numbers    : [" & QTN_DisplayList(fNos) & "]" & vbCrLf & _
                     rowFail
            If rowWarn <> "" Then
                warnN = warnN + 1
                detail = detail & "Warning:" & vbCrLf & rowWarn
            End If
            detail = detail & vbCrLf
        ElseIf rowWarn <> "" Then
            warnN = warnN + 1
            detail = detail & "WARN ROW " & CStr(i) & " | " & ordSize(i) & " | " & _
                     QTN_ShowField(ordName(i)) & " | " & QTN_ShowField(ordNo(i)) & vbCrLf & rowWarn & vbCrLf
        End If

NextRow:
    Next i

    If failN = 0 And warnN = 0 Then
        detail = detail & "Semua row beridentitas PASS." & vbCrLf
    ElseIf failN = 0 Then
        detail = detail & "Tidak ada FAIL, tetapi ada warning yang perlu dicek." & vbCrLf
    End If

    QTN_ValidateRows = detail

End Function

'=========================================================
' ORDER / TEMPLATE / SIZEDB LOADER
'=========================================================
Private Function QTN_LoadOrder( _
    ByVal path As String, _
    ByRef n As Long, _
    ByRef sz() As String, _
    ByRef nm() As String, _
    ByRef no() As String, _
    ByRef nick() As String) As Boolean

    On Error GoTo FAIL

    Dim f As Integer
    Dim ln As String
    Dim p() As String

    n = 0
    f = FreeFile
    Open path For Input As #f

    Do While Not EOF(f)
        Line Input #f, ln
        ln = Trim$(ln)

        If ln <> "" And Left$(ln, 1) <> "@" Then
            p = Split(ln, "|")
            If UBound(p) >= 0 Then
                n = n + 1
                If n = 1 Then
                    ReDim sz(1 To n)
                    ReDim nm(1 To n)
                    ReDim no(1 To n)
                    ReDim nick(1 To n)
                Else
                    ReDim Preserve sz(1 To n)
                    ReDim Preserve nm(1 To n)
                    ReDim Preserve no(1 To n)
                    ReDim Preserve nick(1 To n)
                End If

                sz(n) = ""
                nm(n) = ""
                no(n) = ""
                nick(n) = ""

                If UBound(p) >= 0 Then sz(n) = QTN_NormalizeSize(CStr(p(0)))
                If UBound(p) >= 1 Then nm(n) = QTN_RemoveLigatures(Trim$(CStr(p(1))))
                If UBound(p) >= 2 Then no(n) = QTN_RemoveLigatures(Trim$(CStr(p(2))))
                If UBound(p) >= 3 Then nick(n) = QTN_RemoveLigatures(Trim$(CStr(p(3))))
            End If
        End If
    Loop

    Close #f
    QTN_LoadOrder = True
    Exit Function

FAIL:
    On Error Resume Next
    Close #f
    QTN_LoadOrder = False
End Function

Private Function QTN_LoadKeyValueFile(ByVal path As String) As Object

    On Error GoTo FAIL

    Dim d As Object
    Dim f As Integer
    Dim ln As String
    Dim p As Long
    Dim k As String
    Dim v As String

    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1

    f = FreeFile
    Open path For Input As #f

    Do While Not EOF(f)
        Line Input #f, ln
        ln = Trim$(ln)
        If ln <> "" Then
            p = InStr(1, ln, "=", vbTextCompare)
            If p > 1 Then
                k = UCase$(Trim$(Left$(ln, p - 1)))
                v = Trim$(Mid$(ln, p + 1))
                If d.Exists(k) Then
                    d(k) = v
                Else
                    d.Add k, v
                End If
            End If
        End If
    Loop

    Close #f
    Set QTN_LoadKeyValueFile = d
    Exit Function

FAIL:
    On Error Resume Next
    Close #f
    Set QTN_LoadKeyValueFile = Nothing
End Function

Private Function QTN_LoadSizeDB(ByVal path As String) As Object

    On Error GoTo FAIL

    Dim d As Object
    Dim f As Integer
    Dim ln As String
    Dim p() As String
    Dim s As String

    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1

    f = FreeFile
    Open path For Input As #f

    Do While Not EOF(f)
        Line Input #f, ln
        ln = Trim$(ln)
        If ln <> "" And Left$(ln, 1) <> "@" Then
            p = Split(ln, "|")
            If UBound(p) >= 3 Then
                s = QTN_NormalizeSize(p(0))

                'Regular jersey: SIZE|LEBAR|TINGGI_DEPAN|TINGGI_BELAKANG
                If UBound(p) = 3 Then
                    d(s & "|W") = QTN_Val(p(1))
                    d(s & "|F") = QTN_Val(p(2))
                    d(s & "|B") = QTN_Val(p(3))
                Else
                    'Split front: SIZE|L_BACK|L_FRONT|T_FRONT|T_BACK
                    d(s & "|W") = QTN_Val(p(1))
                    d(s & "|F") = QTN_Val(p(3))
                    d(s & "|B") = QTN_Val(p(4))
                End If
            End If
        End If
    Loop

    Close #f
    Set QTN_LoadSizeDB = d
    Exit Function

FAIL:
    On Error Resume Next
    Close #f
    Set QTN_LoadSizeDB = Nothing
End Function

Private Function QTN_ReadOrderMeta(ByVal orderPath As String, ByVal keyName As String) As String

    On Error GoTo FAIL

    Dim f As Integer
    Dim ln As String
    Dim p As Long
    Dim k As String
    Dim v As String

    keyName = UCase$(Trim$(keyName))

    f = FreeFile
    Open orderPath For Input As #f

    Do While Not EOF(f)
        Line Input #f, ln
        ln = Trim$(ln)
        If Left$(ln, 1) = "@" Then
            p = InStr(1, ln, "=", vbTextCompare)
            If p > 2 Then
                k = UCase$(Trim$(Mid$(ln, 2, p - 2)))
                v = Trim$(Mid$(ln, p + 1))
                If k = keyName Then
                    QTN_ReadOrderMeta = v
                    Close #f
                    Exit Function
                End If
            End If
        End If
    Loop

    Close #f
    Exit Function

FAIL:
    On Error Resume Next
    Close #f
End Function

'=========================================================
' TEMPLATE SLOT HELPERS
'=========================================================
Private Function QTN_TemplateRoleCount(ByVal tpl As Object, ByVal role As String) As Long
    role = UCase$(Trim$(role))
    If tpl.Exists(role & "_COUNT") Then
        QTN_TemplateRoleCount = CLng(Val(CStr(tpl(role & "_COUNT"))))
    End If
End Function

Private Function QTN_TemplateNameSlotCount(ByVal tpl As Object) As Long
    QTN_TemplateNameSlotCount = QTN_TemplateRoleCount(tpl, "NAMA_ATLIT") + QTN_TemplateRoleCount(tpl, "NAMA")
    If QTN_TemplateNameSlotCount = 0 Then QTN_TemplateNameSlotCount = 1
End Function

Private Function QTN_TemplateNumberSideCount(ByVal tpl As Object, ByVal sideName As String) As Long

    Dim n As Long
    Dim i As Long
    Dim h As Double
    Dim label As String
    Dim side As String

    sideName = UCase$(Trim$(sideName))
    n = QTN_TemplateRoleCount(tpl, "NUMBER")

    For i = 1 To n
        h = QTN_TplDbl(tpl, "NUMBER_" & CStr(i) & "_H")
        label = ""
        If tpl.Exists("NUMBER_" & CStr(i) & "_LABEL") Then label = UCase$(CStr(tpl("NUMBER_" & CStr(i) & "_LABEL")))

        side = ""
        If InStr(1, label, "PUNGGUNG", vbTextCompare) > 0 Or h >= QTN_NUMBER_BACK_H Then
            side = "BACK"
        ElseIf InStr(1, label, "DADA", vbTextCompare) > 0 Or h >= QTN_NUMBER_FRONT_H Then
            side = "FRONT"
        End If

        If side = sideName Then QTN_TemplateNumberSideCount = QTN_TemplateNumberSideCount + 1
    Next i

End Function

Private Function QTN_TplDbl(ByVal tpl As Object, ByVal key As String) As Double
    key = UCase$(Trim$(key))
    If tpl.Exists(key) Then QTN_TplDbl = QTN_Val(CStr(tpl(key)))
End Function

'=========================================================
' METADATA / SIDE DETECTION
'=========================================================
Private Function QTN_ParseAMNMeta( _
    ByVal nm As String, _
    ByRef rowNo As Long, _
    ByRef sizeName As String, _
    ByRef panelName As String, _
    ByRef uidName As String) As Boolean

    On Error GoTo FAIL

    Dim p() As String
    Dim i As Long
    Dim kv() As String
    Dim k As String
    Dim v As String

    If Left$(UCase$(Trim$(nm)), Len(QTN_META_PREFIX)) <> QTN_META_PREFIX Then Exit Function

    p = Split(nm, "|")

    For i = LBound(p) To UBound(p)
        If InStr(1, p(i), "=", vbTextCompare) > 0 Then
            kv = Split(p(i), "=", 2)
            k = UCase$(Trim$(kv(0)))
            v = Trim$(kv(1))
            Select Case k
                Case "ROW", "ORDERROW", "REC", "RECIDX"
                    rowNo = CLng(Val(v))
                Case "SIZE", "SZ"
                    sizeName = QTN_NormalizeSize(v)
                Case "PANEL", "BUCKET", "SIDE"
                    panelName = UCase$(Trim$(v))
                Case "UID", "CATUID", "PANEL_UID"
                    uidName = v
            End Select
        End If
    Next i

    QTN_ParseAMNMeta = (rowNo > 0)
    Exit Function

FAIL:
    QTN_ParseAMNMeta = False
End Function

Private Function QTN_DetectSideFromPanelOrDims( _
    ByVal panelHint As String, _
    ByVal sizeName As String, _
    ByVal panelW As Double, _
    ByVal panelH As Double, _
    ByVal db As Object) As String

    On Error Resume Next

    Dim p As String
    p = UCase$(Trim$(panelHint))

    If InStr(1, p, "FRONT", vbTextCompare) > 0 Or InStr(1, p, "DEPAN", vbTextCompare) > 0 Then
        QTN_DetectSideFromPanelOrDims = "FRONT"
        Exit Function
    End If

    If InStr(1, p, "BACK", vbTextCompare) > 0 Or InStr(1, p, "BELAKANG", vbTextCompare) > 0 Then
        QTN_DetectSideFromPanelOrDims = "BACK"
        Exit Function
    End If

    If InStr(1, p, "BODY", vbTextCompare) = 0 Then Exit Function

    If db Is Nothing Then Exit Function
    If Not db.Exists(sizeName & "|F") Or Not db.Exists(sizeName & "|B") Then Exit Function

    Dim h As Double
    Dim fH As Double
    Dim bH As Double
    Dim dF As Double
    Dim dB As Double

    If panelW <= 0 Or panelH <= 0 Then Exit Function

    h = panelH
    If panelW > h Then h = panelW

    fH = CDbl(db(sizeName & "|F"))
    bH = CDbl(db(sizeName & "|B"))

    dF = Abs(h - fH)
    dB = Abs(h - bH)

    If dF <= QTN_SIDE_TOL Or dB <= QTN_SIDE_TOL Then
        If dF <= dB Then
            QTN_DetectSideFromPanelOrDims = "FRONT"
        Else
            QTN_DetectSideFromPanelOrDims = "BACK"
        End If
    End If

End Function


'=========================================================
' GREEN MARKER / UNDO-SAFE OUTLINE
'=========================================================
Private Sub QTN_AddRowShape(ByVal d As Object, ByVal rowKey As String, ByVal shp As Shape)
    On Error Resume Next

    Dim col As Collection
    rowKey = CStr(rowKey)

    If d.Exists(rowKey) Then
        Set col = d(rowKey)
    Else
        Set col = New Collection
        d.Add rowKey, col
    End If

    col.Add shp
End Sub

Private Sub QTN_CopyRowShapes(ByVal dest As Object, ByVal src As Object, ByVal rowKey As String)
    On Error Resume Next

    Dim col As Collection
    Dim shp As Shape

    rowKey = CStr(rowKey)
    If dest Is Nothing Then Exit Sub
    If src Is Nothing Then Exit Sub
    If Not src.Exists(rowKey) Then Exit Sub

    Set col = src(rowKey)
    For Each shp In col
        QTN_AddRowShape dest, rowKey, shp
    Next shp
End Sub

Private Sub QTN_CopyRowShapesPrefer(ByVal dest As Object, ByVal preferred As Object, ByVal fallback As Object, ByVal rowKey As String)
    On Error Resume Next

    rowKey = CStr(rowKey)
    If Not preferred Is Nothing Then
        If preferred.Exists(rowKey) Then
            QTN_CopyRowShapes dest, preferred, rowKey
            Exit Sub
        End If
    End If

    QTN_CopyRowShapes dest, fallback, rowKey
End Sub

Private Function QTN_MarkFailedRowsGreen(ByVal failRows As Object, ByVal rowPanelShapes As Object) As Long
    On Error GoTo SAFE_EXIT

    Dim k As Variant
    Dim col As Collection
    Dim shp As Shape
    Dim cnt As Long
    Dim commandStarted As Boolean

    If failRows Is Nothing Then Exit Function
    If rowPanelShapes Is Nothing Then Exit Function
    If failRows.Count = 0 Then Exit Function

    ActiveDocument.BeginCommandGroup "HADES QC Typo Nest Green Marker"
    commandStarted = True

    For Each k In failRows.Keys
        If rowPanelShapes.Exists(CStr(k)) Then
            Set col = rowPanelShapes(CStr(k))
            For Each shp In col
                QTN_MarkPanelOutlinesGreen shp, cnt
            Next shp
        End If
    Next k

SAFE_EXIT:
    On Error Resume Next
    If commandStarted Then ActiveDocument.EndCommandGroup
    QTN_MarkFailedRowsGreen = cnt
End Function

Private Sub QTN_MarkPanelOutlinesGreen(ByVal shp As Shape, ByRef cnt As Long)
    On Error Resume Next

    Dim c As Shape
    Dim pcShapes As Shapes

    If QTN_IsRedOutline(shp) Then
        QTN_MarkGreen shp
        cnt = cnt + 1
    End If

    If shp.Type = cdrGroupShape Then
        For Each c In shp.Shapes
            QTN_MarkPanelOutlinesGreen c, cnt
        Next c
    End If

    Set pcShapes = shp.PowerClip.Shapes
    If Not pcShapes Is Nothing Then
        For Each c In pcShapes
            QTN_MarkPanelOutlinesGreen c, cnt
        Next c
    End If
End Sub

Private Function QTN_IsRedOutline(ByVal s As Shape) As Boolean
    On Error GoTo SAFE_EXIT

    Dim r As Long
    Dim g As Long
    Dim b As Long

    If s.Outline.Width <= 0 Then Exit Function

    r = s.Outline.Color.RGBRed
    g = s.Outline.Color.RGBGreen
    b = s.Outline.Color.RGBBlue

    If r >= 230 And g <= 80 And b <= 80 Then QTN_IsRedOutline = True

SAFE_EXIT:
End Function

Private Sub QTN_MarkGreen(ByVal s As Shape)
    On Error Resume Next

    Dim c As Color
    Set c = CreateColor
    c.RGBAssign 0, 255, 0
    s.Outline.Color.CopyAssign c
End Sub

'=========================================================
' STRING DICTIONARY HELPERS
'=========================================================
Private Sub QTN_InitDict(ByVal d As Object)
    d.CompareMode = 1
End Sub

Private Sub QTN_AddString(ByVal d As Object, ByVal key As String, ByVal value As String)
    key = CStr(key)
    value = Trim$(value)
    If d.Exists(key) Then
        If value <> "" Then d(key) = CStr(d(key)) & "||" & value
    Else
        d.Add key, value
    End If
End Sub

Private Function QTN_GetString(ByVal d As Object, ByVal key As String) As String
    If d.Exists(key) Then QTN_GetString = CStr(d(key))
End Function

Private Function QTN_CountExactInList(ByVal listText As String, ByVal target As String) As Long

    Dim p() As String
    Dim i As Long
    target = QTN_NormalizeText(target)

    If Len(Trim$(listText)) = 0 Or target = "" Then Exit Function

    p = Split(listText, "||")
    For i = LBound(p) To UBound(p)
        If QTN_NormalizeText(p(i)) = target Then QTN_CountExactInList = QTN_CountExactInList + 1
    Next i

End Function

Private Function QTN_ListValuesNotEqual(ByVal listText As String, ByVal target As String, ByRef cntWrong As Long) As String

    Dim p() As String
    Dim i As Long
    Dim t As String
    Dim v As String

    cntWrong = 0
    target = QTN_NormalizeText(target)

    If Len(Trim$(listText)) = 0 Then Exit Function

    p = Split(listText, "||")
    For i = LBound(p) To UBound(p)
        v = QTN_NormalizeText(p(i))
        If v <> "" And v <> target Then
            cntWrong = cntWrong + 1
            If QTN_ListValuesNotEqual <> "" Then QTN_ListValuesNotEqual = QTN_ListValuesNotEqual & "||"
            QTN_ListValuesNotEqual = QTN_ListValuesNotEqual & v
        End If
    Next i

End Function

Private Function QTN_DisplayList(ByVal listText As String) As String
    If Trim$(listText) = "" Then
        QTN_DisplayList = "-"
    Else
        QTN_DisplayList = Replace(listText, "||", ", ")
    End If
End Function

'=========================================================
' SWAP / OWNER EXPLANATION
'=========================================================
Private Function QTN_ExplainNumberOwners( _
    ByVal wrongNos As String, _
    ByVal ordN As Long, _
    ByRef ordSize() As String, _
    ByRef ordName() As String, _
    ByRef ordNo() As String, _
    ByVal currentRow As Long) As String

    Dim p() As String
    Dim i As Long
    Dim j As Long
    Dim num As String
    Dim out As String

    If Trim$(wrongNos) = "" Then Exit Function

    p = Split(wrongNos, "||")

    For i = LBound(p) To UBound(p)
        num = QTN_NormalizeText(p(i))
        For j = 1 To ordN
            If j <> currentRow Then
                If QTN_NormalizeText(ordNo(j)) = num Then
                    out = out & "  * Nomor " & num & " adalah milik row " & CStr(j) & _
                          " (" & QTN_ShowField(ordName(j)) & ", size " & ordSize(j) & "). Kemungkinan tertukar." & vbCrLf
                End If
            End If
        Next j
    Next i

    QTN_ExplainNumberOwners = out

End Function

'=========================================================
' NORMALIZATION / FILTER
'=========================================================
Private Function QTN_NormalizeText(ByVal s As String) As String
    'V1.1 compatibility:
    'Auto Mass Nesting Ligature Breaker menyisipkan Zero Width Non-Joiner (U+200C)
    'di pasangan F+I/F+F/F+L. Di report ANSI karakter ini bisa terlihat sebagai "?"
    'sehingga ARIFIN terbaca ARIF?IN dan menyebabkan false fail.
    s = QTN_RemoveLigatures(s)
    s = QTN_RemoveVisualBreakers(s)
    s = Replace(s, vbCr, "")
    s = Replace(s, vbLf, "")
    s = Replace(s, Chr(160), " ")

    Do While InStr(1, s, "  ", vbTextCompare) > 0
        s = Replace(s, "  ", " ")
    Loop

    QTN_NormalizeText = UCase$(Trim$(s))
End Function

Private Function QTN_RemoveLigatures(ByVal s As String) As String
    On Error Resume Next
    s = Replace(s, ChrW(&HFB00), "FF")
    s = Replace(s, ChrW(&HFB01), "FI")
    s = Replace(s, ChrW(&HFB02), "FL")
    s = Replace(s, ChrW(&HFB03), "FFI")
    s = Replace(s, ChrW(&HFB04), "FFL")
    s = Replace(s, ChrW(&HFB05), "ST")
    s = Replace(s, ChrW(&HFB06), "ST")
    On Error GoTo 0
    QTN_RemoveLigatures = s
End Function

Private Function QTN_RemoveVisualBreakers(ByVal s As String) As String
    On Error Resume Next

    'Zero-width / invisible controls commonly used to break visual ligatures.
    s = Replace(s, ChrW(&H200B), "") ' Zero Width Space
    s = Replace(s, ChrW(&H200C), "") ' Zero Width Non-Joiner
    s = Replace(s, ChrW(&H200D), "") ' Zero Width Joiner
    s = Replace(s, ChrW(&H2060), "") ' Word Joiner
    s = Replace(s, ChrW(&HFEFF), "") ' BOM / Zero Width No-Break Space
    s = Replace(s, ChrW(&HFFFD), "") ' Replacement char if produced by encoding conversion

    On Error GoTo 0

    'Fallback jika ZWNJ sudah berubah menjadi ASCII "?" oleh Corel/VBA/report encoding.
    'Hanya hapus tanda ? yang berada di antara F + (F/I/L), agar tidak merusak nama valid.
    s = QTN_RemoveQuestionMarkLigatureBreakers(s)

    QTN_RemoveVisualBreakers = s
End Function

Private Function QTN_RemoveQuestionMarkLigatureBreakers(ByVal s As String) As String

    Dim i As Long
    Dim ch As String
    Dim prevCh As String
    Dim nextCh As String
    Dim out As String

    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)

        If ch = "?" And i > 1 And i < Len(s) Then
            prevCh = UCase$(Mid$(s, i - 1, 1))
            nextCh = UCase$(Mid$(s, i + 1, 1))

            If prevCh = "F" Then
                Select Case nextCh
                    Case "F", "I", "L"
                        GoTo SKIP_CHAR
                End Select
            End If
        End If

        out = out & ch

SKIP_CHAR:
    Next i

    QTN_RemoveQuestionMarkLigatureBreakers = out

End Function

Private Function QTN_NormalizeSize(ByVal s As String) As String
    s = UCase$(Trim$(s))
    s = Replace(s, " ", "")
    s = Replace(s, "-", "")
    QTN_NormalizeSize = s
End Function

Private Function QTN_IsNumberText(ByVal s As String) As Boolean
    s = Trim$(s)
    If Len(s) = 0 Then Exit Function
    If Len(s) > 3 Then Exit Function
    If IsNumeric(s) Then QTN_IsNumberText = True
End Function

Private Function QTN_IsSmallIDPO(ByVal txt As String, ByVal h As Double) As Boolean
    txt = QTN_NormalizeText(txt)
    If txt = "IDPO" Then
        QTN_IsSmallIDPO = (h >= QTN_ID_MIN_H And h <= QTN_ID_MAX_H)
        Exit Function
    End If
    If Len(txt) = 6 And IsNumeric(txt) Then
        QTN_IsSmallIDPO = True
    End If
End Function

Private Function QTN_IsMarkerText(ByVal txt As String) As Boolean
    txt = UCase$(Trim$(txt))
    If Left$(txt, 3) = "@A:" Then QTN_IsMarkerText = True
    If Left$(txt, 6) = "@ATTR:" Then QTN_IsMarkerText = True
    If Left$(txt, 12) = "HADES_MARKER" Then QTN_IsMarkerText = True
    If Left$(txt, 10) = "HADES_AMN" Then QTN_IsMarkerText = True
End Function

Private Function QTN_Val(ByVal s As String) As Double
    s = Trim$(s)
    s = Replace(s, ",", ".")
    QTN_Val = Val(s)
End Function

Private Function QTN_ShowField(ByVal s As String) As String
    s = Trim$(s)
    If s = "" Then
        QTN_ShowField = "-"
    Else
        QTN_ShowField = s
    End If
End Function

Private Function QTN_SafeName(ByVal shp As Shape) As String
    On Error Resume Next
    QTN_SafeName = shp.Name
End Function

Private Function QTN_DocumentsPath() As String
    QTN_DocumentsPath = Environ$("USERPROFILE") & "\Documents"
End Function

'=========================================================
' REPORT / ERROR OUTPUT
'=========================================================

Private Sub QTN_ShowFinalPopup( _
    ByVal status As String, _
    ByVal summary As String, _
    ByVal detail As String, _
    ByVal reportPath As String, _
    ByVal markerN As Long)

    On Error Resume Next

    Dim titleStatus As String
    Dim msg As String
    Dim icon As VbMsgBoxStyle

    If UCase$(status) = "PASS" Then
        titleStatus = "PASSED"
        icon = vbInformation
        msg = "QC TYPO NEST PASSED" & vbCrLf & vbCrLf & _
              summary & vbCrLf & vbCrLf & _
              "Semua row beridentitas PASS." & vbCrLf & vbCrLf & _
              "Report:" & vbCrLf & reportPath
    Else
        titleStatus = "FAIL"
        icon = vbCritical
        msg = "QC TYPO NEST FAIL" & vbCrLf & vbCrLf & _
              summary & vbCrLf & vbCrLf & _
              "Panel/outline pada row yang FAIL sudah diberi marker hijau." & vbCrLf & _
              "Jumlah outline hijau: " & CStr(markerN) & vbCrLf & _
              "Undo sekali untuk menghilangkan marker hijau." & vbCrLf & vbCrLf & _
              QTN_TrimPopupDetail(detail) & vbCrLf & vbCrLf & _
              "Report lengkap:" & vbCrLf & reportPath
    End If

    MsgBox msg, icon, "HADES QC TYPO NEST - " & titleStatus
End Sub

Private Function QTN_TrimPopupDetail(ByVal detail As String) As String
    On Error Resume Next

    detail = Trim$(detail)
    If Len(detail) <= QTN_POPUP_MAX_CHARS Then
        QTN_TrimPopupDetail = detail
    Else
        QTN_TrimPopupDetail = Left$(detail, QTN_POPUP_MAX_CHARS) & vbCrLf & _
                              "..." & vbCrLf & _
                              "Detail dipotong di popup. Lihat report TXT untuk detail lengkap."
    End If
End Function

Private Sub QTN_FailOut(ByVal summary As String, ByVal detail As String)
    qtnLastStatus = "FAIL"
    qtnLastSummary = summary
    qtnLastDetail = detail

    If Not qtnReportMode Then
        MsgBox summary & vbCrLf & vbCrLf & detail, vbCritical, "HADES QC TYPO NEST"
    End If
End Sub

Private Sub QTN_WriteReport( _
    ByVal path As String, _
    ByVal status As String, _
    ByVal summary As String, _
    ByVal detail As String, _
    ByVal orderPath As String, _
    ByVal tplPath As String, _
    ByVal dbPath As String)

    On Error Resume Next

    Dim f As Integer
    f = FreeFile
    Open path For Output As #f

    Print #f, "PROJECT H.A.D.E.S. - QC TYPO NEST REPORT"
    Print #f, "STATUS=" & status
    Print #f, "SUMMARY=" & summary
    Print #f, "ORDER=" & orderPath
    Print #f, "TEMPLATE=" & tplPath
    Print #f, "SIZEDB=" & dbPath
    Print #f, ""
    Print #f, detail

    Close #f
End Sub

