Attribute VB_Name = "SPC_Monitor"
' ============================================================
' SPC Calculator v13 - Monitor / folder-watch module
'
' Import:  Alt+F11 -> File -> Import File -> select this .bas
' Save the workbook as .xlsm (macro-enabled) before using it.
'
' Configuration lives on the Settings sheet:
'   F8  Watch Folder            (e.g. \\server\share\cmm-exports)
'   F9  File Pattern            (*.csv  or  *.xlsx  or  *.*)
'   F10 Auto-Refresh (sec)      minimum 10
'   F11 Archive Imported Files  Y / N
'   F12 Archive Subfolder       e.g. archive
'   F13 Kiosk Cycle (sec)
'
' Status is written back to the Monitor sheet (B9:B14).
' Every file already imported is recorded on the hidden _ImportLog sheet,
' so the same file is never counted twice.
' ============================================================
Option Explicit

Private Const SH_SET As String = "Settings"
Private Const SH_MON As String = "Monitor"
Private Const SH_DATA As String = "Data Entry"
Private Const SH_LOG As String = "_ImportLog"

' Monitor sheet status cells
Private Const C_STATUS As String = "B9"
Private Const C_LASTSCAN As String = "B10"
Private Const C_FILES As String = "B11"
Private Const C_ROWS As String = "B12"
Private Const C_LASTFILE As String = "B13"
Private Const C_MSG As String = "B14"

' Data Entry geometry
Private Const DATA_FIRST As Long = 5
Private Const DATA_LAST As Long = 10004
Private Const FEAT_COL As Long = 8        ' column H
Private Const MAX_FEAT As Long = 50

Public NextRun As Double
Public Monitoring As Boolean
Public KioskOn As Boolean
Private NextKiosk As Double
Private KioskSlot As Long

' ============================================================
' STARTUP - Excel runs Auto_Open by itself when the workbook opens,
' so the buttons rebuild themselves with no manual step.
' ============================================================
Sub Auto_Open()
    On Error Resume Next
    SPC_BuildButtons
    SPC_AutoScale
    On Error GoTo 0
End Sub

' ============================================================
' BUTTONS
' Drawn by macro rather than stored in the sheet, so they always point at
' the module that is actually loaded. Safe to re-run at any time.
' ============================================================
Sub SPC_BuildButtons()
    BuildPanel ThisWorkbook.Sheets("How To Use"), "E5", True
    BuildPanel ThisWorkbook.Sheets(SH_MON), "E5", False
End Sub

Private Sub BuildPanel(ws As Worksheet, anchor As String, full As Boolean)
    Const BW As Single = 150
    Const BH As Single = 26
    Const GAPY As Single = 5
    Const GAPX As Single = 12

    Dim shp As Shape, i As Long
    ' remove any buttons from a previous run
    For i = ws.Shapes.Count To 1 Step -1
        If Left$(ws.Shapes(i).Name, 4) = "btn_" Then ws.Shapes(i).Delete
    Next i

    Dim x0 As Single, y0 As Single, y As Single
    x0 = ws.Range(anchor).Left
    y0 = ws.Range(anchor).Top
    y = y0

    y = y + Header(ws, "ACTIONS", x0, y, BW)
    y = y + AddButton(ws, "Import Now", "SPC_ImportNow", x0, y, BW, BH, RGB(0, 130, 60)) + GAPY
    y = y + AddButton(ws, "Start Monitoring", "SPC_StartMonitor", x0, y, BW, BH, RGB(0, 130, 60)) + GAPY
    y = y + AddButton(ws, "Stop Monitoring", "SPC_StopMonitor", x0, y, BW, BH, RGB(170, 40, 40)) + GAPY
    y = y + AddButton(ws, "Refresh All Charts", "SPC_RefreshAll", x0, y, BW, BH, RGB(31, 78, 121)) + GAPY
    y = y + AddButton(ws, "Fit Charts To Data", "SPC_AutoScale", x0, y, BW, BH, RGB(31, 78, 121)) + GAPY
    y = y + AddButton(ws, "Screen / Kiosk Mode", "SPC_KioskMode", x0, y, BW, BH, RGB(112, 48, 160)) + GAPY
    y = y + AddButton(ws, "Stop Kiosk Mode", "SPC_KioskStop", x0, y, BW, BH, RGB(90, 90, 90)) + GAPY

    If full Then
        y = y + AddButton(ws, "Print Collection Form", "SPC_PrintForm", x0, y, BW, BH, RGB(31, 78, 121)) + GAPY
        y = y + AddButton(ws, "Export PDF Report", "SPC_ExportReportPDF", x0, y, BW, BH, RGB(31, 78, 121)) + GAPY
        y = y + AddButton(ws, "Clear All Data", "SPC_ClearData", x0, y, BW, BH, RGB(170, 40, 40)) + GAPY

        Dim x1 As Single
        x1 = x0 + BW + GAPX
        y = y0
        y = y + Header(ws, "GO TO SHEET", x1, y, BW)
        y = y + AddButton(ws, "Settings", "GoToSettings", x1, y, BW, BH, RGB(68, 84, 106)) + GAPY
        y = y + AddButton(ws, "Data Entry", "GoToDataEntry", x1, y, BW, BH, RGB(68, 84, 106)) + GAPY
        y = y + AddButton(ws, "Control Charts", "GoToControlCharts", x1, y, BW, BH, RGB(68, 84, 106)) + GAPY
        y = y + AddButton(ws, "Monitor", "GoToMonitor", x1, y, BW, BH, RGB(68, 84, 106)) + GAPY
        y = y + AddButton(ws, "Visual SPC", "GoToVisualSPC", x1, y, BW, BH, RGB(68, 84, 106)) + GAPY
        y = y + AddButton(ws, "Capability", "GoToCapability", x1, y, BW, BH, RGB(68, 84, 106)) + GAPY
        y = y + AddButton(ws, "Collection Form", "GoToForm", x1, y, BW, BH, RGB(68, 84, 106)) + GAPY
    End If
End Sub

Private Function Header(ws As Worksheet, caption As String, _
                        x As Single, y As Single, w As Single) As Single
    Dim shp As Shape
    Set shp = ws.Shapes.AddShape(msoShapeRectangle, x, y, w, 20)
    shp.Name = "btn_hdr_" & Replace(caption, " ", "_") & "_" & Int(x)
    shp.Fill.ForeColor.RGB = RGB(31, 56, 100)
    shp.Line.Visible = msoFalse
    With shp.TextFrame2
        .TextRange.Text = caption
        .TextRange.Font.Size = 9
        .TextRange.Font.Bold = msoTrue
        .TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .VerticalAnchor = msoAnchorMiddle
        .TextRange.ParagraphFormat.Alignment = msoAlignCenter
    End With
    Header = 20 + 6
End Function

Private Function AddButton(ws As Worksheet, caption As String, macro As String, _
                           x As Single, y As Single, w As Single, h As Single, _
                           colour As Long) As Single
    Dim shp As Shape
    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, x, y, w, h)
    shp.Name = "btn_" & macro
    shp.Fill.ForeColor.RGB = colour
    shp.Line.ForeColor.RGB = RGB(255, 255, 255)
    shp.Line.Weight = 1
    With shp.TextFrame2
        .TextRange.Text = caption
        .TextRange.Font.Size = 10
        .TextRange.Font.Bold = msoTrue
        .TextRange.Font.Name = "Arial"
        .TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .VerticalAnchor = msoAnchorMiddle
        .TextRange.ParagraphFormat.Alignment = msoAlignCenter
        .MarginLeft = 2
        .MarginRight = 2
    End With
    shp.OnAction = macro
    AddButton = h
End Function

' ============================================================
' PUBLIC ENTRY POINTS
' ============================================================
Sub SPC_StartMonitor()
    Dim folder As String
    folder = CfgStr("F8")

    If Len(Trim$(folder)) = 0 Then
        MsgBox "Set the watch folder in Settings cell F8 first.", vbExclamation, "SPC v13"
        Exit Sub
    End If
    If Not FolderExists(folder) Then
        MsgBox "Watch folder not found:" & vbCrLf & folder, vbExclamation, "SPC v13"
        Exit Sub
    End If

    Monitoring = True
    SetStatus "RUNNING", RGB(0, 160, 60)
    ScanFolder                      ' immediate first pass
    ScheduleNext
    MsgBox "Monitoring started." & vbCrLf & vbCrLf & _
           "Folder:   " & folder & vbCrLf & _
           "Pattern:  " & CfgStr("F9") & vbCrLf & _
           "Interval: " & CfgLng("F10", 60) & " s", vbInformation, "SPC v13"
End Sub

Sub SPC_StopMonitor()
    Monitoring = False
    CancelNext
    SetStatus "STOPPED", RGB(255, 80, 80)
End Sub

Sub SPC_ImportNow()
    ' One manual pass, whether or not the timer is running.
    ScanFolder
    SPC_RefreshAll
    MsgBox "Scan complete." & vbCrLf & Mon().Range(C_MSG).Value, vbInformation, "SPC v13"
End Sub

Sub SPC_RefreshAll()
    Application.CalculateFullRebuild
    RefreshCharts ThisWorkbook.Sheets("Control Charts")
    RefreshCharts Mon()
    RefreshCharts ThisWorkbook.Sheets("Visual SPC")
    SPC_AutoScale
    DoEvents
End Sub

' ============================================================
' AUTO-SCALE
' Excel's automatic value axis often anchors at zero. On a process running at,
' say, 7.43 +/- 0.01 that squashes every point into a flat line and you cannot
' see the spread at all. This rescales each X and mR chart to its own data plus
' the control and spec limits, with an 8% margin, so the variation fills the
' plot area.
'
' Histograms are skipped on purpose - a frequency axis belongs at zero.
' ============================================================
Sub SPC_AutoScale()
    AutoScaleSheet ThisWorkbook.Sheets("Control Charts")
    AutoScaleSheet Mon()
    AutoScaleSheet ThisWorkbook.Sheets("Visual SPC")
End Sub

Private Sub AutoScaleSheet(ws As Worksheet)
    Dim co As ChartObject
    On Error Resume Next
    For Each co In ws.ChartObjects
        AutoScaleChart co.Chart
    Next co
    On Error GoTo 0
End Sub

Private Sub AutoScaleChart(ch As Chart)
    Dim sc As Series, v As Variant, i As Long
    Dim lo As Double, hi As Double, pad As Double
    Dim got As Boolean, t As Long

    t = 0
    On Error Resume Next
    t = ch.ChartType
    On Error GoTo 0
    If t = xlColumnClustered Or t = xlColumnStacked Or _
       t = xlBarClustered Or t = xlBarStacked Then Exit Sub

    lo = 1E+308
    hi = -1E+308

    On Error Resume Next
    For Each sc In ch.SeriesCollection
        v = sc.Values
        If IsArray(v) Then
            For i = LBound(v) To UBound(v)
                If Not IsError(v(i)) Then
                    If Not IsEmpty(v(i)) Then
                        If IsNumeric(v(i)) Then
                            If CDbl(v(i)) < lo Then lo = CDbl(v(i))
                            If CDbl(v(i)) > hi Then hi = CDbl(v(i))
                            got = True
                        End If
                    End If
                End If
            Next i
        End If
    Next sc
    On Error GoTo 0

    If Not got Then Exit Sub

    If hi <= lo Then
        ' every point identical - still give it a visible band
        pad = Abs(hi) * 0.002
        If pad = 0 Then pad = 0.5
    Else
        pad = (hi - lo) * 0.08
    End If

    On Error Resume Next
    With ch.Axes(xlValue)
        .MinimumScaleIsAuto = False
        .MaximumScaleIsAuto = False
        .MinimumScale = lo - pad
        .MaximumScale = hi + pad
    End With
    On Error GoTo 0
End Sub

' ============================================================
' TIMER
' ============================================================
Private Sub ScheduleNext()
    Dim iv As Long
    iv = CfgLng("F10", 60)
    If iv < 10 Then iv = 10
    NextRun = Now + TimeSerial(0, 0, iv)
    Application.OnTime NextRun, "SPC_Tick"
End Sub

Private Sub CancelNext()
    On Error Resume Next
    Application.OnTime NextRun, "SPC_Tick", , False
    On Error GoTo 0
End Sub

Sub SPC_Tick()
    If Not Monitoring Then Exit Sub
    ScanFolder
    SPC_RefreshAll
    ScheduleNext
End Sub

' ============================================================
' FOLDER SCAN
' ============================================================
Private Sub ScanFolder()
    Dim folder As String, pattern As String, f As String
    Dim nFiles As Long, nRows As Long, added As Long
    Dim lastFile As String

    folder = AddSlash(CfgStr("F8"))
    pattern = CfgStr("F9")
    If Len(pattern) = 0 Then pattern = "*.csv"
    If Len(folder) = 0 Or Not FolderExists(folder) Then
        SetStatus "ERROR", RGB(255, 165, 0)
        Mon().Range(C_MSG).Value = "Watch folder missing"
        Exit Sub
    End If

    On Error GoTo ScanFail
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    ' Collect the whole file list BEFORE importing anything. Dir() keeps a
    ' single global enumeration, and ArchiveFile / FolderExists call Dir()
    ' themselves - doing that inside a "f = Dir" loop resets the walk, so
    ' files get skipped or visited twice.
    Dim names() As String, count As Long, k As Long
    ReDim names(0 To 511)
    f = Dir(folder & pattern)
    Do While Len(f) > 0
        If count > UBound(names) Then ReDim Preserve names(0 To UBound(names) + 512)
        names(count) = f
        count = count + 1
        f = Dir
    Loop

    For k = 0 To count - 1
        f = names(k)
        If Not AlreadyImported(f, FileDateTime(folder & f)) Then
            added = ImportFile(folder & f)
            If added > 0 Then
                LogImport f, FileDateTime(folder & f), added
                nFiles = nFiles + 1
                nRows = nRows + added
                lastFile = f
                If UCase$(CfgStr("F11")) = "Y" Then ArchiveFile folder, f
            ElseIf added = 0 Then
                ' nothing usable in it - log it so it is not retried forever
                LogImport f, FileDateTime(folder & f), 0
            End If
        End If
    Next k

    Mon().Range(C_LASTSCAN).Value = Format$(Now, "yyyy-mm-dd hh:nn:ss")
    Mon().Range(C_FILES).Value = nFiles
    Mon().Range(C_ROWS).Value = nRows
    If nFiles > 0 Then
        Mon().Range(C_LASTFILE).Value = lastFile
        Mon().Range(C_MSG).Value = nFiles & " file(s), " & nRows & " row(s) imported"
    Else
        Mon().Range(C_MSG).Value = "No new files"
    End If

    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub

ScanFail:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    SetStatus "ERROR", RGB(255, 165, 0)
    Mon().Range(C_MSG).Value = "Scan failed: " & Err.Description
End Sub

' ------------------------------------------------------------
' Import one file. Returns rows added, or -1 on failure.
'
' Accepted layouts (per row):
'   a) Date, Shift, Operator, ToolID, Cluster, F1, F2, ... Fn
'   b) F1, F2, ... Fn                       (first field numeric)
' A non-numeric first data row is treated as a header and skipped.
' ------------------------------------------------------------
Private Function ImportFile(path As String) As Long
    Dim ext As String
    ext = LCase$(Mid$(path, InStrRev(path, ".") + 1))
    Select Case ext
        Case "csv", "txt", "prn": ImportFile = ImportDelimited(path)
        Case "xlsx", "xlsm", "xls": ImportFile = ImportWorkbook(path)
        Case Else: ImportFile = -1
    End Select
End Function

Private Function ImportDelimited(path As String) As Long
    Dim ff As Integer, line As String, parts() As String
    Dim ws As Worksheet, r As Long, added As Long

    Set ws = ThisWorkbook.Sheets(SH_DATA)
    r = NextFreeRow(ws)
    If r > DATA_LAST Then
        ImportDelimited = -1                 ' Data Entry is full
        Exit Function
    End If

    ff = FreeFile
    Open path For Input As #ff
    Do While Not EOF(ff)
        Line Input #ff, line
        line = Trim$(line)
        If Len(line) > 0 Then
            ' tab-separated and semicolon-separated exports are common; treat
            ' both as comma-separated. Semicolons only when there is no comma,
            ' so a genuine comma file is never mangled.
            line = Replace(line, vbTab, ",")
            If InStr(line, ",") = 0 And InStr(line, ";") > 0 Then
                line = Replace(line, ";", ",")
            End If
            parts = Split(line, ",")
            If WriteRow(ws, r, parts) Then
                r = r + 1
                added = added + 1
                If r > DATA_LAST Then Exit Do
            End If
        End If
    Loop
    Close #ff
    ImportDelimited = added
End Function

Private Function ImportWorkbook(path As String) As Long
    Dim wbSrc As Workbook, wsSrc As Worksheet, ws As Worksheet
    Dim r As Long, sr As Long, lastSrc As Long, lastCol As Long
    Dim parts() As String, c As Long, added As Long

    Set ws = ThisWorkbook.Sheets(SH_DATA)
    r = NextFreeRow(ws)
    If r > DATA_LAST Then
        ImportWorkbook = -1                  ' Data Entry is full
        Exit Function
    End If

    On Error GoTo Fail
    Set wbSrc = Workbooks.Open(path, ReadOnly:=True, UpdateLinks:=0)
    Set wsSrc = wbSrc.Sheets(1)
    lastSrc = wsSrc.UsedRange.Row + wsSrc.UsedRange.Rows.Count - 1
    lastCol = wsSrc.UsedRange.Column + wsSrc.UsedRange.Columns.Count - 1

    For sr = 1 To lastSrc
        ReDim parts(0 To lastCol - 1)
        For c = 1 To lastCol
            parts(c - 1) = CStr(wsSrc.Cells(sr, c).Value)
        Next c
        If WriteRow(ws, r, parts) Then
            r = r + 1
            added = added + 1
            If r > DATA_LAST Then Exit For
        End If
    Next sr

    wbSrc.Close SaveChanges:=False
    ImportWorkbook = added
    Exit Function
Fail:
    On Error Resume Next
    If Not wbSrc Is Nothing Then wbSrc.Close SaveChanges:=False
    On Error GoTo 0
    ImportWorkbook = -1
End Function

' Writes one parsed row. Returns True if it was a data row.
Private Function WriteRow(ws As Worksheet, r As Long, parts() As String) As Boolean
    Dim i As Long, n As Long, firstFeat As Long, lo As Long
    lo = LBound(parts)
    n = UBound(parts) - lo + 1
    If n = 0 Then Exit Function

    ' strip stray spaces and surrounding quotes from every field first
    For i = lo To UBound(parts)
        parts(i) = CleanField(parts(i))
    Next i

    If IsNumeric(parts(lo)) Then
        firstFeat = lo                            ' layout (b): readings only
    Else
        ' layout (a): 5 metadata fields then readings.
        ' A header row has text where the first reading belongs, so it is skipped.
        If n < 6 Then Exit Function
        If Not IsNumeric(parts(lo + 5)) Then Exit Function
        ws.Cells(r, 3).Value = parts(lo)          ' Date
        ws.Cells(r, 4).Value = parts(lo + 1)      ' Shift
        ws.Cells(r, 5).Value = parts(lo + 2)      ' Operator
        ws.Cells(r, 6).Value = parts(lo + 3)      ' Tool ID
        ws.Cells(r, 7).Value = parts(lo + 4)      ' Cluster
        firstFeat = lo + 5
    End If

    Dim k As Long, written As Long
    k = 0
    For i = firstFeat To UBound(parts)
        If k >= MAX_FEAT Then Exit For
        If Len(parts(i)) > 0 Then
            If IsNumeric(parts(i)) Then
                ws.Cells(r, FEAT_COL + k).Value = CDbl(parts(i))
                written = written + 1
            End If
        End If
        k = k + 1
    Next i
    ' only count it as a data row if at least one real reading landed
    WriteRow = (written > 0)
End Function

Private Function CleanField(s As String) As String
    Dim t As String
    t = Trim$(s)
    If Len(t) >= 2 Then
        If Left$(t, 1) = """" And Right$(t, 1) = """" Then
            t = Mid$(t, 2, Len(t) - 2)
        End If
    End If
    CleanField = Trim$(t)
End Function

' ------------------------------------------------------------
' First row after ALL existing data.
'
' This must never come back too low: whatever it returns is written over.
' Two traps, both of which used to destroy previously collected data:
'   * looking only at column H - a file that carries no feature 1 reading
'     leaves that column empty, so the answer came back as row 5 and the
'     import overwrote everything already there;
'   * Range.End(xlUp) skips rows hidden by the AutoFilter on Data Entry,
'     so with a filter applied it lands somewhere in the middle.
' MATCH over each column ignores hidden rows and filters entirely.
' ------------------------------------------------------------
Private Function NextFreeRow(ws As Worksheet) As Long
    Dim c As Long, pos As Variant, last As Long

    last = DATA_FIRST - 1
    For c = FEAT_COL To FEAT_COL + MAX_FEAT - 1
        pos = Empty
        On Error Resume Next
        pos = Application.Match(1E+307, _
                  ws.Range(ws.Cells(DATA_FIRST, c), ws.Cells(DATA_LAST, c)), 1)
        On Error GoTo 0
        If IsNumeric(pos) And Not IsEmpty(pos) Then
            If CLng(pos) + DATA_FIRST - 1 > last Then last = CLng(pos) + DATA_FIRST - 1
        End If
    Next c

    NextFreeRow = FirstEmptyRowFrom(ws, last + 1)
End Function

' Belt and braces: never hand back a row that already holds a reading.
Private Function FirstEmptyRowFrom(ws As Worksheet, startRow As Long) As Long
    Dim r As Long
    r = startRow
    If r < DATA_FIRST Then r = DATA_FIRST
    Do While r <= DATA_LAST
        If Application.CountA(ws.Range(ws.Cells(r, FEAT_COL), _
                                       ws.Cells(r, FEAT_COL + MAX_FEAT - 1))) = 0 Then
            FirstEmptyRowFrom = r
            Exit Function
        End If
        r = r + 1
    Loop
    FirstEmptyRowFrom = DATA_LAST + 1        ' sheet full
End Function

' ============================================================
' IMPORT LOG (hidden sheet)
' ============================================================
Private Function LogSheet() As Worksheet
    On Error Resume Next
    Set LogSheet = ThisWorkbook.Sheets(SH_LOG)
    On Error GoTo 0
    If LogSheet Is Nothing Then
        Set LogSheet = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        LogSheet.Name = SH_LOG
        LogSheet.Range("A1:D1").Value = Array("File", "Modified", "Rows", "Imported at")
        LogSheet.Visible = xlSheetHidden
    End If
End Function

Private Function AlreadyImported(fileName As String, modified As Date) As Boolean
    Dim ws As Worksheet, last As Long, i As Long
    Set ws = LogSheet()
    last = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For i = 2 To last
        If StrComp(CStr(ws.Cells(i, 1).Value), fileName, vbTextCompare) = 0 Then
            If CDate(ws.Cells(i, 2).Value) >= modified Then
                AlreadyImported = True
                Exit Function
            End If
        End If
    Next i
End Function

Private Sub LogImport(fileName As String, modified As Date, rows As Long)
    Dim ws As Worksheet, r As Long
    Set ws = LogSheet()
    r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    ws.Cells(r, 1).Value = fileName
    ws.Cells(r, 2).Value = modified
    ws.Cells(r, 3).Value = rows
    ws.Cells(r, 4).Value = Now
End Sub

Private Sub ArchiveFile(folder As String, fileName As String)
    Dim sub_ As String, dest As String
    sub_ = CfgStr("F12")
    If Len(sub_) = 0 Then sub_ = "archive"
    dest = AddSlash(folder) & sub_
    On Error Resume Next
    If Not FolderExists(dest) Then MkDir dest
    Name AddSlash(folder) & fileName As AddSlash(dest) & fileName
    On Error GoTo 0
End Sub

' ============================================================
' KIOSK / SHARED-SCREEN MODE
' ============================================================
Sub SPC_KioskMode()
    KioskOn = True
    KioskSlot = 0
    Application.DisplayFullScreen = True
    ThisWorkbook.Windows(1).DisplayHeadings = False
    ThisWorkbook.Windows(1).DisplayWorkbookTabs = False
    Mon().Activate
    KioskStep
    MsgBox "Kiosk mode on. Run SPC_KioskStop to exit.", vbInformation, "SPC v13"
End Sub

Sub SPC_KioskStop()
    KioskOn = False
    On Error Resume Next
    Application.OnTime NextKiosk, "SPC_KioskTick", , False
    On Error GoTo 0
    Application.DisplayFullScreen = False
    ThisWorkbook.Windows(1).DisplayHeadings = True
    ThisWorkbook.Windows(1).DisplayWorkbookTabs = True
End Sub

Sub SPC_KioskTick()
    If Not KioskOn Then Exit Sub
    KioskSlot = (KioskSlot + 1) Mod 4
    KioskStep
End Sub

Private Sub KioskStep()
    Dim iv As Long, topRow As Long
    Mon().Activate
    topRow = 42 + KioskSlot * 32          ' first row of each slot band
    ActiveWindow.ScrollRow = topRow
    iv = CfgLng("F13", 20)
    If iv < 5 Then iv = 5
    NextKiosk = Now + TimeSerial(0, 0, iv)
    Application.OnTime NextKiosk, "SPC_KioskTick"
End Sub

' ============================================================
' HOUSEKEEPING
' ============================================================
Sub SPC_ClearData()
    If MsgBox("Delete every measurement on Data Entry?" & vbCrLf & _
              "This cannot be undone.", vbYesNo + vbExclamation, "SPC v13") <> vbYes Then Exit Sub
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SH_DATA)
    ws.Range(ws.Cells(DATA_FIRST, 3), ws.Cells(DATA_LAST, FEAT_COL + MAX_FEAT - 1)).ClearContents
    On Error Resume Next
    LogSheet().Range("A2:D100000").ClearContents
    On Error GoTo 0
    SPC_RefreshAll
End Sub

Sub SPC_PrintForm()
    ThisWorkbook.Sheets("Data Collection Form").PrintOut Copies:=1
End Sub

Sub SPC_ExportReportPDF()
    Dim path As String
    path = ThisWorkbook.Path & Application.PathSeparator & _
           "SPC_Report_" & Format$(Now, "yyyymmdd_hhnnss") & ".pdf"
    ThisWorkbook.Sheets(Array("Capability", "Control Charts", "Monitor")) _
        .Select
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, fileName:=path, _
        Quality:=xlQualityStandard, OpenAfterPublish:=False
    ThisWorkbook.Sheets("Monitor").Select
    MsgBox "Saved:" & vbCrLf & path, vbInformation, "SPC v13"
End Sub

' ============================================================
' NAVIGATION
' ============================================================
Sub GoToHowToUse()
    ThisWorkbook.Sheets("How To Use").Activate
End Sub

Sub GoToSettings()
    ThisWorkbook.Sheets(SH_SET).Activate
End Sub

Sub GoToDataEntry()
    ThisWorkbook.Sheets(SH_DATA).Activate
End Sub

Sub GoToControlCharts()
    ThisWorkbook.Sheets("Control Charts").Activate
End Sub

Sub GoToMonitor()
    ThisWorkbook.Sheets(SH_MON).Activate
End Sub

Sub GoToVisualSPC()
    ThisWorkbook.Sheets("Visual SPC").Activate
End Sub

Sub GoToCapability()
    ThisWorkbook.Sheets("Capability").Activate
End Sub

Sub GoToForm()
    ThisWorkbook.Sheets("Data Collection Form").Activate
End Sub

' ============================================================
' UTILITY
' ============================================================
Private Function Mon() As Worksheet
    Set Mon = ThisWorkbook.Sheets(SH_MON)
End Function

Private Function CfgStr(cell As String) As String
    CfgStr = Trim$(CStr(ThisWorkbook.Sheets(SH_SET).Range(cell).Value))
End Function

Private Function CfgLng(cell As String, dflt As Long) As Long
    Dim v As Variant
    v = ThisWorkbook.Sheets(SH_SET).Range(cell).Value
    If IsNumeric(v) Then CfgLng = CLng(v) Else CfgLng = dflt
    If CfgLng <= 0 Then CfgLng = dflt
End Function

Private Sub SetStatus(text As String, colour As Long)
    With Mon().Range(C_STATUS)
        .Value = text
        .Interior.Color = colour
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
    End With
End Sub

Private Sub RefreshCharts(ws As Worksheet)
    Dim co As ChartObject
    On Error Resume Next
    For Each co In ws.ChartObjects
        co.Chart.Refresh
    Next co
    On Error GoTo 0
End Sub

Private Function FolderExists(p As String) As Boolean
    On Error Resume Next
    FolderExists = (Len(Dir(AddSlash(p), vbDirectory)) > 0)
    On Error GoTo 0
End Function

Private Function AddSlash(p As String) As String
    If Len(p) = 0 Then Exit Function
    If Right$(p, 1) = "\" Or Right$(p, 1) = "/" Then
        AddSlash = p
    Else
        AddSlash = p & Application.PathSeparator
    End If
End Function
