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
    DoEvents
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

    f = Dir(folder & pattern)
    Do While Len(f) > 0
        If Not AlreadyImported(f, FileDateTime(folder & f)) Then
            added = ImportFile(folder & f)
            If added >= 0 Then
                LogImport f, FileDateTime(folder & f), added
                nFiles = nFiles + 1
                nRows = nRows + added
                lastFile = f
                If UCase$(CfgStr("F11")) = "Y" Then ArchiveFile folder, f
            End If
        End If
        f = Dir
    Loop

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

    ff = FreeFile
    Open path For Input As #ff
    Do While Not EOF(ff)
        Line Input #ff, line
        line = Trim$(line)
        If Len(line) > 0 Then
            parts = Split(Replace(line, vbTab, ","), ",")
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
    Dim i As Long, n As Long, firstFeat As Long
    n = UBound(parts) - LBound(parts) + 1
    If n = 0 Then Exit Function

    If IsNumeric(parts(LBound(parts))) Then
        firstFeat = LBound(parts)                ' layout (b)
    Else
        ' layout (a) - but a header row has no numeric fields at all
        If n < 6 Then Exit Function
        If Not IsNumeric(parts(LBound(parts) + 5)) Then Exit Function
        ws.Cells(r, 3).Value = parts(LBound(parts))       ' Date
        ws.Cells(r, 4).Value = parts(LBound(parts) + 1)   ' Shift
        ws.Cells(r, 5).Value = parts(LBound(parts) + 2)   ' Operator
        ws.Cells(r, 6).Value = parts(LBound(parts) + 3)   ' Tool ID
        ws.Cells(r, 7).Value = parts(LBound(parts) + 4)   ' Cluster
        firstFeat = LBound(parts) + 5
    End If

    Dim k As Long
    k = 0
    For i = firstFeat To UBound(parts)
        If k >= MAX_FEAT Then Exit For
        If IsNumeric(parts(i)) Then
            ws.Cells(r, FEAT_COL + k).Value = CDbl(parts(i))
        End If
        k = k + 1
    Next i
    WriteRow = (k > 0)
End Function

Private Function NextFreeRow(ws As Worksheet) As Long
    Dim r As Long
    r = ws.Cells(DATA_LAST, FEAT_COL).End(xlUp).Row
    If r < DATA_FIRST Then
        NextFreeRow = DATA_FIRST
    Else
        NextFreeRow = r + 1
    End If
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
Sub GoToSettings():     ThisWorkbook.Sheets(SH_SET).Activate:                  End Sub
Sub GoToDataEntry():    ThisWorkbook.Sheets(SH_DATA).Activate:                 End Sub
Sub GoToControlCharts(): ThisWorkbook.Sheets("Control Charts").Activate:       End Sub
Sub GoToMonitor():      ThisWorkbook.Sheets(SH_MON).Activate:                  End Sub
Sub GoToVisualSPC():    ThisWorkbook.Sheets("Visual SPC").Activate:            End Sub
Sub GoToCapability():   ThisWorkbook.Sheets("Capability").Activate:            End Sub
Sub GoToForm():         ThisWorkbook.Sheets("Data Collection Form").Activate:  End Sub

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
