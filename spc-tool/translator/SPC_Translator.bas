Attribute VB_Name = "SPC_Translator"
' ============================================================
' SPC TRANSLATOR v13 - Excel edition
'
' Reads arbitrary CSV / Excel exports from an input folder, one file at a
' time, works out which columns hold readings and which hold traceability,
' fills in anything missing with defaults, and writes a clean SPC Calculator
' file to the output folder. The SPC Calculator watches the output folder
' only, so it never sees a malformed file.
'
' Import:  Alt+F11 -> File -> Import File -> select this file (rename to
'          .bas first, or paste the contents into a new module and drop this
'          first Attribute line).
' Save the workbook as .xlsm before using it.
'
' Configuration lives on the Config sheet. Buttons are drawn automatically.
' ============================================================
Option Explicit

Private Const SH_CFG As String = "Config"
Private Const SH_LOG As String = "Log"
Private Const SH_HELP As String = "How To Use"

' Config cells
Private Const C_IN As String = "B4"
Private Const C_OUT As String = "B5"
Private Const C_ARCH As String = "B6"
Private Const C_REJ As String = "B7"
Private Const C_PATTERN As String = "B8"
Private Const C_INTERVAL As String = "B9"
Private Const C_MOVE As String = "B10"
Private Const C_MAXFEAT As String = "B11"
Private Const C_DEF_SHIFT As String = "B12"
Private Const C_DEF_OP As String = "B13"
Private Const C_DEF_TOOL As String = "B14"
Private Const C_DEF_DATE As String = "B15"
Private Const C_STATUS As String = "B17"
Private Const C_LASTRUN As String = "B18"
Private Const C_TRANSLATED As String = "B19"
Private Const C_REJECTED As String = "B20"
Private Const C_ROWS As String = "B21"

' Feature name map (blank = take source column order)
Private Const MAP_FIRST As Long = 25
Private Const MAP_LAST As Long = 74

Private Const MAXCOL As Long = 400

Public TR_Watching As Boolean
Public TR_NextRun As Double

' ============================================================
' STARTUP
' ============================================================
Sub Auto_Open()
    On Error Resume Next
    TR_BuildButtons
    On Error GoTo 0
End Sub

' ============================================================
' PUBLIC ENTRY POINTS
' ============================================================
Sub TR_TranslateNow()
    RunPass False
End Sub

Sub TR_Preview()
    RunPass True
End Sub

Sub TR_StartWatch()
    If Not FolderExists(Cfg(C_IN)) Then
        MsgBox "Input folder not found:" & vbCrLf & Cfg(C_IN), vbExclamation, "SPC Translator"
        Exit Sub
    End If
    TR_Watching = True
    SetStatus "WATCHING", RGB(0, 150, 60)
    RunPass False
    ScheduleNext
    MsgBox "Watching " & Cfg(C_IN) & vbCrLf & _
           "every " & CfgLng(C_INTERVAL, 30) & " seconds.", vbInformation, "SPC Translator"
End Sub

Sub TR_StopWatch()
    TR_Watching = False
    On Error Resume Next
    Application.OnTime TR_NextRun, "TR_Tick", , False
    On Error GoTo 0
    SetStatus "IDLE", RGB(120, 120, 120)
End Sub

Sub TR_Tick()
    If Not TR_Watching Then Exit Sub
    RunPass False
    ScheduleNext
End Sub

Private Sub ScheduleNext()
    Dim iv As Long
    iv = CfgLng(C_INTERVAL, 30)
    If iv < 10 Then iv = 10
    TR_NextRun = Now + TimeSerial(0, 0, iv)
    Application.OnTime TR_NextRun, "TR_Tick"
End Sub

Sub TR_ClearLog()
    If MsgBox("Clear the translation log?", vbYesNo + vbQuestion, _
              "SPC Translator") <> vbYes Then Exit Sub
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SH_LOG)
    If ws.Cells(ws.Rows.count, 1).End(xlUp).Row > 3 Then
        ws.Range("A4:G" & ws.Cells(ws.Rows.count, 1).End(xlUp).Row).ClearContents
    End If
End Sub

Sub TR_OpenInput()
    OpenFolder Cfg(C_IN)
End Sub

Sub TR_OpenOutput()
    OpenFolder Cfg(C_OUT)
End Sub

Private Sub OpenFolder(p As String)
    If Len(p) = 0 Then Exit Sub
    On Error Resume Next
    Shell "explorer.exe """ & p & """", vbNormalFocus
    On Error GoTo 0
End Sub

' ============================================================
' MAIN PASS
' ============================================================
Private Sub RunPass(previewOnly As Boolean)
    Dim inDir As String, outDir As String, pattern As String
    Dim names() As String, count As Long, k As Long, f As String
    Dim okCount As Long, badCount As Long, rowTotal As Long
    Dim msg As String

    inDir = AddSlash(Cfg(C_IN))
    outDir = AddSlash(Cfg(C_OUT))
    pattern = Cfg(C_PATTERN)
    If Len(pattern) = 0 Then pattern = "*.csv"

    If Len(inDir) = 0 Or Not FolderExists(inDir) Then
        MsgBox "Set a valid Input Folder on the Config sheet (B4).", _
               vbExclamation, "SPC Translator"
        Exit Sub
    End If
    If Not previewOnly Then
        If Len(outDir) = 0 Then
            MsgBox "Set an Output Folder on the Config sheet (B5).", _
                   vbExclamation, "SPC Translator"
            Exit Sub
        End If
        EnsureFolder outDir
    End If

    ' collect the file list first - Dir() has one global enumeration and the
    ' folder helpers below call Dir() themselves
    ReDim names(0 To 511)
    f = Dir(inDir & pattern)
    Do While Len(f) > 0
        If count > UBound(names) Then ReDim Preserve names(0 To UBound(names) + 512)
        names(count) = f
        count = count + 1
        f = Dir
    Loop

    If count = 0 Then
        MsgBox "No files matching " & pattern & " in:" & vbCrLf & inDir, _
               vbInformation, "SPC Translator"
        Exit Sub
    End If

    Application.ScreenUpdating = False
    For k = 0 To count - 1
        Dim outLines As Collection, featNames() As String
        Dim notes As String, reason As String, nOut As Long

        Set outLines = New Collection
        notes = ""
        reason = ""
        nOut = TranslateFile(inDir & names(k), outLines, featNames, notes, reason)

        If nOut <= 0 Then
            badCount = badCount + 1
            If Len(reason) = 0 Then reason = "no data rows with readings"
            msg = msg & "REJECT  " & names(k) & " - " & reason & vbCrLf
            If Not previewOnly Then
                WriteLog names(k), "REJECTED", 0, "", "", reason & IIf(Len(notes) > 0, "; " & notes, "")
                MoveFile inDir & names(k), RejectFolder(inDir), names(k)
            End If
        Else
            okCount = okCount + 1
            rowTotal = rowTotal + nOut
            Dim outName As String
            outName = UniqueName(outDir, BaseStem(names(k)) & "_spc", ".csv")
            msg = msg & "OK      " & names(k) & " - " & nOut & " row(s), " & _
                  (UBound(featNames) + 1) & " feature(s)"
            If previewOnly Then
                msg = msg & vbCrLf
            Else
                msg = msg & " -> " & outName & vbCrLf
                WriteAllLines outDir & outName, outLines
                WriteLog names(k), "OK", nOut, JoinNames(featNames), outName, notes
                If UCase$(Cfg(C_MOVE)) = "Y" Then
                    MoveFile inDir & names(k), ArchiveFolder(inDir), names(k)
                End If
            End If
            If Len(notes) > 0 Then msg = msg & "        " & notes & vbCrLf
        End If
    Next k
    Application.ScreenUpdating = True

    If Not previewOnly Then
        ThisWorkbook.Sheets(SH_CFG).Range(C_LASTRUN).Value = Format$(Now, "yyyy-mm-dd hh:nn:ss")
        ThisWorkbook.Sheets(SH_CFG).Range(C_TRANSLATED).Value = okCount
        ThisWorkbook.Sheets(SH_CFG).Range(C_REJECTED).Value = badCount
        ThisWorkbook.Sheets(SH_CFG).Range(C_ROWS).Value = rowTotal
    End If

    MsgBox IIf(previewOnly, "PREVIEW - nothing was written or moved" & vbCrLf & _
               String(46, "-") & vbCrLf, "") & msg & String(46, "-") & vbCrLf & _
           okCount & " translated (" & rowTotal & " rows), " & badCount & " rejected", _
           vbInformation, "SPC Translator"
End Sub

' ============================================================
' TRANSLATE ONE FILE
' Returns rows written, 0 or -1 if the file is unusable.
' ============================================================
Private Function TranslateFile(path As String, outLines As Collection, _
                               ByRef featNames() As String, ByRef notes As String, _
                               ByRef reason As String) As Long
    Dim rows As Collection
    Dim headerIdx As Long
    Dim fieldCol(0 To 4) As Long          ' date, shift, operator, tool, cluster
    Dim measCol() As Long, measCount As Long
    Dim i As Long, j As Long, r As Variant
    Dim defDate As String, line As String
    Dim rowsOut As Long, skipped As Long
    Dim hadTime As Boolean

    Set rows = LoadRows(path, reason)
    If rows Is Nothing Then
        TranslateFile = -1
        Exit Function
    End If
    If rows.count = 0 Then
        reason = "file is empty"
        TranslateFile = -1
        Exit Function
    End If

    headerIdx = FindHeader(rows)
    ClassifyColumns rows, headerIdx, fieldCol, measCol, measCount, featNames, notes, hadTime

    If measCount = 0 Then
        reason = "no column of readings found"
        TranslateFile = -1
        Exit Function
    End If

    defDate = Cfg(C_DEF_DATE)
    If Len(defDate) = 0 Then defDate = Format$(Date, "yyyy-mm-dd")

    ' header line
    line = "Date,Shift,Operator,Tool ID,Cluster"
    For j = 0 To measCount - 1
        line = line & "," & CsvSafe(featNames(j))
    Next j
    outLines.Add line

    Dim startAt As Long
    startAt = IIf(headerIdx > 0, headerIdx + 1, 1)

    For i = startAt To rows.count
        r = rows(i)
        Dim vals As String, anyVal As Boolean
        vals = ""
        anyVal = False
        For j = 0 To measCount - 1
            Dim cell As String
            cell = FieldAt(r, measCol(j))
            If IsNum(cell) Then
                vals = vals & "," & NumStr(ToNum(cell))
                anyVal = True
            Else
                vals = vals & ","
            End If
        Next j

        If anyVal Then
            line = NormDate(FieldAt(r, fieldCol(0)), defDate) & "," & _
                   CsvSafe(PickOr(FieldAt(r, fieldCol(1)), CfgOr(C_DEF_SHIFT, "A"))) & "," & _
                   CsvSafe(PickOr(FieldAt(r, fieldCol(2)), CfgOr(C_DEF_OP, "Operator 1"))) & "," & _
                   CsvSafe(PickOr(FieldAt(r, fieldCol(3)), CfgOr(C_DEF_TOOL, "Tool 1"))) & "," & _
                   CsvSafe(FieldAt(r, fieldCol(4))) & vals
            outLines.Add line
            rowsOut = rowsOut + 1
        Else
            skipped = skipped + 1
        End If
    Next i

    If skipped > 0 Then AddNote notes, skipped & " row(s) had no readings, dropped"
    If fieldCol(0) = 0 Then AddNote notes, "date not in source, set to " & defDate
    If fieldCol(1) = 0 Then AddNote notes, "shift set to " & CfgOr(C_DEF_SHIFT, "A")
    If fieldCol(2) = 0 Then AddNote notes, "operator set to " & CfgOr(C_DEF_OP, "Operator 1")
    If fieldCol(3) = 0 Then AddNote notes, "tool set to " & CfgOr(C_DEF_TOOL, "Tool 1")
    If hadTime Then AddNote notes, "time column present; SPC has no time field, not carried over"

    If rowsOut = 0 Then
        reason = "no data rows with readings"
        TranslateFile = 0
    Else
        TranslateFile = rowsOut
    End If
End Function

' ============================================================
' LOADING
' ============================================================
Private Function LoadRows(path As String, ByRef reason As String) As Collection
    Dim ext As String
    ext = LCase$(Mid$(path, InStrRev(path, ".") + 1))
    Select Case ext
        Case "csv", "txt", "prn", "tsv"
            Set LoadRows = LoadDelimited(path, reason)
        Case "xlsx", "xlsm", "xlsb"
            Set LoadRows = LoadWorkbook(path, reason)
        Case "xls"
            reason = "old .xls format - re-save as .xlsx"
            Set LoadRows = Nothing
        Case Else
            reason = "unsupported file type ." & ext
            Set LoadRows = Nothing
    End Select
End Function

Private Function LoadDelimited(path As String, ByRef reason As String) As Collection
    Dim content As String, lines_() As String, i As Long
    Dim delim As String, res As Collection, line As String

    On Error GoTo Fail
    content = ReadAllText(path)
    If Len(Trim$(content)) = 0 Then
        reason = "file is empty"
        Set LoadDelimited = Nothing
        Exit Function
    End If

    ' Normalise every line ending. Line Input would only break on CR/CRLF, so
    ' a Unix or CMM export arrives as one enormous line and its readings end
    ' up scattered across the sheet.
    content = Replace(content, vbCrLf, vbLf)
    content = Replace(content, vbCr, vbLf)
    lines_ = Split(content, vbLf)

    delim = DetectDelimiter(content)

    Set res = New Collection
    For i = LBound(lines_) To UBound(lines_)
        line = Trim$(lines_(i))
        If Len(line) > 0 Then res.Add SplitLine(line, delim)
    Next i
    Set LoadDelimited = res
    Exit Function
Fail:
    reason = "could not read file: " & Err.Description
    Set LoadDelimited = Nothing
End Function

' A cell's .Text is what is displayed - "####" when the column is narrow, and
' formatted for the local decimal separator. .Value is the real content.
Private Function CellText(c As Range) As String
    Dim v As Variant
    v = c.Value
    If IsEmpty(v) Then Exit Function
    If IsError(v) Then Exit Function
    If IsDate(v) Then
        CellText = Format$(v, "yyyy-mm-dd")
    ElseIf VarType(v) = vbDouble Or VarType(v) = vbSingle Or _
           VarType(v) = vbLong Or VarType(v) = vbInteger Or VarType(v) = vbCurrency Then
        CellText = Replace(Trim$(CStr(v)), ",", ".")
    Else
        CellText = Trim$(CStr(v))
    End If
End Function

Private Function LoadWorkbook(path As String, ByRef reason As String) As Collection
    Dim wb As Workbook, ws As Worksheet, res As Collection
    Dim r As Long, c As Long, firstR As Long, firstC As Long
    Dim lastR As Long, lastC As Long, arr() As String, n As Long

    On Error GoTo Fail
    Set wb = Workbooks.Open(path, ReadOnly:=True, UpdateLinks:=0)
    Set ws = wb.Sheets(1)

    firstR = ws.UsedRange.Row
    firstC = ws.UsedRange.Column
    lastR = firstR + ws.UsedRange.rows.count - 1
    lastC = firstC + ws.UsedRange.Columns.count - 1
    If lastC - firstC + 1 > MAXCOL Then lastC = firstC + MAXCOL - 1

    Set res = New Collection
    For r = firstR To lastR
        n = lastC - firstC
        ReDim arr(0 To n)
        Dim anyText As Boolean
        anyText = False
        For c = firstC To lastC
            arr(c - firstC) = CellText(ws.Cells(r, c))
            If Len(arr(c - firstC)) > 0 Then anyText = True
        Next c
        If anyText Then res.Add arr
    Next r

    wb.Close SaveChanges:=False
    Set LoadWorkbook = res
    Exit Function
Fail:
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    On Error GoTo 0
    reason = "could not read workbook: " & Err.Description
    Set LoadWorkbook = Nothing
End Function

' ============================================================
' SHAPE DETECTION
' ============================================================
' First essentially-text row that is followed by a more numeric row.
' Taking the FIRST such row matters: a data row beginning with a date, an
' operator and a tool is itself mostly text, so a looser rule swallows it as
' a header and quietly loses a reading.
Private Function FindHeader(rows As Collection) As Long
    Dim i As Long, j As Long
    Dim shareHere As Double, shareNext As Double

    For i = 1 To WorksheetFunction.Min(25, rows.count)
        shareHere = NumericShare(rows(i))
        If CellCount(rows(i)) >= 2 And shareHere <= 0.2 Then
            For j = i + 1 To WorksheetFunction.Min(i + 3, rows.count)
                If CellCount(rows(j)) > 0 Then
                    shareNext = NumericShare(rows(j))
                    If shareNext > 0 And shareNext > shareHere Then
                        FindHeader = i
                        Exit Function
                    End If
                    Exit For
                End If
            Next j
        End If
    Next i
    FindHeader = 0
End Function

Private Sub ClassifyColumns(rows As Collection, headerIdx As Long, _
                            ByRef fieldCol() As Long, ByRef measCol() As Long, _
                            ByRef measCount As Long, ByRef featNames() As String, _
                            ByRef notes As String, ByRef hadTime As Boolean)
    Dim ncols As Long, i As Long, c As Long, k As Long
    Dim hdr As Variant, key As String
    Dim startAt As Long, total As Long, numeric As Long
    Dim maxFeat As Long

    maxFeat = CfgLng(C_MAXFEAT, 50)
    If maxFeat < 1 Then maxFeat = 50

    ncols = 0
    For i = 1 To rows.count
        If UBound(rows(i)) + 1 > ncols Then ncols = UBound(rows(i)) + 1
    Next i
    If ncols > MAXCOL Then ncols = MAXCOL

    For i = 0 To 4
        fieldCol(i) = 0
    Next i

    Dim hdrName() As String
    ReDim hdrName(0 To ncols)
    If headerIdx > 0 Then
        hdr = rows(headerIdx)
        For c = 0 To WorksheetFunction.Min(UBound(hdr), ncols - 1)
            hdrName(c) = Trim$(CStr(hdr(c)))
            key = NormKey(hdrName(c))
            If Len(key) > 0 Then
                If fieldCol(0) = 0 Then If MatchesAlias(key, "date|measured|measurementdate|day|datetime|timestamp|inspectiondate|recorded") Then fieldCol(0) = c + 1
                If fieldCol(1) = 0 Then If MatchesAlias(key, "shift|turn|team") Then fieldCol(1) = c + 1
                If fieldCol(2) = 0 Then If MatchesAlias(key, "operator|op|inspector|user|employee|operatorid") Then fieldCol(2) = c + 1
                If fieldCol(3) = 0 Then If MatchesAlias(key, "tool|toolid|machine|machineid|equipment|gauge|gage|gaugeid|device|cell|station") Then fieldCol(3) = c + 1
                If fieldCol(4) = 0 Then If MatchesAlias(key, "cluster|batch|lot|serial|part|partno|serialno|job|order|cavity") Then fieldCol(4) = c + 1
                If MatchesAlias(key, "time|clock|hour") Then hadTime = True
            End If
        Next c
    End If

    startAt = IIf(headerIdx > 0, headerIdx + 1, 1)

    ReDim measCol(0 To ncols)
    ReDim featNames(0 To ncols)
    measCount = 0

    For c = 0 To ncols - 1
        If Not IsTaken(fieldCol, c + 1) Then
            total = 0
            numeric = 0
            For i = startAt To rows.count
                Dim v As String
                v = FieldAt(rows(i), c + 1)
                If Len(v) > 0 Then
                    total = total + 1
                    If IsNum(v) Then numeric = numeric + 1
                End If
            Next i
            If total > 0 Then
                If numeric / total >= 0.8 Then
                    key = NormKey(hdrName(c))
                    If MatchesAlias(key, "no|num|number|index|id|seq|sequence|sample|sampleno|item|row|count|n") And Len(key) > 0 Then
                        AddNote notes, "column '" & hdrName(c) & "' looks like a counter, ignored"
                    ElseIf IsCounterColumn(rows, startAt, c + 1) Then
                        AddNote notes, "column " & (c + 1) & " is a 1,2,3... counter, ignored"
                    Else
                        If measCount < maxFeat Then
                            measCol(measCount) = c + 1
                            If Len(hdrName(c)) > 0 Then
                                featNames(measCount) = hdrName(c)
                            Else
                                featNames(measCount) = "F" & (measCount + 1)
                            End If
                            measCount = measCount + 1
                        End If
                    End If
                End If
            End If
        End If
    Next c

    ' optional: pin the feature order from the Config sheet map
    Dim wanted() As String, wantCount As Long
    wantCount = ReadFeatureMap(wanted)
    If wantCount > 0 Then
        Dim newCol() As Long, newNames() As String, hit As Long, anyHit As Boolean
        ReDim newCol(0 To wantCount)
        ReDim newNames(0 To wantCount)
        For k = 0 To wantCount - 1
            hit = 0
            For i = 0 To measCount - 1
                If NormKey(featNames(i)) = NormKey(wanted(k)) Then
                    hit = measCol(i)
                    anyHit = True
                    Exit For
                End If
            Next i
            newCol(k) = hit
            newNames(k) = wanted(k)
        Next k
        If anyHit Then
            measCol = newCol
            featNames = newNames
            measCount = wantCount
        Else
            AddNote notes, "no source column matched the feature map; using column order"
        End If
    End If

    If measCount > 0 Then
        ReDim Preserve featNames(0 To measCount - 1)
    End If
End Sub

Private Function IsCounterColumn(rows As Collection, startAt As Long, colNo As Long) As Boolean
    Dim i As Long, n As Long, prev As Double, v As String, d As Double
    Dim first As Boolean
    first = True
    For i = startAt To rows.count
        v = FieldAt(rows(i), colNo)
        If Len(v) > 0 Then
            If Not IsNum(v) Then Exit Function
            d = ToNum(v)
            If d <> Int(d) Then Exit Function
            If Not first Then
                If d <> prev + 1 Then Exit Function
            End If
            prev = d
            first = False
            n = n + 1
        End If
    Next i
    IsCounterColumn = (n >= 5)
End Function

Private Function ReadFeatureMap(ByRef wanted() As String) As Long
    Dim ws As Worksheet, r As Long, n As Long, v As String
    Set ws = ThisWorkbook.Sheets(SH_CFG)
    ReDim wanted(0 To MAP_LAST - MAP_FIRST)
    For r = MAP_FIRST To MAP_LAST
        v = Trim$(CStr(ws.Cells(r, 2).Value))
        If Len(v) > 0 Then
            wanted(n) = v
            n = n + 1
        End If
    Next r
    ReadFeatureMap = n
End Function

' ============================================================
' FIELD HELPERS
' ============================================================
Private Function FieldAt(r As Variant, colNo As Long) As String
    If colNo <= 0 Then Exit Function
    If colNo - 1 > UBound(r) Then Exit Function
    FieldAt = Trim$(CStr(r(colNo - 1)))
End Function

Private Function CellCount(r As Variant) As Long
    Dim i As Long, n As Long
    For i = LBound(r) To UBound(r)
        If Len(Trim$(CStr(r(i)))) > 0 Then n = n + 1
    Next i
    CellCount = n
End Function

Private Function NumericShare(r As Variant) As Double
    Dim i As Long, n As Long, num As Long, v As String
    For i = LBound(r) To UBound(r)
        v = Trim$(CStr(r(i)))
        If Len(v) > 0 Then
            n = n + 1
            If IsNum(v) Then num = num + 1
        End If
    Next i
    If n = 0 Then Exit Function
    NumericShare = num / n
End Function

Private Function IsTaken(fieldCol() As Long, colNo As Long) As Boolean
    Dim i As Long
    For i = 0 To 4
        If fieldCol(i) = colNo Then
            IsTaken = True
            Exit Function
        End If
    Next i
End Function

Private Function MatchesAlias(key As String, aliases As String) As Boolean
    Dim parts() As String, i As Long
    If Len(key) = 0 Then Exit Function
    parts = Split(aliases, "|")
    For i = LBound(parts) To UBound(parts)
        If key = parts(i) Then
            MatchesAlias = True
            Exit Function
        End If
        If Len(parts(i)) >= 4 Then
            If Left$(key, Len(parts(i))) = parts(i) Then
                MatchesAlias = True
                Exit Function
            End If
        End If
    Next i
End Function

Private Function NormKey(s As String) As String
    Dim i As Long, ch As String, res As String
    For i = 1 To Len(s)
        ch = LCase$(Mid$(s, i, 1))
        If (ch >= "a" And ch <= "z") Or (ch >= "0" And ch <= "9") Then res = res & ch
    Next i
    NormKey = res
End Function

' Numeric test that tolerates a decimal comma
Private Function IsNum(ByVal s As String) As Boolean
    s = Replace(Trim$(s), " ", "")
    If Len(s) = 0 Then Exit Function
    If InStr(s, ",") > 0 And InStr(s, ".") = 0 Then s = Replace(s, ",", ".")
    If InStr(s, ",") > 0 Then Exit Function
    Dim i As Long, ch As String, digits As Long
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch >= "0" And ch <= "9" Then
            digits = digits + 1
        ElseIf ch = "." Or ch = "-" Or ch = "+" Or LCase$(ch) = "e" Then
            ' allowed
        Else
            Exit Function
        End If
    Next i
    If digits = 0 Then Exit Function
    IsNum = IsNumeric(s)
End Function

' Format$ uses the machine's decimal separator, so on a comma-decimal locale
' it would emit 25,0343 and split the CSV into an extra column. "0.0000" never
' produces a thousands separator, so the only comma it can contain is the
' decimal point - swapping it back for a dot is safe.
Private Function NumStr(d As Double) As String
    NumStr = Replace(Format$(d, "0.0000"), ",", ".")
End Function

Private Function ToNum(ByVal s As String) As Double
    s = Replace(Trim$(s), " ", "")
    If InStr(s, ",") > 0 And InStr(s, ".") = 0 Then s = Replace(s, ",", ".")
    ToNum = Val(s)
End Function

Private Function PickOr(v As String, dflt As String) As String
    If Len(Trim$(v)) = 0 Then
        PickOr = dflt
    Else
        PickOr = Trim$(v)
    End If
End Function

' Everything out as YYYY-MM-DD; anything unparseable falls back
Private Function NormDate(ByVal s As String, fallback As String) As String
    Dim t As String, d As Date
    t = Trim$(s)
    If Len(t) = 0 Then
        NormDate = fallback
        Exit Function
    End If
    If InStr(t, " ") > 0 Then t = Left$(t, InStr(t, " ") - 1)
    If InStr(t, "T") > 0 And Len(t) > 10 Then t = Left$(t, InStr(t, "T") - 1)

    ' already ISO?
    If Len(t) = 10 Then
        If Mid$(t, 5, 1) = "-" And Mid$(t, 8, 1) = "-" Then
            NormDate = t
            Exit Function
        End If
    End If
    ' 8 digit yyyymmdd
    If Len(t) = 8 And IsNumeric(t) Then
        NormDate = Left$(t, 4) & "-" & Mid$(t, 5, 2) & "-" & Right$(t, 2)
        Exit Function
    End If

    On Error Resume Next
    d = 0
    d = CDate(t)
    On Error GoTo 0
    If d > 0 Then
        NormDate = Format$(d, "yyyy-mm-dd")
    Else
        NormDate = fallback
    End If
End Function

' ============================================================
' TEXT / FILE HELPERS
' ============================================================
Private Function ReadAllText(path As String) As String
    Dim ff As Integer, s As String
    ff = FreeFile
    Open path For Binary Access Read As #ff
    If LOF(ff) > 0 Then
        s = Space$(LOF(ff))
        Get #ff, 1, s
    End If
    Close #ff
    If Len(s) >= 3 Then
        If Left$(s, 3) = Chr$(239) & Chr$(187) & Chr$(191) Then s = Mid$(s, 4)
    End If
    ReadAllText = s
End Function

Private Function DetectDelimiter(content As String) As String
    Dim sample As String
    Dim nComma As Long, nSemi As Long, nTab As Long, nPipe As Long
    sample = Left$(content, 8192)
    nComma = CountChar(sample, ",")
    nSemi = CountChar(sample, ";")
    nTab = CountChar(sample, vbTab)
    nPipe = CountChar(sample, "|")

    DetectDelimiter = ","
    If nSemi > nComma And nSemi >= nTab And nSemi >= nPipe Then DetectDelimiter = ";"
    If nTab > nComma And nTab > nSemi And nTab >= nPipe Then DetectDelimiter = vbTab
    If nPipe > nComma And nPipe > nSemi And nPipe > nTab Then DetectDelimiter = "|"
End Function

Private Function CountChar(s As String, ch As String) As Long
    CountChar = Len(s) - Len(Replace(s, ch, ""))
End Function

' Quote-aware split, so a comma inside "25.03, nominal" does not break the row
Private Function SplitLine(ByVal s As String, ByVal delim As String) As Variant
    Dim res() As String, n As Long, i As Long
    Dim ch As String, cur As String, inQ As Boolean

    ReDim res(0 To 255)
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch = """" Then
            If inQ Then
                If i < Len(s) Then
                    If Mid$(s, i + 1, 1) = """" Then
                        cur = cur & """"
                        i = i + 1
                    Else
                        inQ = False
                    End If
                Else
                    inQ = False
                End If
            Else
                inQ = True
            End If
        ElseIf ch = delim And Not inQ Then
            If n > UBound(res) Then ReDim Preserve res(0 To UBound(res) + 256)
            res(n) = Trim$(cur)
            n = n + 1
            cur = ""
        Else
            cur = cur & ch
        End If
    Next i
    If n > UBound(res) Then ReDim Preserve res(0 To UBound(res) + 256)
    res(n) = Trim$(cur)
    ReDim Preserve res(0 To n)
    SplitLine = res
End Function

Private Function CsvSafe(ByVal s As String) As String
    s = Replace(s, """", "'")
    If InStr(s, ",") > 0 Then s = Replace(s, ",", " ")
    CsvSafe = s
End Function

' Always CRLF - the SPC importer handles any ending, but other Excel-side
' readers (and Line Input) do not.
Private Sub WriteAllLines(path As String, lines_ As Collection)
    Dim ff As Integer, i As Long
    ff = FreeFile
    Open path For Output As #ff
    For i = 1 To lines_.count
        Print #ff, lines_(i)
    Next i
    Close #ff
End Sub

Private Function BaseStem(fileName As String) As String
    Dim p As Long
    p = InStrRev(fileName, ".")
    If p > 1 Then
        BaseStem = Left$(fileName, p - 1)
    Else
        BaseStem = fileName
    End If
End Function

Private Function UniqueName(folder As String, stem As String, ext As String) As String
    Dim n As Long, candidate As String
    candidate = stem & ext
    Do While Len(Dir(AddSlash(folder) & candidate)) > 0
        n = n + 1
        candidate = stem & "_" & n & ext
    Loop
    UniqueName = candidate
End Function

Private Sub MoveFile(src As String, destFolder As String, fileName As String)
    Dim target As String, ext As String, dot As Long
    EnsureFolder destFolder
    dot = InStrRev(fileName, ".")
    If dot > 1 Then ext = Mid$(fileName, dot)      ' Mid$(s, 0) would error
    target = AddSlash(destFolder) & UniqueName(destFolder, BaseStem(fileName), ext)
    On Error Resume Next
    Name src As target
    On Error GoTo 0
End Sub

Private Function ArchiveFolder(inDir As String) As String
    ArchiveFolder = Cfg(C_ARCH)
    If Len(ArchiveFolder) = 0 Then ArchiveFolder = AddSlash(inDir) & "archive"
End Function

Private Function RejectFolder(inDir As String) As String
    RejectFolder = Cfg(C_REJ)
    If Len(RejectFolder) = 0 Then RejectFolder = AddSlash(inDir) & "rejects"
End Function

Private Sub EnsureFolder(p As String)
    Dim parts() As String, i As Long, build As String
    If Len(p) = 0 Then Exit Sub
    p = Replace(p, "/", "\")
    If Right$(p, 1) = "\" Then p = Left$(p, Len(p) - 1)
    parts = Split(p, "\")
    For i = LBound(parts) To UBound(parts)
        If i = 0 Then
            build = parts(i)
        Else
            build = build & "\" & parts(i)
        End If
        If Len(build) > 0 And InStr(build, ":") <> Len(build) Then
            If Not FolderExists(build) And Len(parts(i)) > 0 Then
                On Error Resume Next
                MkDir build
                On Error GoTo 0
            End If
        End If
    Next i
End Sub

Private Function FolderExists(p As String) As Boolean
    If Len(p) = 0 Then Exit Function
    On Error Resume Next
    FolderExists = (Len(Dir(AddSlash(p), vbDirectory)) > 0)
    On Error GoTo 0
End Function

Private Function AddSlash(p As String) As String
    If Len(p) = 0 Then Exit Function
    If Right$(p, 1) = "\" Or Right$(p, 1) = "/" Then
        AddSlash = p
    Else
        AddSlash = p & "\"
    End If
End Function

Private Sub AddNote(ByRef notes As String, s As String)
    If Len(notes) = 0 Then
        notes = s
    Else
        notes = notes & "; " & s
    End If
End Sub

Private Function JoinNames(names() As String) As String
    Dim i As Long, s As String
    On Error Resume Next
    For i = LBound(names) To UBound(names)
        If Len(names(i)) > 0 Then
            If Len(s) = 0 Then
                s = names(i)
            Else
                s = s & " | " & names(i)
            End If
        End If
    Next i
    On Error GoTo 0
    JoinNames = s
End Function

' ============================================================
' CONFIG / LOG / STATUS
' ============================================================
Private Function Cfg(cell As String) As String
    Cfg = Trim$(CStr(ThisWorkbook.Sheets(SH_CFG).Range(cell).Value))
End Function

Private Function CfgOr(cell As String, dflt As String) As String
    CfgOr = Cfg(cell)
    If Len(CfgOr) = 0 Then CfgOr = dflt
End Function

Private Function CfgLng(cell As String, dflt As Long) As Long
    Dim v As Variant
    v = ThisWorkbook.Sheets(SH_CFG).Range(cell).Value
    If IsNumeric(v) Then
        CfgLng = CLng(v)
    Else
        CfgLng = dflt
    End If
    If CfgLng <= 0 Then CfgLng = dflt
End Function

Private Sub SetStatus(text As String, colour As Long)
    With ThisWorkbook.Sheets(SH_CFG).Range(C_STATUS)
        .Value = text
        .Interior.Color = colour
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
    End With
End Sub

Private Sub WriteLog(srcName As String, result As String, rowsOut As Long, _
                     features As String, outName As String, notes As String)
    Dim ws As Worksheet, r As Long
    Set ws = ThisWorkbook.Sheets(SH_LOG)
    r = ws.Cells(ws.rows.count, 1).End(xlUp).Row + 1
    If r < 4 Then r = 4
    ws.Cells(r, 1).Value = Format$(Now, "yyyy-mm-dd hh:nn:ss")
    ws.Cells(r, 2).Value = srcName
    ws.Cells(r, 3).Value = result
    ws.Cells(r, 4).Value = rowsOut
    ws.Cells(r, 5).Value = features
    ws.Cells(r, 6).Value = outName
    ws.Cells(r, 7).Value = notes
End Sub

' ============================================================
' BUTTONS
' ============================================================
Sub TR_BuildButtons()
    Dim ws As Worksheet, i As Long
    Set ws = ThisWorkbook.Sheets(SH_HELP)

    For i = ws.Shapes.count To 1 Step -1
        If Left$(ws.Shapes(i).Name, 4) = "btn_" Then ws.Shapes(i).Delete
    Next i

    Dim x As Single, y As Single
    x = ws.Range("E5").Left
    y = ws.Range("E5").Top

    y = y + AddButton(ws, "Preview (writes nothing)", "TR_Preview", x, y, 170, 26, RGB(31, 78, 121)) + 5
    y = y + AddButton(ws, "Translate Now", "TR_TranslateNow", x, y, 170, 26, RGB(0, 130, 60)) + 5
    y = y + AddButton(ws, "Start Watching", "TR_StartWatch", x, y, 170, 26, RGB(0, 130, 60)) + 5
    y = y + AddButton(ws, "Stop Watching", "TR_StopWatch", x, y, 170, 26, RGB(170, 40, 40)) + 5
    y = y + AddButton(ws, "Open Input Folder", "TR_OpenInput", x, y, 170, 26, RGB(68, 84, 106)) + 5
    y = y + AddButton(ws, "Open Output Folder", "TR_OpenOutput", x, y, 170, 26, RGB(68, 84, 106)) + 5
    y = y + AddButton(ws, "Clear Log", "TR_ClearLog", x, y, 170, 26, RGB(120, 120, 120)) + 5
End Sub

Private Function AddButton(ws As Worksheet, caption As String, macro As String, _
                           x As Single, y As Single, w As Single, h As Single, _
                           colour As Long) As Single
    Dim shp As Shape
    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, x, y, w, h)
    shp.Name = "btn_" & macro
    shp.Fill.ForeColor.RGB = colour
    shp.Line.ForeColor.RGB = RGB(255, 255, 255)
    With shp.TextFrame2
        .TextRange.text = caption
        .TextRange.Font.Size = 10
        .TextRange.Font.Bold = msoTrue
        .TextRange.Font.Name = "Arial"
        .TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .VerticalAnchor = msoAnchorMiddle
        .TextRange.ParagraphFormat.Alignment = msoAlignCenter
    End With
    shp.OnAction = macro
    AddButton = h
End Function
