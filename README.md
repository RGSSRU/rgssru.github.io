VBA

Option Explicit
Option Base 1

' =========================================================
' CONFIG
' =========================================================

Private Const SOURCE_FOLDER As String = "d:\\"
Private Const TARGET_FOLDER As String = "d:\\Target\"
Private Const BACKUP_FOLDER As String = "d:\\Target\Backup\"

Private Const FILE_2026 As String = ".xlsx"
Private Const FILE_2526 As String = ".xlsx"
Private Const FILE_2324 As String = ".xlsx"
Private Const FILE_2022 As String = ".xlsx"

' =========================================================
' CHECKS (True = enabled, False = disabled)
' =========================================================

Private Const ENABLE_SOURCE_DUP_CHECK As Boolean = True
Private Const ENABLE_TARGET_DUP_CHECK As Boolean = False

' =========================================================
' LOGGING (True = enabled, False = disabled)
' =========================================================

Private Const ENABLE_LOGGING As Boolean = False

' =========================================================
' GLOBAL STATISTICS
' =========================================================

Private statsUpdated As Long
Private statsCreated As Long

' =========================================================
' LOGGING WRAPPER
' =========================================================

Private Sub LogMessage(ByVal msg As String)
    If ENABLE_LOGGING Then
        Debug.Print msg
    End If
End Sub

' =========================================================
' ENTRY POINT
' =========================================================

Public Sub RunInkassoCockpitUpdate()

    Dim sourcePath As String
    Dim startTime As Double
    Dim endTime As Double
    Dim msg As String

    startTime = Timer

    On Error GoTo FAIL

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.DisplayAlerts = False

    sourcePath = SOURCE_FOLDER & Format(Date, "dd.mm") & "\"

    If Dir(sourcePath, vbDirectory) = "" Then
        Err.Raise vbObjectError + 1, , "Source folder not found: " & sourcePath
    End If

    ' Create backup folder if not exists
    If Dir(BACKUP_FOLDER, vbDirectory) = "" Then
        MkDir BACKUP_FOLDER
    End If

    LogMessage String(60, "=")
    LogMessage "=== INKASSO COCKPIT UPDATE ==="
    LogMessage "=== START " & Now
    LogMessage String(60, "=")
    LogMessage "SOURCE DUP CHECK: " & IIf(ENABLE_SOURCE_DUP_CHECK, "ON", "OFF")
    LogMessage "TARGET DUP CHECK: " & IIf(ENABLE_TARGET_DUP_CHECK, "ON", "OFF")
    LogMessage "BACKUP FOLDER: " & BACKUP_FOLDER
    LogMessage ""

    statsUpdated = 0
    statsCreated = 0

    ProcessAllSourceFiles sourcePath

    endTime = Timer

    LogMessage ""
    LogMessage String(60, "=")
    LogMessage "=== FINISH ==="
    LogMessage "Total rows UPDATED: " & statsUpdated
    LogMessage "Total rows CREATED: " & statsCreated
    LogMessage "Total time: " & Format(endTime - startTime, "0.00") & " sec"
    LogMessage String(60, "=")

    msg = "FINISH" & vbCrLf & vbCrLf
    msg = msg & "Total rows UPDATED: " & statsUpdated & vbCrLf
    msg = msg & "Total rows CREATED: " & statsCreated & vbCrLf
    msg = msg & "Total time: " & Format(endTime - startTime, "0.00") & " sec"

    MsgBox msg, vbInformation, "Inkasso Cockpit Update"

CLEAN_EXIT:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Exit Sub

FAIL:
    Debug.Print "[ERROR] " & Err.Description
    MsgBox Err.Description, vbCritical
    Resume CLEAN_EXIT

End Sub

' =========================================================
' MAIN PIPELINE
' =========================================================

Private Sub ProcessAllSourceFiles(ByVal sourcePath As String)

    Dim d25_26 As Object, d23_24 As Object, d20_22 As Object
    Dim loadStart As Double, loadEnd As Double
    Dim validateStart As Double, validateEnd As Double
    Dim processStart As Double, processEnd As Double

    Set d25_26 = CreateObject("Scripting.Dictionary")
    Set d23_24 = CreateObject("Scripting.Dictionary")
    Set d20_22 = CreateObject("Scripting.Dictionary")

    ' === LOAD ===
    loadStart = Timer
    LoadAllFiles sourcePath, d25_26, d23_24, d20_22
    loadEnd = Timer

    LogMessage ""
    LogMessage "[LOAD] Completed in " & Format(loadEnd - loadStart, "0.00") & " sec"
    LogMessage "  25-26: " & d25_26.Count & " groups"
    LogMessage "  23-24: " & d23_24.Count & " groups"
    LogMessage "  20-22: " & d20_22.Count & " groups"

    ' === VALIDATE ===
    validateStart = Timer
    If ENABLE_SOURCE_DUP_CHECK Then
        ValidateSourceGroups d25_26
        ValidateSourceGroups d23_24
        ValidateSourceGroups d20_22
    Else
        LogMessage "[WARN] Source duplicate checks are DISABLED"
    End If
    validateEnd = Timer

    LogMessage "[VALIDATE] Completed in " & Format(validateEnd - validateStart, "0.00") & " sec"

    ' === PROCESS ===
    processStart = Timer
    ProcessTargetGroup d25_26, "2025-2026"
    ProcessTargetGroup d23_24, "2023-2024"
    ProcessTargetGroup d20_22, "2020-2022"
    processEnd = Timer

    LogMessage "[PROCESS] Completed in " & Format(processEnd - processStart, "0.00") & " sec"

End Sub

' =========================================================
' LOAD FILES ENTRY
' =========================================================

Private Sub LoadAllFiles(ByVal sourcePath As String, _
                        ByRef d25_26 As Object, _
                        ByRef d23_24 As Object, _
                        ByRef d20_22 As Object)

    Dim f As String
    Dim fileCount As Long

    f = Dir(sourcePath & "*.xlsx")
    fileCount = 0

    Do While f <> ""

        If Left$(f, 2) <> "~$" Then
            fileCount = fileCount + 1
            LogMessage "  Loading: " & f
            LoadFile sourcePath & f, d25_26, d23_24, d20_22
        End If

        f = Dir()
    Loop

    LogMessage "  Files loaded: " & fileCount

End Sub

' =========================================================
' FILE LOADER
' =========================================================

Private Sub LoadFile(ByVal filePath As String, _
                     ByRef d25_26 As Object, _
                     ByRef d23_24 As Object, _
                     ByRef d20_22 As Object)

    Dim wb As Workbook, ws As Worksheet
    Dim lastRow As Long, lastCol As Long
    Dim i As Long
    Dim rowCount As Long

    Set wb = Workbooks.Open(filePath, ReadOnly:=True, UpdateLinks:=False)
    Set ws = wb.Sheets(1)

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastCol = 86 'ws.Cells(3, ws.Columns.Count).End(xlToLeft).Column

    rowCount = 0

    For i = 3 To lastRow

        Dim prefix As String
        Dim arr As Variant
        Dim g As String

        prefix = Trim$(ws.Cells(i, 1).Value)

        If prefix <> "" Then

            g = GetYearGroup(prefix)
            If g = "" Then GoTo NextI

            arr = ws.Range(ws.Cells(i, 1), ws.Cells(i, lastCol)).Value

            Select Case g
                Case "2025-2026": AddSourceRow d25_26, prefix, arr, filePath
                Case "2023-2024": AddSourceRow d23_24, prefix, arr, filePath
                Case "2020-2022": AddSourceRow d20_22, prefix, arr, filePath
            End Select

            rowCount = rowCount + 1

        End If

NextI:
    Next i

    wb.Close False

    LogMessage "    Rows loaded: " & rowCount

End Sub

' =========================================================
' ADD SOURCE ROW (GROUPING BASE/INDEX)
' =========================================================

Private Sub AddSourceRow(ByRef groups As Object, _
                         ByVal fullPrefix As String, _
                         ByVal rowData As Variant, _
                         ByVal filePath As String)

    Dim basePrefix As String
    Dim idxText As String
    Dim idx As Long
    Dim d As Object

    SplitPrefix fullPrefix, basePrefix, idxText

    If idxText = "" Then
        idx = 0
    Else
        If Not IsNumeric(idxText) Then
            Err.Raise vbObjectError + 1000, , _
                "Invalid index in source: " & fullPrefix & vbCrLf & "File: " & filePath
        End If
        idx = CLng(idxText)
    End If

    If Not groups.Exists(basePrefix) Then
        Set d = CreateObject("Scripting.Dictionary")
        groups.Add basePrefix, d
    End If

    Set d = groups(basePrefix)

    If ENABLE_SOURCE_DUP_CHECK Then
        If d.Exists(idx) Then
            Err.Raise vbObjectError + 1001, , _
                "Duplicate prefix in source files: " & fullPrefix & vbCrLf & "File: " & filePath
        End If
    Else
        If d.Exists(idx) Then
            LogMessage "[WARN] Duplicate source prefix ignored (check disabled): " & fullPrefix
            Exit Sub
        End If
    End If

    d.Add idx, rowData

End Sub

' =========================================================
' VALIDATE SOURCE GROUPS
' =========================================================

Private Sub ValidateSourceGroups(ByRef groups As Object)

    Dim basePrefix As Variant
    Dim d As Object
    Dim k As Variant
    Dim hasBase As Boolean
    Dim maxIdx As Long
    Dim i As Long

    For Each basePrefix In groups.Keys

        Set d = groups(basePrefix)

        hasBase = d.Exists(0)

        maxIdx = 0
        For Each k In d.Keys
            If CLng(k) > maxIdx Then maxIdx = CLng(k)
        Next k

        If hasBase And maxIdx > 0 Then
            Err.Raise vbObjectError + 1100, , _
                "Invalid source: base and indexed rows exist together: " & basePrefix
        End If

        If maxIdx > 0 Then
            For i = 1 To maxIdx
                If Not d.Exists(i) Then
                    Err.Raise vbObjectError + 1101, , _
                        "Missing index /" & i & " for base: " & basePrefix
                End If
            Next i
        End If

    Next basePrefix

End Sub

' =========================================================
' CREATE BACKUP
' =========================================================

Private Function CreateBackup(ByVal filePath As String) As String

    Dim fileName As String
    Dim backupName As String
    Dim backupPath As String
    Dim timestamp As String

    ' Extract file name from full path
    fileName = Right(filePath, Len(filePath) - InStrRev(filePath, "\"))

    ' Create timestamp: YYYYMMDD_HHMM
    timestamp = Format(Year(Date), "0000") & _
                Format(Month(Date), "00") & _
                Format(Day(Date), "00") & "_" & _
                Format(Hour(Now), "00") & _
                Format(Minute(Now), "00")

    ' Remove extension for backup name
    If InStr(fileName, ".") > 0 Then
        backupName = Left(fileName, InStrRev(fileName, ".") - 1) & "_backup_" & timestamp & ".xlsx"
    Else
        backupName = fileName & "_backup_" & timestamp & ".xlsx"
    End If

    backupPath = BACKUP_FOLDER & backupName

    ' Create backup folder if not exists
    If Dir(BACKUP_FOLDER, vbDirectory) = "" Then
        MkDir BACKUP_FOLDER
    End If

    ' Copy file
    FileCopy filePath, backupPath

    LogMessage "  Backup created: " & backupName

    CreateBackup = backupPath

End Function

' =========================================================
' GROUP ROUTER
' =========================================================

Private Sub ProcessTargetGroup(ByRef dict As Object, ByVal groupName As String)

    Dim wb2026 As Workbook, ws2026 As Worksheet
    Dim wb2526 As Workbook, ws2526 As Worksheet
    Dim wb2324 As Workbook, ws2324 As Worksheet
    Dim wb2022 As Workbook, ws2022 As Worksheet

    Dim map2026 As Object, map2526 As Object, map2324 As Object, map2022 As Object

    Dim stats2026 As String, stats2526 As String
    Dim stats2324 As String, stats2022 As String

    Dim filePath2026 As String, filePath2526 As String
    Dim filePath2324 As String, filePath2022 As String

    If dict.Count = 0 Then Exit Sub

    LogMessage ""
    LogMessage "[PROCESS] Group: " & groupName & " (" & dict.Count & " base prefixes)"

    Select Case groupName

        Case "2025-2026"

            filePath2026 = TARGET_FOLDER & FILE_2026
            filePath2526 = TARGET_FOLDER & FILE_2526

            ' Create backups before opening
            CreateBackup filePath2026
            CreateBackup filePath2526

            Set wb2026 = OpenCockpit(FILE_2026)
            Set wb2526 = OpenCockpit(FILE_2526)

            Set ws2026 = wb2026.Sheets(1)
            Set ws2526 = wb2526.Sheets(1)

            stats2026 = GetFileStats(ws2026)
            stats2526 = GetFileStats(ws2526)

            Set map2026 = BuildCockpitMap(ws2026)
            Set map2526 = BuildCockpitMap(ws2526)

            If ENABLE_TARGET_DUP_CHECK Then
                ValidateCockpitMap map2026
                ValidateCockpitMap map2526
            Else
                LogMessage "[WARN] Target duplicate checks are DISABLED"
            End If

            Dim basePrefix As Variant
            For Each basePrefix In dict.Keys

                Select Case FindCockpitForPrefix(CStr(basePrefix), map2026, map2526)

                    Case 1
                        ProcessGroup ws2026, CStr(basePrefix), dict(basePrefix), map2026, groupName
                        Set map2026 = BuildCockpitMap(ws2026)

                    Case 2
                        ProcessGroup ws2526, CStr(basePrefix), dict(basePrefix), map2526, groupName
                        Set map2526 = BuildCockpitMap(ws2526)

                End Select

            Next basePrefix

            wb2026.Save
            wb2526.Save
            wb2026.Close False
            wb2526.Close False

            LogMessage "  " & FILE_2026 & ": " & stats2026
            LogMessage "  " & FILE_2526 & ": " & stats2526

        Case "2023-2024"

            filePath2324 = TARGET_FOLDER & FILE_2324

            CreateBackup filePath2324

            Set wb2324 = OpenCockpit(FILE_2324)
            Set ws2324 = wb2324.Sheets(1)

            stats2324 = GetFileStats(ws2324)

            Set map2324 = BuildCockpitMap(ws2324)

            If ENABLE_TARGET_DUP_CHECK Then
                ValidateCockpitMap map2324
            Else
                LogMessage "[WARN] Target duplicate checks are DISABLED"
            End If

            For Each basePrefix In dict.Keys
                ProcessGroup ws2324, CStr(basePrefix), dict(basePrefix), map2324, groupName
                Set map2324 = BuildCockpitMap(ws2324)
            Next basePrefix

            wb2324.Save
            wb2324.Close False

            LogMessage "  " & FILE_2324 & ": " & stats2324

        Case "2020-2022"

            filePath2022 = TARGET_FOLDER & FILE_2022

            CreateBackup filePath2022

            Set wb2022 = OpenCockpit(FILE_2022)
            Set ws2022 = wb2022.Sheets(1)

            stats2022 = GetFileStats(ws2022)

            Set map2022 = BuildCockpitMap(ws2022)

            If ENABLE_TARGET_DUP_CHECK Then
                ValidateCockpitMap map2022
            Else
                LogMessage "[WARN] Target duplicate checks are DISABLED"
            End If

            For Each basePrefix In dict.Keys
                ProcessGroup ws2022, CStr(basePrefix), dict(basePrefix), map2022, groupName
                Set map2022 = BuildCockpitMap(ws2022)
            Next basePrefix

            wb2022.Save
            wb2022.Close False

            LogMessage "  " & FILE_2022 & ": " & stats2022

    End Select

End Sub

' =========================================================
' GET FILE STATS
' =========================================================

Private Function GetFileStats(ws As Worksheet) As String

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    GetFileStats = lastRow - 2 & " rows"

End Function

' =========================================================
' OPEN COCKPIT
' =========================================================

Private Function OpenCockpit(ByVal fileName As String) As Workbook

    Dim fullPath As String
    fullPath = TARGET_FOLDER & fileName

    If Dir(fullPath) = "" Then
        Err.Raise vbObjectError + 2000, , "Cockpit file not found: " & fileName
    End If

    Set OpenCockpit = Workbooks.Open(fullPath)

End Function

' =========================================================
' BUILD COCKPIT MAP
' =========================================================

Private Function BuildCockpitMap(ws As Worksheet) As Object

    Dim result As Object
    Dim info As Object
    Dim idxRows As Object

    Dim lastRow As Long
    Dim r As Long

    Dim fullPrefix As String
    Dim basePrefix As String
    Dim idxText As String
    Dim idx As Long

    Set result = CreateObject("Scripting.Dictionary")

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    For r = 3 To lastRow

        fullPrefix = Trim$(ws.Cells(r, 1).Value)

        If fullPrefix <> "" Then

            SplitPrefix fullPrefix, basePrefix, idxText

            If Not result.Exists(basePrefix) Then

                Set info = CreateObject("Scripting.Dictionary")
                Set idxRows = CreateObject("Scripting.Dictionary")

                info.Add "HasBase", False
                info.Add "BaseRow", 0
                info.Add "MaxIndex", 0
                info.Add "IndexRows", idxRows

                result.Add basePrefix, info

            End If

            Set info = result(basePrefix)
            Set idxRows = info("IndexRows")

            If idxText = "" Then

                If info("HasBase") Then
                    If ENABLE_TARGET_DUP_CHECK Then
                        Err.Raise vbObjectError + 2100, , _
                            "Duplicate base prefix in cockpit: " & basePrefix
                    Else
                        LogMessage "[WARN] Duplicate base prefix in cockpit (ignored): " & basePrefix
                    End If
                End If

                info("HasBase") = True
                info("BaseRow") = r

            Else

                If Not IsNumeric(idxText) Then
                    Err.Raise vbObjectError + 2200, , _
                        "Invalid index in cockpit: " & fullPrefix
                End If

                idx = CLng(idxText)

                If idxRows.Exists(idx) Then
                    If ENABLE_TARGET_DUP_CHECK Then
                        Err.Raise vbObjectError + 2101, , _
                            "Duplicate indexed prefix in cockpit: " & fullPrefix
                    Else
                        LogMessage "[WARN] Duplicate indexed prefix in cockpit (ignored): " & fullPrefix
                        GoTo NextR
                    End If
                End If

                idxRows.Add idx, r

                If idx > CLng(info("MaxIndex")) Then
                    info("MaxIndex") = idx
                End If

            End If

        End If

NextR:
    Next r

    Set BuildCockpitMap = result

End Function

' =========================================================
' VALIDATE COCKPIT MAP
' =========================================================

Private Sub ValidateCockpitMap(ByRef cockpitMap As Object)

    Dim basePrefix As Variant
    Dim info As Object

    For Each basePrefix In cockpitMap.Keys

        Set info = cockpitMap(basePrefix)

        If info("HasBase") And info("MaxIndex") > 0 Then

            Err.Raise vbObjectError + 2300, , _
                "Invalid cockpit structure. Base and indexed rows exist together: " _
                & basePrefix

        End If

    Next basePrefix

End Sub

' =========================================================
' FIND COCKPIT FOR PREFIX (2526 only)
' =========================================================

Private Function FindCockpitForPrefix(ByVal basePrefix As String, _
                                      ByRef map2026 As Object, _
                                      ByRef map2526 As Object) As Long

    Dim found2026 As Boolean
    Dim found2526 As Boolean

    found2026 = map2026.Exists(basePrefix)
    found2526 = map2526.Exists(basePrefix)

    If found2026 And found2526 Then

        If ENABLE_TARGET_DUP_CHECK Then
            Err.Raise vbObjectError + 3000, , _
                "Prefix exists in both cockpit files: " & basePrefix
        Else
            LogMessage "[WARN] Prefix exists in both cockpit files (using first): " & basePrefix
            FindCockpitForPrefix = 1
            Exit Function
        End If

    End If

    If Not found2026 And Not found2526 Then

        Err.Raise vbObjectError + 3001, , _
            "Prefix not found in cockpit: " & basePrefix

    End If

    If found2026 Then
        FindCockpitForPrefix = 1
    Else
        FindCockpitForPrefix = 2
    End If

End Function

' =========================================================
' PROCESS GROUP
' =========================================================

Private Sub ProcessGroup(ByVal ws As Worksheet, _
                         ByVal basePrefix As String, _
                         ByVal srcGroup As Object, _
                         ByRef cockpitMap As Object, _
                         ByVal groupName As String)

    Dim info As Object

    If Not cockpitMap.Exists(basePrefix) Then

        Err.Raise vbObjectError + 3100, , _
            "Prefix not found in Cockpit: " & basePrefix

    End If

    Set info = cockpitMap(basePrefix)

    If srcGroup.Exists(0) Then

        ProcessBaseRow ws, basePrefix, srcGroup, info, groupName

    Else

        ProcessIndexedGroup ws, basePrefix, srcGroup, info, groupName

    End If

End Sub

' =========================================================
' PROCESS BASE ROW
' =========================================================

Private Sub ProcessBaseRow(ByVal ws As Worksheet, _
                           ByVal basePrefix As String, _
                           ByVal srcGroup As Object, _
                           ByVal info As Object, _
                           ByVal groupName As String)

    Dim rowNum As Long

    If Not info("HasBase") Then

        Err.Raise vbObjectError + 3200, , _
            "Expected base row in cockpit: " & basePrefix

    End If

    rowNum = info("BaseRow")

    CopyRow ws, rowNum, srcGroup(0), groupName

    statsUpdated = statsUpdated + 1

    LogMessage "[UPDATE] Base: " & basePrefix

End Sub

' =========================================================
' PROCESS INDEXED GROUP
' =========================================================

Private Sub ProcessIndexedGroup(ByVal ws As Worksheet, _
                                ByVal basePrefix As String, _
                                ByVal srcGroup As Object, _
                                ByVal info As Object, _
                                ByVal groupName As String)

    Dim srcMax As Long
    Dim cockpitMax As Long

    srcMax = GetSourceMaxIndex(srcGroup)
    cockpitMax = CLng(info("MaxIndex"))

    If info("HasBase") Then

        ConvertBaseToGroup ws, basePrefix, info, srcMax

        Set info = BuildCockpitMap(ws)(basePrefix)

        cockpitMax = CLng(info("MaxIndex"))

    End If

    If cockpitMax < srcMax Then

        ValidateExpansion srcGroup, cockpitMax, basePrefix

        ExpandGroup ws, basePrefix, info, srcMax

        Set info = BuildCockpitMap(ws)(basePrefix)

    End If

    UpdateGroupRows ws, srcGroup, info, groupName

End Sub

' =========================================================
' GET SOURCE MAX INDEX
' =========================================================

Private Function GetSourceMaxIndex(ByVal srcGroup As Object) As Long

    Dim k As Variant
    Dim m As Long

    For Each k In srcGroup.Keys

        If CLng(k) > m Then
            m = CLng(k)
        End If

    Next k

    GetSourceMaxIndex = m

End Function

' =========================================================
' CONVERT BASE TO GROUP
' =========================================================

Private Sub ConvertBaseToGroup(ByVal ws As Worksheet, _
                               ByVal basePrefix As String, _
                               ByVal info As Object, _
                               ByVal srcMax As Long)

    Dim baseRow As Long
    Dim i As Long

    baseRow = info("BaseRow")

    ws.Cells(baseRow, 1).Value = basePrefix & "/1"

    For i = 2 To srcMax

        ws.Rows(baseRow + i - 2).Copy

        ws.Rows(baseRow + i - 1).Insert Shift:=xlDown

        ws.Cells(baseRow + i - 1, 1).Value = basePrefix & "/" & i

        LogMessage "[CREATE] " & basePrefix & "/" & i
        statsCreated = statsCreated + 1

    Next i

End Sub

' =========================================================
' EXPAND GROUP
' =========================================================

Private Sub ExpandGroup(ByVal ws As Worksheet, _
                        ByVal basePrefix As String, _
                        ByVal info As Object, _
                        ByVal targetMax As Long)

    Dim idxRows As Object
    Dim currentMax As Long
    Dim r As Long
    Dim i As Long

    Set idxRows = info("IndexRows")

    currentMax = info("MaxIndex")

    r = idxRows(currentMax)

    For i = currentMax + 1 To targetMax

        ws.Rows(r).Copy

        ws.Rows(r + 1).Insert Shift:=xlDown

        ws.Cells(r + 1, 1).Value = basePrefix & "/" & i

        LogMessage "[EXTEND] " & basePrefix & "/" & i
        statsCreated = statsCreated + 1

        r = r + 1

    Next i

End Sub

' =========================================================
' VALIDATE EXPANSION (no gaps)
' =========================================================

Private Sub ValidateExpansion(ByVal srcGroup As Object, _
                              ByVal cockpitMax As Long, _
                              ByVal basePrefix As String)

    Dim srcMax As Long
    Dim i As Long

    srcMax = GetSourceMaxIndex(srcGroup)

    If srcMax <= cockpitMax Then Exit Sub

    For i = cockpitMax + 1 To srcMax

        If Not srcGroup.Exists(i) Then

            Err.Raise vbObjectError + 4000, , _
                "Missing index /" & i & " before creation for base " & basePrefix

        End If

    Next i

End Sub

' =========================================================
' UPDATE GROUP ROWS
' =========================================================

Private Sub UpdateGroupRows(ByVal ws As Worksheet, _
                            ByVal srcGroup As Object, _
                            ByVal info As Object, _
                            ByVal groupName As String)

    Dim idxRows As Object
    Dim idx As Variant

    Set idxRows = info("IndexRows")

    For Each idx In srcGroup.Keys

        If CLng(idx) = 0 Then GoTo NextIdx

        If Not idxRows.Exists(CLng(idx)) Then

            Err.Raise vbObjectError + 4100, , _
                "Target index not found in cockpit: /" & idx

        End If

        CopyRow ws, idxRows(CLng(idx)), srcGroup(idx), groupName

        statsUpdated = statsUpdated + 1

        LogMessage "[UPDATE] " & ws.Cells(idxRows(CLng(idx)), 1).Value

NextIdx:
    Next idx

End Sub

' =========================================================
' COPY ROW
' =========================================================

Private Sub CopyRow(ByVal ws As Worksheet, _
                    ByVal r As Long, _
                    ByVal dataRow As Variant, _
                    ByVal groupName As String)

    If Not IsArray(dataRow) Then Exit Sub

    If groupName = "2025-2026" Then
        CopyBySpec ws, r, dataRow, "H,Q-AE,AG-AY,BA"
    Else
        CopyBySpec ws, r, dataRow, "O,W-AL,AO-BO,BY,CH"
    End If

End Sub

' =========================================================
' COPY BY SPEC
' =========================================================

Private Sub CopyBySpec(ByVal ws As Worksheet, _
                       ByVal r As Long, _
                       ByVal dataRow As Variant, _
                       ByVal spec As String)

    Dim parts, p
    Dim c1 As String, c2 As String
    Dim col As Long, startCol As Long, endCol As Long

    parts = Split(spec, ",")

    For Each p In parts

        p = Trim$(p)

        If InStr(p, "-") > 0 Then

            c1 = Split(p, "-")(0)
            c2 = Split(p, "-")(1)

            startCol = ws.Range(c1 & "1").Column
            endCol = ws.Range(c2 & "1").Column

            For col = startCol To endCol
                ws.Cells(r, col).Value = dataRow(1, col)
            Next col

        Else

            startCol = ws.Range(p & "1").Column
            ws.Cells(r, startCol).Value = dataRow(1, startCol)

        End If

    Next p

End Sub

' =========================================================
' SPLIT PREFIX
' =========================================================

Private Sub SplitPrefix(ByVal s As String, _
                        ByRef base As String, _
                        ByRef idx As String)

    Dim p As Long
    p = InStr(1, s, "/")

    If p > 0 Then
        base = Left$(s, p - 1)
        idx = Mid$(s, p + 1)
    Else
        base = s
        idx = ""
    End If

End Sub

' =========================================================
' GET YEAR GROUP
' =========================================================

Private Function GetYearGroup(ByVal p As String) As String

    Select Case Left$(p, 4)

        Case "2025", "2026": GetYearGroup = "2025-2026"
        Case "2023", "2024": GetYearGroup = "2023-2024"
        Case "2020", "2021", "2022": GetYearGroup = "2020-2022"
        Case Else: GetYearGroup = ""

    End Select

End Function
