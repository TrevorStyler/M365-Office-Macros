Attribute VB_Name = "SG_Metadata_Export_Migration"
Option Explicit

'===============================================================================
' SG_Metadata_Export_Migration: ShareGate metadata export cleanup for migration mapping
'
' Version: 1
' Date: August 15, 2025
' Authors: Trevor Styler; ChatGPT 5
' Website: trevor.styler.ca
' Repo: https://github.com/TrevorStyler/M365-Office-Macros/tree/main/Excel
'
' What it does (whitelist flow):
'   • Normalize a raw export (no detection/aliasing)
'   • Rename key headers to canonical names
'   • Create "Destination Folder" (if missing)
'   • Clear values in "Destination Library"
'   • Keep ONLY the columns listed in `keep`, in that order (skip missing)
'   • Delete every other column
'===============================================================================

Sub SG_Metadata_Export_Migration()
    Dim ws As Worksheet
    Dim firstRow As Long, firstCol As Long, lastRow As Long, lastCol As Long
    Dim src As Range, rng As Range
    Dim lo As ListObject
    Dim c As Long, r As Long
    Dim oldCalc As XlCalculation, oldScreen As Boolean, oldEvents As Boolean, oldAlerts As Boolean

    Set ws = ActiveSheet
    If Application.WorksheetFunction.CountA(ws.Cells) = 0 Then Exit Sub

    ' Speed / safety
    oldCalc = Application.Calculation
    oldScreen = Application.ScreenUpdating
    oldEvents = Application.EnableEvents
    oldAlerts = Application.DisplayAlerts
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False

    On Error GoTo CleanExit

    ' Unhide & clear filters
    On Error Resume Next
    ws.Rows.Hidden = False
    ws.Columns.Hidden = False
    If ws.AutoFilterMode Then ws.AutoFilterMode = False
    On Error GoTo 0

    ' Remove fully empty columns/rows
    If Not ws.Cells.Find(What:="*", LookIn:=xlFormulas) Is Nothing Then
        lastCol = ws.Cells.Find(What:="*", LookIn:=xlFormulas, _
                                SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
        For c = lastCol To 1 Step -1
            If Application.WorksheetFunction.CountA(ws.Columns(c)) = 0 Then ws.Columns(c).Delete
        Next c

        lastRow = ws.Cells.Find(What:="*", LookIn:=xlFormulas, _
                                SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
        For r = lastRow To 1 Step -1
            If Application.WorksheetFunction.CountA(ws.Rows(r)) = 0 Then ws.Rows(r).Delete
        Next r
    End If

    ' Shift used block to A1 if needed
    firstRow = ws.Cells.Find(What:="*", LookIn:=xlFormulas, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
                             After:=ws.Cells(ws.Rows.Count, ws.Columns.Count)).Row
    firstCol = ws.Cells.Find(What:="*", LookIn:=xlFormulas, SearchOrder:=xlByColumns, SearchDirection:=xlNext, _
                             After:=ws.Cells(ws.Rows.Count, ws.Columns.Count)).Column
    lastRow = ws.Cells.Find(What:="*", LookIn:=xlFormulas, SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    lastCol = ws.Cells.Find(What:="*", LookIn:=xlFormulas, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column

    Set src = ws.Range(ws.Cells(firstRow, firstCol), ws.Cells(lastRow, lastCol))
    If firstRow > 1 Or firstCol > 1 Then
        src.Cut Destination:=ws.Cells(1, 1)
        lastRow = ws.Cells.Find(What:="*", LookIn:=xlFormulas, SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
        lastCol = ws.Cells.Find(What:="*", LookIn:=xlFormulas, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
    End If
    Set rng = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol))

    ' Build a fresh table
    Do While ws.ListObjects.Count > 0
        ws.ListObjects(1).Unlist
    Loop
    Set lo = ws.ListObjects.Add(xlSrcRange, rng, , xlYes)
    lo.name = "Table1"
    lo.TableStyle = "TableStyleMedium13"

    ' ---- Canonicalize exact headers (no fuzzy aliasing) ----
    ' Folder or Filename
    If Not TryRename(lo, "Column 1", "Folder or Filename") Then
        Dim idIdx As Long
        idIdx = FindHeaderIndex(lo, "ID")
        If idIdx > 0 And idIdx < lo.ListColumns.Count Then
            lo.ListColumns(idIdx + 1).name = "Folder or Filename"
        End If
    End If
    ' Other exact renames
    TryRename lo, "ContentType", "Content Type"
    TryRename lo, "SourcePath", "Source Location"
    TryRename lo, "DestinationPath", "Destination Library"

    ' Ensure Destination Folder exists right after Destination Library
    Dim dLibIdx As Long
    dLibIdx = FindHeaderIndex(lo, "Destination Library")
    If dLibIdx > 0 Then
        If FindHeaderIndex(lo, "Destination Folder") = 0 Then
            Dim newCol As ListColumn
            Set newCol = lo.ListColumns.Add(Position:=dLibIdx + 1)
            newCol.name = "Destination Folder"
        End If
    End If

    ' Clear values in Destination Library (keep header)
    Dim dLibCol As ListColumn
    On Error Resume Next
    Set dLibCol = lo.ListColumns("Destination Library")
    On Error GoTo 0
    If Not dLibCol Is Nothing Then
        If Not dLibCol.DataBodyRange Is Nothing Then dLibCol.DataBodyRange.ClearContents
    End If

    ' ---- Keep ONLY these (in this order). Missing ones are skipped. ----
    Dim keep As Variant
    keep = Array( _
        "Content Type", _
        "Source Location", _
        "Folder or Filename", _
        "Destination Library", _
        "Destination Folder", _
        "Created By", _
        "Created", _
        "Modified By", _
        "Modified" _
    )

    ' Order the kept columns that exist
    OrderColumnsPresent lo, keep

    ' Delete everything NOT in the keep list
    KeepOnlyColumns lo, keep

    rng.EntireColumn.AutoFit

CleanExit:
    Application.Calculation = oldCalc
    Application.ScreenUpdating = oldScreen
    Application.EnableEvents = oldEvents
    Application.DisplayAlerts = oldAlerts
    Application.CutCopyMode = False
End Sub

'================ helpers ================

Private Sub OrderColumnsPresent(lo As ListObject, desired As Variant)
    Dim i As Long, leftMost As String
    ' Move the first present desired column to column 1
    For i = LBound(desired) To UBound(desired)
        If FindHeaderIndex(lo, CStr(desired(i))) > 0 Then
            leftMost = HeaderNameAtIndex(lo, 1)
            If LCase$(leftMost) <> LCase$(CStr(desired(i))) Then
                MoveListColumnBefore lo, CStr(desired(i)), leftMost
            End If
            Exit For
        End If
    Next i

    ' Then chain the rest (place each existing one after the previous existing kept)
    Dim prevName As String, curName As String
    prevName = ""
    For i = LBound(desired) To UBound(desired)
        curName = CStr(desired(i))
        If FindHeaderIndex(lo, curName) > 0 Then
            If prevName <> "" Then
                If FindHeaderIndex(lo, prevName) > 0 Then
                    MoveListColumnAfter lo, curName, prevName
                End If
            End If
            prevName = curName
        End If
    Next i
End Sub

Private Sub KeepOnlyColumns(lo As ListObject, keep As Variant)
    Dim j As Long
    For j = lo.ListColumns.Count To 1 Step -1
        If Not IsInList(lo.ListColumns(j).name, keep) Then
            lo.ListColumns(j).Delete
        End If
    Next j
End Sub

Private Function IsInList(ByVal name As String, keep As Variant) As Boolean
    Dim i As Long
    For i = LBound(keep) To UBound(keep)
        If LCase$(Trim$(CStr(keep(i)))) = LCase$(Trim$(name)) Then
            IsInList = True
            Exit Function
        End If
    Next i
End Function

Private Function TryRename(lo As ListObject, ByVal fromName As String, ByVal toName As String) As Boolean
    On Error Resume Next
    lo.ListColumns(fromName).name = toName
    TryRename = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
End Function

Private Function FindHeaderIndex(lo As ListObject, ByVal headerName As String) As Long
    Dim c As Range
    For Each c In lo.HeaderRowRange.Cells
        If LCase$(Trim$(CStr(c.Value2))) = LCase$(Trim$(headerName)) Then
            FindHeaderIndex = c.Column - lo.HeaderRowRange.Column + 1
            Exit Function
        End If
    Next c
End Function

Private Function HeaderNameAtIndex(lo As ListObject, idx As Long) As String
    If idx >= 1 And idx <= lo.ListColumns.Count Then
        HeaderNameAtIndex = CStr(lo.ListColumns(idx).name)
    End If
End Function

Private Function MoveListColumnBefore(lo As ListObject, ByVal sourceName As String, ByVal beforeName As String) As Boolean
    Dim iSrc As Long, iBefore As Long
    iSrc = FindHeaderIndex(lo, sourceName)
    iBefore = FindHeaderIndex(lo, beforeName)
    If iSrc = 0 Or iBefore = 0 Or iSrc = iBefore Then Exit Function
    lo.ListColumns(iSrc).Range.Cut
    lo.ListColumns(iBefore).Range.Insert Shift:=xlToRight
    MoveListColumnBefore = True
End Function

' Robust mover that avoids 1004 by using captured ranges and no-op when already adjacent
Private Function MoveListColumnAfter(lo As ListObject, ByVal sourceName As String, ByVal afterName As String) As Boolean
    Dim iSrc As Long, iAfter As Long
    Dim srcR As Range, beforeR As Range
    Dim tmp As ListColumn

    iSrc = FindHeaderIndex(lo, sourceName)
    iAfter = FindHeaderIndex(lo, afterName)
    If iSrc = 0 Or iAfter = 0 Then Exit Function

    ' No-op if source already immediately follows target
    If iSrc = iAfter + 1 Then
        MoveListColumnAfter = True
        Exit Function
    End If

    ' Decide the anchor range BEFORE the cut
    If iAfter + 1 > lo.ListColumns.Count Then
        Set tmp = lo.ListColumns.Add(Position:=lo.ListColumns.Count + 1)
        tmp.name = "__tmp_anchor__"
        Set beforeR = tmp.Range
    Else
        Set beforeR = lo.ListColumns(iAfter + 1).Range
    End If

    ' Capture source range BEFORE cutting
    Set srcR = lo.ListColumns(iSrc).Range

    ' Cut & insert
    srcR.Cut
    beforeR.Insert Shift:=xlToRight

    ' Clean up temp anchor (if any)
    On Error Resume Next
    lo.ListColumns("__tmp_anchor__").Delete
    On Error GoTo 0

    MoveListColumnAfter = True
End Function


