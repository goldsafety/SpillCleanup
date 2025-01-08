Attribute VB_Name = "sgSpillCleanup"
Option Explicit

Public Sub SpillCleanup()
    'SpillCleanup (C)Copyright Stephen Goldsmith 2024-2025. All rights reserved.
    'Version 1.0.3 last updated January 2025
    'Distributed at https://github.com/goldsafety/ and https://aircraftsystemsafety.com/code/
    
    'Excel VBA script to cleanup spilling from dynamic arrays by resolving #SPILL! errors and removing blank rows.
    
    'Eclipse Public License - v 2.0
    'THE ACCOMPANYING PROGRAM IS PROVIDED UNDER THE TERMS OF THIS ECLIPSE PUBLIC LICENSE (“AGREEMENT”).
    'ANY USE, REPRODUCTION OR DISTRIBUTION OF THE PROGRAM CONSTITUTES RECIPIENT'S ACCEPTANCE OF THIS AGREEMENT.
    'https://www.eclipse.org/legal/epl-2.0/
    
    'Dynamic arrays and 'spilling' data from functions is a great new feature in Microsoft Excel that is much
    'easier to use than Control+Shift+Enter (CSE) array formulas. Many of the built-in functions now return
    'multiple rows of data which automatically 'spills' into neighbouring blank cells. As the data changes the
    'spilled data will automatically update, which may affect the size of the spilled range. If other data
    'already exists in the neigbouring cells you will see a #SPILL! error. As data updates, the spill range
    'may also get smaller leaving potentially large blank areas in your worksheet.
    
    'This procedure simply runs down the active worksheet and inserts or deletes rows to push data outside of
    'the spill range to resolve #SPILL! errors and then to insert or delete rows immediately below the spill
    'range to ensure only a single blank row exists after the spill range. A blank row is defined as a row
    'without any data, even if it contains formatting.
    
    'This procedure is particularly intended to be used with the built-in FILTER function, and only considers
    'that the number of rows returned is changing. If a #SPILL! error is being caused by data to the right of
    'a spill range rather than below it, unexpected results may occur.
    
    Dim i As Integer, l As Long, s As String, lSpillRows As Long, rCell As Range, lBlocksInserted As Long, lBlocksDeleted As Long
    Dim bStatusBarState As Boolean, lSpillRanges As Long, lSpillRange As Long
    
    i = MsgBox("This procedure will insert or delete entire rows to resolve #SPILL! errors and to ensure only one blank row exists after a spill range. Unexpected results can occur if you have data either side of a spill range, including the deletion of that data. Make sure you save a copy before running this procedure as this cannot be undone. If you have a large number of dynamic arrays, this procedure can take several minutes. Are you sure you want to proceed?", vbYesNo Or vbQuestion Or vbDefaultButton2, "Spill Cleanup")
    If i <> vbYes Then
        Exit Sub
    End If
    
    bStatusBarState = Application.DisplayStatusBar
    Application.DisplayStatusBar = True
    
    'There are a couple of methods to find the overall data range in the current worksheet, see:
    'https://learn.microsoft.com/en-us/office/vba/excel/concepts/cells-and-ranges/select-a-range
    'This version of the procedure uses ActiveSheet.UsedRange but another option is to find the last data
    'using the following:
    'lLastRow = ActiveSheet.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    'lLastCol = ActiveSheet.Cells.Find("*", SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
    
    'Find out how many spill ranges there are
    lSpillRanges = 0
    For Each rCell In ActiveSheet.UsedRange
        If IsError(rCell.Value) Then
            If rCell.Value = CVErr(2045) Then 'Error 2045 is #SPILL!
                lSpillRanges = lSpillRanges + 1
            End If
        End If
    Next
    
    'Insert rows to resolve any spill errors
    lSpillRange = 0
    For Each rCell In ActiveSheet.UsedRange
        If IsError(rCell.Value) Then
            If rCell.Value = CVErr(2045) Then 'Error 2045 is #SPILL!
                s = rCell.Offset(0, 1).Address & ":" & ActiveSheet.Cells(rCell.Row + 1, ActiveSheet.UsedRange.Column + ActiveSheet.UsedRange.Columns.Count - 1).Address
                If Application.WorksheetFunction.CountA(ActiveSheet.Range(s)) > 0 Then
                    MsgBox "Cannot complete spill cleanup. There is data to the right of a spill range (" & s & "). A total of " & lBlocksInserted & " row(s) have already been inserted and " & lBlocksDeleted & " row(s) deleted.", vbCritical, "Spill Cleanup"
                    Exit Sub
                End If
                Do
                    'Inserting one row at a time takes a long time to recalculate (we can't set calculation to manual as we need to know if
                    'the spill error has been resolved). As such, insert in blocks of 100 (we will delete the extra rows later) but first
                    'make a copy of the first row (clearing contents) to ensure we copy the formatting (such as gridlines).
                    ActiveSheet.Rows(rCell.Row + 1).Insert xlShiftDown, xlFormatFromLeftOrAbove
                    ActiveSheet.Rows(rCell.Row).Copy ActiveSheet.Rows(rCell.Row + 1)
                    ActiveSheet.Rows(rCell.Row + 1).ClearContents
                    ActiveSheet.Rows(rCell.Row + 1 & ":" & rCell.Row + 100).Insert xlShiftDown, xlFormatFromLeftOrAbove
                    lBlocksInserted = lBlocksInserted + 1
                    If lBlocksInserted Mod 100 = 0 Then
                        i = MsgBox("A large number of rows (in " & lBlocksInserted & " blocks) have been inserted, which might indicate a problem. Do you wish to continue?", vbYesNo Or vbQuestion Or vbDefaultButton2, "Spill Cleanup")
                        If i <> vbYes Then
                            Exit Sub
                        End If
                    End If
                Loop Until IsError(rCell.Value) = False
                lSpillRange = lSpillRange + 1
                Application.StatusBar = "SpillCleanup: " & Int((lSpillRange / lSpillRanges) * 50) & "% complete"
                DoEvents
            End If
        End If
    Next
    
    'Find out how many spill ranges there are
    lSpillRanges = 0
    For Each rCell In ActiveSheet.UsedRange
        lSpillRows = 0 'This variable holds how many rows are being spilled or zero if this is not a spilling range
        If rCell.HasSpill = True Then
            If rCell.Address = rCell.SpillParent.Address Then
                lSpillRows = rCell.SpillingToRange.Rows.Count
            End If
        ElseIf rCell.Formula <> "" Then
            If Left(rCell.Formula, 8) = "=FILTER(" Or Left(rCell.Formula, 6) = "=SORT(" Then
                lSpillRows = 1
            End If
        End If
        If lSpillRows > 0 Then
            lSpillRanges = lSpillRanges + 1
        End If
    Next
    
    'Make sure there is one blank row after every spilling range, or FILTER or SORT dynamic array that is not spilling (either only one match or no matches)
    lSpillRange = 0
    For Each rCell In ActiveSheet.UsedRange
        lSpillRows = 0 'This variable holds how many rows are being spilled or zero if this is not a spilling range
        If rCell.HasSpill = True Then
            If rCell.Address = rCell.SpillParent.Address Then
                lSpillRows = rCell.SpillingToRange.Rows.Count
            End If
        ElseIf rCell.Formula <> "" Then
            If Left(rCell.Formula, 8) = "=FILTER(" Or Left(rCell.Formula, 6) = "=SORT(" Then
                lSpillRows = 1
            End If
        End If
        If lSpillRows > 0 Then
            l = rCell.Row + lSpillRows
            If Application.WorksheetFunction.CountA(ActiveSheet.Rows(l)) <> 0 Then
                'Insert a blank row
                ActiveSheet.Rows(rCell.Row + 1).Insert xlShiftDown, xlFormatFromLeftOrAbove
                lBlocksInserted = lBlocksInserted + 1
            Else
                'Delete extra blank rows
                Do
                    l = l + 1
                Loop Until l > ActiveSheet.UsedRange.Row + ActiveSheet.UsedRange.Rows.Count Or Application.WorksheetFunction.CountA(ActiveSheet.Rows(l)) <> 0
                If (l - rCell.Row - lSpillRows - 1) > 0 Then
                    ActiveSheet.Rows((rCell.Row + 1) & ":" & (l - lSpillRows - 1)).Delete
                    lBlocksDeleted = lBlocksDeleted + 1
                End If
            End If
            
            'Excel does not seem to update the word wrapping for the dynamic arrays so we must toggle it manually, however
            'we only seem to need to toggle the first column and the other columns will update also.
            If rCell.WrapText = True Then
                If rCell.HasSpill = True Then
                    ActiveSheet.Rows(rCell.Row & ":" & rCell.Row + rCell.SpillingToRange.Rows.Count - 1).WrapText = False
                    ActiveSheet.Rows(rCell.Row & ":" & rCell.Row + rCell.SpillingToRange.Rows.Count - 1).WrapText = True
                Else
                    rCell.WrapText = False
                    rCell.WrapText = True
                End If
            End If
            
            'Update status bar
            lSpillRange = lSpillRange + 1
            Application.StatusBar = "SpillCleanup: " & Int(50 + (lSpillRange / lSpillRanges) * 50) & "% complete"
            DoEvents
        End If
    Next
    
    Application.StatusBar = False
    Application.DisplayStatusBar = bStatusBarState
    
    MsgBox "Cleanup completed. A total of " & lBlocksInserted & " block(s) of rows were inserted and " & lBlocksDeleted & " block(s) of rows were deleted.", vbInformation, "Spill Cleanup"
End Sub
