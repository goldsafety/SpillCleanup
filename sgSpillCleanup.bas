Attribute VB_Name = "SpillCleanup"
Option Explicit

Public Sub sgSpillCleanup()
    'sgSpillCleanup (C)Copyright Stephen Goldsmith 2024-2025. All rights reserved.
    'Version 1.0.1 last updated January 2025
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
    
    Dim i As Integer, l As Long, s As String, lSpillRows As Long, rCell As Range, lRowsInserted As Long, lRowsDeleted As Long
    
    i = MsgBox("This procedure will insert or delete entire rows to resolve #SPILL! errors and to ensure only one blank row exists after a spill range. Unexpected results can occur if you have data either side of a spill range, including the deletion of that data. Make sure you save a copy before running this procedure as this cannot be undone. Are you sure you want to proceed?", vbYesNo Or vbQuestion Or vbDefaultButton2, "Spill Cleanup")
    If i <> vbYes Then
        Exit Sub
    End If
    
    'There are a couple of methods to find the overall data range in the current worksheet, see:
    'https://learn.microsoft.com/en-us/office/vba/excel/concepts/cells-and-ranges/select-a-range
    'This version of the procedure uses ActiveSheet.UsedRange but another option is to find the last data
    'using the following:
    'lLastRow = ActiveSheet.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    'lLastCol = ActiveSheet.Cells.Find("*", SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
    
    'Insert rows to resolve any spill errors
    For Each rCell In ActiveSheet.UsedRange
        If IsError(rCell.Value) Then
            If rCell.Value = CVErr(2045) Then 'Error 2045 is #SPILL!
                s = rCell.Offset(0, 1).Address & ":" & ActiveSheet.Cells(rCell.Row + 1, ActiveSheet.UsedRange.Column + ActiveSheet.UsedRange.Columns.Count - 1).Address
                If Application.WorksheetFunction.CountA(ActiveSheet.Range(s)) > 0 Then
                    MsgBox "Cannot complete spill cleanup. There is data to the right of a spill range (" & s & "). A total of " & lRowsInserted & " row(s) have already been inserted and " & lRowsDeleted & " row(s) deleted.", vbCritical, "Spill Cleanup"
                    Exit Sub
                End If
                Do
                    ActiveSheet.Rows(rCell.Row + 1).Insert xlShiftDown, xlFormatFromLeftOrAbove
                    lRowsInserted = lRowsInserted + 1
                    If lRowsInserted Mod 1000 = 0 Then
                        i = MsgBox("A large number of rows (" & lRowsInserted & ") have been inserted, which might indicate a problem. Do you wish to continue?", vbYesNo Or vbQuestion Or vbDefaultButton2, "Spill Cleanup")
                        If i <> vbYes Then
                            Exit Sub
                        End If
                    End If
                Loop Until IsError(rCell.Value) = False
            End If
        End If
    Next
    
    'Make sure there is one blank row after every spilling range, or FILTER or SORT dynamic array that is not spilling (either only one match or no matches)
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
                lRowsInserted = lRowsInserted + 1
            Else
                'Delete extra blank rows
                Do
                    l = l + 1
                Loop Until l > ActiveSheet.UsedRange.Row + ActiveSheet.UsedRange.Rows.Count Or Application.WorksheetFunction.CountA(ActiveSheet.Rows(l)) <> 0
                If (l - rCell.Row - lSpillRows - 1) > 0 Then
                    ActiveSheet.Range((rCell.Row + 1) & ":" & (l - lSpillRows - 1)).Delete
                    lRowsDeleted = lRowsDeleted + (l - rCell.Row - lSpillRows - 1)
                End If
            End If
        End If
    Next
    
    MsgBox "Cleanup completed. A total of " & lRowsInserted & " row(s) were inserted and " & lRowsDeleted & " row(s) were deleted.", vbInformation, "Spill Cleanup"
End Sub
