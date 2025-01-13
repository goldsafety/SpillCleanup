Attribute VB_Name = "sgSpillCleanup"
Option Explicit

Private Const KSpillInsertRowCount = 1

Public Sub SpillCleanup()
    'SpillCleanup (C)Copyright Stephen Goldsmith 2024-2025. All rights reserved.
    'Version 1.0.6 last updated January 2025
    'Distributed at https://github.com/goldsafety/ and https://aircraftsystemsafety.com/code/
    
    'Microsoft(R) Excel(R) VBA script to cleanup spilled ranges from dynamic array formulas by resolving
    '#SPILL! errors and removing blank rows.
    
    'Eclipse Public License - v 2.0
    'THE ACCOMPANYING PROGRAM IS PROVIDED UNDER THE TERMS OF THIS ECLIPSE PUBLIC LICENSE (“AGREEMENT”).
    'ANY USE, REPRODUCTION OR DISTRIBUTION OF THE PROGRAM CONSTITUTES RECIPIENT'S ACCEPTANCE OF THIS AGREEMENT.
    'https://www.eclipse.org/legal/epl-2.0/
    
    'This script has been written to automate the process of resolving #SPILL! errors and removing blank rows
    'below spilled ranges where data used to reside. It assumes that spilled ranges will only change the number
    'of rows being returned, and inserts or deletes entire rows below selected dynamic array formulas until
    'each #SPILL! error has been resolved and only a single blank row exists underneath it. A blank row is
    'defined as a row without any data, even if it contains formatting. The script further handles disabling
    'worksheet protection if set, ensuring formatting of the first row is reflected in inserted rows, and
    'wrapping text where it has been set.
    
    'Dynamic array formulas have some significant limitations, as they can be slow when you have many formulas
    'and they cannot spill into merged cells, which can make vertical layout of a report challenging. When you
    'need more control to produce a dynamic report, try ProtoSheet. This Excel VBA script requires you to
    'layout a prototype worksheet which it then uses to construct a completed version in a new worksheet on
    'demand. To find out more and to try this option, visit the following site:
    'https://github.com/goldsafety/ProtoSheet/
    
    'Known limitations of SpillCleanup are that it currently only resolves #SPILL! errors caused by data in
    'rows below the formula. If the #SPILL! error is caused by data or merged cells to the right of the dynamic
    'array formula, it will either raise an error or attempt to insert many rows until it figures something has
    'gone wrong (at which point a hundred or more rows will already have been inserted). In addition, only
    '#SPILL! errors caused by the FILTER or SORT dynamic array functions will be resolved, though it is
    'relatively simple to add support for others (please let me know).
    
    'History
    '1.0.0  First public release
    '1.0.1  Delete rows in one block rather than line by line
    '1.0.2  Insert large blocks of 100 rows rather than line by line until the spill error is resolved
    '1.0.3  Added progress message to statusbar and toggle wrap text
    '1.0.4  Changed wrap text to first column only instead of entire row
    '1.0.5  Set calculation mode to manual and only calculate the spill range being updated to improve speed
    '1.0.6  Add global error handler, and remove worksheet protection if set
    
    Dim i As Integer, l As Long, s As String
    Dim rCell As Range, lBlocksInserted As Long, lBlocksDeleted As Long, lSpillRows As Long, lSpillRanges As Long, lSpillRange As Long
    Dim bStatusBarState As Boolean, lCalcMode As Long, sStartTime As Single
    
    'Confirmation dialog
    i = MsgBox("This procedure will insert or delete entire rows to resolve #SPILL! errors and to ensure only one blank row exists after a spilled range. Unexpected results can occur if you have data either side of a spilled range, including the deletion of that data. Make sure you save a copy before running this procedure as this cannot be undone. If you have a large number of dynamic array formulas, this procedure can take a few minutes. Are you sure you want to proceed?", vbYesNo Or vbQuestion Or vbDefaultButton2, "Spill Cleanup")
    If i <> vbYes Then
        Exit Sub
    End If
    
    'Record start time
    sStartTime = Timer
    
    'Set the status bar to report progress
    bStatusBarState = Application.DisplayStatusBar
    Application.DisplayStatusBar = True
    Application.StatusBar = "Spill Cleanup: 0% complete (elapsed time 0 seconds)"
    
    'Set the calculation mode to manual in order to reduce the time to complete
    lCalcMode = Application.Calculation
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    Application.Calculate
    
    'If the worksheet is protected, modify the protection to allow editing by a macro (this resets on next open)
    If ActiveSheet.ProtectContents = True And ActiveSheet.ProtectionMode = False Then
        s = ""
        On Error Resume Next
            ActiveSheet.Unprotect ""
        If Err.Number = 1004 Then
            s = InputBox("Password:", "Unprotect Sheet")
            ActiveSheet.Unprotect s
            If Err.Number > 0 Then
                On Error GoTo Err_SpillCleanup
                Err.Raise vbObjectError + 1, , "The password you supplied is not correct"
            End If
        End If
        On Error GoTo Err_SpillCleanup
        If s = "" Then
            ActiveSheet.Protect UserInterfaceOnly:=True
        Else
            ActiveSheet.Protect Password:=s, UserInterfaceOnly:=True
        End If
    End If
    
    'Set a global error handler
    On Error GoTo Err_SpillCleanup
    
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
                    Err.Raise vbObjectError + 1, , "There is data to the right of a spill range (within " & s & ")"
                End If
                l = 0
                Do
                    l = l + 1
                    ActiveSheet.Rows(rCell.Row + 1).Insert xlShiftDown, xlFormatFromLeftOrAbove
                    If l = 1 Then
                        'If this is the first row being inserted, copy the first row but clear the contents in order to copy the formatting
                        '(such as font styles, gridlines, and word wrapping settings).
                        ActiveSheet.Rows(rCell.Row).Copy ActiveSheet.Rows(rCell.Row + 1)
                        ActiveSheet.Rows(rCell.Row + 1).ClearContents
                    End If
                    If KSpillInsertRowCount > 1 Then
                        'Inserting one row at a time can take a longer time as we need to recalculate after each insert to see if the #SPILL!
                        'error has been resolved. To reduce the time taken, we can optionally insert multiple rows before recalculating, and allow
                        'any unnecessary empty rows to be deleted in the next phase of this procedure.
                        ActiveSheet.Rows(rCell.Row + 1 & ":" & rCell.Row + KSpillInsertRowCount - 1).Insert xlShiftDown, xlFormatFromLeftOrAbove
                    End If
                    'Update just this spill range before checking if no longer an error
                    rCell.Calculate
                    'Ask to continue if we keep inserting lots of rows but the #SPILL! error has not been resolved
                    If IsError(rCell.Value) And ((KSpillInsertRowCount < 10 And l Mod 500 = 0) Or (KSpillInsertRowCount >= 10 And KSpillInsertRowCount < 100 And l Mod 50 = 0) Or (KSpillInsertRowCount >= 100 And l Mod 5 = 0)) Then
                        i = MsgBox("A large number of rows (" & (l * KSpillInsertRowCount) & ") have been inserted whilst trying to resolve a #SPILL! error, which might indicate a problem. Do you wish to continue?", vbYesNo Or vbQuestion Or vbDefaultButton2, "Spill Cleanup")
                        If i <> vbYes Then
                            Err.Raise vbObjectError + 1, , "User cancelled"
                        End If
                    End If
                Loop Until IsError(rCell.Value) = False
                lBlocksInserted = lBlocksInserted + 1
                lSpillRange = lSpillRange + 1
                Application.StatusBar = "Spill Cleanup: " & Int((lSpillRange / lSpillRanges) * 50) & "% complete (elapsed time " & Int(Timer - sStartTime) & " seconds)"
                DoEvents
            End If
        End If
    Next
    
    'Do a full recalculate before moving on to deleting blank rows
    Application.Calculate
    
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
                    ActiveSheet.Range(rCell.Address & ":" & rCell.Offset(rCell.SpillingToRange.Rows.Count - 1).Address).WrapText = False
                    ActiveSheet.Range(rCell.Address & ":" & rCell.Offset(rCell.SpillingToRange.Rows.Count - 1).Address).WrapText = True
                Else
                    rCell.WrapText = False
                    rCell.WrapText = True
                End If
            End If
            
            'Update status bar
            lSpillRange = lSpillRange + 1
            Application.StatusBar = "Spill Cleanup: " & Int(50 + (lSpillRange / lSpillRanges) * 50) & "% complete (elapsed time " & Int(Timer - sStartTime) & " seconds)"
            DoEvents
        End If
    Next
    
    On Error GoTo 0
Err_SpillCleanup:
    
    'Do a full update and then restore calculation mode
    Application.Calculate
    Application.Calculation = lCalcMode
    Application.ScreenUpdating = True
    
    'Display final progress on status bar
    'Application.StatusBar = "Spill Cleanup: 100% complete (elapsed time " & Int(Timer - sStartTime) & " seconds)"
    
    If Err.Number = 0 Then
        MsgBox "Cleanup completed. A total of " & lBlocksInserted & " block(s) of rows were inserted and " & lBlocksDeleted & " block(s) of rows were deleted.", vbInformation, "Spill Cleanup"
    Else
        'MsgBox "Cannot complete spill cleanup. " & Err.Description & ". A total of " & lBlocksInserted & " row(s) have already been inserted and " & lBlocksDeleted & " row(s) deleted.", vbCritical, "Spill Cleanup"
        MsgBox "Cannot complete spill cleanup. " & Err.Description & ". A total of " & lBlocksInserted & " block(s) of rows were inserted and " & lBlocksDeleted & " block(s) of rows were deleted.", vbCritical, "Spill Cleanup"
    End If
    
    'Reset status bar
    Application.StatusBar = False
    Application.DisplayStatusBar = bStatusBarState
End Sub
