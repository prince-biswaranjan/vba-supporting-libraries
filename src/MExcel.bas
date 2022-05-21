Attribute VB_Name = "MExcel"
'@IgnoreModule ProcedureNotUsed, LineLabelNotUsed, ConstantNotUsed
'==================================================================================================
'   Module          :   MExcel
'   Type            :   Module
'   Description     :   Procedures to manipulate excel worksheet
'--------------------------------------------------------------------------------------------------
'   Procedures      :   HideUnusedRowsInSheet       Void
'                       GetNextEmptyRange           Range
'                       ReplaceFormulaWithValue     Void
'--------------------------------------------------------------------------------------------------
'   References      :   NA
'   Dependencies    :   NA
'==================================================================================================

'--------------------------------------------------------------------------------------------------
'   Option Statements
'--------------------------------------------------------------------------------------------------
Option Explicit


'--------------------------------------------------------------------------------------------------
'   Constant Declarations
'--------------------------------------------------------------------------------------------------

'   Module Level
Private Const moduleName    As String = "MExcel"

Public Sub HideUnusedRowsInSheet(ByVal targetWorksheet As Worksheet)
'==================================================================================================
'   Description :   Hides rows outside of used range
'   Parameters  :   targetWorksheet    -   The worksheet where unused rows are to be hidded
'   Returns     :   NA
'==================================================================================================
    
    Const procName      As String = "HideUnusedRows"
    Dim rangeWithData   As Range

    '----------------------------------------------------------------------------------------------
    
    With targetWorksheet
        .Cells.EntireColumn.Hidden = False
        .Cells.EntireRow.Hidden = False
        Set rangeWithData = .Range("A1").Offset(.UsedRange.Rows.Count + 1, .UsedRange.columns.Count + 1)
    End With
    
    With targetWorksheet.Range(rangeWithData, rangeWithData.End(xlDown).End(xlToRight))
        .EntireColumn.Hidden = True
        .EntireRow.Hidden = True
    End With
    
    '----------------------------------------------------------------------------------------------
    
End Sub

Public Function GetNextEmptyCellInColumn(ByVal startRange As Range) As Range
'==================================================================================================
'   Description :   Returns the next empty cell in column
'   Parameters  :   startRange  -   Range from where the next empty cell will be calculated
'   Returns     :   Range
'==================================================================================================

    GetNextEmptyRange startRange, True
    
End Function

Public Function GetNextEmptyCellInRow(ByVal startRange As Range) As Range
'==================================================================================================
'   Description :   Returns the next empty cell in row
'   Parameters  :   startRange  -   Range from where the next empty cell will be calculated
'   Returns     :   Range
'==================================================================================================
    
    GetNextEmptyRange startRange, False
    
End Function

Private Function GetNextEmptyRange(ByVal startRange As Range, ByVal inColumn As Boolean) As Range
'==================================================================================================
'   Description :   Returns the next empty range in column or row
'   Parameters  :   startRange  -   Range from where the next empty range will be calculated
'                   inColumn    -   True finds in the same column and false finds in the same row
'   Returns     :   Range
'==================================================================================================
    
    Const procName  As String = "GetNextEmptyRange"
    Dim rngEmpty    As Range
    
    On Error GoTo PROC_ERR
    
    '----------------------------------------------------------------------------------------------
    
    If inColumn Then
        If IsEmpty(startRange.Offset(1, 0)) Then
            Set rngEmpty = startRange.Offset(1, 0)
        Else
            Set rngEmpty = startRange.End(xlDown).Offset(1, 0)
        End If
    Else
        If IsEmpty(startRange.Offset(0, 1)) Then
            Set rngEmpty = startRange.Offset(0, 1)
        Else
            Set rngEmpty = startRange.End(xlToRight).Offset(0, 1)
        End If
    End If
    
    '----------------------------------------------------------------------------------------------
    
PROC_EXIT:
    Set GetNextEmptyRange = rngEmpty
    Set rngEmpty = Nothing
    Exit Function
    
PROC_ERR:
    Resume PROC_EXIT

End Function


Public Sub ReplaceFormulaWithValue(ByVal targetRange As Range)
'==================================================================================================
'   Description :   Replaces all formula in the range with values
'   Parameters  :   targetRange  -   Range where formulas will be replaced with values
'   Returns     :   Void
'==================================================================================================
    
    Const procName  As String = "ReplaceFormulaWithValue"
    
    On Error GoTo PROC_ERR
    
    '----------------------------------------------------------------------------------------------
    'Wait till calculation is complete
    Do Until Application.CalculationState = xlDone
        DoEvents
    Loop
    
    'Replace with values
    targetRange.value = targetRange.value
    
    '----------------------------------------------------------------------------------------------
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    Resume PROC_EXIT
    
End Sub

Public Sub FreezeSheetPanes(ByVal sheet As Worksheet, ByVal rowNumber As Long, ByVal columnNumber As Long)
' =================================================================================================
' Description : Procedure to Freeze panes in a sheet
'
' Parameter : sheet (Worksheet): The worksheet whose panes have to be frozen
' Parameter : rowNumber (Long): Row Number of the last row that's always visible. Use 0 for no rowsfrozen
' Parameter : columnNumber (Long): Columne number of the last row that's always visible. Use 0 for no columns frozen
'
' Comments    :
' =================================================================================================

    Const PROCEDURE_NAME    As String = "FreezeSheetPanes"
    
    Dim cellRow             As Long
    Dim cellColumn          As Long
    Dim sheetVisibility     As Long
    Dim currentSheet        As Worksheet
    Dim updateScreen        As Boolean


    '----------------------------------------------------------------------------------------------
    
    'Sanitize row & column numbers
    '-----------------------------
    cellRow = IIf(rowNumber > 0, rowNumber + 1, 1)
    cellColumn = IIf(columnNumber > 0, columnNumber + 1, 1)
    
    'Save current status
    '-------------------------
    sheetVisibility = sheet.Visible
    Set currentSheet = ActiveSheet
    updateScreen = Application.ScreenUpdating
    
    'Goto the cell to be forzen
    '--------------------------
    If Not updateScreen Then Application.ScreenUpdating = True
    
    With sheet
        .Visible = xlSheetVisible
        .Activate
        .Cells.Item(cellRow, cellColumn).Select
    End With
    
    ActiveWindow.FreezePanes = True

    '----------------------------------------------------------------------------------------------

PROC_EXIT:
    
    'Restore original status
    '-----------------------
    sheet.Visible = sheetVisibility
    currentSheet.Activate
    Application.ScreenUpdating = updateScreen

    Exit Sub
    
    '----------------------------------------------------------------------------------------------
    
PROC_ERR:

    Resume PROC_EXIT
    
End Sub


'Returns column header as dictionary
Public Function GetColumnHeaderNumberDictionary(ByVal wsSheet As Worksheet, ByVal headerRowNum As Long) As Scripting.Dictionary
' =================================================================================================
' Description : Returns the header names & column numbers as a dictionary
'
' Parameter : wsSheet (Worksheet): The worksheet containing the headers
' Parameter : headerRowNum (Integer): The row numbeer of header row

' Return Type : Dictionary
'
' Comments    :
' =================================================================================================

    Const PROCEDURE_NAME    As String = "GetColHeadNumDic"

    Dim headerDictionary    As Scripting.Dictionary
    Dim headerCell          As Range

    '----------------------------------------------------------------------------------------------
    
    'Variable(s) Initialization
    Set headerDictionary = New Scripting.Dictionary
    Set headerCell = wsSheet.Cells.Item(headerRowNum, 1)
    
    Do While Not IsEmpty(headerCell)
        If headerDictionary.Exists(Trim$(headerCell.value)) Then
            Err.Raise 0, vbNullString, FormatString("Found duplicate key: {0}, Sheet: {1}, Address: {2}", _
            Trim$(headerCell.value), wsSheet.Name, headerCell.Address)
        Else
            headerDictionary.Add Trim$(headerCell.value), headerCell.column
        End If
        Set headerCell = headerCell.Offset(0, 1)
    Loop
    
    'Add sheet name
    headerDictionary.Add "Sheet Name", wsSheet.Name

    '----------------------------------------------------------------------------------------------

PROC_EXIT:
    
    Set GetColumnHeaderNumberDictionary = headerDictionary
    Set headerCell = Nothing
    Set headerDictionary = Nothing

    Exit Function
    
    '----------------------------------------------------------------------------------------------
    
PROC_ERR:

    Resume PROC_EXIT
    
End Function

Public Function GetKeyValueDictionary(ByVal sourceWorksheet As Worksheet, ByVal keyCol As Long, ByVal valCol As Long) As Scripting.Dictionary
' =================================================================================================
' Description : Function to return the dictionary containing key value pairs of data in 2 columns
'
' Parameter : ws (Worksheet): Source Worksheet
' Parameter : keyCol (Integer): Column number of the Key
' Parameter : valCol (Integer): Column number of the Value

' Return Type : Scripting.Dictionary
'
' Comments    :
' =================================================================================================

    Const PROCEDURE_NAME    As String = "GetKeyValueDictionary"
    
    Dim rngKey              As Range
    Dim keyValueDictionary  As Scripting.Dictionary

    '----------------------------------------------------------------------------------------------
    
    'Variable(s) Initialization
    Set rngKey = sourceWorksheet.Cells.Item(1, keyCol)
    Set keyValueDictionary = New Scripting.Dictionary
    
    'Loop through all rng
    Do While Not IsEmpty(rngKey)
        
        'Add key value pair to dictionary
        keyValueDictionary.Add Trim$(rngKey.value), Trim$(rngKey.Offset(0, valCol - keyCol).value)
        
        'Move to next row
        Set rngKey = rngKey.Offset(1, 0)
    Loop

    '----------------------------------------------------------------------------------------------

PROC_EXIT:
    
    Set GetKeyValueDictionary = keyValueDictionary

    Exit Function
    
    '----------------------------------------------------------------------------------------------
    
PROC_ERR:

    Resume PROC_EXIT

End Function


