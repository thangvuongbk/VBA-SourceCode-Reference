Attribute VB_Name = "MGenSkeleton"
Option Explicit
Const OK = 1
Const ERROR = -1
'$$$$$$$$$$$$ DEFINITIONS FOR "MCDC" SHEET $$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
Const MCDC_TC_NO_COL = 2
Const MCDC_MCDC_COL = 3
Const MCDC_INPUT_COL = 4
'$$$$$$$$$$$$ DEFINITIONS FOR "Testcases" SHEET $$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
Const TC_TC_NO_COLUMN = 1
Const TC_DATASET_PARAMETER_STRING = "PARAMETERS"
Const TC_TC_NO_STRING = "TC No."
Const TC_INPUT_STRING = "INPUTS"
Const TC_LOCAL_VARIABLE_STRING = "LOCAL VARIABLES"
Const TC_OUTPUT_STRING = "OUTPUTS"
Const TC_DESCRIPTION_STRING = "DESCRIPTIONS"

Const COMMON_COLOR_BLACK = 1
Const COMMON_COLOR_WHITE = 2
Const COMMON_COLOR_RED = 3
Const COMMON_COLOR_GREEN = 4
Const COMMON_COLOR_BLUE = 5
Const COMMON_COLOR_YELLOW = 6
Const COMMON_COLOR_MAGENTA = 7
Const COMMON_COLOR_CYAN = 8



Sub GenTDSkeleton()
    '***********************************************************
    '
    'AUTHOR:            Le Thai (RBVH\EMB3)
    'DATE CREATED:      12/Dec/2013
    'FUNCTION NAME:     GenTDSkeleton
    'DESCRIPTION:
    '       INPUT: Testcases sheet
    '
    '       OUTPUT: Actual signal colum with names are filled
    '
    '***********************************************************
    Dim TCSheet As String
    Dim TCNoRow, TCNoCol As Integer
    Dim signalNameRow As Integer
    Dim parameterRow As Integer
    Dim startingRow, startingCol As Integer
    Dim toleranceRow As Integer
    Dim typeRow As Integer
    Dim MaxRow As Integer
    Dim MinRow As Integer
    Dim startInputCol, endInputCol As Integer
    Dim startLocalVarCol, endLocalVarCol As Integer
    Dim startOutputCol, endOutputCol As Integer
    Dim i, j As Integer
    Dim localVarColIdx As Integer
    Dim doneFlag As Boolean
    Dim temp As Variant
    
    TCSheet = "Testcases"
    'Back up
    Backup
    'Confirmation
    temp = MsgBox("GenTDSkeleton! Generate local variables?", _
                    vbYesNo + vbExclamation, "Attention!")
    'OK
    If (temp = vbYes) Then
        'OK
    Else
        'Exit sub
        Exit Sub
    End If
    
    startingRow = 1
    startingCol = 1
    TCNoCol = 1
    'Find row number for the string "Tolerance"
    toleranceRow = FindRowAll(TCSheet, TCNoCol, "Tolerance", startingRow)
    If (toleranceRow < 0) Then
        MsgBox "GenTDSkeleton! Unable to find String 'Tolerance' in column '" & TCNoCol & "' in sheet '" & TCSheet & "'!"
        Exit Sub
    End If
    typeRow = toleranceRow + 1
    MaxRow = toleranceRow + 2
    MinRow = toleranceRow + 3
    'Find row number for the string "TC No."
    TCNoRow = FindRowAll(TCSheet, TCNoCol, TC_TC_NO_STRING, startingRow)
    If (TCNoRow < 0) Then
        MsgBox "GenTDSkeleton! Unable to find String '" & TC_TC_NO_STRING & "' in column '" & TCNoCol & "' in sheet '" & TCSheet & "'!"
        Exit Sub
    End If
    signalNameRow = TCNoRow + 1
    'Find row number for the string "PARAMETERS"
    parameterRow = FindRowAll(TCSheet, TCNoCol + 1, TC_DATASET_PARAMETER_STRING, startingRow)
    If (parameterRow < 0) Then
        MsgBox "GenTDSkeleton! Unable to find String '" & TC_DATASET_PARAMETER_STRING & "' in column '" & TCNoCol & "' in sheet '" & TCSheet & "'!"
        Exit Sub
    End If
    parameterRow = parameterRow + 1
    'Find colum number for the string "INPUTS"
    startInputCol = FindColAll(TCSheet, TCNoRow, TC_INPUT_STRING, startingCol)
    If (startInputCol < 0) Then
        MsgBox "GenTDSkeleton! Unable to find String '" & TC_INPUT_STRING & "' in row '" & TCNoRow & "' in sheet '" & TCSheet & "'!"
        Exit Sub
    End If
    'Find colum number for the string "LOCAL VARIABLES"
    startLocalVarCol = FindColAll(TCSheet, TCNoRow, TC_LOCAL_VARIABLE_STRING, startingCol)
    If (startLocalVarCol < 0) Then
        MsgBox "GenTDSkeleton! Unable to find String '" & TC_LOCAL_VARIABLE_STRING & "' in row '" & TCNoRow & "' in sheet '" & TCSheet & "'!"
        Exit Sub
    End If
    'Find colum number for the string "OUTPUTS"
    startOutputCol = FindColAll(TCSheet, TCNoRow, TC_OUTPUT_STRING, startingCol)
    If (startOutputCol < 0) Then
        MsgBox "GenTDSkeleton! Unable to find String '" & TC_OUTPUT_STRING & "' in row '" & TCNoRow & "' in sheet '" & TCSheet & "'!"
        Exit Sub
    End If
    'Find colum number for the string "DESCRIPTIONS"
    endOutputCol = FindColAll(TCSheet, TCNoRow, TC_DESCRIPTION_STRING, startingCol)
    If (endOutputCol < 0) Then
        MsgBox "GenTDSkeleton! Unable to find String '" & TC_DESCRIPTION_STRING & "' in row '" & TCNoRow & "' in sheet '" & TCSheet & "'!"
        Exit Sub
    End If
    endOutputCol = endOutputCol - 1
    'Check Sequence INPUTS --> LOCAL VARIABLES --> OUTPUTS
    If (startInputCol < startLocalVarCol And startLocalVarCol < startOutputCol) Then
        'endInputCol
        endInputCol = startLocalVarCol - 1
        'endLocalVarCol
        endLocalVarCol = startOutputCol - 1
    Else
        MsgBox "GenTDSkeleton! Signal sequence must be 'INPUTS --> LOCAL VARIABLES --> OUTPUTS'"
        Exit Sub
    End If
    
    'Loop for all inputs
    For i = startInputCol To endInputCol
        If (Worksheets(TCSheet).Cells(signalNameRow, i) <> vbNullString) Then
            'Loop for all local variable cells
            doneFlag = False
            For j = startLocalVarCol To endLocalVarCol
                If (doneFlag = False) Then
                    'Empty --> Insert data
                    If (Worksheets(TCSheet).Cells(signalNameRow, j) = vbNullString) Then
                        'Signal name
                        Worksheets(TCSheet).Cells(signalNameRow, j) = "actual_" & Worksheets(TCSheet).Cells(signalNameRow, i).value
                        'Tolerance
                        Worksheets(TCSheet).Cells(toleranceRow, j) = Worksheets(TCSheet).Cells(toleranceRow, i).value
                        'Type
                        Worksheets(TCSheet).Cells(typeRow, j) = Worksheets(TCSheet).Cells(typeRow, i).value
                        'Max
                        Worksheets(TCSheet).Cells(MaxRow, j) = Worksheets(TCSheet).Cells(MaxRow, i).value
                        'Min
                        Worksheets(TCSheet).Cells(MinRow, j) = Worksheets(TCSheet).Cells(MinRow, i).value
                        'Update doneFlag
                        doneFlag = True
                    Else
                        'Check whether the signal was already filled
                        If (Worksheets(TCSheet).Cells(signalNameRow, j) = "actual_" & Worksheets(TCSheet).Cells(signalNameRow, i).value) Then
                            'Update doneFlag
                            doneFlag = True
                        End If
                    End If
                'doneFlag = True
                Else
                    'Do nothing
                End If
            Next j
            'doneFlag = False --> Insert new column and fill data
            If (doneFlag = False) Then
                'Insert new column
                Worksheets(TCSheet).Columns(endLocalVarCol + 1).Insert
                'Update endLocalVarCol
                endLocalVarCol = endLocalVarCol + 1
                'Fill data
                'Signal name
                Worksheets(TCSheet).Cells(signalNameRow, endLocalVarCol) = "actual_" & Worksheets(TCSheet).Cells(signalNameRow, i).value
                'Tolerance
                Worksheets(TCSheet).Cells(toleranceRow, endLocalVarCol) = Worksheets(TCSheet).Cells(toleranceRow, i).value
                'Type
                Worksheets(TCSheet).Cells(typeRow, endLocalVarCol) = Worksheets(TCSheet).Cells(typeRow, i).value
                'Max
                Worksheets(TCSheet).Cells(MaxRow, endLocalVarCol) = Worksheets(TCSheet).Cells(MaxRow, i).value
                'Min
                Worksheets(TCSheet).Cells(MinRow, endLocalVarCol) = Worksheets(TCSheet).Cells(MinRow, i).value
                'Update doneFlag
                doneFlag = True
            'doneFlag = True
            Else
                'Do nothing
            End If
        Else
            MsgBox "GenTDSkeleton! Input name is empty at cell(" & signalNameRow & ", " & i & ")!"
            Exit Sub
        End If
    Next i
    'Loop for all parameter
    i = TCNoCol + 2
    Do While (Worksheets(TCSheet).Cells(parameterRow, i) <> vbNullString)
        'Loop for all local variable cells
        doneFlag = False
        For j = startLocalVarCol To endLocalVarCol
            If (doneFlag = False) Then
                'Empty --> Insert data
                If (Worksheets(TCSheet).Cells(signalNameRow, j) = vbNullString) Then
                    'Signal name
                    Worksheets(TCSheet).Cells(signalNameRow, j) = "actual_" & Worksheets(TCSheet).Cells(parameterRow, i).value
                    'Update doneFlag
                    doneFlag = True
                Else
                    'Check whether the signal was already filled
                    If (Worksheets(TCSheet).Cells(signalNameRow, j) = "actual_" & Worksheets(TCSheet).Cells(parameterRow, i).value) Then
                        'Update doneFlag
                        doneFlag = True
                    End If
                End If
            'doneFlag = True
            Else
                'Do nothing
            End If
        Next j
        'doneFlag = False --> Insert new column and fill data
        If (doneFlag = False) Then
            'Insert new column
            Worksheets(TCSheet).Columns(endLocalVarCol + 1).Insert
            'Update endLocalVarCol
            endLocalVarCol = endLocalVarCol + 1
            'Fill data
            'Signal name
            Worksheets(TCSheet).Cells(signalNameRow, endLocalVarCol) = "actual_" & Worksheets(TCSheet).Cells(parameterRow, i).value
            'Update doneFlag
            doneFlag = True
        'doneFlag = True
        Else
            'Do nothing
        End If
        'Increasement
        i = i + 1
    Loop
End Sub
