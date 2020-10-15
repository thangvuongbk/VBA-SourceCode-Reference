Attribute VB_Name = "MCommon"
Option Explicit
Const OK = 1
Const ERROR = -1
Const MAX_ROW = 1000
Const MAX_COL = 255
'Const NO_INPUT = 2
'Const NO_INPUT_OUTPUT = 3
'Const NO_INPUT_OUTPUT_LOCALOUTPUT = 4
'Const NO_INPUT_LOCALOUTPUT = 5
'Const NO_OUTPUT = 6
'Const NO_OUTPUT_LOCALOUTPUT = 7
'Const NO_LOCALOUTPUT = 8
Sub Backup()
    'Backup Input sheet
    Worksheets("MCDC_Backup").Cells.Clear
    Worksheets("MCDC").Cells.Copy
    Worksheets("MCDC_Backup").Range("A1").PasteSpecial Paste:=xlPasteAll
    'Backup Result sheet
    Worksheets("Testcases_Backup").Cells.Clear
    Worksheets("Testcases").Cells.Copy
    Worksheets("Testcases_Backup").Range("A1").PasteSpecial Paste:=xlPasteAll
    'Backup Constants sheet
    'Worksheets("Constants_Backup").Cells.Clear
    'Worksheets("Constants").Cells.Copy
    'Worksheets("Constants_Backup").Range("A1").PasteSpecial Paste:=xlPasteAll
End Sub
Sub Restore()
    'Restore Input sheet
    Worksheets("MCDC").Cells.Clear
    Worksheets("MCDC_Backup").Cells.Copy
    Worksheets("MCDC").Range("A1").PasteSpecial Paste:=xlPasteAll
    'Restore Result sheet
    Worksheets("Testcases").Cells.Clear
    Worksheets("Testcases_Backup").Cells.Copy
    Worksheets("Testcases").Range("A1").PasteSpecial Paste:=xlPasteAll
    'Restore Constants sheet
    'Worksheets("Constants").Cells.Clear
    'Worksheets("Constants_Backup").Cells.Copy
    'Worksheets("Constants").Range("A1").PasteSpecial Paste:=xlPasteAll
End Sub
Sub Clear()
    If (MsgBox("Clear all data of all sheets! Are you sure?", vbYesNo + vbExclamation, "CLEAR") <> vbYes) Then
        Exit Sub
    End If
    'Back up
    Call Backup
    'Clear
    Worksheets("MCDC").Cells.Clear
    Worksheets("Testcases").Cells.Clear
    'Worksheets("Constants").Cells.Clear
    Worksheets("Temporary").Cells.Clear
    'Put initial data
    'MCDC Sheet
    'Worksheets("Guide").Range("MCDC").Copy
    'Worksheets("MCDC").Range("A1").PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
    '            False, Transpose:=False
    'Testcases Sheet
    Worksheets("Guide").Range("Testcases").Copy
    Worksheets("Testcases").Range("A1").PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
                False, Transpose:=False
    'Constants Sheet
    'Worksheets("Guide").Range("Constants").Copy
    'Worksheets("Constants").Range("A1").PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
                False, Transpose:=False
                
    'Set font
    With Worksheets("MCDC").Cells.Font
        .name = "Arial"
        .Size = 12
    End With
    With Worksheets("Testcases").Cells.Font
        .name = "Arial"
        .Size = 12
    End With
    'With Worksheets("Constants").Cells.Font
    '    .name = "Arial"
    '    .Size = 12
    'End With
    With Worksheets("Temporary").Cells.Font
        .name = "Arial"
        .Size = 12
    End With
    
    Worksheets("MCDC").Activate
    Worksheets("MCDC").Range("B2").Select
End Sub
Sub Border(ByVal sheet As Variant, _
                ByVal startRow As Integer, ByVal startCol As Integer, _
                ByVal endRow As Integer, ByVal endCol As Integer)
    '***********************************************************
    '
    'AUTHOR:            Le Thai (RBVH\EMB3)
    'DATE CREATED:      28/May/2013
    'FUNCTION NAME:     FindRow2
    'DESCRIPTION:
    '       INPUT:
    '               sheet:  sheet number or sheet name
    '               startRow: row number of starting cell to border
    '               startCol: column number of starting cell to border
    '               endRow: row number of ending cell to border
    '               endRow: row number of ending cell to border
    '       OUTPUT: Range(Cells(startRow, startCol), Cells(endRow, endCol)) will be all bordered and bordered around with thick line
    '
    '***********************************************************
    'Bordering
    Worksheets(sheet).Range(Cells(startRow, startCol), Cells(endRow, endCol)).Borders.LineStyle = xlContinuous
    Worksheets(sheet).Range(Cells(startRow, startCol), Cells(endRow, endCol)).BorderAround Weight:=xlThick
End Sub
Sub AutofitCellsActivesheet()
    Dim i As Integer
    For i = 0 To 1
        ActiveSheet.UsedRange.Columns.AutoFit
        ActiveSheet.UsedRange.Rows.AutoFit
    Next i
End Sub
Sub AutofitColumns(ByVal sheet As Variant)
    Worksheets(sheet).Activate
    Dim i As Integer
    For i = 0 To 5
        ActiveSheet.UsedRange.Columns.AutoFit
    Next i
End Sub
Function FindRowAll(ByVal sheet As Variant, _
                    ByVal col As Integer, _
                    ByVal key As String, _
                    ByVal startRow As Integer) As Integer
    '***********************************************************
    '
    'AUTHOR:            Le Thai (RBVH\EMB3)
    'DATE CREATED:      28/May/2013
    'FUNCTION NAME:     FindRow2
    'DESCRIPTION:
    '       INPUT:
    '               sheet:  sheet number or sheet name
    '               col:    column number
    '               key:
    '                    " "    : find first nullstring cell
    '                    "*"    : find first non-nullstring cell
    '                    "xxx"  : contents of the cell
    '               startRow: starting row to search
    '       OUTPUT: row of the cell containing the string key
    '
    '***********************************************************
    Dim row As Integer
    
    If (key = vbNullString) Then
        FindRowAll = -1
        Exit Function
    ElseIf (key = " ") Then
        row = startRow
        Do While (Worksheets(sheet).Cells(row, col) <> vbNullString)
            If (row = MAX_ROW) Then
                FindRowAll = -1
                Exit Function
            End If
            row = row + 1
        Loop
    ElseIf (key = "*") Then
        row = startRow
        Do While (Worksheets(sheet).Cells(row, col) = vbNullString)
            If (row = MAX_ROW) Then
                FindRowAll = -1
                Exit Function
            End If
            row = row + 1
        Loop
    Else
        row = startRow
        Do While (Worksheets(sheet).Cells(row, col) <> key)
            If (row = MAX_ROW) Then
                FindRowAll = -1
                Exit Function
            End If
            row = row + 1
        Loop
    End If
    FindRowAll = row
End Function
Function FindColAll(ByVal sheet As Variant, _
                    ByVal row As Integer, _
                    ByVal key As String, _
                    ByVal startCol As Integer) As Integer
    '***********************************************************
    '
    'AUTHOR:            Le Thai (RBVH\EMB3)
    'DATE CREATED:      28/May/2013
    'FUNCTION NAME:     FindRow2
    'DESCRIPTION:
    '       INPUT:
    '               sheet:  sheet number or sheet name
    '               row:    row number
    '               key:
    '                    "c*c"  : find first colorized cell
    '                    "c c"  : find first uncolorized cell
    '                    " "    : find first nullstring cell
    '                    "*"    : find first non-nullstring cell
    '                    "xxx"  : contents of the cell
    '               startCol: starting col to search
    '       OUTPUT: column of the cell containing the string key
    '
    '***********************************************************
    Dim col As Integer
    On Error GoTo FindColAllErrorHandler
    
    If (IsNumeric(key)) Then
        FindColAll = -1
        Exit Function
    End If
    
    If (key = vbNullString) Then
        FindColAll = -1
        Exit Function
    ElseIf (key = "c c") Then
        col = startCol
        Do While (Worksheets(sheet).Cells(row, col).Interior.ColorIndex <> xlNone)
            If (col = MAX_COL) Then
                FindColAll = -1
                Exit Function
            End If
            col = col + 1
        Loop
    ElseIf (key = "c*c") Then
        col = startCol
        Do While (Worksheets(sheet).Cells(row, col).Interior.ColorIndex = xlNone)
            If (col = MAX_COL) Then
                FindColAll = -1
                Exit Function
            End If
            col = col + 1
        Loop
    ElseIf (key = " ") Then
        col = startCol
        Do While (Worksheets(sheet).Cells(row, col) <> vbNullString)
            If (col = MAX_COL) Then
                FindColAll = -1
                Exit Function
            End If
            col = col + 1
        Loop
    ElseIf (key = "*") Then
        col = startCol
        Do While (Worksheets(sheet).Cells(row, col) = vbNullString)
            If (col = MAX_COL) Then
                FindColAll = -1
                Exit Function
            End If
            col = col + 1
        Loop
    Else
        col = startCol
        Do While (Worksheets(sheet).Cells(row, col) <> key)
            If (col = MAX_COL) Then
                FindColAll = -1
                Exit Function
            End If
            col = col + 1
        Loop
    End If
    FindColAll = col
    Exit Function
    
FindColAllErrorHandler:
    MsgBox "ERROR! FindColAll" & vbNewLine & _
            Err.Number & vbCr & Err.description
    FindColAll = -1
End Function


Function SaveVar(ByVal name As String, ByVal value As Variant)
    Dim varSheet As String
    Dim row, col As Integer
    '***********************************************************
    '
    'AUTHOR:            Le Thai (RBVH\EMB3)
    'DATE CREATED:      11/Jun/2013
    'FUNCTION NAME:     SaveVar
    'DESCRIPTION:
    '       INPUT:
    '               name:  variable name
    '               value: variable value
    '       OUTPUT: the variable 'name ' will be save with value 'value' in "SavedVariables" Sheet
    '       RETURN VALUE:
    '               1:  Update the value of available variable
    '               2:  Create new variable and update the value of this variable
    '               -1: Saving failed
    '
    '***********************************************************
    varSheet = "SavedVariables"
    col = 1
    row = FindRowAll(varSheet, col, name, 1)
    'Find row
    If (row > 0) Then
        Worksheets(varSheet).Cells(row, col + 1).value = value
        SaveVar = 1
        Exit Function
    'Not found
    Else
        row = FindRowAll(varSheet, col, " ", 1)
        If (row < 0) Then
            SaveVar = -1
            Exit Function
        End If
        'Set name and value for new var
        Worksheets(varSheet).Cells(row, col).value = name
        Worksheets(varSheet).Cells(row, col + 1).value = value
        SaveVar = 2
        Exit Function
    End If
End Function
Function GetVar(ByVal name As String, ByRef value As Variant)
    '***********************************************************
    '
    'AUTHOR:            Le Thai (RBVH\EMB3)
    'DATE CREATED:      11/Jun/2013
    'FUNCTION NAME:     GetVar
    'DESCRIPTION:
    '       INPUT:
    '               name:  variable name
    '       OUTPUT:
    '               value: variable value
    '       RETURN VALUE:
    '               1:  Get value sucessfully
    '               -1: Getting failed
    '
    '***********************************************************
    
    
    Dim varSheet As String
    Dim row, col As Integer
    
    varSheet = "SavedVariables"
    col = 1
    row = FindRowAll(varSheet, col, name, 1)
    'Found
    If (row > 0) Then
        value = Worksheets(varSheet).Cells(row, col + 1).value
        GetVar = 1
        Exit Function
    'Not found
    Else
        GetVar = -1
        Exit Function
    End If
End Function
Function WriteTCAll(ByVal sheet As Variant, _
                    ByRef condition As String, _
                    ByVal oRowIdx As Integer, _
                    ByVal TCType As Integer, _
                    ByRef log As String, _
                    Optional signalNameRowIdx As Integer)
    '***********************************************************
    '
    'AUTHOR:            Le Thai (RBVH\EMB3)
    'DATE CREATED:      28/May/2013
    'FUNCTION NAME:     WriteTCAll
    'DESCRIPTION:
    '       INPUT:
    '           sheet:          output sheet (Testcase_ATT sheet)
    '           condition:      conditions to be written
    '           oRowIdx:        row index to write
    '           TCType:         Type of TC
    '                        0: Legal
    '                        1: Illegal
    '                        2: Guarded
    '       OUTPUT:
    '           Testcase design in sheet
    '***********************************************************
    Dim oColIdx As Integer
    Dim inputName As String
    Dim inputValue As String
    Dim temp As Integer
    Dim inputCol, outputCol, localOutputCol, DesCol As Integer
    '*****************************Writing the output TC***************************
    log = ""
    'Check optional parameter
    If (IsMissing(signalNameRowIdx)) Then
        signalNameRowIdx = 2
    End If
    'Loop to write output
    Do While (ExtractFirstConditionAll(condition, inputName, inputValue) > 0)
        'Find column number if inputName
        oColIdx = FindColAll(sheet, signalNameRowIdx, inputName, 2)
        'Found inputName
        If (oColIdx > 0) Then
            'Set value
            If (Worksheets(sheet).Cells(oRowIdx, oColIdx) <> vbNullString) Then
                'Check inputValue: Replace if (inputValue <> "Current Value")
                If (inputValue <> "Current Value") Then
                    'For Legal TC, replace current values
                    If (TCType = 0) Then
                        log = log & vbNewLine & _
                                "Current value of signal '" & inputName & "' is " & Worksheets(sheet).Cells(oRowIdx, oColIdx).value & vbNewLine & _
                                "Replaced this value with value: " & inputValue & vbNewLine
                        Worksheets(sheet).Cells(oRowIdx, oColIdx).value = inputValue
                    'For ILlegal TC, replace current values
                    ElseIf (TCType = 1) Then
                        log = log & vbNewLine & _
                                "Current value of signal '" & inputName & "' is " & Worksheets(sheet).Cells(oRowIdx, oColIdx).value & vbNewLine & _
                                "Replaced this value with value: " & inputValue & vbNewLine
                        Worksheets(sheet).Cells(oRowIdx, oColIdx).value = inputValue
                    'replace current values
                    Else
                        log = log & vbNewLine & _
                                "Current value of signal '" & inputName & "' is " & Worksheets(sheet).Cells(oRowIdx, oColIdx).value & vbNewLine & _
                                "Replaced this value with value: " & inputValue & vbNewLine
                        Worksheets(sheet).Cells(oRowIdx, oColIdx).value = inputValue
                    End If
                'Not replace
                Else
                    'Do nothing
                End If
            Else
                Worksheets(sheet).Cells(oRowIdx, oColIdx).value = inputValue
            End If

        'Not Found inputName
        Else
            'Check the signal name: input or local output
            'Local output
            If (InStr(inputName, "exp_asp") > 0) Then
                localOutputCol = FindColAll(sheet, signalNameRowIdx - 1, "LOCAL OUTPUTS", 1)
                If (localOutputCol > 0) Then
                    'DesCol
                    DesCol = FindColAll(sheet, signalNameRowIdx - 1, "DESCRIPTIONS", 1)
                    If (DesCol < 0) Then
                        log = "Unable to find cell 'DESCRIPTIONS' in Testcases sheet to insert column"
                        WriteTCAll = ERROR
                        Exit Function
                    End If
                    'Find Empty cell for Local output
                    localOutputCol = FindColAll(sheet, signalNameRowIdx, " ", localOutputCol)
                    If (localOutputCol < 0) Then
                        log = "WriteTCAll! Unable to find empty cell to fill local output!"
                        WriteTCAll = ERROR
                        Exit Function
                    ElseIf (localOutputCol >= DesCol) Then
                        'Insert new column for this inputName
                        Worksheets(sheet).Columns(DesCol).Insert
                        'Modify input name for signal in new column
                        Worksheets(sheet).Cells(signalNameRowIdx, DesCol).value = inputName
                        'Set value
                        Worksheets(sheet).Cells(oRowIdx, DesCol).value = inputValue
                    Else
                        'Modify input name for signal in new column
                        Worksheets(sheet).Cells(signalNameRowIdx, localOutputCol).value = inputName
                        'Set value
                        Worksheets(sheet).Cells(oRowIdx, localOutputCol).value = inputValue
                    End If
                Else
                    'Set oColIdx
                    oColIdx = FindColAll(sheet, signalNameRowIdx - 1, "DESCRIPTIONS", 1)
                    If (oColIdx < 0) Then
                        log = "Unable to find cell 'DESCRIPTIONS' in Testcases sheet to insert column"
                        WriteTCAll = ERROR
                        Exit Function
                    End If
                    'Insert new column for this inputName
                    Worksheets(sheet).Columns(oColIdx).Insert
                    'Insert 'LOCAL OUTPUTS'
                    Worksheets(sheet).Cells(signalNameRowIdx - 1, oColIdx).value = "LOCAL OUTPUTS"
                    'Modify input name for signal in new column
                    Worksheets(sheet).Cells(signalNameRowIdx, oColIdx).value = inputName
                    'Set value
                    Worksheets(sheet).Cells(oRowIdx, oColIdx).value = inputValue
                End If
            'output
            ElseIf (InStr(inputName, "expected") > 0) Then
                outputCol = FindColAll(sheet, signalNameRowIdx - 1, "OUTPUTS", 1)
                If (outputCol > 0) Then
                    'localOutputCol
                    localOutputCol = FindColAll(sheet, signalNameRowIdx - 1, "LOCAL OUTPUTS", 1)
                    If (localOutputCol < 0) Then
                        log = "Unable to find cell 'LOCAL OUTPUTS' in Testcases sheet to insert column"
                        'WriteTCAll = ERROR
                        'Exit Function
                        'Find "DESCRIPTIONS" as "LOCAL OUTPUTS"
                        localOutputCol = FindColAll(sheet, signalNameRowIdx - 1, "DESCRIPTIONS", 1)
                        If (localOutputCol < 0) Then
                            log = log & vbNewLine & "Unable to find cell 'DESCRIPTIONS' in Testcases sheet to insert column"
                            WriteTCAll = ERROR
                            Exit Function
                        End If
                    End If
                    'Find Empty cell for output
                    outputCol = FindColAll(sheet, signalNameRowIdx, " ", outputCol)
                    If (outputCol < 0) Then
                        log = "WriteTCAll! Unable to find empty cell to fill output!"
                        WriteTCAll = ERROR
                        Exit Function
                    ElseIf (outputCol >= localOutputCol) Then
                        'Insert new column for this inputName
                        Worksheets(sheet).Columns(localOutputCol).Insert
                        'Modify input name for signal in new column
                        Worksheets(sheet).Cells(signalNameRowIdx, localOutputCol).value = inputName
                        'Set value
                        Worksheets(sheet).Cells(oRowIdx, localOutputCol).value = inputValue
                    Else
                        'Modify input name for signal in new column
                        Worksheets(sheet).Cells(signalNameRowIdx, outputCol).value = inputName
                        'Set value
                        Worksheets(sheet).Cells(oRowIdx, outputCol).value = inputValue
                    End If
                'Input
                Else
                    'Set oColIdx
                    oColIdx = FindColAll(sheet, signalNameRowIdx - 1, "LOCAL OUTPUTS", 1)
                    If (oColIdx < 0) Then
                        log = "Unable to find cell 'LOCAL OUTPUTS' in Testcases sheet to insert column"
                        WriteTCAll = ERROR
                        Exit Function
                    End If
                    'Insert new column for this inputName
                    Worksheets(sheet).Columns(oColIdx).Insert
                    'Modify input name for signal in new column
                    Worksheets(sheet).Cells(signalNameRowIdx, oColIdx).value = inputName
                    'Set value
                    Worksheets(sheet).Cells(oRowIdx, oColIdx).value = inputValue
                End If
            'Input
            Else
                inputCol = FindColAll(sheet, signalNameRowIdx - 1, "INPUTS", 1)
                If (inputCol > 0) Then
                    'outputCol
                    'outputCol = FindColAll(sheet, signalNameRowIdx - 1, "LOCAL VARIABLES", 1)
                    'If (outputCol < 0) Then
                    '    log = "Unable to find cell 'LOCAL VARIABLES' in Testcases sheet to insert column"
                    '    'WriteTCAll = ERROR
                    '    'Exit Function
                    '    '
                        outputCol = FindColAll(sheet, signalNameRowIdx - 1, "OUTPUTS", 1)
                        If (outputCol < 0) Then
                            log = log & vbNewLine & "Unable to find cell 'OUTPUTS' in Testcases sheet to insert column"
                            WriteTCAll = ERROR
                            Exit Function
                        End If
                    'End If
                    'Find Empty cell for output
                    inputCol = FindColAll(sheet, signalNameRowIdx, " ", inputCol)
                    If (inputCol < 0) Then
                        log = "WriteTCAll! Unable to find empty cell to fill input!"
                        WriteTCAll = ERROR
                        Exit Function
                    ElseIf (inputCol >= outputCol) Then
                        'Insert new column for this inputName
                        Worksheets(sheet).Columns(outputCol).Insert
                        'Modify input name for signal in new column
                        Worksheets(sheet).Cells(signalNameRowIdx, outputCol).value = inputName
                        'Set value
                        Worksheets(sheet).Cells(oRowIdx, outputCol).value = inputValue
                    Else
                        'Modify input name for signal in new column
                        Worksheets(sheet).Cells(signalNameRowIdx, inputCol).value = inputName
                        'Set value
                        Worksheets(sheet).Cells(oRowIdx, inputCol).value = inputValue
                    End If
                Else
                    'Set oColIdx
                    oColIdx = FindColAll(sheet, signalNameRowIdx - 1, "OUTPUTS", 1)
                    If (oColIdx < 0) Then
                        log = "Unable to find cell 'OUTPUTS' in Testcases sheet to insert column"
                        WriteTCAll = ERROR
                        Exit Function
                    End If
                    'Insert new column for this inputName
                    Worksheets(sheet).Columns(oColIdx).Insert
                    'Modify input name for signal in new column
                    Worksheets(sheet).Cells(signalNameRowIdx, oColIdx).value = inputName
                    'Set value
                    Worksheets(sheet).Cells(oRowIdx, oColIdx).value = inputValue
                End If
            End If
        End If
    Loop
    WriteTCAll = OK
End Function
Function ExtractFirstConditionAll(ByRef condition As String, _
                            ByRef inputName As String, _
                            ByRef inputValue As String) As Integer
    '***********************************************************
    '
    'AUTHOR:            Le Thai (RBVH\EMB3)
    'DATE CREATED:      28/May/2013
    'FUNCTION NAME:     ExtractFirstConditionAll
    'DESCRIPTION:
    '       INPUT:
    '           condition:  condition of test case
    '       OUTPUT:
    '           condition: remaining condition
    '           inputName:  input name
    '           inputValue: input value
    '
    '***********************************************************
    Dim andPos, equalPos As Integer
    Dim firstCondition As String
    'Check input
    If (condition = vbNullString Or Len(condition) = vbNull) Then
        'MsgBox "ExtractFirstConditionAll! input condition is null."
        ExtractFirstConditionAll = 0
        Exit Function
    End If
    
    'get position of "&&"
    andPos = InStr(condition, "&&")
    'Check andPos
    If (andPos > 0) Then
        firstCondition = Mid(condition, 1, andPos - 2)
        condition = Mid(condition, andPos + 3)
    Else
        firstCondition = condition
        condition = ""
    End If
    'get position of "="
    equalPos = InStr(firstCondition, "=")
    'Check equalPos
    If (equalPos > 0) Then
        inputName = Mid(firstCondition, 1, equalPos - 1)
        inputValue = Mid(firstCondition, equalPos + 1)
    'invalid condition
    Else
        MsgBox "ExtractFirstConditionAll! First condition is invalid, no '=' character."
        ExtractFirstConditionAll = -1
        Exit Function
    End If
    ExtractFirstConditionAll = 1
End Function
Function GetSignalDataAll(ByVal sheet As Variant, _
                        ByVal row As Integer, _
                        ByVal col As Integer, _
                        ByRef inputs() As String, _
                        ByRef outputs() As String, _
                        ByRef inputValues() As String, _
                        ByRef outputValues() As String, _
                        ByVal kind As Integer, _
                        ByRef log As String) As Integer
    '***********************************************************
    '
    'AUTHOR:            Le Thai (RBVH\EMB3)
    'DATE CREATED:      24/May/2013
    'FUNCTION NAME:     GetSignalDataAll
    'DESCRIPTION:
    '       INPUT:
    '               sheet   :   sheet number
    '               row     :   row number of starting cell of the table
    '               col     :   column number of starting cell of the table
    '               kind    :   to indicate that reserved values are used or not
    '                           1   :   used
    '                           0   :   not used
    '
    '       OUTPUT:
    '               inputs()        :   array contains all the inputs names
    '               outputs()       :   array contains all the inputs names
    '               inputValues()   :   array contains all the values of all inputs
    '               inputValues()   :   array contains all the values of all inputs
    '               log             :   log
    '
    '***********************************************************
    Dim nInputs, nOutputs, nValues As Integer
    Dim nInputValues, nOutputValues As Integer
    Dim inputReservedValue As String
    Dim outputReservedValue As String
    Dim signalFlag As String
    Dim i, j As Integer
    Dim rowIdx, colIdx As Integer
    
    log = ""
    
    'Check for signalFlag
    If (Worksheets(sheet).Cells(row - 1, col) = "OUTPUTS") Then
        signalFlag = "output"
    Else
        signalFlag = "input"
    End If
    'Count the number of inputs and outputs
    nInputs = 0
    nOutputs = 0
    colIdx = col
    Do While (Worksheets(sheet).Cells(row, colIdx) <> vbNullString)
        If (signalFlag = "input") Then
            nInputs = nInputs + 1
            If (Worksheets(sheet).Cells(row - 1, colIdx + 1) = "OUTPUTS") Then
                signalFlag = "output"
            End If
        Else
            nOutputs = nOutputs + 1
        End If
        colIdx = colIdx + 1
    Loop
    'For debugging
    'MsgBox "nInputs: " & nInputs & vbNewLine & _
            "nOutputs: " & nOutputs


    'Count the number of values
    nValues = 0
    rowIdx = row + 1
    Do While (Worksheets(sheet).Cells(rowIdx, col) <> vbNullString)
        nValues = nValues + 1
        rowIdx = rowIdx + 1
    Loop
    'For debugging
    'MsgBox "nInputs: " & nInputs & vbNewLine & _
            "nOutputs: " & nOutputs & vbNewLine & _
            "nValues: " & nValues
            
    If (nValues > 0) Then
        If (nInputs > 0) Then
            ReDim Preserve inputs(0 To nInputs - 1)
            ReDim inputValues(0 To nInputs - 1, 0 To nValues - 1)
        Else
            log = log & "GetSignalDataAll! Unable to find 'INPUTS'"
            GetSignalDataAll = ERROR
            Exit Function
        End If
         If (nOutputs > 0) Then
            ReDim Preserve outputs(0 To nOutputs - 1)
            ReDim outputValues(0 To nOutputs - 1, 0 To nValues - 1)
        Else
            log = log & "GetSignalDataAll! Unable to find 'OUTPUTS'"
            GetSignalDataAll = ERROR
            Exit Function
        End If
    Else
        log = log & "GetSignalDataAll! No values"
        GetSignalDataAll = ERROR
        Exit Function
    End If

    'Get input data
    For i = 0 To (nInputs - 1)
        inputs(i) = Worksheets(sheet).Cells(row, col + i).value
        'Set initial values
        inputReservedValue = "0"
        'Get input value
        For j = 0 To (nValues - 1)
            'Get value
            'kind = 1 --> used reserved value
            If (kind = 1) Then
                'Normal value
                If (Worksheets(sheet).Cells(row + j + 1, col + i) <> vbNullString And _
                    Worksheets(sheet).Cells(row + j + 1, col + i) <> "X" And _
                    Worksheets(sheet).Cells(row + j + 1, col + i) <> "x") Then
                    inputValues(i, j) = Worksheets(sheet).Cells(row + j + 1, col + i).value
                    'Reserve the value
                    inputReservedValue = Worksheets(sheet).Cells(row + j + 1, col + i).value
                'Undefined value
                Else
                    inputValues(i, j) = inputReservedValue
                End If
            'kind = 0 --> not used reserved value
            Else
                'Cell is not null
                If (Worksheets(sheet).Cells(row + j + 1, col + i) <> vbNullString) Then
                    inputValues(i, j) = Worksheets(sheet).Cells(row + j + 1, col + i).value
                'Cell is null
                Else
                    inputValues(i, j) = ""
                End If
            End If
        Next j
    Next i
    'Get output data
    For i = 0 To (nOutputs - 1)
        outputs(i) = Worksheets(sheet).Cells(row, col + nInputs + i).value
        'Set initial values
        outputReservedValue = "0"
        'Get output value
        For j = 0 To (nValues - 1)
            'Get value
            'kind = 1 --> used reserved value
            If (kind = 1) Then
                'Normal value
                If (Worksheets(sheet).Cells(row + j + 1, col + nInputs + i) <> vbNullString And _
                    Worksheets(sheet).Cells(row + j + 1, col + nInputs + i) <> "X" And _
                    Worksheets(sheet).Cells(row + j + 1, col + nInputs + i) <> "x") Then
                    outputValues(i, j) = Worksheets(sheet).Cells(row + j + 1, col + nInputs + i).value
                    'Reserve the value
                    outputReservedValue = Worksheets(sheet).Cells(row + j + 1, col + nInputs + i).value
                Else
                    outputValues(i, j) = outputReservedValue
                End If
            'kind = 0 --> not used reserved value
            'Undefined value
            Else
                'Cell is not null
                If (Worksheets(sheet).Cells(row + j + 1, col + nInputs + i) <> vbNullString) Then
                    outputValues(i, j) = Worksheets(sheet).Cells(row + j + 1, col + nInputs + i).value
                'Cell is null
                Else
                    outputValues(i, j) = ""
                End If
            End If
        Next j
    Next i
    GetSignalDataAll = OK
End Function
Sub CSVDisplay2DStr(ByRef a2() As String, ByVal startRow As Integer)
    Dim i, j As Integer
    For i = 0 To UBound(a2, 1)
        For j = 0 To UBound(a2, 2)
            'MsgBox a2(i, j)
            Worksheets("Temporary").Cells(startRow + j + 1, i + 1) = a2(i, j)
        Next j
    Next i
End Sub
Sub CSVDisplay2DInt(ByRef a2() As Integer, ByVal startRow As Integer)
    Dim i, j As Integer
    For i = 0 To UBound(a2, 1)
        For j = 0 To UBound(a2, 2)
            'MsgBox a2(i, j)
            Worksheets("Temporary").Cells(startRow + i + 1, j + 1) = a2(i, j)
        Next j
    Next i
End Sub
Sub CSVDisplay1DInt(ByRef a() As Integer, ByVal startRow As Integer)
    Dim i As Integer
    For i = 0 To UBound(a)
        Worksheets("Temporary").Cells(startRow, i + 1) = a(i)
    Next i
End Sub
Sub CSVDisplay1DStr(ByRef a() As String, ByVal startRow As Integer)
    Dim i As Integer
    For i = 0 To UBound(a)
        Worksheets("Temporary").Cells(startRow, i + 1) = a(i)
    Next i
End Sub
Function ReadDirAll(ByRef filepath() As String, _
                ByRef log As String) As Integer
    '***********************************************************
    '
    'AUTHOR:            Le Thai (RBVH\EMB3)
    'DATE CREATED:      19/August/2013
    'FUNCTION NAME:     ReadDirAll
    'DESCRIPTION:
    '       INPUT: Multiple files
    '
    '       OUTPUT:
    '               filePath()        :   array of selected file names
    '               log             :   log
    '
    '***********************************************************
    Dim fileToOpen As Variant
    Dim writefile As Integer
    Dim i As Integer
    'Get file path
    fileToOpen = Application.GetOpenFilename(FileFilter:="Testcase csv Files (*.csv), *.csv", _
                                            MultiSelect:=True)
    'fileToOpen = Application.GetOpenFilename(MultiSelect:=True)
    If IsNumeric(fileToOpen) Then
        log = "ReadDirAll! File not found "
        ReadDirAll = ERROR
        Exit Function
    End If
    ReDim Preserve filepath(0 To (UBound(fileToOpen) - 1))
    For i = 1 To UBound(fileToOpen)
        filepath(i - 1) = fileToOpen(i)
    Next i
    ReadDirAll = 1
End Function
Function ReadFileToLinesAll(ByVal filepath As String, _
                    ByRef lines() As String, _
                    ByRef log As String) As Integer
    '***********************************************************
    '
    'AUTHOR:            Le Thai (RBVH\EMB3)
    'DATE CREATED:      19/August/2013
    'FUNCTION NAME:     ReadFileToLinesAll
    'DESCRIPTION:
    '       INPUT: File Path
    '
    '       OUTPUT:
    '               lines        :   array of lines in file path
    '               log          :   log
    '
    '***********************************************************
    Dim writefile As Integer
    Dim i As Integer
    If (Dir(filepath, vbNormal) = vbNullString) Then
        log = "File not found: '" & filepath & "'"
        ReadFileToLinesAll = ERROR
        Exit Function
    End If
    'Initialize writefile
    writefile = FreeFile
    'Open file
    Open filepath For Input As writefile
    'Read file line by line
    i = 0
    Do While Not EOF(writefile)
        ReDim Preserve lines(0 To i)
        Line Input #writefile, lines(i)
        'DEBUG
        'MsgBox "line(" & i & "): " & lines(i)
        i = i + 1
    Loop
    'Close file
    Close writefile
    'DEBUG
    'MsgBox lines(3)
    
    ReadFileToLinesAll = OK
End Function
Function SplitAll(ByVal inString As Variant, _
                  ByRef patterns() As String) As String()
    '***********************************************************
    '
    'AUTHOR:            Le Thai (RBVH\EMB3)
    'DATE CREATED:      12/Feb/2014
    'FUNCTION NAME:     SplitAll
    'DESCRIPTION: split inString with delimeters in patterns
    '       INPUT:
    '               inString       :   input string to be replaced
    '               patterns()  :   array of patterns in inString to be replaced
    '
    '       OUTPUT:
    '               output      :   string arry after splitting inString with delimeters in patterns
    '
    '***********************************************************
    Dim l_condition As String
    Dim result() As String
    Dim i  As Integer
    Dim l_inString As String
    Dim pattern_0 As String
    l_inString = inString
    'Only 1 patterns
    If (UBound(patterns) = 0) Then
        result = Split(l_inString, patterns(0))
    'Multiple patterns
    Else
        'Replace all patterns in inString by patterns(0)
        pattern_0 = patterns(0)
        patterns(0) = patterns(1)
        l_inString = ReplaceAll(l_inString, patterns(), pattern_0)
        patterns(0) = pattern_0
        'Split with delimeter is patterns(0)
        result = Split(l_inString, patterns(0))
    End If
    SplitAll = result
End Function
Function ReplaceAll(ByVal inString As Variant, _
                        ByRef patterns() As String, _
                        ByVal repStr As String) As String
    '***********************************************************
    '
    'AUTHOR:            Le Thai (RBVH\EMB3)
    'DATE CREATED:      12/Feb/2014
    'FUNCTION NAME:     ReplaceAll
    'DESCRIPTION: all patterns in inString will be replaced by repStr
    '       INPUT:
    '               inString       :   input string to be replaced
    '               patterns()  :   array of patterns in inString to be replaced
    '               repStr()    :   all string in patterns will be replaced by repStr
    '
    '       OUTPUT:
    '               output      :   string after all patterns in inString will be replaced by repStr
    '
    '***********************************************************
    Dim result As String
    Dim i  As Integer
    Dim pos As Integer
    result = inString
    For i = 0 To UBound(patterns)
        pos = InStr(result, patterns(i))
        Do While (pos > 0)
            result = Replace(result, patterns(i), repStr)
            pos = InStr(result, patterns(i))
        Loop
    Next i
    ReplaceAll = result
End Function

