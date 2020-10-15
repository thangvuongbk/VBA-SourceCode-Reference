Attribute VB_Name = "MGenCSV"
Option Explicit
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$ COMMON $$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
Const ERROR = -1
Const OK = 1
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$FOR GetSignalNames $$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
Const NO_INPUT = 2
Const NO_INPUT_OUTPUT = 3
Const NO_INPUT_OUTPUT_LOCALOUTPUT = 4
Const NO_INPUT_LOCALOUTPUT = 5
Const NO_OUTPUT = 6
Const NO_OUTPUT_LOCALOUTPUT = 7
Const NO_LOCALOUTPUT = 8
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$ FOR CheckSignalNames $$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
Const CHECKSIGNALNAME_SPACE = -2
Const CHECKSIGNALNAME_DUPLICATE = -3


Sub GenCSV()
    '***********************************************************
    '
    'AUTHOR:            Le Thai (RBVH\EMB3)
    'DATE CREATED:      24/May/2013
    'FUNCTION NAME:     GenCSV
    'DESCRIPTION:
    '       INPUT: sheet Testcases_ATT
    '
    '       OUTPUT: testcase files in csv format
    '
    '***********************************************************
    Dim inputs() As String
    Dim inputValues() As String
    Dim outputs() As String
    Dim outputValues() As String
    Dim localOutputs() As String
    Dim localOutputValues() As String
    Dim ConstantSheet As String
    Dim ATTSheet As String
    Dim mainData As String
    Dim line1, line2, line3, line4 As String
    Dim directoryPath As String
    Dim fPath As String
    Dim writefile As Integer
    Dim i, j As Integer
    Dim testModuleName As String
    Dim cycleTime As String
    Dim moduleIndex As Integer
    Dim constantSet As String
    Dim form As frmCSV
    Dim TCDigit1, TCDigit2, TCDigit3 As Integer
    Dim retGetSignalData As Integer
    Dim ret As Integer
    Dim log As String
    Dim startingCol As Integer
    Dim consRow, consCol As Integer
    Dim nameRow, nameCol As Integer
    
    On Error GoTo GenCSVErrorHandler
    
    'Initial
    ConstantSheet = "Constants"
    ATTSheet = "Testcases"
    testModuleName = ""
    cycleTime = "0.02"
    moduleIndex = 1
    constantSet = "DEFAULT"
    'Get input from user
    Set form = New frmCSV
    form.txtPath = ActiveWorkbook.path & "\CSV"
    'Get Vaiable
    If (GetVar("form.txtTMName", testModuleName) > 0 And _
        GetVar("form.txtModuleIndex", moduleIndex) > 0 And _
        GetVar("form.txtConstantSet", constantSet) > 0) Then
        form.txtTMName = testModuleName
        form.txtModuleIndex = moduleIndex
        form.txtConstantSet = constantSet
        'MsgBox "Get 'form.txtTMName' with value: " & form.txtTMName & vbNewLine & _
                "'form.txtModuleIndex' with value: " & form.txtModuleIndex & vbNewLine & _
                "'form.txtConstantSet' with value: " & form.txtConstantSet
    Else
        form.txtTMName = "TM_ClassName"
        form.txtModuleIndex = 0
        form.txtConstantSet = "DEFAULT"
    End If

    form.Show
    If (form.txtPath = "") Then
        Exit Sub
    End If
    testModuleName = form.txtTMName
    directoryPath = form.txtPath & "\" & testModuleName
    moduleIndex = form.txtModuleIndex
    constantSet = form.txtConstantSet
    'Save variable 'form.txtTMName'
    'Save OK
    If (SaveVar("form.txtTMName", testModuleName) > 0) Then
        'MsgBox "Save 'form.txtTMName' with value: " & form.txtTMName
    'Save failed
    Else
        'MsgBox "Unable to save 'form.txtTMName' with value: " & form.txtTMName
    End If
    'Save variable 'form.txtModuleIndex'
    'Save OK
    If (SaveVar("form.txtModuleIndex", moduleIndex) > 0) Then
        'MsgBox "Save 'form.txtModuleIndex' with value: " & form.txtModuleIndex
    'Save failed
    Else
        'MsgBox "Unable to save 'form.txtModuleIndex' with value: " & form.txtModuleIndex
    End If
    'Save variable 'form.txtConstantSet'
    'Save OK
    If (SaveVar("form.txtConstantSet", constantSet) > 0) Then
        'MsgBox "Save 'form.txtConstantSet' with value: " & form.txtConstantSet
    'Save failed
    Else
        'MsgBox "Unable to save 'form.txtConstantSet' with value: " & form.txtConstantSet
    End If
    
    'Create folder for TM inside form.txtPath
    If (Dir(directoryPath, vbDirectory) = "") Then
        MkDir directoryPath
    End If
    
    'Find column number of starting input name
    startingCol = FindColAll(ATTSheet, 2, "*", 1)
    'Exit Sub
    If (startingCol < 1) Then
        MsgBox "GenCSV! Unable to find singnal name in row 2 of Testcases sheet!"
        Exit Sub
    End If
      
    'Get Signal data
    retGetSignalData = GetSignalData(ATTSheet, 2, startingCol, inputs(), inputValues(), outputs(), outputValues(), localOutputs(), localOutputValues(), 1, log)
    If (retGetSignalData < 0) Then
        MsgBox "ERROR: GetSignalData" & vbNewLine & log
        Exit Sub
    End If
    'FOR DEBUG
    'If (log <> "") Then
    '    MsgBox "GenCSV!GetSignalData log: " & vbNewLine & log
    'End If
    'For debugging
    'CSVDisplay inputValues(), 1
    'CSVDisplay outputValues(), 20
    'Get Constans value for inputs and change input names (if needed)
    If (OK <= retGetSignalData Or retGetSignalData >= NO_OUTPUT) Then
        'Find constant row number and column number
        'Check constantSet
        consRow = FindRowAll(ConstantSheet, 1, constantSet, 1)
        If (consRow < 0) Then
            MsgBox "Unable to find constant set '" & constantSet & "' in Constants sheet"
        End If
        'Find "CONSTANTS"
        consRow = FindRowAll(ConstantSheet, 2, "CONSTANTS", consRow)
        If (consRow < 0) Then
            MsgBox "Unable to find cell 'CONSTANTS' of contant set '" & constantSet & "' in Constants sheet"
        End If
        consRow = consRow + 1
        consCol = 3
        'Get constant values
        ret = GetConstantValues(ConstantSheet, consRow, consCol, inputs(), inputValues())
        If (ret < 0) Then
            Exit Sub
        End If
        'Find signals row number and column number
         'Check constantSet
        nameRow = FindRowAll(ConstantSheet, 1, constantSet, 1)
        If (nameRow < 0) Then
            MsgBox "Unable to find constant set '" & constantSet & "' in Constants sheet"
        End If
        'Find "SIGNALS"
        nameRow = FindRowAll(ConstantSheet, 2, "SIGNALS", nameRow)
        If (nameRow < 0) Then
            MsgBox "Unable to find cell 'SIGNALS' of contant set '" & constantSet & "' in Constants sheet"
        End If
        nameRow = nameRow + 1
        nameCol = 3
        'Change inputs name
        ret = ChangeSignalNames(ConstantSheet, nameRow, nameCol, inputs(), log)
        If (ret < 0) Then
            MsgBox log
            Exit Sub
        End If
        If (log <> "") Then
            MsgBox log
        End If
    End If
    'Get Constans value for outputs and change output names (if needed)
    If (retGetSignalData <> NO_INPUT_OUTPUT And _
        retGetSignalData <> NO_INPUT_OUTPUT_LOCALOUTPUT And _
        retGetSignalData <> NO_OUTPUT And _
        retGetSignalData <> NO_OUTPUT_LOCALOUTPUT) Then
        'Find constant row number and column number
        'Check constantSet
        consRow = FindRowAll(ConstantSheet, 1, constantSet, 1)
        If (consRow < 0) Then
            MsgBox "Unable to find constant set '" & constantSet & "' in Constants sheet"
        End If
        'Find "CONSTANTS"
        consRow = FindRowAll(ConstantSheet, 2, "CONSTANTS", consRow)
        If (consRow < 0) Then
            MsgBox "Unable to find cell 'CONSTANTS' of contant set '" & constantSet & "' in Constants sheet"
        End If
        consRow = consRow + 1
        consCol = 3
        ret = GetConstantValues(ConstantSheet, consRow, consCol, outputs(), outputValues())
        If (ret < 0) Then
            Exit Sub
        End If
        'Find signals row number and column number
         'Check constantSet
        nameRow = FindRowAll(ConstantSheet, 1, constantSet, 1)
        If (nameRow < 0) Then
            MsgBox "Unable to find constant set '" & constantSet & "' in Constants sheet"
        End If
        'Find "SIGNALS"
        nameRow = FindRowAll(ConstantSheet, 2, "SIGNALS", nameRow)
        If (nameRow < 0) Then
            MsgBox "Unable to find cell 'SIGNALS' of contant set '" & constantSet & "' in Constants sheet"
        End If
        nameRow = nameRow + 1
        nameCol = 3
        'Change output names
        ret = ChangeSignalNames(ConstantSheet, nameRow, nameCol, outputs(), log)
        If (ret < 0) Then
            MsgBox log
            Exit Sub
        End If
        If (log <> "") Then
            MsgBox log
        End If
    End If
    'Get Constans value for local outputs and change local output names (if needed)
    If (retGetSignalData <> NO_INPUT_OUTPUT_LOCALOUTPUT And _
            retGetSignalData <> NO_INPUT_LOCALOUTPUT And _
            retGetSignalData <> NO_OUTPUT_LOCALOUTPUT And _
            retGetSignalData <> NO_LOCALOUTPUT) Then
        'Find constant row number and column number
        'Check constantSet
        consRow = FindRowAll(ConstantSheet, 1, constantSet, 1)
        If (consRow < 0) Then
            MsgBox "Unable to find constant set '" & constantSet & "' in Constants sheet"
        End If
        'Find "CONSTANTS"
        consRow = FindRowAll(ConstantSheet, 2, "CONSTANTS", consRow)
        If (consRow < 0) Then
            MsgBox "Unable to find cell 'CONSTANTS' of contant set '" & constantSet & "' in Constants sheet"
        End If
        consRow = consRow + 1
        consCol = 3
        ret = GetConstantValues(ConstantSheet, consRow, consCol, localOutputs(), localOutputValues())
        If (ret < 0) Then
            Exit Sub
        End If
        'Find signals row number and column number
         'Check constantSet
        nameRow = FindRowAll(ConstantSheet, 1, constantSet, 1)
        If (nameRow < 0) Then
            MsgBox "Unable to find constant set '" & constantSet & "' in Constants sheet"
        End If
        'Find "SIGNALS"
        nameRow = FindRowAll(ConstantSheet, 2, "SIGNALS", nameRow)
        If (nameRow < 0) Then
            MsgBox "Unable to find cell 'SIGNALS' of contant set '" & constantSet & "' in Constants sheet"
        End If
        nameRow = nameRow + 1
        nameCol = 3
        'Change local output names
        ret = ChangeSignalNames(ConstantSheet, nameRow, nameCol, localOutputs(), log)
        If (ret < 0) Then
            MsgBox log
            Exit Sub
        End If
        If (log <> "") Then
            MsgBox log
        End If
    End If
    'For debugging
    'CSVDisplay inputValues(), 40
    'CSVDisplay outputValues(), 60
    
    'Start Writing out csv files
    'Loop for all TCs
    For i = 0 To UBound(inputValues, 2)
        writefile = FreeFile
        TCDigit1 = Int(((i + 1) Mod 1000) / 100)
        TCDigit2 = Int(((i + 1) Mod 100) / 10)
        TCDigit3 = Int(((i + 1) Mod 10))
        fPath = directoryPath & "\" & testModuleName & "_TC" & TCDigit1 & TCDigit2 & TCDigit3 & ".csv"
        line1 = "'Time';'moduleIndex';"
        line2 = "'s';'-';"
        line3 = "0;" & moduleIndex & ";"
        line4 = cycleTime & ";" & moduleIndex & ";"
        mainData = ""
        'Check whether inputs exist
        If (OK <= retGetSignalData Or retGetSignalData >= NO_OUTPUT) Then
            'Loop for all input names
            For j = 0 To UBound(inputs)
                line1 = line1 & "'" & inputs(j) & "." & testModuleName & "';"
                line2 = line2 & "'-';"
                line3 = line3 & inputValues(j, i) & ";"
                line4 = line4 & inputValues(j, i) & ";"
            Next j
        End If
        'Check whether outputs exist
        If (retGetSignalData <> NO_INPUT_OUTPUT And _
            retGetSignalData <> NO_INPUT_OUTPUT_LOCALOUTPUT And _
            retGetSignalData <> NO_OUTPUT And _
            retGetSignalData <> NO_OUTPUT_LOCALOUTPUT) Then
            'Loop for all output names
            For j = 0 To UBound(outputs)
                'line1 = line1 & "'expected_" & outputs(j) & "." & testModuleName & "'"
                line1 = line1 & "'" & outputs(j) & "." & testModuleName & "'"
                line2 = line2 & "'-'"
                line3 = line3 & outputValues(j, i)
                line4 = line4 & outputValues(j, i)
                If j <> UBound(outputs) Then
                    line1 = line1 & ";"
                    line2 = line2 & ";"
                    line3 = line3 & ";"
                    line4 = line4 & ";"
                End If
            Next j
            'Check whether local outputs exist
            If (retGetSignalData <> NO_INPUT_OUTPUT_LOCALOUTPUT And _
                retGetSignalData <> NO_INPUT_LOCALOUTPUT And _
                retGetSignalData <> NO_OUTPUT_LOCALOUTPUT And _
                retGetSignalData <> NO_LOCALOUTPUT) Then
                line1 = line1 & ";"
                line2 = line2 & ";"
                line3 = line3 & ";"
                line4 = line4 & ";"
                 'Loop for all output names
                For j = 0 To UBound(localOutputs)
                    'line1 = line1 & ";'exp_asp_" & localOutputs(j) & "." & testModuleName & "'"
                    line1 = line1 & "'" & localOutputs(j) & "." & testModuleName & "'"
                    line2 = line2 & "'-'"
                    line3 = line3 & localOutputValues(j, i)
                    line4 = line4 & localOutputValues(j, i)
                    If j <> UBound(localOutputs) Then
                        line1 = line1 & ";"
                        line2 = line2 & ";"
                        line3 = line3 & ";"
                        line4 = line4 & ";"
                    End If
                Next j
            End If
        'Check whether local outputs exist
        ElseIf (retGetSignalData <> NO_INPUT_OUTPUT_LOCALOUTPUT And _
                retGetSignalData <> NO_INPUT_LOCALOUTPUT And _
                retGetSignalData <> NO_OUTPUT_LOCALOUTPUT And _
                retGetSignalData <> NO_LOCALOUTPUT) Then
             'Loop for all output names
            For j = 0 To UBound(localOutputs)
                'line1 = line1 & "'exp_asp_" & localOutputs(j) & "." & testModuleName & "'"
                line1 = line1 & "'" & localOutputs(j) & "." & testModuleName & "'"
                line2 = line2 & "'-'"
                line3 = line3 & localOutputValues(j, i)
                line4 = line4 & localOutputValues(j, i)
                If j <> UBound(localOutputs) Then
                    line1 = line1 & ";"
                    line2 = line2 & ";"
                    line3 = line3 & ";"
                    line4 = line4 & ";"
                End If
            Next j
        End If
        'mainData
        mainData = line1 & vbNewLine & _
                    line2 & vbNewLine & _
                    line3 & vbNewLine & _
                    line4
        'Write main data to csv file
        Open fPath For Output As writefile
        Print #writefile, mainData
        Close writefile
    Next i
   
    'Promt to user: sucessful
    MsgBox "Testcases generated: " & (UBound(inputValues, 2) + 1) & vbNewLine & _
            "Please check the folder: " & directoryPath
    Exit Sub
GenCSVErrorHandler:
    MsgBox "ERROR! GenCSV" & vbNewLine & _
            Err.Number & vbCr & Err.description
End Sub
Function GetConstantValues(ByVal sheet As Variant, _
                        ByVal row As Integer, _
                        ByVal col As Integer, _
                        ByRef signals() As String, _
                        ByRef signalValues() As String)
    '***********************************************************
    '
    'AUTHOR:            Le Thai (RBVH\EMB3)
    'DATE CREATED:      24/May/2013
    'FUNCTION NAME:     GetConstantValues
    'DESCRIPTION:
    '       INPUT:
    '               sheet:    sheet number
    '               row:   row number of starting cell of the constants table
    '               col:   column number of starting cell of the constants table
    '
    '       OUTPUT:
    '               signals()    :   array contains all the signals names
    '               signalValues()  :   array contains all the values of all signals
    '
    '***********************************************************
    Dim i, j As Integer
    Dim colIdx As Integer
    Dim temp As Variant
    'Loop for signals
    For i = 0 To UBound(signals)
        For j = 0 To UBound(signalValues, 2)
            'Find the column number of the constant name in Constants sheet
            colIdx = FindColAll(sheet, row, signalValues(i, j), col)
            'Found
            If colIdx > 0 Then
                signalValues(i, j) = Worksheets(sheet).Cells(row + 1, colIdx).value
            'is numeric
            ElseIf IsNumeric(signalValues(i, j)) Then
                
            'Not found
            Else
                'Ask user to create constant value
                temp = MsgBox("There is no constant named '" & signalValues(i, j) & "'" & vbNewLine & _
                                "signal name: '" & signals(i) & "'" & vbNewLine & _
                                "Test case: TC" & (j + 1) & vbNewLine & _
                                "Create value for this contant name?", vbYesNo + vbExclamation, "GetConstantValues")
                'YES
                If (temp = vbYes) Then
                    'Get constant value
                    temp = Application.InputBox(Prompt:="Please enter a numeric value for constant name '" & signalValues(i, j) & "'!", _
                                                Title:="GET VALUE", _
                                                Default:=0, _
                                                Type:=1)
                    If (IsNumeric(temp)) Then
                        'Find colum number for first empty cell
                        colIdx = FindColAll(sheet, row, " ", col)
                        'Write constant name to the sheet
                        Worksheets(sheet).Cells(row, colIdx).value = signalValues(i, j)
                        'Write constant value to the sheet
                        Worksheets(sheet).Cells(row + 1, colIdx).value = temp
                        'Update signalValues
                        signalValues(i, j) = temp
                    'Invalid input
                    Else
                        MsgBox "Invalid! Input must be a numeric value." & vbNewLine & _
                                "Please run the macro again!"
                        'Exit
                        GetConstantValues = -1
                        Exit Function
                    End If
                'NO
                Else
                    'Exit
                    GetConstantValues = -1
                    Exit Function
                End If
            End If
        Next j
    Next i
    GetConstantValues = 1
End Function
Function ChangeSignalNames(ByVal sheet As Variant, _
                        ByVal row As Integer, _
                        ByVal col As Integer, _
                        ByRef signals() As String, _
                        ByRef log As String)
    '***********************************************************
    '
    'AUTHOR:            Le Thai (RBVH\EMB3)
    'DATE CREATED:      24/May/2013
    'FUNCTION NAME:     ChangeSignalNames
    'DESCRIPTION:
    '       INPUT:
    '               sheet:    sheet number
    '               row:   row number of starting cell of the constants table
    '               col:   column number of starting cell of the constants table
    '
    '       OUTPUT:
    '               signals()    :   array contains all the signals names
    '               log          :   log
    '
    '***********************************************************
    Dim i, j As Integer
    Dim colIdx As Integer
    'Loop for signals
    log = ""
    For i = 0 To UBound(signals)
        colIdx = FindColAll(sheet, row, signals(i), col)
        'Found
        If (colIdx > 0) Then
            'Check whether replacing value exists
            'Replacing value exist
            If (Worksheets(sheet).Cells(row + 1, colIdx) <> vbNullString) Then
                'log
                log = log & "Signal name '" & signals(i) & "' was replaced by '" & Worksheets(sheet).Cells(row + 1, colIdx).value & "'!" & vbNewLine
                'Change name
                signals(i) = Worksheets(sheet).Cells(row + 1, colIdx).value
            'Replacing value does not exist
            Else
                log = "Replacing name for signal name '" & signals(i) & "' is empty!" & vbNewLine & _
                        "Please check and provide replacing name for the same in 'Constant' sheet!"
                ChangeSignalNames = -1
                Exit Function
            End If
        'Not found
        Else
        
        End If
    Next i
    ChangeSignalNames = 1
End Function
Function GetSignalData(ByVal sheet As Variant, _
                        ByVal row As Integer, _
                        ByVal col As Integer, _
                        ByRef inputs() As String, _
                        ByRef inputValues() As String, _
                        ByRef outputs() As String, _
                        ByRef outputValues() As String, _
                        ByRef localOutputs() As String, _
                        ByRef localOutputValues() As String, _
                        ByVal kind As Integer, _
                        ByRef log As String) As Integer
    '***********************************************************
    '
    'AUTHOR:            Le Thai (RBVH\EMB3)
    'DATE CREATED:      24/May/2013
    'FUNCTION NAME:     GetSignalData
    'DESCRIPTION:
    '       INPUT:
    '               sheet:    sheet number
    '               row:   row number of starting cell of the table
    '               col:   column number of starting cell of the table
    '               kind    :   to indicate that reserved values are used or not
    '                           1   :   used
    '                           0   :   not used
    '       OUTPUT:
    '               inputs()    :   array contains all the inputs names
    '               outputs()    :   array contains all the inputs names
    '               inputValues()  :   array contains all the values of all inputs
    '               inputValues()  :   array contains all the values of all inputs
    '               log             :   log
    '
    '***********************************************************
    Dim nInputs, nOutputs, nValues As Integer
    Dim nLocalOutputs, nLocalOutputValues As Integer
    Dim nInputValues, nOutputValues As Integer
    Dim inputReservedValue As String
    Dim outputReservedValue As String
    Dim localOutputReservedValue As String
    Dim signalFlag As String
    Dim i, j As Integer
    Dim rowIdx, colIdx As Integer
    Dim iName, oName As String
    log = ""
    
    'Check for signalFlag
    If (Worksheets(sheet).Cells(row - 1, col) = "OUTPUTS") Then
        signalFlag = "output"
    ElseIf (Worksheets(sheet).Cells(row - 1, col) = "LOCAL OUTPUTS") Then
        signalFlag = "local_output"
    Else
        signalFlag = "input"
    End If
    'Initial inputs()
    nInputs = 0
    ReDim inputs(0 To 0)
    'Initial outputs()
    nOutputs = 0
    ReDim outputs(0 To 0)
    'Initial localOutputs()
    nLocalOutputs = 0
    ReDim localOutputs(0 To 0)
    colIdx = col
    'Count the number of inputs, outputs and local outputs
    Do While (Worksheets(sheet).Cells(row, colIdx) <> vbNullString)
        If (signalFlag = "input") Then
            nInputs = nInputs + 1
            If (Worksheets(sheet).Cells(row - 1, colIdx + 1) = "OUTPUTS") Then
                signalFlag = "output"
            ElseIf (Worksheets(sheet).Cells(row - 1, colIdx + 1) = "LOCAL OUTPUTS") Then
                signalFlag = "local_output"
            End If
        ElseIf (signalFlag = "output") Then
            nOutputs = nOutputs + 1
            If (Worksheets(sheet).Cells(row - 1, colIdx + 1) = "LOCAL OUTPUTS") Then
                signalFlag = "local_output"
            End If
        Else
            nLocalOutputs = nLocalOutputs + 1
        End If
        colIdx = colIdx + 1
    Loop
    'Count the number of values
    nValues = 0
    rowIdx = row + 1
    Do While (Worksheets(sheet).Cells(rowIdx, col) <> vbNullString)
        nValues = nValues + 1
        rowIdx = rowIdx + 1
    Loop
    
    If (nValues > 0) Then
        If (nInputs > 0) Then
            ReDim Preserve inputs(0 To nInputs - 1)
            ReDim Preserve inputValues(0 To nInputs - 1, 0 To nValues - 1)
        Else
            log = "WARNING: nInputs = " & nInputs & vbNewLine
        End If
        If (nOutputs > 0) Then
            ReDim Preserve outputs(0 To nOutputs - 1)
            ReDim Preserve outputValues(0 To nOutputs - 1, 0 To nValues - 1)
        Else
            log = log & "WARNING: outputs = " & nOutputs & vbNewLine
        End If
        If (nLocalOutputs > 0) Then
            ReDim Preserve localOutputs(0 To nLocalOutputs - 1)
            ReDim Preserve localOutputValues(0 To nLocalOutputs - 1, 0 To nValues - 1)
        Else
            log = log & "INFO: There is no local output."
        End If
    Else
        log = log & "WARNING: nValues = " & nValues & vbNewLine
        GetSignalData = ERROR
        Exit Function
    End If
    'FOR DEBUG
    'MsgBox "nInputs: " & nInputs & vbNewLine & _
            "nOutputs: " & nOutputs & vbNewLine & _
            "nValues: " & nValues
    'Get input data
    For i = 0 To (nInputs - 1)
        iName = Worksheets(sheet).Cells(row, col + i).value
        'Check name
        If (CheckSignalName(inputs(), i - 1, iName, oName, log) < 0) Then
            log = "GetSignalData!" & log
            GetSignalData = ERROR
            Exit Function
        End If
        inputs(i) = iName
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
        iName = Worksheets(sheet).Cells(row, col + nInputs + i).value
        'Check name
        If (CheckSignalName(outputs(), i - 1, iName, oName, log) < 0) Then
            log = "GetSignalData!" & log
            GetSignalData = ERROR
            Exit Function
        End If
        outputs(i) = iName
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
    'Get local output data
    For i = 0 To (nLocalOutputs - 1)
        'Get name
        iName = Worksheets(sheet).Cells(row, col + nInputs + nOutputs + i).value
        'Check name
        If (CheckSignalName(localOutputs(), i - 1, iName, oName, log) < 0) Then
            log = "GetSignalData!" & log
            GetSignalData = ERROR
            Exit Function
        End If
        localOutputs(i) = iName
        'Set initial values
        localOutputReservedValue = "0"
        'Get value
        For j = 0 To (nValues - 1)
            'Get value
            'kind = 1 --> used reserved value
            If (kind = 1) Then
                'Normal value
                If (Worksheets(sheet).Cells(row + j + 1, col + nInputs + nOutputs + i) <> vbNullString And _
                    Worksheets(sheet).Cells(row + j + 1, col + nInputs + nOutputs + i) <> "X" And _
                    Worksheets(sheet).Cells(row + j + 1, col + nInputs + nOutputs + i) <> "x") Then
                    localOutputValues(i, j) = Worksheets(sheet).Cells(row + j + 1, col + nInputs + nOutputs + i).value
                    'Reserve the value
                    localOutputReservedValue = Worksheets(sheet).Cells(row + j + 1, col + nInputs + nOutputs + i).value
                Else
                    localOutputValues(i, j) = localOutputReservedValue
                End If
            'kind = 0 --> not used reserved value
            'Undefined value
            Else
                'Cell is not null
                If (Worksheets(sheet).Cells(row + j + 1, col + nInputs + nOutputs + i) <> vbNullString) Then
                    localOutputValues(i, j) = Worksheets(sheet).Cells(row + j + 1, col + nInputs + nOutputs + i).value
                'Cell is null
                Else
                    localOutputValues(i, j) = ""
                End If
            End If
        Next j
    Next i
    If (nInputs = 0) Then
        GetSignalData = NO_INPUT
        If (nOutputs = 0) Then
            GetSignalData = NO_INPUT_OUTPUT
            If (nLocalOutputs = 0) Then
                GetSignalData = NO_INPUT_OUTPUT_LOCALOUTPUT
            End If
        ElseIf (nLocalOutputs = 0) Then
            GetSignalData = NO_INPUT_LOCALOUTPUT
        End If
    ElseIf (nOutputs = 0) Then
        GetSignalData = NO_OUTPUT
        If (nLocalOutputs = 0) Then
            GetSignalData = NO_OUTPUT_LOCALOUTPUT
        End If
    ElseIf (nLocalOutputs = 0) Then
        GetSignalData = NO_LOCALOUTPUT
    Else
        GetSignalData = OK
    End If
End Function
Function CheckSignalName(ByRef signals() As String, _
                        ByVal lastIndex As Integer, _
                        ByVal iName As String, _
                        ByRef oName As String, _
                        ByRef log As String) As Integer
    '***********************************************************
    '
    'AUTHOR:            Le Thai (RBVH\EMB3)
    'DATE CREATED:      24/May/2013
    'FUNCTION NAME:     GetSignalData
    'DESCRIPTION:
    '       INPUT:
    '               signals      :   Array of signal names
    '               lastIndex    :   last index in array to check
    '               iName        :   name to check
    '       OUTPUT:
    '               log          :   log
    '               oName        :   iName after removing "."
    '               return values:
    '                               OK      : OK
    '                               ERROR   : ERROR
    '
    '***********************************************************
    Dim i As Integer
    Dim dotPos As Integer
    Dim tmpName As String
    'First element, no need to check
    If (lastIndex < 0) Then
        '----------- Check whether iName contains space character --
        If (InStr(iName, " ") > 0) Then
            log = "CheckSignalName! Signal name '" & iName & "' contains space charater ' '!"
            CheckSignalName = CHECKSIGNALNAME_SPACE
            Exit Function
        End If
         oName = iName
        'Check "." character
        dotPos = InStr(oName, ".")
        If (dotPos > 1) Then
            'Remove ".", get name only
            oName = Mid(oName, 1, dotPos - 1)
        End If
    Else
        '----------- Check whether iName contains space character --
        If (InStr(iName, " ") > 0) Then
            log = "CheckSignalName! Signal name '" & iName & "' contains space charater ' '!"
            CheckSignalName = CHECKSIGNALNAME_SPACE
            Exit Function
        End If
        '----------- Check whether the iName is duplicated ---------
        '----------- E.g "ay" and "ay.abs()" ----------------------
        oName = iName
        'Check "." character
        dotPos = InStr(oName, ".")
        If (dotPos > 1) Then
            'Remove ".", get name only
            oName = Mid(oName, 1, dotPos - 1)
        End If
        'Check duplicated
        For i = 0 To lastIndex
            tmpName = signals(i)
            'Check "." character
            dotPos = InStr(tmpName, ".")
            If (dotPos > 1) Then
                'Remove ".", get name only
                tmpName = Mid(tmpName, 1, dotPos - 1)
            End If
        
            If (tmpName = oName) Then
                log = "CheckSignalName! Signal name '" & iName & "' is duplicated with '" & signals(i) & "'!"
                CheckSignalName = CHECKSIGNALNAME_DUPLICATE
                Exit Function
            End If
        Next i
    End If
    CheckSignalName = OK
End Function

