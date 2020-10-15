Attribute VB_Name = "MGenTC"
Option Explicit
Const OK = 1
Const ERROR = -1
'$$$$$$$$$$$$ DEFINITIONS FOR "MCDC" SHEET $$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
Const MCDC_TC_NO_COL = 2
Const MCDC_MCDC_COL = 3
Const MCDC_INPUT_COL = 4
'$$$$$$$$$$$$ DEFINITIONS FOR "Testcases" SHEET $$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
Const TC_TC_NO_COLUMN = 1
Sub GenTC()
    '***********************************************************
    '
    'AUTHOR:            Le Thai (RBVH\EMB3)
    'DATE CREATED:      04/Jul/2013
    'FUNCTION NAME:     GenTC
    'DESCRIPTION:
    '       INPUT: sheet MDCC sheet
    '
    '       OUTPUT: testcase design in Testcases sheet
    '
    '***********************************************************
    Dim MCDCSheet, TCSheet As String
    Dim iRow, iRowIdx, iCol, iColIdx As Integer
    Dim oRow, oRowIdx, oCol, oColIdx As Integer
    Dim outcomeRow, outcomeCol As Integer
    Dim startingORow As Integer
    Dim inputs() As String
    Dim inputValues() As String
    Dim outputs() As String
    Dim outputValues() As String
    'Dim localOutputs() As String
    'Dim localOutputValues() As String
    Dim TCNo() As String
    Dim outcome() As String
    Dim i, j As Integer
    Dim condition As String
    Dim copiedCondition As String
    Dim log As String
    Dim TCID As Integer
    Dim temp As Variant
    Dim TCNoValue As String
    Dim subTCNoValue As String
    Dim prefixDesc As String
    Dim description, finalDescription As String
    Dim retGetOutcome As Integer
    Dim specifiedTC As String
    
    MCDCSheet = "MCDC"
    TCSheet = "Testcases"
    description = ""
    
    'Back up
    Backup
    'Confirmation
    temp = MsgBox("GenTC! Generate all TCs or specified TC?" & vbNewLine & _
                    "Yes: Generate all TCs" & vbNewLine & _
                    "No: Generate specified TC", _
                    vbYesNoCancel + vbExclamation, "Attention!")
    'OK
    If (temp = vbYes) Then
    'CANCLE
        specifiedTC = "*"
    ElseIf (temp = vbNo) Then
        'Get Certain TC
        specifiedTC = Application.InputBox(Prompt:="Please enter a TCID '", _
                                                    Title:="GET TCID", _
                                                    Default:="TCX, TCY", _
                                                    Type:=2)
    Else
        'Exit sub
        Exit Sub
    End If

    'Find the row number for requirement
    iCol = 1
    iRow = FindRowAll(MCDCSheet, iCol, "*", 1)
    'Find "TC No." in Testcases sheet
    startingORow = FindRowAll(TCSheet, TC_TC_NO_COLUMN, "TC No.", 1)
    If (startingORow < 0) Then
        MsgBox "GenTC! Unable to to find cell '" & TCSheet & "' sheet"
        Exit Sub
    End If
    startingORow = startingORow + 2
    'Loop for all requirements
    'oRow = 3
    oRow = FindRowAll(TCSheet, TC_TC_NO_COLUMN, " ", startingORow)
    TCID = 1
    Do While (iRow > 0)
        'FOR DEBUG
        'MsgBox "Requirement: " & Worksheets(MCDCSheet).Cells(iRow, iCol).value
        'Get description
        'If (Worksheets(MCDCSheet).Cells(iRow + 1, iCol + 1) <> vbNullString) Then
        '    description = "+ " & Worksheets(MCDCSheet).Cells(iRow + 1, iCol + 1).Value
        'End If
        description = "+ Check [" & Worksheets(MCDCSheet).Cells(iRow, iCol).value & "]"
        'Find row number for the string "TC No."
        iColIdx = 2
        iRowIdx = FindRowAll(MCDCSheet, iColIdx, "TC No.", iRow)
        If (iRowIdx < 0) Then
            MsgBox "GenTC! Unable to find String 'TC No.' for the requirement '" & Sheets(MCDCSheet).Cells(iRow, iCol) & "'"
            Exit Sub
        End If
        'For debugging
        'MsgBox "iRow: " & iRow & vbNewLine & _
                "iCol: " & iCol & vbNewLine & _
                "iRowIdx: " & iRowIdx & vbNewLine & _
                "iColIdx: " & iColIdx
        'Get signal data
        temp = GetSignalDataAll(MCDCSheet, iRowIdx, 4, inputs(), outputs(), inputValues(), outputValues(), 0, log)
        'temp = GetSignalData(MCDCSheet, iRowIdx, 4, inputs(), inputValues(), outputs(), outputValues(), localOutputs(), localOutputValues(), 0, log)
        If (temp < 0) Then
            MsgBox "GenTC! Requirement: " & Worksheets(MCDCSheet).Cells(iRow, iCol).value & vbNewLine & log
            Exit Sub
        End If
        'For debugging
        'CSVDisplay inputValues(), 40
        'CSVDisplay outputValues(), 60
        'Get TCNo information
        temp = GetTCNoInfo(MCDCSheet, iRowIdx + 1, 2, TCNo(), log)
        If (temp < 0) Then
            MsgBox "Requirement: '" & Sheets(MCDCSheet).Cells(iRow, iCol) & "'" & vbNewLine & log
            Exit Sub
        End If
        'Check the size of inputs, outputs, and TCNo
        If (UBound(inputValues, 2) = UBound(outputValues, 2) And UBound(inputValues, 2) = UBound(TCNo)) Then
            'Do nothing
        Else
            MsgBox "GenTC! Requirement '" & Sheets(MCDCSheet).Cells(iRow, iCol) & "'" & vbNewLine & _
                    "TCNo, INPUTS, OUTPUTS in MCDC table must be fulfil." & vbNewLine & _
                    "Please fulfil the table and try again!"
            Exit Sub
        End If
        
        'Get outcome information
        'Get column number and row number for "OUTCOME"
        outcomeRow = iRowIdx - 1
        outcomeCol = FindColAll(MCDCSheet, outcomeRow, "OUTCOME", 1)
        If (outcomeCol < 0) Then
            MsgBox "GenTC! Unable to find String 'OUTCOME' for the requirement '" & Sheets(MCDCSheet).Cells(iRow, iCol) & "'"
            'Exit Sub
        End If
        'Start Getting outcome information
        retGetOutcome = GetOutcomeInfo(MCDCSheet, outcomeRow + 2, outcomeCol, outcome(), log)
        If (retGetOutcome < 0) Then
            MsgBox "GenTC! Requirement '" & Sheets(MCDCSheet).Cells(iRow, iCol) & "'" & vbNewLine & log
            'Exit Sub
        End If
        'For debugging
        'MsgBox "iRow: " & iRow & vbNewLine & _
                "iCol: " & iCol & vbNewLine & _
                "iRowIdx: " & iRowIdx & vbNewLine & _
                "iColIdx: " & iColIdx
        'Loop for all TC
        For i = 0 To UBound(inputValues, 2)
            condition = ""
            'Condition for inputs
            For j = 0 To UBound(inputs)
                If (j = 0) Then
                    condition = inputs(j) & "=" & inputValues(j, i)
                Else
                    condition = condition & " && " & inputs(j) & "=" & inputValues(j, i)
                End If
            Next j
            'Condition for outputs
            For j = 0 To UBound(outputs)
                condition = condition & " && " & outputs(j) & "=" & outputValues(j, i)
            Next j
            'add outcome info to description
            finalDescription = description
            If (description <> "" And retGetOutcome >= 0) Then
                 finalDescription = description & " with outcome " & outcome(i)
            End If
            'Extract TCNo
            'FOR DEBUG
            'MsgBox "Requirement: " & Worksheets(MCDCSheet).Cells(iRow, iCol).value & vbNewLine & _
                    "TCNo: " & TCNo(i)
            'TCNoValue
            TCNoValue = TCNo(i)
            Do While (ExtractTCNo(TCNoValue, subTCNoValue, log) > 0)
                If (CheckSpecifiedTC(specifiedTC, subTCNoValue) = True Or specifiedTC = "*") Then
                    oRowIdx = FindRowAll(TCSheet, TC_TC_NO_COLUMN, subTCNoValue, startingORow)
                    copiedCondition = condition
                    'Prefix for testcase description
                    prefixDesc = "Please refer to " & subTCNoValue & " in " & ActiveWorkbook.name
                    'FOR DEBUG
                    'MsgBox "subTCNoValue: " & subTCNoValue & vbNewLine & _
                            "oRowIdx: " & oRowIdx
                    'Found
                    If (oRowIdx > 0) Then
                        'Write TC
                        temp = WriteTCAll(TCSheet, copiedCondition, oRowIdx, 0, log, startingORow - 1)
                        If (temp < 0) Then
                            MsgBox log
                            Exit Sub
                        End If
                        'FOR DEBUG
                        'If (log <> "") Then
                        '    MsgBox log
                        'End If
                        'Find column "DESCRIPTIONS"
                        temp = FindColAll(TCSheet, startingORow - 2, "DESCRIPTIONS", 1)
                        'Start writting description
                        If (temp > 0) Then
                            ' Description <> ""
                            If (finalDescription <> "") Then
                                If (Worksheets(TCSheet).Cells(oRowIdx, temp) <> vbNullString) Then
                                    Worksheets(TCSheet).Cells(oRowIdx, temp).value = Worksheets(TCSheet).Cells(oRowIdx, temp).value & vbNewLine & finalDescription
                                Else
                                    Worksheets(TCSheet).Cells(oRowIdx, temp).value = prefixDesc & vbNewLine & finalDescription
                                End If
                            'No description
                            Else
                                'FOR DEBUG
                                'MsgBox "description = null"
                            End If
                        'Not found column "DESCRIPTIONS"
                        Else
                            'FOR DEBUG
                            'MsgBox "Not found column 'DESCRIPTIONS'"
                        End If
                    'Not found
                    Else    'If (oRowIdx > 0) Then
                        'Write TC No.
                        'Worksheets(TCSheet).Cells(oRow, TC_TC_NO_COLUMN).value = "TC" & TCID
                        Worksheets(TCSheet).Cells(oRow, TC_TC_NO_COLUMN).value = subTCNoValue
                        'Write TC
                        temp = WriteTCAll(TCSheet, copiedCondition, oRow, 0, log, startingORow - 1)
                        If (temp < 0) Then
                            MsgBox log
                            Exit Sub
                        End If
                        'For Debugging
                        'If (log <> vbNullString) Then
                        '    MsgBox log
                        'End If
                        'Stick "X" for empty cells
                        'Find starting column number for signal
                        oColIdx = FindColAll(TCSheet, startingORow - 1, "*", 1)
                        If (oColIdx < 0) Then
                            MsgBox "GenTC! Unable to stick 'X' for dont care signal"
                            Exit Sub
                        End If
                        'Stick "X"
                        Do While (Worksheets(TCSheet).Cells(startingORow - 1, oColIdx) <> vbNullString)
                            If (Worksheets(TCSheet).Cells(oRow, oColIdx) = vbNullString) Then
                                Worksheets(TCSheet).Cells(oRow, oColIdx).value = "X"
                            End If
                            oColIdx = oColIdx + 1
                        Loop
                        'Write description
                        'Find column "DESCRIPTIONS"
                        temp = FindColAll(TCSheet, startingORow - 2, "DESCRIPTIONS", 1)
                        'Start writting description
                        If (temp > 0) Then
                            ' Description <> ""
                            If (finalDescription <> "") Then
                                Worksheets(TCSheet).Cells(oRow, temp).value = prefixDesc & vbNewLine & finalDescription
                            'No description
                            Else
                                'FOR DEBUG
                                MsgBox "finalDescription = null"
                            End If
                        'Not found column "DESCRIPTIONS"
                        Else
                            'FOR DEBUG
                            MsgBox "Not found column 'DESCRIPTIONS'"
                        End If
                        'Next oRow
                        oRow = oRow + 1
                        'TCID
                        TCID = TCID + 1
                    End If 'If (oRowIdx > 0) Then ... Else ...
                End If  'If (subTCNoValue = specifiedTC) Then
            Loop    'Do While (ExtractTCNo(TCNoValue, subTCNoValue, log) > 0)
            If (log <> "") Then
                MsgBox "Requirement: " & Worksheets(MCDCSheet).Cells(iRow, iCol).value & vbNewLine & log
                Exit Sub
            End If
        Next i
        'Next row for requirement
            'Find the row number for requirement
        iCol = 1
        iRow = FindRowAll(MCDCSheet, iCol, "*", iRow + 1)
    Loop
    AutofitColumns (TCSheet)
End Sub
Function CheckSpecifiedTC(ByVal specifiedTCs As String, _
                        ByVal TCStr As String) As Boolean
    '***********************************************************
    '
    'AUTHOR:            Le Thai (RBVH\EMB3)
    'DATE CREATED:      24/May/2013
    'FUNCTION NAME:     CheckSpecifiedTC
    'DESCRIPTION:
    '       INPUT:
    '               specifiedTCs:    sheet number
    '               TCStr:   row number of starting cell of the table
    '
    '       OUTPUT:
    '
    '***********************************************************
    Dim TCIDs() As String
    Dim patterns() As String
    Dim trimmedSpecifiedTCs As String
    Dim i As Integer
    'Remove " ", tab
    'Note: chr(9) = tab charater
    patterns = Split(" ," & Chr(9), ",")
    trimmedSpecifiedTCs = ReplaceAll(specifiedTCs, patterns, "")
    
    TCIDs = Split(trimmedSpecifiedTCs, ",")
    For i = 0 To UBound(TCIDs)
        If (TCIDs(i) = TCStr) Then
            CheckSpecifiedTC = True
            Exit Function
        End If
    Next i
    CheckSpecifiedTC = False
End Function
Function GetOutcomeInfo(ByVal sheet As Variant, _
                        ByVal row As Integer, _
                        ByVal col As Integer, _
                        ByRef outcome() As String, _
                        ByRef log As String) As Integer
    '***********************************************************
    '
    'AUTHOR:            Le Thai (RBVH\EMB3)
    'DATE CREATED:      24/May/2013
    'FUNCTION NAME:     GetOutcomeInfo
    'DESCRIPTION:
    '       INPUT:
    '               sheet:    sheet number
    '               row:   row number of starting cell of the table
    '               col:   column number of starting cell of the table
    '
    '       OUTPUT:
    '               TCNo()  :   array contains all the values of all TCNo
    '               log     :   log
    '
    '***********************************************************
    Dim nOutcome As Integer
    Dim rowIdx As Integer
    Dim i, j As Integer
    Dim firstFlag As Boolean
    log = ""
    'Count the number of TCNo
    nOutcome = 0
    rowIdx = row
    Do While (Worksheets(sheet).Cells(rowIdx, col) <> vbNullString)
        nOutcome = nOutcome + 1
        rowIdx = rowIdx + 1
    Loop
    'Check nOutcome to declare nOutcome array
    If (nOutcome > 0) Then
        ReDim Preserve outcome(0 To nOutcome - 1)
    Else
        log = log & "There is no value in OUTCOME column" & vbNewLine
        GetOutcomeInfo = ERROR
        Exit Function
    End If
    'Get nOutcome data
    For i = 0 To nOutcome - 1
        outcome(i) = Worksheets(sheet).Cells(row + i, col).value
        firstFlag = True
        For j = MCDC_INPUT_COL To col
            If (Worksheets(sheet).Cells(row + i, j).Interior.ColorIndex <> xlNone) Then
                 If (firstFlag = True) Then
                    outcome(i) = outcome(i) & _
                                   ", focus on " & _
                                   Worksheets(sheet).Cells(row - 1, j).value & _
                                   " is set " & _
                                   Worksheets(sheet).Cells(row + i, j).value
                    firstFlag = False
                Else
                    outcome(i) = outcome(i) & _
                                   ", " & _
                                   Worksheets(sheet).Cells(row - 1, j).value & _
                                   " is set " & _
                                   Worksheets(sheet).Cells(row + i, j).value
                End If
            End If
        Next j
    Next i
    GetOutcomeInfo = OK
End Function
Function GetTCNoInfo(ByVal sheet As Variant, _
                        ByVal row As Integer, _
                        ByVal col As Integer, _
                        ByRef TCNo() As String, _
                        ByRef log As String) As Integer
    '***********************************************************
    '
    'AUTHOR:            Le Thai (RBVH\EMB3)
    'DATE CREATED:      24/May/2013
    'FUNCTION NAME:     GetTCNoInfo
    'DESCRIPTION:
    '       INPUT:
    '               sheet:    sheet number
    '               row:   row number of starting cell of the table
    '               col:   column number of starting cell of the table
    '
    '       OUTPUT:
    '               TCNo()  :   array contains all the values of all TCNo
    '               log     :   log
    '
    '***********************************************************
    Dim nTCNos As Integer
    Dim rowIdx As Integer
    Dim i As Integer
    log = ""
    'Count the number of TCNo
    nTCNos = 0
    rowIdx = row
    Do While (Worksheets(sheet).Cells(rowIdx, col) <> vbNullString)
        nTCNos = nTCNos + 1
        rowIdx = rowIdx + 1
    Loop
    'Check nTCNos to declare TCNo array
    If (nTCNos > 0) Then
        ReDim Preserve TCNo(0 To nTCNos - 1)
    Else
        log = log & "There is no TCNo." & vbNewLine
        GetTCNoInfo = ERROR
        Exit Function
    End If
    'Get TCNo data
    For i = 0 To nTCNos - 1
        TCNo(i) = Worksheets(sheet).Cells(row + i, col).value
    Next i
    GetTCNoInfo = OK
End Function
Function ExtractTCNo(ByRef TCNoValue As String, _
                            ByRef subTCNoValue As String, _
                            ByRef log As String) As Integer
    '***********************************************************
    '
    'AUTHOR:            Le Thai (RBVH\EMB3)
    'DATE CREATED:      28/May/2013
    'FUNCTION NAME:     ExtractTCNo
    'DESCRIPTION:
    '       INPUT:
    '           TCNo    :   TCNo
    '       OUTPUT:
    '           subTCNo :  subTCNo
    '           log     :  log
    '
    '***********************************************************
    Dim commaPos As Integer
    Dim copiedTCNoValue As String
    Dim patterns() As String
    copiedTCNoValue = TCNoValue
    log = ""
    'Check input
    If (TCNoValue = vbNullString Or Len(TCNoValue) = vbNull) Then
        'MsgBox "ExtractTCNo! input condition is null."
        ExtractTCNo = ERROR
        Exit Function
    End If
    'Remove space
    patterns = Split(" | ", "|")
    TCNoValue = ReplaceAll(TCNoValue, patterns(), "")
    'get position of ","
    commaPos = InStr(TCNoValue, ",")
    'Check commaPos
    If (commaPos > 0) Then
        subTCNoValue = Mid(TCNoValue, 1, commaPos - 1)
        TCNoValue = Mid(TCNoValue, commaPos + 1)
        'Check to ensure the format of TCNoValue is "TCx, TCy, ..."
        'Check space character in subTCNoValue
        If (InStr(subTCNoValue, " ") > 0) Then
            log = log & "Found space character (' ')" & vbNewLine & _
                        "Please check the TC '" & subTCNoValue & "' in TCNo '" & copiedTCNoValue & "'!" & vbNewLine & _
                    "Format of TCNo must be 'TCx, TCy, ...'"
            ExtractTCNo = ERROR
            Exit Function
        End If
         'Check 'T' character must be the first character in TCNoValue
        If (InStr(TCNoValue, "T") <> 1) Then
            log = log & "First character of TC must be 'T'" & vbNewLine & _
                    "Please check the TC after TC '" & subTCNoValue & "' in TCNo '" & copiedTCNoValue & "'!" & vbNewLine & _
                    "Format of TCNo must be 'TCx, TCy, ...'"
            ExtractTCNo = ERROR
            Exit Function
        End If
    Else
        subTCNoValue = TCNoValue
        TCNoValue = ""
        'Check to ensure the format of TCNoValue is "TCx, TCy, ..."
        'Check space character in subTCNoValue
        If (InStr(subTCNoValue, " ") < 0) Then
            log = log & "Please check the TCNo '" & copiedTCNoValue & "'!" & vbNewLine & _
                    "Format of TCNo must be 'TCx, TCy, ...'"
            ExtractTCNo = ERROR
            Exit Function
        End If
    End If
    ExtractTCNo = OK
End Function



