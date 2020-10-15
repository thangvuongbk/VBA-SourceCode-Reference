Attribute VB_Name = "MAddReq"
Option Explicit
Const OK = 1
Const WARNING = 0
Const ERROR = -1

Const BLACK = 1
Const WHITE = 2
Const RED = 3
Const GREEN = 4
Const LIGHT_BLUE = 5
Const YELLOW = 6
Const LIGHT_BLUE_2 = 20
Const BLUE = 33

Const OPERATOR_AND = "&&"
Const OPERATOR_OR = "||"
Const OPERATOR_SINGLE_AND = "&"
Const OPERATOR_SINGLE_OR = "|"
Const OPERATOR_EQU = "=="
Const OPERATOR_DIF = "!="
Const OPERATOR_GT = ">"
Const OPERATOR_LT = "<"
Const OPERATOR_GE = ">="
Const OPERATOR_LE = "<="
Const PRIORITY_LOWEST = 100
Const PRIORITY_HIGHEST = 0
'$$$$$$$$$$$$ DEFINITIONS FOR "MCDC" SHEET $$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
Const MCDC_SHEET_NAME = "MCDC"
Const MCDC_REQ_COL = 1
Const MCDC_TC_NO_COL = 2
Const MCDC_MCDC_COL = 3
Const MCDC_INPUT_COL = 4

Const MCDC_REQ_STR = "Requirement"
Const MCDC_TC_NO_STR = "TC No."
Const MCDC_MCDC_STR = "MCDC"
Const MCDC_INPUT_STR = "INPUTS"
Const MCDC_OUTPUT_STR = "OUTPUTS"
Const MCDC_OUTCOME_STR = "OUTCOME"
Const MCDC_CONDITION_1_STR = "Check MCDC for condition "
Const MCDC_CONDITION_2_STR = "Check condition "
Const MCDC_TCX_STR = "TCX"

Const MCDC_MAX_INPUTS = 50
Const MCDC_MAX_EXPRESSIONS = 50
'$$$$$$$$$$$$ DEFINITIONS FOR "Testcases" SHEET $$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
Const TC_TC_NO_COLUMN = 1
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
Type MCDC_Expression
    level As String
    levelInt As Integer
    elsePartLevelCharCode As Integer
    expression As String
    lineNumber As Integer
End Type
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
Type Signal_Input
    expression As String
    name As String
    value As String
End Type

Sub AddReq()
    '***********************************************************
    '
    'AUTHOR:            Le Thai (RBVH\EMB3)
    'DATE CREATED:      04/Jul/2013
    'FUNCTION NAME:     AddReq
    'DESCRIPTION:
    '
    '       OUTPUT: Add one more requirement template in MCDC sheet
    '
    '***********************************************************
    Dim temp As Variant
    Dim row, col As Integer
    Dim lastReqRow, INPUTSRow As Integer
    Dim startOutputRow As Integer
    Dim reqName As String
    Dim reqCount As Integer
    Dim expression As String
    Dim expressions(0 To MCDC_MAX_EXPRESSIONS - 1) As MCDC_Expression
    Dim nExpressions As Integer
    Dim inputs(0 To MCDC_MAX_INPUTS - 1) As Signal_Input
    Dim nInputs As Integer
    Dim operators(0 To MCDC_MAX_INPUTS - 2) As String
    Dim nOperators As Integer
    Dim priorities(0 To MCDC_MAX_INPUTS - 2) As Integer
    Dim nPriorities As Integer
    Dim rawMCDCResults(0 To 2 * MCDC_MAX_INPUTS - 1, 0 To MCDC_MAX_INPUTS) As Integer
    Dim nRawMCDCResults_1 As Integer
    Dim nRawMCDCResults_2 As Integer
    Dim finalMCDCResults(0 To 2 * MCDC_MAX_INPUTS - 1, 0 To MCDC_MAX_INPUTS) As Integer
    Dim nFinalMCDCResults_1 As Integer
    Dim nFinalMCDCResults_2 As Integer
    Dim highlights(0 To 2 * MCDC_MAX_INPUTS - 1, 0 To MCDC_MAX_INPUTS) As Integer
    Dim nHighlights_1 As Integer
    Dim nHighlights_2 As Integer
    Dim duplicatedRowStatus(0 To 2 * MCDC_MAX_INPUTS - 1) As Integer
    Dim nDuplicatedRowStatus As Integer
    Dim summaryMCDCExpression As String
    Dim summaryMCDCResults(0 To 2 * MCDC_MAX_INPUTS - 1) As String
    Dim nSummaryMCDCResults As Integer
    Dim log As String
    Dim i As Integer
    Dim defaultExpression As String
    Dim defaultPath As String
    Dim filepath As String
    Dim form As frmReq
    Dim lines() As String
    Dim patterns() As String
    'Back up
    Backup
    'Get input from user
    Set form = New frmReq
    'Get Vaiable
    If (GetVar("defaultPath", defaultPath) > 0) Then

    Else
        defaultPath = "D:\TEMP"
    End If
    'Set default value
    form.txtPath = defaultPath
    'Get expression
    'temp = Application.InputBox(Prompt:="Please enter the expression.", _
    '                                        Title:="EXPRESSION", _
    '                                        Default:=defaultExpression, _
    '                                        Left:=1, _
    '                                        Top:=1000, _
    '                                        Type:=2)
    'If (temp = False) Then
    '    Exit Sub
    'End If
    'expression = temp
    
    form.Show
    If (form.cancelButtonClicked = True) Then
        Exit Sub
    End If
    filepath = form.txtPath
    'SaveVar
    If (SaveVar("defaultPath", filepath) > 0) Then
        'MsgBox "Save 'defaultPath' with value: " & filePath
    'Save failed
    Else
        'MsgBox "Unable to save 'defaultPath' with value: " & filePath
    End If
    'Read all lines
    If (ReadFileToLinesAll(filepath, lines(), log) < 0) Then
        MsgBox "ERROR! AddReq! Unable to read lines in file '" & filepath & "'!"
        Exit Sub
    End If
    'Get expression
    'expression = ""
    'For i = 0 To UBound(lines)
    '    expression = expression & lines(i)
    'Next i
    'Remove " ", tab
    'Note: chr(9) = tab charater
    'patterns = Split(" ," & Chr(9), ",")
    'expression = ReplaceAll(expression, patterns, "")
    If (GetMCDCExpressions(lines(), expressions(), nExpressions, log) < 0) Then
        MsgBox "ERROR! AddReq! Unable to get expressions from lines in file '" & filepath & "'!"
        Exit Sub
    End If
    'Loop for all expressions
    'For i = 0 To UBound(expressions)
    For i = 0 To nExpressions - 1
        'Expression
        expression = expressions(i).expression
        'MsgBox expression
        'Check expression
        If (CheckExpression(expression, log) = False) Then
            'MsgBox "ERROR! Invalid expression!"
            MsgBox log
            Call WriteReq(vbNullString, _
                vbNullString, _
                0, _
                inputs(), nInputs, _
                finalMCDCResults(), nFinalMCDCResults_1, nFinalMCDCResults_2, _
                highlights(), nHighlights_1, nHighlights_2, _
                summaryMCDCExpression, _
                summaryMCDCResults(), _
                nSummaryMCDCResults)
            Exit Sub
        End If
        'Get inputs
        If (GetInputs(expression, inputs(), nInputs, operators(), nOperators, priorities(), nPriorities, log) < 0) Then
            MsgBox log
            Exit Sub
        End If
        'Do MCDC
        If (DoMCDC(inputs(), nInputs, operators(), nOperators, priorities(), nPriorities, rawMCDCResults(), nRawMCDCResults_1, nRawMCDCResults_2, log) < 0) Then
            MsgBox log
            Exit Sub
        End If
        'DEBUG
        'CSVDisplay2DInt rawMCDCResults(), 2
        'GetDuplicatedRowStatus
        nDuplicatedRowStatus = nRawMCDCResults_1
        If (GetDuplicatedRowStatus(rawMCDCResults(), nRawMCDCResults_1, nRawMCDCResults_2, duplicatedRowStatus(), nDuplicatedRowStatus, log) < 0) Then
            MsgBox log
            Exit Sub
        End If
        'GetFinalMCDCResults
        If (GetFinalMCDCResults(rawMCDCResults(), _
                                nRawMCDCResults_1, _
                                nRawMCDCResults_2, _
                                duplicatedRowStatus(), nDuplicatedRowStatus, _
                                finalMCDCResults(), nFinalMCDCResults_1, nFinalMCDCResults_2, _
                                highlights(), nHighlights_1, nHighlights_2, _
                                log) < 0) Then
            MsgBox log
            Exit Sub
        End If
        'GetSummaryMCDCResults
        nSummaryMCDCResults = nFinalMCDCResults_1
        If (GetSummaryMCDCResults(inputs(), nInputs, _
                                    operators(), nOperators, _
                                    priorities(), nPriorities, _
                                    finalMCDCResults(), nFinalMCDCResults_1, nFinalMCDCResults_2, _
                                    summaryMCDCExpression, _
                                    summaryMCDCResults(), nSummaryMCDCResults, _
                                    log) < 0) Then
            MsgBox log
            Exit Sub
        End If
            
        'FOR DEBUG
        'CSVDisplay2DInt rawMCDCResults(), 2
        'CSVDisplay1DInt duplicatedRowStatus(), 2 + UBound(rawMCDCResults, 1) + 3
        
        'CSVDisplay2DInt finalMCDCResults(), 30
        'CSVDisplay2DInt highlights(), 40
        'Write expression
        Call WriteReq(expression, _
                        expressions(i).level, _
                        expressions(i).lineNumber, _
                        inputs(), nInputs, _
                        finalMCDCResults(), nFinalMCDCResults_1, nFinalMCDCResults_2, _
                        highlights(), nHighlights_1, nHighlights_2, _
                        summaryMCDCExpression, _
                        summaryMCDCResults(), nSummaryMCDCResults)
    Next i
    'Auto fit
    AutofitColumns (MCDC_SHEET_NAME)
End Sub
Function GetMCDCExpressions(ByRef lines() As String, _
                    ByRef expressions() As MCDC_Expression, _
                    ByRef nExpressions As Integer, _
                    ByRef log As String) As Integer
    '***********************************************************
    '
    'AUTHOR:            Le Thai (RBVH\EMB3)
    'DATE CREATED:      12/Feb/2014
    'FUNCTION NAME:     GetMCDCExpressions
    'DESCRIPTION:
    '       INPUT:
    '               lines()             :   array of lines read from file
    '
    '       OUTPUT:
    '               expressions()       :   array contains all MCDC expressions
    '               nExpressions        :   array size
    '
    '***********************************************************
    Dim i As Integer
    Dim levelInt As Integer
    Dim elsePartLevelCharCodes(0 To 20) As Integer
    Dim expression As String
    Dim enableFlag As Boolean
    Dim isCurrentLineGot As Boolean
    Dim isElsePart As Boolean
    Dim patterns() As String
    Dim tmpStr As String
    Dim commentFlag As Boolean
    'Loop for all lines
    levelInt = 0
    'Note: &H61 --> 0x61 --> Character 'a' in ASCII table
    For i = 0 To 20
        elsePartLevelCharCodes(i) = &H61
    Next i
    expression = ""
    enableFlag = False
    isCurrentLineGot = False
    isElsePart = False
    commentFlag = False
    For i = 0 To UBound(lines)
        'Remove " ", tab
        'Note: chr(9) = tab charater
        patterns = Split(" ," & Chr(9), ",")
        lines(i) = ReplaceAll(lines(i), patterns, "")
        tmpStr = ""
        'Check "//"
        If (InStr(lines(i), "//") > 0) Then
            If (InStr(lines(i), "//") = 1) Then
                lines(i) = ""
            Else
                lines(i) = Mid(lines(i), 1, InStr(lines(i), "//") - 1)
            End If
        'Check "/*"
        ElseIf (InStr(lines(i), "/*") > 0) Then
            If (commentFlag = False) Then
                commentFlag = True
                If (InStr(lines(i), "/*") = 1) Then
                    tmpStr = ""
                    If (InStr(lines(i), "*/") > 0) Then
                        commentFlag = False
                        lines(i) = tmpStr & Mid(lines(i), InStr(lines(i), "*/") + Len("/*"))
                    End If
                Else
                    tmpStr = Mid(lines(i), 1, InStr(lines(i), "/*") - 1)
                    If (InStr(lines(i), "*/") > 0) Then
                        commentFlag = False
                        lines(i) = tmpStr & Mid(lines(i), InStr(lines(i), "*/") + Len("/*"))
                    End If
                End If
            Else
                If (InStr(lines(i), "*/") > 0) Then
                    commentFlag = False
                    lines(i) = tmpStr & Mid(lines(i), InStr(lines(i), "*/") + Len("/*"))
                End If
            End If
        'Check "*/"
        ElseIf (InStr(lines(i), "*/") > 0) Then
            commentFlag = False
            lines(i) = Mid(lines(i), InStr(lines(i), "*/") + Len("/*"))
        End If
        'Check commentFlag
        If (commentFlag = True) Then
            lines(i) = tmpStr
        End If
        'Check "if"
        If (InStr(lines(i), "if(") = 1 Or InStr(lines(i), "elseif(") = 1) Then
            expressions(nExpressions).lineNumber = i + 1
            enableFlag = True
            isCurrentLineGot = True
            If (InStr(lines(i), "else") > 0) Then
                isElsePart = True
            Else
                isElsePart = False
                'Reset elsePartLevelCharCode --> 0x61 ("a")
                elsePartLevelCharCodes(levelInt + 1) = &H61
            End If
            'Get line
            expression = Mid(lines(i), InStr(lines(i), "if") + Len("if"))
            'Check "{"
            If (InStr(expression, "{") > 0) Then
                levelInt = levelInt + 1
                If (isElsePart = True) Then
                    elsePartLevelCharCodes(levelInt) = elsePartLevelCharCodes(levelInt) + 1
                Else
                    
                End If
                'Get line
                expression = Mid(expression, 1, InStr(expression, "{") - 1)
                'Remove " ", tab
                'Note: chr(9) = tab charater
                patterns = Split(" ," & Chr(9), ",")
                expression = ReplaceAll(expression, patterns, "")
                'Get expression element
                expressions(nExpressions).expression = expression
                expressions(nExpressions).levelInt = levelInt
                expressions(nExpressions).elsePartLevelCharCode = elsePartLevelCharCodes(levelInt)
                expressions(nExpressions).level = levelInt & Chr(elsePartLevelCharCodes(levelInt))
                nExpressions = nExpressions + 1
                enableFlag = False
            End If
        'Check "{"
        ElseIf (InStr(lines(i), "{") > 0) Then
            levelInt = levelInt + 1
            If (enableFlag = True) Then
                If (isElsePart = True) Then
                    elsePartLevelCharCodes(levelInt) = elsePartLevelCharCodes(levelInt) + 1
                Else
                    
                End If
                'Get line
                expression = expression & Mid(expression, 1, InStr(expression, "{"))
                'Remove " ", tab
                'Note: chr(9) = tab charater
                patterns = Split(" ," & Chr(9), ",")
                expression = ReplaceAll(expression, patterns, "")
                'Get expression element
                expressions(nExpressions).expression = expression
                expressions(nExpressions).levelInt = levelInt
                expressions(nExpressions).elsePartLevelCharCode = elsePartLevelCharCodes(levelInt)
                expressions(nExpressions).level = levelInt & Chr(elsePartLevelCharCodes(levelInt))
                nExpressions = nExpressions + 1
                enableFlag = False
            End If
        'Check "}"
        ElseIf (InStr(lines(i), "}") > 0) Then
            levelInt = levelInt - 1
        End If
        'Check enable flag
        If (enableFlag = True) Then
            If (isCurrentLineGot = False) Then
                expression = expression & lines(i)
            Else
                isCurrentLineGot = False
            End If
        End If
    Next i
    GetMCDCExpressions = OK
End Function
Function GetInputs(ByVal condition As String, _
                    ByRef inputs() As Signal_Input, _
                    ByRef nInputs As Integer, _
                    ByRef operators() As String, _
                    ByRef nOperators As Integer, _
                    ByRef priorities() As Integer, _
                    ByRef nPriorities As Integer, _
                    ByRef log As String) As Integer
    '***********************************************************
    '
    'AUTHOR:            Le Thai (RBVH\EMB3)
    'DATE CREATED:      12/Feb/2014
    'FUNCTION NAME:     GetInputs
    'DESCRIPTION:
    '       INPUT:
    '               condition   :   condition
    '
    '       OUTPUT:
    '               inputs()        :   array contains all the inputs names and value
    '               nInputs         :   array size
    '               operators()     :   array contains all operators in condition
    '               nOperators      :   array size
    '               priorities()    :   array contains all priority of operators
    '               nPriorities     :   array size
    '
    '***********************************************************
    Dim trimmedCondition As String
    Dim l_condition As String
    Dim patterns() As String
    Dim repStr As String
    Dim conditions() As String
    Dim i As Integer
    Dim pos As Integer
    Dim buff() As String
    Dim priority As Integer
    
    trimmedCondition = condition
    'Remove " ", "(", ")"
    patterns = Split(" |(|)", "|")
    repStr = ""
    trimmedCondition = ReplaceAll(trimmedCondition, patterns(), repStr)
    'Split to get single conditions
    patterns = Split(OPERATOR_AND & "," & OPERATOR_OR, ",")
    conditions = SplitAll(trimmedCondition, patterns())
    'Get inputs
    'ReDim Preserve inputs(0 To UBound(conditions))
    nInputs = UBound(conditions) + 1
    For i = 0 To UBound(conditions)
        inputs(i).expression = conditions(i)
        '"=="
        If (InStr(conditions(i), OPERATOR_EQU) > 0) Then
            patterns = Split(conditions(i), OPERATOR_EQU)
            inputs(i).name = patterns(0)
            inputs(i).value = patterns(1)
        '"!="
        ElseIf (InStr(conditions(i), OPERATOR_DIF) > 0) Then
            patterns = Split(conditions(i), OPERATOR_DIF)
            inputs(i).name = patterns(0)
            inputs(i).value = OPERATOR_DIF & patterns(1)
        '">="
        ElseIf (InStr(conditions(i), OPERATOR_GE) > 0) Then
            patterns = Split(conditions(i), OPERATOR_GE)
            inputs(i).name = patterns(0)
            inputs(i).value = OPERATOR_GE & patterns(1)
        '"<="
        ElseIf (InStr(conditions(i), OPERATOR_LE) > 0) Then
            patterns = Split(conditions(i), OPERATOR_LE)
            inputs(i).name = patterns(0)
            inputs(i).value = OPERATOR_LE & patterns(1)
        '">"
        ElseIf (InStr(conditions(i), OPERATOR_GT) > 0) Then
            patterns = Split(conditions(i), OPERATOR_GT)
            inputs(i).name = patterns(0)
            inputs(i).value = OPERATOR_GT & patterns(1)
        '"<"
        ElseIf (InStr(conditions(i), OPERATOR_LT) > 0) Then
            patterns = Split(conditions(i), OPERATOR_LT)
            inputs(i).name = patterns(0)
            inputs(i).value = OPERATOR_LT & patterns(1)
        'LOGICAL
        Else
            If (InStr(conditions(i), "!") > 0) Then
                inputs(i).name = Replace(conditions(i), "!", "")
                inputs(i).value = "FALSE"
            Else
                inputs(i).name = conditions(i)
                inputs(i).value = "TRUE"
            End If
        End If
    Next i
    'Check there is only 1 input, no need to get operators and priorities
    'If (UBound(inputs) <= 0) Then
    If (nInputs <= 1) Then
        log = "WARNING!GetInputs! There is only 1 input!"
        GetInputs = WARNING
        Exit Function
    End If
    'Get operators
    'ReDim Preserve operators(0 To UBound(conditions) - 1)
    nOperators = nInputs - 1
    For i = 0 To UBound(conditions) - 1
        pos = InStr(trimmedCondition, conditions(i))
        If (pos > 0) Then
            operators(i) = Mid(trimmedCondition, pos + Len(conditions(i)), 2)
        Else
            log = "ERROR! Unable to get operator!"
            GetInputs = ERROR
            Exit Function
        End If
    Next i
    'Get priorities
    l_condition = condition
    'Remove " "
    patterns = Split(" | ", "|")
    repStr = ""
    l_condition = ReplaceAll(l_condition, patterns(), repStr)
    'Replace conditions(i) by i in l_condition
    For i = 0 To UBound(conditions)
        l_condition = Replace(l_condition, conditions(i), i)
    Next i
    'l_condition: replace OPERATOR_AND by OPERATOR_SINGLE_AND
    patterns = Split(OPERATOR_AND & "," & OPERATOR_AND, ",")
    repStr = OPERATOR_SINGLE_AND
    l_condition = ReplaceAll(l_condition, patterns(), repStr)
    'l_condition: replace OPERATOR_OR by OPERATOR_SINGLE_OR
    patterns = Split(OPERATOR_OR & "," & OPERATOR_OR, ",")
    repStr = OPERATOR_SINGLE_OR
    l_condition = ReplaceAll(l_condition, patterns(), repStr)
    'Split l_condition to array of characters
    buff = Split(StrConv(l_condition, vbUnicode), Chr$(0))
    'Init array
    priority = 0
    nPriorities = 0
    For i = 0 To UBound(buff)
        If (buff(i) = "(") Then
            priority = priority + 1
        ElseIf (buff(i) = ")") Then
            priority = priority - 1
        End If
        If (buff(i) = OPERATOR_SINGLE_AND Or buff(i) = OPERATOR_SINGLE_OR) Then
            'ReDim Preserve priorities(0 To nPri)
            priorities(nPriorities) = priority
            nPriorities = nPriorities + 1
        End If
    Next i
    GetInputs = OK
End Function
Function DoMCDC(ByRef inputs() As Signal_Input, _
                    ByVal nInputs As Integer, _
                    ByRef operators() As String, _
                    ByVal nOperators As Integer, _
                    ByRef priorities() As Integer, _
                    ByVal nPriorities As Integer, _
                    ByRef MCDCResults() As Integer, _
                    ByRef nMCDCResults_1 As Integer, _
                    ByRef nMCDCResults_2 As Integer, _
                    ByRef log As String) As Integer
    '***********************************************************
    '
    'AUTHOR:            Le Thai (RBVH\EMB3)
    'DATE CREATED:      12/Feb/2014
    'FUNCTION NAME:     DoMCDC
    'DESCRIPTION:
    '       INPUT:
    '               inputs()        :   array contains all the inputs names and value
    '               nInputs         :   array size
    '               operators()     :   array contains all operators in condition
    '               nOperators      :   array size
    '               priorities()    :   array contains all priority of operators
    '               nPriorities     :   array size
    '
    '       OUTPUT:
    '               MCDCResults()   : 2D array result with end column is outcome, the sequence for MCDC row is from outcome "TRUE" to outcome "FALSE"
    '               nMCDCResults_1     :   array size, dimension 1
    '               nMCDCResults_2     :   array size, dimension 2
    '
    '***********************************************************
    Dim tmpMCDC(0 To MCDC_MAX_INPUTS) As Integer
    Dim nTmpMCDC As Integer
    Dim MCDCRow(0 To MCDC_MAX_INPUTS) As Integer
    Dim nMCDCRow As Integer
    Dim i, j, k As Integer
    Dim highestPri As Integer
    Dim lowestPri As Integer
    Dim nCurPriConditions As Integer
    Dim previousPri As Integer
    Dim curOperator As String
    Dim startPriIdx As Integer
    Dim endPriIdx As Integer
    Dim tmpRet As Integer
    Dim outcome As Integer
    Dim curOutcome As Integer
    Dim curFocusIdx As Integer
    
    nTmpMCDC = nInputs
    nMCDCRow = nInputs
    'MCDCResults
    'ReDim Preserve MCDCResults(0 To 2 * UBound(inputs) + 1, 0 To UBound(inputs) + 1)
    nMCDCResults_1 = 2 * nInputs
    nMCDCResults_2 = nInputs + 1
    'Check case only 1 input
    'If (UBound(inputs) = 0) Then
    If (nInputs <= 1) Then
        'GetFullMCDC
        If (GetFullMCDC(OPERATOR_AND, 1, MCDCResults(), nMCDCResults_1, nMCDCResults_2, log) < 0) Then
            MsgBox log
            DoMCDC = ERROR
            Exit Function
        End If
        DoMCDC = OK
        Exit Function
    End If
    'Find the highest priority (smallest number) and lowest priority (biggest number)
    highestPri = PRIORITY_LOWEST
    lowestPri = PRIORITY_HIGHEST
    'For i = 0 To UBound(priorities())
    For i = 0 To nPriorities - 1
        If (priorities(i) < highestPri) Then
            highestPri = priorities(i)
        End If
        If (priorities(i) > lowestPri) Then
            lowestPri = priorities(i)
        End If
    Next i
    'OUTCOME: TRUE
    outcome = 1
    'Loop for all inputs
    'For i = 0 To UBound(inputs())
    For i = 0 To nInputs - 1
        'ReDim Preserve MCDCRow(0 To UBound(inputs))
        nMCDCRow = nInputs
        'Init MCDCRow for highest priority
        'Priority = (int)(x/10)
        'Value = x%10
        'For j = 0 To UBound(MCDCRow)
        For j = 0 To nMCDCRow - 1
            MCDCRow(j) = highestPri * 10 + outcome
        Next j
        'Loop for all priorities
        For j = highestPri To lowestPri
            'Get MCDCInfo at priority j
            startPriIdx = 0
            endPriIdx = 0
            tmpRet = GetInfoAtPriority(priorities(), nPriorities, operators(), nOperators, j, startPriIdx, endPriIdx, i, nCurPriConditions, curOperator, curFocusIdx, log)
            Do While (tmpRet > 0)
                'curOutcome
                curOutcome = MCDCRow(startPriIdx) Mod 10
                'GetSingleRowMCDC
                If (GetSingleRowMCDC(curOperator, nCurPriConditions, curOutcome, curFocusIdx, tmpMCDC(), nTmpMCDC, log) < 0) Then
                    MsgBox log
                End If
                'ExpandMCDCRowResult
                If (ExpandMCDCRowResult(priorities(), nPriorities, tmpMCDC(), nTmpMCDC, j, startPriIdx, endPriIdx, MCDCRow(), nMCDCRow, log) < 0) Then
                    MsgBox log
                End If
                'Increasement
                startPriIdx = endPriIdx + 1
                tmpRet = GetInfoAtPriority(priorities(), nPriorities, operators(), nOperators, j, startPriIdx, endPriIdx, i, nCurPriConditions, curOperator, curFocusIdx, log)
            Loop
            'Check error
            If (tmpRet < 0) Then
                DoMCDC = ERROR
                Exit Function
            End If
        Next j
        'Get MCDC Row Result
        'For j = 0 To UBound(inputs)
        For j = 0 To nInputs - 1
            MCDCResults(i, j) = MCDCRow(j) Mod 10
        Next j
        'MCDCResults(i, UBound(inputs) + 1) = outcome
        MCDCResults(i, nInputs) = outcome
    Next i
    'OUTCOME: FALSE
    outcome = 0
    'Loop for all inputs
    'For i = 0 To UBound(inputs())
    For i = 0 To nInputs - 1
        'ReDim Preserve MCDCRow(0 To UBound(inputs))
        nMCDCRow = nInputs
        'Init MCDCRow for highest priority
        'Priority = (int)(x/10)
        'Value = x%10
        'For j = 0 To UBound(MCDCRow)
        For j = 0 To nMCDCRow - 1
            MCDCRow(j) = highestPri * 10 + outcome
        Next j
        'Loop for all priorities
        For j = highestPri To lowestPri
            'Get MCDCInfo at priority j
            startPriIdx = 0
            endPriIdx = 0
            tmpRet = GetInfoAtPriority(priorities(), nPriorities, operators(), nOperators, j, startPriIdx, endPriIdx, i, nCurPriConditions, curOperator, curFocusIdx, log)
            Do While (tmpRet > 0)
                'curOutcome
                curOutcome = MCDCRow(startPriIdx) Mod 10
                'GetSingleRowMCDC
                If (GetSingleRowMCDC(curOperator, nCurPriConditions, curOutcome, curFocusIdx, tmpMCDC(), nTmpMCDC, log) < 0) Then
                    MsgBox log
                End If
                'ExpandMCDCRowResult
                If (ExpandMCDCRowResult(priorities(), nPriorities, tmpMCDC(), nTmpMCDC, j, startPriIdx, endPriIdx, MCDCRow(), nMCDCRow, log) < 0) Then
                    MsgBox log
                End If
                'Increasement
                startPriIdx = endPriIdx + 1
                tmpRet = GetInfoAtPriority(priorities(), nPriorities, operators(), nOperators, j, startPriIdx, endPriIdx, i, nCurPriConditions, curOperator, curFocusIdx, log)
            Loop
            'Check error
            If (tmpRet < 0) Then
                DoMCDC = ERROR
                Exit Function
            End If
        Next j
        'Get MCDC Row Result
        'For j = 0 To UBound(inputs)
        For j = 0 To nInputs - 1
            'MCDCResults(UBound(inputs) + i + 1, j) = MCDCRow(j) Mod 10
            MCDCResults(nInputs + i, j) = MCDCRow(j) Mod 10
        Next j
        'MCDCResults(UBound(inputs) + i + 1, UBound(inputs) + 1) = outcome
        MCDCResults(nInputs + i, nInputs) = outcome
    Next i
    DoMCDC = OK
End Function
Function GetInfoAtPriority(ByRef priorities() As Integer, _
                    ByVal nPriorities As Integer, _
                    ByRef operators() As String, _
                    ByVal nOperators As String, _
                    ByVal priority As Integer, _
                    ByRef startPriIdx As Integer, _
                    ByRef endPriIdx As Integer, _
                    ByVal inFocusIdx As Integer, _
                    ByRef count As Integer, _
                    ByRef operator As String, _
                    ByRef outFocusIdx As Integer, _
                    ByRef log As String) As Integer
    '***********************************************************
    '
    'AUTHOR:            Le Thai (RBVH\EMB3)
    'DATE CREATED:      17/Feb/2014
    'FUNCTION NAME:     GetInfoAtPriority
    'DESCRIPTION:       Find the [operator] and the number of input for MCDC at priority [priority]
    '       INPUT:
    '               priorities()    :   array of priorities
    '               nPriorities     :   array size
    '               operators()     :   array of priorities
    '               nOperators      :   array size
    '               priority        :   priority to be calculated
    '               startPriIdx   :   starting index of priorities() to be calculated (i/o)
    '               inFocusIdx      :   original index of focused input
    '
    '       OUTPUT:
    '               startPriIdx   :   starting index of priorities() to be calculated (i/o)
    '               endPriIdx     :   end index of priorities() to be calculated (i/o)
    '               count           :   the number of conditions
    '               operator        :   operator
    '               outFocusIdx     :   index of focused input after calculation
    '***********************************************************
    Dim i As Integer
    'Dim previousPri As Integer
    Dim copiedStartPriIdx As Integer
    Dim isStartPriIdxFound As Boolean
    Dim isEndPriIdxFound As Boolean
    Dim isOperatorFound As Boolean
    Dim isFocusIdxFound  As Boolean
    
    'Check startPriIdx out of range
    'If (startPriIdx < LBound(priorities) Or startPriIdx > UBound(priorities)) Then
    If ((startPriIdx < 0) Or (startPriIdx > (nPriorities - 1))) Then
        log = "WARNING!GetInfoAtPriority! Start input index out of range!"
        GetInfoAtPriority = WARNING
        Exit Function
    End If
    'Check focus index
    'Calculation
    copiedStartPriIdx = startPriIdx
    endPriIdx = copiedStartPriIdx
    isStartPriIdxFound = False
    isEndPriIdxFound = False
    'Find start input index and end input index
    'Loop for all input priorities
    'For i = copiedStartPriIdx To UBound(priorities())
    For i = copiedStartPriIdx To nPriorities - 1
        'Find while not found end index
        If (isEndPriIdxFound = False) Then
            'If current priority is lower or equal to input priority --> Keep finding
            If ((priorities(i) >= priority)) Then
                If (isStartPriIdxFound = False) Then
                    startPriIdx = i
                    isStartPriIdxFound = True
                End If
            'If current priority is higher than input priority --> check to get end index
            Else
                'If start index found --> Get end index
                If (isStartPriIdxFound = True) Then
                    endPriIdx = i - 1
                    isEndPriIdxFound = True
                'Else keep finding
                Else
                    'Do nothing
                End If
            End If
        End If
    Next i
    'Check if start input index is found
    If (isStartPriIdxFound = False) Then
        log = "WARNING!GetInfoAtPriority! Not found start input index from index '" & copiedStartPriIdx & "' for priority 'priority" & "'!"
        GetInfoAtPriority = WARNING
        Exit Function
    End If
    'Check case end input index is the last index
    If (isEndPriIdxFound = False) Then
        'endPriIdx = UBound(priorities())
        endPriIdx = nPriorities - 1
        isEndPriIdxFound = True
    End If
    'Find outFocusIdx, count and operator
    'Note:  outFocusIdx is focus index of input
    '       priorities is priotities of operators
    count = 0
    'previousPri = -1
    isOperatorFound = False
    isFocusIdxFound = False
    outFocusIdx = -1
    For i = startPriIdx To endPriIdx
        'Get outFocusIdx
        'Note: when (i = inFocusIdx), get outFocusIdx before count is calculated
        If (i = inFocusIdx) Then
            'If (count = 0) Then
            '    outFocusIdx = 0
            'Else
            '    outFocusIdx = count - 1
            'End If
            outFocusIdx = count
            isFocusIdxFound = True
        End If
        'Get count and operator
        'Check if priority matched
        If (priorities(i) = priority) Then
            'Get operator
            If (isOperatorFound = False) Then
                operator = operators(i)
                isOperatorFound = True
            End If
            'Get count
            count = count + 1
            'If (previousPri = priority) Then
            '    count = count + 1
            'Else
            '    count = count + 2
            'End If
        End If
        'previousPri = priorities(i)
    Next i
    'Final count
    count = count + 1
    'Re-Check inFocusIdx
    If (isFocusIdxFound = False) Then
        'Last index
        If (inFocusIdx = endPriIdx + 1) Then
            If (count > 0) Then
                outFocusIdx = count - 1
            Else
                log = "ERROR! GetInfoAtPriority! Unable to identify focus input index duo to count <= 0!"
                GetInfoAtPriority = ERROR
                Exit Function
            End If
        Else
            outFocusIdx = 0
        End If
        isFocusIdxFound = True
    End If
    GetInfoAtPriority = OK
End Function
Function GetSingleRowMCDC(ByVal operator As String, _
                    ByVal count As Integer, _
                    ByVal outcome As Integer, _
                    ByVal focusIdx As Integer, _
                    ByRef resultRow() As Integer, _
                    ByRef nResultRow As Integer, _
                    ByRef log As String) As Integer
    '***********************************************************
    '
    'AUTHOR:            Le Thai (RBVH\EMB3)
    'DATE CREATED:      17/Feb/2014
    'FUNCTION NAME:     GetSingleRowMCDC
    'DESCRIPTION:       Calculate MCDC result line for [count] conditions using operator [operator]
    '                   with outcome [outcome] and focus on input index [focusIdx] equal to [outcome]
    '       INPUT:
    '               operator    :   operator
    '               count       :   the number of conditions
    '               outcome     :   outcome
    '               focusIdx    :   index of input need to focus on, start from 0
    '
    '       OUTPUT:
    '               resultRow()   : 1D array result row
    '               nResultRow()   : 1D array size
    '               E.g: count = 3, operator = "&&", focusIdx = 1, outcome = 0
    '                   ==> resultRow() = {1    0   1}
    '               E.g: count = 3, operator = "&&", focusIdx = 1, outcome = 1
    '                   ==> resultRow() = {1    1   1}
    '
    '                      (A &&B &&C)
    '                       1   1   1
    '                       0   1   1
    '                       1   0   1
    '                       1   1   0
    '***********************************************************
    Dim i, j As Integer
    Dim rowIdx As Integer
    Dim tmpMCDCResults(0 To 2 * MCDC_MAX_INPUTS - 1, 0 To MCDC_MAX_INPUTS) As Integer
    Dim nTmpMCDCResults_1 As Integer
    Dim nTmpMCDCResults_2 As Integer
    
    nTmpMCDCResults_1 = count + 1
    nTmpMCDCResults_2 = count + 1
    'Get MCDC
    If (GetFullMCDC(operator, count, tmpMCDCResults(), nTmpMCDCResults_1, nTmpMCDCResults_2, log) < 0) Then
        MsgBox log
    End If
    'Find row index for MCDC result row, first matched row
    rowIdx = -1
    For i = 0 To count
        If (rowIdx < 0 And tmpMCDCResults(i, focusIdx) = outcome And tmpMCDCResults(i, count) = outcome) Then
            rowIdx = i
        End If
    Next i
    If (rowIdx >= 0) Then
        'ReDim Preserve resultRow(0 To count - 1)
        nResultRow = count
        For i = 0 To count - 1
            resultRow(i) = tmpMCDCResults(rowIdx, i)
        Next i
    Else
        log = "ERROR!GetSingleRowMCDC! Unable to find row index satisfied with input condition!" & vbNewLine & _
            "operator: " & operator & vbNewLine & _
            "count:" & count & vbNewLine & _
            "outcome: " & outcome & vbNewLine & _
            "focusIdx:" & focusIdx
    End If
    GetSingleRowMCDC = OK
End Function
Function GetFullMCDC(ByVal operator As String, _
                    ByVal count As Integer, _
                    ByRef MCDCResults() As Integer, _
                    ByVal nMCDCResults_1 As Integer, _
                    ByVal nMCDCResults_2 As Integer, _
                    ByRef log As String) As Integer
    '***********************************************************
    '
    'AUTHOR:            Le Thai (RBVH\EMB3)
    'DATE CREATED:      17/Feb/2014
    'FUNCTION NAME:     GetMCDC
    'DESCRIPTION:       Calculate MCDC result for [count] conditions using operator [operator]
    '       INPUT:
    '               count           :   the number of conditions
    '               operator        :   operator
    '               nMCDCResults_1  :   array size
    '               nMCDCResults_2  :   array size
    '
    '       OUTPUT:
    '               MCDCResults()   : 2D array result with end column is outcome, the sequence for MCDC row is from outcome "TRUE" to outcome "FALSE"
    '               E.g: count = 3, operator = "&&"
    '                   ==> MCDCResults()
    '                       1   1   1   1
    '                       0   1   1   0
    '                       1   0   1   0
    '                       1   1   0   0
    '               E.g: count = 3, operator = "||"
    '                   ==> MCDCResults()
    '                       1   0   0   1
    '                       0   1   0   1
    '                       0   0   1   1
    '                       0   0   0   0
    '
    '***********************************************************
    Dim i, j As Integer
    'ReDim Preserve MCDCResults(0 To count, 0 To count)
    Select Case (operator)
        Case Is = OPERATOR_AND
            For i = 0 To count
                For j = 0 To count
                    If (j <> count) Then
                        'MCDC
                        If (j = i - 1) Then
                            MCDCResults(i, j) = 0
                        Else
                            MCDCResults(i, j) = 1
                        End If
                    Else
                        'OUTCOME
                        If (i = 0) Then
                            MCDCResults(i, j) = 1
                        Else
                            MCDCResults(i, j) = 0
                        End If
                    End If
                    
                Next j
            Next i
        Case Is = OPERATOR_OR
            For i = 0 To count
                For j = 0 To count
                    If (j <> count) Then
                        'MCDC
                        If (j = i) Then
                            MCDCResults(i, j) = 1
                        Else
                            MCDCResults(i, j) = 0
                        End If
                    Else
                        'OUTCOME
                        If (i = count) Then
                            MCDCResults(i, j) = 0
                        Else
                            MCDCResults(i, j) = 1
                        End If
                    End If
                Next j
            Next i
        Case Else
            log = "ERROR!GetMCDC! Invalid operator '" & operator & "'!"
            GetFullMCDC = ERROR
            Exit Function
    End Select
    GetFullMCDC = OK
End Function
Function ExpandMCDCRowResult(ByRef priorities() As Integer, _
                    ByVal nPriorities As Integer, _
                    ByRef inMCDCRow() As Integer, _
                    ByVal nInMCDCRow As Integer, _
                    ByVal priority As Integer, _
                    ByRef startPriIdx As Integer, _
                    ByRef endPriIdx As Integer, _
                    ByRef outMCDCRow() As Integer, _
                    ByVal nOutMCDCRow As Integer, _
                    ByRef log As String) As Integer
    '***********************************************************
    '
    'AUTHOR:            Le Thai (RBVH\EMB3)
    'DATE CREATED:      17/Feb/2014
    'FUNCTION NAME:     ExpandMCDCRowResult
    'DESCRIPTION:       Process array of MCDC row after expanding inMCDCRow() (i/o)
    '       INPUT:
    '               priorities()    :   array of priorities
    '               nPriorities    :   array size
    '               inMCDCRow()     :   array of single row MCDC
    '               nInMCDCRow    :   array size
    '               priority        :   priority to be calculated
    '               startPriIdx   :   starting index of priorities() to be calculated
    '               endPriIdx     :   end index of priorities() to be calculated
    '               outMCDCRow      :   array of MCDC row before expanding inMCDCRow() (i/o)
    '               nOutMCDCRow    :   array size
    '
    '       OUTPUT:
    '               outMCDCRow      :   array of MCDC row after expanding inMCDCRow() (i/o)
    '***********************************************************
    Dim i As Integer
    Dim inMCDCRowIdx As Integer
    
    inMCDCRowIdx = 0
    For i = startPriIdx To endPriIdx
        outMCDCRow(i) = (priority + 1) * 10 + inMCDCRow(inMCDCRowIdx)
        'Match priority
        If (priorities(i) = priority) Then
            inMCDCRowIdx = inMCDCRowIdx + 1
        End If
    Next i
    'Last input
    outMCDCRow(endPriIdx + 1) = (priority + 1) * 10 + inMCDCRow(inMCDCRowIdx)
    
    ExpandMCDCRowResult = OK
End Function
Function CheckExpression(ByVal expression As String, _
                            ByRef log As String) As Boolean
    '***********************************************************
    '
    'AUTHOR:            Le Thai (RBVH\EMB3)
    'DATE CREATED:      14/Feb/2014
    'FUNCTION NAME:     IsLetter
    'DESCRIPTION:       Check whether input expression is correctly inputted
    '       INPUT:
    '               expression   :   expression
    '
    '       OUTPUT:
    '
    '***********************************************************
    Dim tmp1, tmp2 As String
    Dim patterns() As String
    'Check null
    If (expression = vbNullString) Then
        CheckExpression = False
        log = "WARNING!CheckExpression! Expression is empty!"
        Exit Function
    End If
    'Check the number of "(" = the number of ")"
    patterns = Split("(,(", ",")
    tmp1 = ReplaceAll(expression, patterns(), "")
    patterns = Split("),)", ",")
    tmp2 = ReplaceAll(expression, patterns(), "")
    If (Len(tmp1) <> Len(tmp2)) Then
        CheckExpression = False
        log = "WARNING! The number of character '(' is not equal to the number of character ')'!"
        Exit Function
    End If
    CheckExpression = True
End Function
Function IsLetter(strValue As String) As Boolean
    '***********************************************************
    '
    'AUTHOR:            Le Thai (RBVH\EMB3)
    'DATE CREATED:      14/Feb/2014
    'FUNCTION NAME:     IsLetter
    'DESCRIPTION:       Check whether input string contains characters only
    '       ASCII table:
    '               65 – 90 (uppercase A-Z)
    '               97 – 122 (lowercase a-z)
    '       INPUT:
    '               strValue   :   input string
    '
    '       OUTPUT:
    '
    '***********************************************************
    Dim intPos As Integer
    For intPos = 1 To Len(strValue)
    Select Case Asc(Mid(strValue, intPos, 1))
        Case 65 To 90, 97 To 122
            IsLetter = True
        Case Else
            IsLetter = False
            Exit For
    End Select
    Next
End Function
Function GetDuplicatedRowStatus(ByRef inMCDCResults() As Integer, _
                                ByVal nInMCDCResults_1 As Integer, _
                                ByVal nInMCDCResults_2 As Integer, _
                                    ByRef duplicatedRowStatus() As Integer, _
                                    ByVal nDuplicatedRowStatus As Integer, _
                                    ByRef log As String) As Integer
    '***********************************************************
    '
    'AUTHOR:            Le Thai (RBVH\EMB3)
    'DATE CREATED:      17/Feb/2014
    'FUNCTION NAME:     GetDuplicatedRowStatus
    'DESCRIPTION:       Get Duplicated Row Status
    '       INPUT:
    '               inMCDCResults()    :   inMCDCResults
    '               nInMCDCResults_1    :   array size
    '               nInMCDCResults_2    :   array size
    '               nDuplicatedRowStatus    :   array size
    '
    '       OUTPUT:
    '               duplicatedRowStatus()      :   duplicatedRowStatus
    '***********************************************************
    Dim i, j, k As Integer
    Dim duplicatedFlag As Boolean
    'ReDim Preserve duplicatedRowStatus(0 To UBound(inMCDCResults, 1)) As Integer
    'Init duplicatedRowStatus
    'For i = 0 To UBound(duplicatedRowStatus)
    For i = 0 To nDuplicatedRowStatus - 1
        duplicatedRowStatus(i) = -1
    Next i
    'Find duplicated rows
    'For i = 0 To UBound(inMCDCResults, 1)
    For i = 0 To nInMCDCResults_1 - 1
        'For j = i + 1 To UBound(inMCDCResults, 1)
        For j = i + 1 To nInMCDCResults_1 - 1
            'Check if status of current row is not duplicated
            If (duplicatedRowStatus(j) = -1) Then
                duplicatedFlag = True
                'For k = 0 To UBound(inMCDCResults, 2)
                For k = 0 To nInMCDCResults_2 - 1
                    If (duplicatedFlag = True) Then
                        If (inMCDCResults(i, k) <> inMCDCResults(j, k)) Then
                            duplicatedFlag = False
                        End If
                    End If
                Next k
                'Get duplicated result
                If (duplicatedFlag = True) Then
                    duplicatedRowStatus(j) = i
                End If
            End If
        Next j
    Next i
    GetDuplicatedRowStatus = OK
End Function
Function GetFinalMCDCResults(ByRef inMCDCResults() As Integer, _
                                ByVal nInMCDCResults_1 As Integer, _
                                ByVal nInMCDCResults_2 As Integer, _
                                ByRef duplicatedRowStatus() As Integer, _
                                ByVal nDuplicatedRowStatus As Integer, _
                                ByRef outMCDCResults() As Integer, _
                                ByRef nOutMCDCResults_1 As Integer, _
                                ByRef nOutMCDCResults_2 As Integer, _
                                ByRef highlights() As Integer, _
                                ByRef nHighlights_1 As Integer, _
                                ByRef nHighlights_2 As Integer, _
                                ByRef log As String) As Integer
    '***********************************************************
    '
    'AUTHOR:            Le Thai (RBVH\EMB3)
    'DATE CREATED:      17/Feb/2014
    'FUNCTION NAME:     GetFinalMCDCResults
    'DESCRIPTION:       Get Final MCDC Results
    '       INPUT:
    '               inMCDCResults()         :   2D array
    '               nInMCDCResults_1    :   array size
    '               nInMCDCResults_2    :   array size
    '               duplicatedRowStatus()   :   1D array
    '               nDuplicatedRowStatus    :   array size
    '
    '       OUTPUT:
    '               outMCDCResults()        :   2D array
    '               nOutMCDCResults_1    :   array size
    '               nOutMCDCResults_2    :   array size
    '               highlights()            :   2D array
    '               nHighlights_1    :   array size
    '               nHighlights_2    :   array size
    '***********************************************************
    Dim i, j As Integer
    Dim count As Integer
    Dim tmpHighlights(0 To 2 * MCDC_MAX_INPUTS - 1, 0 To MCDC_MAX_INPUTS) As Integer
    Dim nTmpHighlights_1 As Integer
    Dim nTmpHighlights_2 As Integer
    Dim outMCDCResultIdx As Integer
    Dim highlightIdx As Integer
    'Count the number of non-duplicated rows
    count = 0
    'For i = 0 To UBound(inMCDCResults, 1)
    For i = 0 To nInMCDCResults_1 - 1
        If (duplicatedRowStatus(i) < 0) Then
            count = count + 1
        End If
    Next i
    'Check count
    If (count <= 0) Then
        log = "ERROR"
        GetFinalMCDCResults = "ERROR! GetFinalMCDCResults! Unable to count the number of non-duplicated rows!"
        Exit Function
    End If
    'Calculate outMCDCResults
    'ReDim Preserve outMCDCResults(0 To count - 1, 0 To UBound(inMCDCResults, 2))
    nOutMCDCResults_1 = count
    nOutMCDCResults_2 = nInMCDCResults_2
    'Loop for all rows
    outMCDCResultIdx = 0
    'For i = 0 To UBound(inMCDCResults, 1)
    For i = 0 To nInMCDCResults_1 - 1
        'Check duplicated
        If (duplicatedRowStatus(i) < 0) Then
            'Loop for all inputs
            'For j = 0 To UBound(inMCDCResults, 2)
            For j = 0 To nInMCDCResults_2 - 1
                outMCDCResults(outMCDCResultIdx, j) = inMCDCResults(i, j)
            Next j
            outMCDCResultIdx = outMCDCResultIdx + 1
        End If
    Next i
    
    'Init tmpHighlights:
    'Note:  there is no outcome column for tmpHighlights
    '       nRows = 2*nCol for tmpHighlights
    'E.g 4 inputs, tmpHighlights should be 4X8 as below
    '           a   b   c   d           outcome(for reference only)
    'Row 0      1   0   0   0               1
    'Row 1      0   1   0   0               1
    'Row 2      0   0   1   0               1
    'Row 3      0   0   0   1               1
    'Row 4      1   0   0   0               0
    'Row 5      0   1   0   0               0
    'Row 6      0   0   1   0               0
    'Row 7      0   0   0   1               0
    'ReDim Preserve tmpHighlights(0 To UBound(inMCDCResults, 1), 0 To UBound(inMCDCResults, 2) - 1)
    nTmpHighlights_1 = nInMCDCResults_1
    nTmpHighlights_2 = nInMCDCResults_2 - 1
    'Loop for all rows
    'For i = 0 To UBound(tmpHighlights, 1)
    For i = 0 To nTmpHighlights_1 - 1
        'Loop for all input values
        'For j = 0 To UBound(tmpHighlights, 2)
        For j = 0 To nTmpHighlights_2 - 1
            'Case i <= UBound(tmpHighlights, 2)
            'If (i <= UBound(tmpHighlights, 2)) Then
            If (i < nTmpHighlights_2) Then
                If (i = j) Then
                    tmpHighlights(i, j) = 1
                Else
                    tmpHighlights(i, j) = 0
                End If
            'Case i > UBound(tmpHighlights, 2)
            Else
                'If ((i - UBound(tmpHighlights, 2) - 1) = j) Then
                If ((i - (nTmpHighlights_2 - 1) - 1) = j) Then
                    tmpHighlights(i, j) = 1
                Else
                    tmpHighlights(i, j) = 0
                End If
            
            End If

        Next j
    Next i
    'Move highlights at duplicated row to remain row
    'Note:  there is no outcome column for tmpHighlights
    '       nRows = 2*nCol for tmpHighlights
    'E.g 4 inputs, tmpHighlights and duplicatedRowStatus as below
    '
    '           a   b   c   d           duplicatedRowStatus
    'Row 0      1   0   0   0               -1
    'Row 1      0   1   0   0               0
    'Row 2      0   0   1   0               0
    'Row 3      0   0   0   1               -1
    'Row 4      1   0   0   0               3
    'Row 5      0   1   0   0               -1
    'Row 6      0   0   1   0               -1
    'Row 7      0   0   0   1               6
    '
    'After processing, tmpHighlights should be as below
    '
    '           a   b   c   d           duplicatedRowStatus
    'Row 0      1   1   1   0               -1
    '               ^   ^
    'Row 1      0   1   ^0   0               0
    '                   ^
    'Row 2      0   0   1   0               0
    'Row 3      1   0   0   1               -1
    '           ^
    'Row 4      1   0   0   0               3
    'Row 5      0   1   0   0               -1
    'Row 6      0   0   1   1               -1
    '                       ^
    'Row 7      0   0   0   1               6
    'Loop for all rows
    'For i = 0 To UBound(tmpHighlights, 1)
    For i = 0 To nTmpHighlights_1 - 1
        'Check duplicated
        If (duplicatedRowStatus(i) >= 0) Then
            'Mark highlight position
            'If (i <= UBound(tmpHighlights, 2)) Then
            If (i < nTmpHighlights_2) Then
                tmpHighlights(duplicatedRowStatus(i), i) = 1
            Else
                'tmpHighlights(duplicatedRowStatus(i), i - UBound(tmpHighlights, 2) - 1) = 1
                tmpHighlights(duplicatedRowStatus(i), i - (nTmpHighlights_2 - 1) - 1) = 1
            End If
        End If
    Next i
    
    'Calculate highlights
    'ReDim Preserve highlights(0 To count - 1, 0 To UBound(tmpHighlights, 2))
    nHighlights_1 = count
    nHighlights_2 = nTmpHighlights_2
    'Loop for all rows
    highlightIdx = 0
    'For i = 0 To UBound(tmpHighlights, 1)
    For i = 0 To nTmpHighlights_1 - 1
        'Check duplicated
        If (duplicatedRowStatus(i) < 0) Then
            'Loop for all inputs
            'For j = 0 To UBound(tmpHighlights, 2)
            For j = 0 To nTmpHighlights_2 - 1
                highlights(highlightIdx, j) = tmpHighlights(i, j)
            Next j
            highlightIdx = highlightIdx + 1
        End If
    Next i
    GetFinalMCDCResults = OK
End Function
Function GetSummaryMCDCResults(ByRef inputs() As Signal_Input, _
                                ByVal nInputs As Integer, _
                                ByRef operators() As String, _
                                ByVal nOperators As Integer, _
                                ByRef priorities() As Integer, _
                                ByVal nPriorities As Integer, _
                                ByRef finalMCDCResults() As Integer, _
                                ByVal nFinalMCDCResults_1 As Integer, _
                                ByVal nFinalMCDCResults_2 As Integer, _
                                ByRef summaryMCDCExpression As String, _
                                ByRef summaryMCDCResults() As String, _
                                ByVal nSummaryMCDCResults As Integer, _
                                ByRef log As String) As Integer
    '***********************************************************
    '
    'AUTHOR:            Le Thai (RBVH\EMB3)
    'DATE CREATED:      17/Feb/2014
    'FUNCTION NAME:     GetSumaryMCDCResults
    'DESCRIPTION:       Get Summary MCDC Results for MCDC column in MCDC table in MCDC sheet
    '       INPUT:
    '               inputs()                    :   1D array
    '               operators()                 :   1D array
    '               priorities()                :   1D array
    '               finalMCDCResults()          :   2D array
    '               nInputs                     :   array size
    '               nOperators                  :   array size
    '               nPriorities                 :   array size
    '               nFinalMCDCResults_1         :   array size
    '               nFinalMCDCResults_2         :   array size
    '               nSummaryMCDCResults         :   array size
    '
    '       OUTPUT:
    '               summaryMCDCExpression        :      eg. (a&&b&&c)
    '               summaryMCDCResults()         :      eg. {(--111--), (--100--), ...}
    '***********************************************************
    Dim alphabets() As String
    Dim highestPri, lowestPri As Integer
    Dim i, j As Integer
    Dim prevPri As Integer
    Dim bracketCarrier As Integer
    
    'FIND summaryMCDCExpression
    'Check case only 1 input
    'If (UBound(inputs) = 0) Then
    If (nInputs <= 1) Then
        summaryMCDCExpression = "a"
    Else
        'Alphabet charater array
        alphabets = Split("a,b,c,d,e,f,g,h,i,j,k,l,m,n,o,p,q,r,s,t,u,v,w,x,y,z,aa,ab,ac,ad,ae,af,ag,ah,ai,aj,ak,al,am,an,ao,ap,aq,ar,as,at,au,av,aw,ax,ay,az", ",")
        'Find the highest priority (smallest number) and lowest priority (biggest number)
        highestPri = PRIORITY_LOWEST
        lowestPri = PRIORITY_HIGHEST
        'For i = 0 To UBound(priorities())
        For i = 0 To nPriorities - 1
            If (priorities(i) < highestPri) Then
                highestPri = priorities(i)
            End If
            If (priorities(i) > lowestPri) Then
                lowestPri = priorities(i)
            End If
        Next i
        'Get summaryMCDCExpression
        summaryMCDCExpression = ""
        prevPri = highestPri
        bracketCarrier = 0
        'For i = 0 To UBound(operators)
        For i = 0 To nOperators - 1
            'Priority no change
            If (prevPri = priorities(i)) Then
                summaryMCDCExpression = summaryMCDCExpression & alphabets(i) & operators(i)
            'Priority decrease
            ElseIf (prevPri < priorities(i)) Then
                bracketCarrier = bracketCarrier + (priorities(i) - prevPri)
                'Add (priorities(i) - prevPri) character "("
                For j = 1 To priorities(i) - prevPri
                     summaryMCDCExpression = summaryMCDCExpression & "("
                Next j
                summaryMCDCExpression = summaryMCDCExpression & alphabets(i) & operators(i)
            'Priority increase
            Else
                bracketCarrier = bracketCarrier + (priorities(i) - prevPri)
                summaryMCDCExpression = summaryMCDCExpression & alphabets(i)
                'Add (prevPri - priorities(i)) character ")"
                For j = 1 To prevPri - priorities(i)
                     summaryMCDCExpression = summaryMCDCExpression & ")"
                Next j
                summaryMCDCExpression = summaryMCDCExpression & operators(i)
            End If
            prevPri = priorities(i)
        Next i
        'Last input for summaryMCDCExpression
        'summaryMCDCExpression = summaryMCDCExpression & alphabets(UBound(operators) + 1)
        summaryMCDCExpression = summaryMCDCExpression & alphabets(nOperators)
        If (bracketCarrier > 0) Then
            'Add bracketCarrier charater ")"
            For i = 1 To bracketCarrier
                 summaryMCDCExpression = summaryMCDCExpression & ")"
            Next i
        End If
    End If 'End if: FIND summaryMCDCExpression
    'FIND summaryMCDCResults
    'ReDim Preserve summaryMCDCResults(0 To UBound(finalMCDCResults, 1))
    'For i = 0 To UBound(finalMCDCResults, 1)
    For i = 0 To nFinalMCDCResults_1 - 1
        summaryMCDCResults(i) = "(--"
        'For j = 0 To UBound(finalMCDCResults, 2) - 1
        For j = 0 To nFinalMCDCResults_2 - 2
            summaryMCDCResults(i) = summaryMCDCResults(i) & finalMCDCResults(i, j)
        Next j
        summaryMCDCResults(i) = summaryMCDCResults(i) & "--)"
    Next i
    GetSummaryMCDCResults = OK
End Function
Sub WriteReq(ByVal condition As String, _
            ByVal level As String, _
            ByVal lineNumber As Integer, _
            ByRef inputs() As Signal_Input, _
            ByVal nInputs As Integer, _
            ByRef MCDCResults() As Integer, _
            ByVal nMCDCResults_1 As Integer, _
            ByVal nMCDCResults_2 As Integer, _
            ByRef highlights() As Integer, _
            ByVal nHighlights_1 As Integer, _
            ByVal nHighlights_2 As Integer, _
            ByVal summaryMCDCExpression As String, _
            ByRef summaryMCDCResults() As String, _
            ByVal nSummaryMCDCResults As Integer)
    '***********************************************************
    '
    'AUTHOR:            Le Thai (RBVH\EMB3)
    'DATE CREATED:      14/Feb/2014
    'FUNCTION NAME:     WriteReq
    'DESCRIPTION: Write requirement table in MCDC sheet
    '       INPUT:
    '               -   inputs                  :   array of input signals in condition
    '               -   condition               :   condition
    '               -   level                   :   level
    '               -   MCDCResults             :   MCDCResults
    '               -   highlights              :   highlights
    '       OUTPUT:
    '
    '***********************************************************
    Dim row, col As Integer
    Dim lastReqRow, INPUTSRow As Integer
    Dim startOutputRow As Integer
    Dim reqName As String
    Dim reqCount As Integer
    Dim i, j As Integer
    
    'Find last requirement row number
    row = FindRowAll(MCDC_SHEET_NAME, 1, "*", 1)
    If (row < 0) Then
        reqCount = 0
        lastReqRow = 1
        row = 1
        startOutputRow = 1
    Else
        reqCount = 0
        lastReqRow = 1
        Do While (row > 0)
            reqCount = reqCount + 1
            lastReqRow = row
            row = FindRowAll(MCDC_SHEET_NAME, 1, "*", row + 1)
        Loop
        'Find ending row number for last requirement
        'Find row number of cell "INPUTS" in last req
        INPUTSRow = FindRowAll(MCDC_SHEET_NAME, 4, "INPUTS", lastReqRow)
        If (INPUTSRow < 0) Then
            MsgBox "AddReq! Unable to find cell 'INPUTS' in the last requirement"
            Exit Sub
        End If
        'Find ending row number for last requirement
        row = FindRowAll(MCDC_SHEET_NAME, 4, " ", INPUTSRow)
        If (row < 0) Then
            MsgBox "AddReq! Unable to find ending row number for last requirement"
            Exit Sub
        End If
        startOutputRow = row + 2
    End If
    'OUTPUT DATA
    If (condition = vbNullString) Then
        'Copy MCDC table in Guide sheet to MCDC sheet
        Worksheets("Guide").Range(MCDC_SHEET_NAME).Copy
        Worksheets(MCDC_SHEET_NAME).Cells(startOutputRow, 1).PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
                False, Transpose:=False
        reqName = Worksheets(MCDC_SHEET_NAME).Cells(startOutputRow, MCDC_REQ_COL).value
        Worksheets(MCDC_SHEET_NAME).Cells(startOutputRow, MCDC_REQ_COL).value = Replace(reqName, "1", reqCount + 1)
    Else
        'OUTPUT DATA
        'Requirement
        row = startOutputRow
        Worksheets(MCDC_SHEET_NAME).Cells(row, MCDC_REQ_COL).value = MCDC_REQ_STR & " " & (reqCount + 1)
        'Level
        'Worksheets(MCDC_SHEET_NAME).Cells(row, MCDC_REQ_COL + 1).value = level
        'Line number
        Worksheets(MCDC_SHEET_NAME).Cells(row, MCDC_REQ_COL + 2).value = "#Line " & lineNumber
        'Check MCDC for condition
        row = row + 1
        'If (UBound(inputs) > 0) Then
        If (nInputs > 1) Then
            Worksheets(MCDC_SHEET_NAME).Cells(row, MCDC_TC_NO_COL).value = MCDC_CONDITION_1_STR & vbNewLine & condition
            'Worksheets(MCDC_SHEET_NAME).Cells(row, MCDC_TC_NO_COL).value = MCDC_CONDITION_1_STR & condition
        Else
            Worksheets(MCDC_SHEET_NAME).Cells(row, MCDC_TC_NO_COL).value = MCDC_CONDITION_2_STR & vbNewLine & condition
            'Worksheets(MCDC_SHEET_NAME).Cells(row, MCDC_TC_NO_COL).value = MCDC_CONDITION_2_STR & condition
        End If
        'Table: Header
        row = row + 2
        Worksheets(MCDC_SHEET_NAME).Cells(row, MCDC_MCDC_COL).value = MCDC_MCDC_STR
        Worksheets(MCDC_SHEET_NAME).Cells(row, MCDC_INPUT_COL).value = MCDC_INPUT_STR
        'Worksheets(MCDC_SHEET_NAME).Cells(row, MCDC_INPUT_COL + UBound(inputs) + 1).value = MCDC_OUTPUT_STR
        Worksheets(MCDC_SHEET_NAME).Cells(row, MCDC_INPUT_COL + nInputs).value = MCDC_OUTPUT_STR
        'Worksheets(MCDC_SHEET_NAME).Cells(row, MCDC_INPUT_COL + UBound(inputs) + 2).value = MCDC_OUTCOME_STR
        Worksheets(MCDC_SHEET_NAME).Cells(row, MCDC_INPUT_COL + nInputs + 1).value = MCDC_OUTCOME_STR
        'Table: "TC No.", summary MCDC expression, signal names
        row = row + 1
        Worksheets(MCDC_SHEET_NAME).Cells(row, MCDC_TC_NO_COL).value = MCDC_TC_NO_STR
        '
        Worksheets(MCDC_SHEET_NAME).Cells(row, MCDC_MCDC_COL).value = summaryMCDCExpression
        'For i = 0 To UBound(inputs)
        For i = 0 To nInputs - 1
            Worksheets(MCDC_SHEET_NAME).Cells(row, MCDC_INPUT_COL + i).value = inputs(i).name
        Next i
        'Table: value
        row = row + 1
        'Loop for all testcase
        'For i = 0 To UBound(MCDCResults, 1)
        For i = 0 To nMCDCResults_1 - 1
            '"TCX"
            Worksheets(MCDC_SHEET_NAME).Cells(row, MCDC_TC_NO_COL).value = MCDC_TCX_STR
            'Summary MCDC
            Worksheets(MCDC_SHEET_NAME).Cells(row, MCDC_MCDC_COL).value = summaryMCDCResults(i)
            'Loop for all inputs
            'For j = 0 To UBound(MCDCResults, 2) - 1
            For j = 0 To nMCDCResults_2 - 2
                'Check value: TRUE
                If (MCDCResults(i, j) = 1) Then
                    Worksheets(MCDC_SHEET_NAME).Cells(row, MCDC_INPUT_COL + j).value = inputs(j).value
                'Check value: FALSE
                Else
                    'Check "!="
                    If (InStr(inputs(j).value, OPERATOR_DIF) > 0) Then
                        Worksheets(MCDC_SHEET_NAME).Cells(row, MCDC_INPUT_COL + j).value = Replace(inputs(j).value, OPERATOR_DIF, "")
                    'Check ">="
                    ElseIf (InStr(inputs(j).value, OPERATOR_GE) > 0) Then
                        Worksheets(MCDC_SHEET_NAME).Cells(row, MCDC_INPUT_COL + j).value = Replace(inputs(j).value, OPERATOR_GE, OPERATOR_LT)
                    'Check "<="
                    ElseIf (InStr(inputs(j).value, OPERATOR_LE) > 0) Then
                        Worksheets(MCDC_SHEET_NAME).Cells(row, MCDC_INPUT_COL + j).value = Replace(inputs(j).value, OPERATOR_LE, OPERATOR_GT)
                    'Check ">"
                    ElseIf (InStr(inputs(j).value, OPERATOR_GT) > 0) Then
                        Worksheets(MCDC_SHEET_NAME).Cells(row, MCDC_INPUT_COL + j).value = Replace(inputs(j).value, OPERATOR_GT, OPERATOR_LE)
                    'Check "<"
                    ElseIf (InStr(inputs(j).value, OPERATOR_LT) > 0) Then
                        Worksheets(MCDC_SHEET_NAME).Cells(row, MCDC_INPUT_COL + j).value = Replace(inputs(j).value, OPERATOR_LT, OPERATOR_GE)
                    'Check "TRUE"
                    ElseIf (InStr(inputs(j).value, "TRUE") > 0 Or _
                            InStr(inputs(j).value, "True") > 0 Or _
                            InStr(inputs(j).value, "true") > 0) Then
                        Worksheets(MCDC_SHEET_NAME).Cells(row, MCDC_INPUT_COL + j).value = "FALSE"
                    'Check "FALSE"
                    ElseIf (InStr(inputs(j).value, "FALSE") > 0 Or _
                            InStr(inputs(j).value, "False") > 0 Or _
                            InStr(inputs(j).value, "false") > 0) Then
                        Worksheets(MCDC_SHEET_NAME).Cells(row, MCDC_INPUT_COL + j).value = "TRUE"
                    Else
                        Worksheets(MCDC_SHEET_NAME).Cells(row, MCDC_INPUT_COL + j).value = "!=" & inputs(j).value
                    End If
                End If
                'Check highlights
                If (highlights(i, j) = 1) Then
                    Worksheets(MCDC_SHEET_NAME).Cells(row, MCDC_INPUT_COL + j).Interior.ColorIndex = LIGHT_BLUE_2
                End If
            Next j
            'OUTCOME: TRUE
            'If (MCDCResults(i, UBound(MCDCResults, 2)) = 1) Then
            If (MCDCResults(i, nMCDCResults_2 - 1) = 1) Then
                'Worksheets(MCDC_SHEET_NAME).Cells(row, MCDC_INPUT_COL + UBound(inputs) + 2).value = "TRUE"
                Worksheets(MCDC_SHEET_NAME).Cells(row, MCDC_INPUT_COL + nInputs + 1).value = "TRUE"
            'OUTCOME: FALSE
            Else
                'Worksheets(MCDC_SHEET_NAME).Cells(row, MCDC_INPUT_COL + UBound(inputs) + 2).value = "FALSE"
                Worksheets(MCDC_SHEET_NAME).Cells(row, MCDC_INPUT_COL + nInputs + 1).value = "FALSE"
            End If
            'increasement
            row = row + 1
        Next i
        'Font, Color, Border
        'Activate sheet MCDC_SHEET_NAME
        Worksheets(MCDC_SHEET_NAME).Activate
        Worksheets(MCDC_SHEET_NAME).Cells(startOutputRow, MCDC_REQ_COL).Select
        'Req
        Worksheets(MCDC_SHEET_NAME).Cells(startOutputRow, MCDC_REQ_COL).Font.Bold = True
        Worksheets(MCDC_SHEET_NAME).Cells(startOutputRow, MCDC_REQ_COL).Interior.ColorIndex = BLUE
        'Table: Header
        'Worksheets(MCDC_SHEET_NAME).Range(Cells(startOutputRow + 3, MCDC_TC_NO_COL), Cells(startOutputRow + 4, MCDC_INPUT_COL + UBound(inputs) + 2)).Font.Bold = True
        Worksheets(MCDC_SHEET_NAME).Range(Cells(startOutputRow + 3, MCDC_TC_NO_COL), Cells(startOutputRow + 4, MCDC_INPUT_COL + nInputs + 1)).Font.Bold = True
        'Border
        'Worksheets(MCDC_SHEET_NAME).Range(Cells(startOutputRow + 3, MCDC_TC_NO_COL), Cells(row - 1, MCDC_INPUT_COL + UBound(inputs) + 2)).Borders.LineStyle = xlContinuous
        Worksheets(MCDC_SHEET_NAME).Range(Cells(startOutputRow + 3, MCDC_TC_NO_COL), Cells(row - 1, MCDC_INPUT_COL + nInputs + 1)).Borders.LineStyle = xlContinuous
        'Allignment
        'Worksheets(MCDC_SHEET_NAME).Range(Cells(startOutputRow + 3, MCDC_TC_NO_COL), Cells(row - 1, MCDC_INPUT_COL + UBound(inputs) + 2)).HorizontalAlignment = xlCenter
        Worksheets(MCDC_SHEET_NAME).Range(Cells(startOutputRow + 3, MCDC_TC_NO_COL), Cells(row - 1, MCDC_INPUT_COL + nInputs + 1)).HorizontalAlignment = xlCenter
        'Worksheets(MCDC_SHEET_NAME).Range(Cells(startOutputRow + 3, MCDC_TC_NO_COL), Cells(row - 1, MCDC_INPUT_COL + UBound(inputs) + 2)).VerticalAlignment = xlCenter
        Worksheets(MCDC_SHEET_NAME).Range(Cells(startOutputRow + 3, MCDC_TC_NO_COL), Cells(row - 1, MCDC_INPUT_COL + nInputs + 1)).VerticalAlignment = xlCenter
        'Font Red
        Worksheets(MCDC_SHEET_NAME).Range(Cells(startOutputRow + 5, MCDC_TC_NO_COL), Cells(row - 1, MCDC_TC_NO_COL)).Font.Color = vbRed
        'Worksheets(MCDC_SHEET_NAME).Range(Cells(startOutputRow + 5, MCDC_INPUT_COL + UBound(inputs) + 2), Cells(row - 1, MCDC_INPUT_COL + UBound(inputs) + 2)).Font.Color = vbRed
        Worksheets(MCDC_SHEET_NAME).Range(Cells(startOutputRow + 5, MCDC_INPUT_COL + nInputs + 1), Cells(row - 1, MCDC_INPUT_COL + nInputs + 1)).Font.Color = vbRed
    End If
End Sub


