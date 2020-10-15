Attribute VB_Name = "MExpressionFormular"

Option Explicit
Const OK = 1
Const ERROR = -1
'$$$$$$$$$$$$ DEFINITIONS FOR "Testcases" SHEET $$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
Const TC_TC_NO_COLUMN = 1
Const EXPRESSION_OPERATORS = "+,-,*,/"
Const EXPRESSION_MAX_INPUTS = 50
Const EXPRESSION_NAME_PREFIX_1 = "actual_"
Const EXPRESSION_NAME_PREFIX_2 = "exp_"
Const EXPRESSION_NAME_SUFFIX = ""
Const EXPRESSION_GT_STR = ">"
Const EXPRESSION_GE_STR = ">="
Const EXPRESSION_LT_STR = "<"
Const EXPRESSION_LE_STR = "<="
Const EXPRESSION_EQ_STR = "="
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
Type Expression_Input
    name As String
    cell As String
End Type
Sub InsertExpressionFormula()
    '***********************************************************
    '
    'AUTHOR:            Le Thai (RBVH\EMB3)
    'DATE CREATED:      04/Jul/2013
    'FUNCTION NAME:     InsertExpressionFormula
    'DESCRIPTION: Fill  formula for active cell with specific expression
    '             E.g:  expression  : "input1+input2",
    '                   active cell : "T20"
    '                   "input1" is at column H
    '                   "input2" is at column A
    '           --> Output: Value at cell "T20" is "=H20+A20"
    '       INPUT: Expression
    '
    '       OUTPUT: Insert Expression Formula for active cell
    '
    '***********************************************************
    Dim patterns() As String
    Dim TCNoRow As Integer
    Dim toleranceRow As Integer
    Dim TCSheet As String
    Dim activeRow As Integer
    Dim activeCol As Integer
    Dim formula As String
    Dim temp As Variant
    Dim defaultExpression As String
    Dim expression As String
    Dim inputs(0 To EXPRESSION_MAX_INPUTS - 1) As Expression_Input
    Dim nInputs As Integer
    Dim log As String
    Dim toleranceCellAddress As String
    Dim cmpSign As String
    'Back up
    Backup
    'TCSheet
    TCSheet = "Testcases"
    If (ActiveCell <> vbNullString) Then
        expression = ActiveCell.value
    Else
        'Get Vaiable
        If (GetVar("defaultExpression", defaultExpression) > 0) Then
    
        Else
            defaultExpression = "(A+B)*C"
        End If
        'Get expression
        temp = Application.InputBox(Prompt:="Please enter the expression.", _
                                                Title:="EXPRESSION", _
                                                Default:=defaultExpression, _
                                                Left:=1, _
                                                Top:=1000, _
                                                Type:=2)
        If (temp = False) Then
            Exit Sub
        End If
        expression = temp
    End If
    'SaveVar
    If (SaveVar("defaultExpression", expression) > 0) Then
        'MsgBox "Save 'defaultPath' with value: " & filePath
    'Save failed
    Else
        'MsgBox "Unable to save 'defaultPath' with value: " & filePath
    End If
    'Remove " ", tab
    'Note: chr(9) = tab charater
    patterns = Split(" ," & Chr(9), ",")
    expression = ReplaceAll(expression, patterns(), "")
    'Check ">"
    If (Mid(expression, 1, 1) = ">") Then
        'Remove ">"
        expression = Mid(expression, 2)
        'Check "="
        If (Mid(expression, 1, 1) = "=") Then
            'Remove "="
            expression = Mid(expression, 2)
            cmpSign = EXPRESSION_GE_STR
        Else
            cmpSign = EXPRESSION_GT_STR
        End If
    'Check "<"
    ElseIf (Mid(expression, 1, 1) = "<") Then
        'Remove "<"
        expression = Mid(expression, 2)
        'Check "="
        If (Mid(expression, 1, 1) = "=") Then
            'Remove "="
            expression = Mid(expression, 2)
            cmpSign = EXPRESSION_LE_STR
        Else
            cmpSign = EXPRESSION_LT_STR
        End If
    Else
        cmpSign = EXPRESSION_EQ_STR
    End If
    
    'Find "TC No."
    TCNoRow = FindRowAll(TCSheet, TC_TC_NO_COLUMN, "TC No.", 1)
    If (TCNoRow < 0) Then
        MsgBox "InsertExpressionFormula! Unable to to find cell 'TC No' in '" & TCSheet & "' sheet"
        Exit Sub
    End If
    'Tolerance Row
    If (TCNoRow > 4) Then
        toleranceRow = TCNoRow - 4
    Else
        MsgBox "InsertExpressionFormula! 'TC No.' row number must be greater than 4!"
        Exit Sub
    End If
    activeRow = ActiveCell.row
    activeCol = ActiveCell.Column
    'Get input names
    If (Expression_GetInputNames(expression, inputs(), nInputs, log) < 0) Then
        MsgBox "InsertExpressionFormula!" & log
        Exit Sub
    End If
    'Get input cell addresses
    If (Expression_GetInputCells(TCSheet, activeRow, TCNoRow + 1, inputs(), nInputs, log) < 0) Then
        MsgBox "InsertExpressionFormula!" & log
        Exit Sub
    End If
    'Get Formula
    toleranceCellAddress = Worksheets(ActiveSheet.name).Cells(toleranceRow, activeCol).Address
    If (Expression_GetFormula(expression, inputs(), nInputs, formula, log) < 0) Then
        MsgBox "InsertExpressionFormula!" & log
        Exit Sub
    End If
    'Check comparision sign
    If (cmpSign = EXPRESSION_EQ_STR Or _
        cmpSign = EXPRESSION_GE_STR Or _
        cmpSign = EXPRESSION_LE_STR) Then
        ActiveCell.value = "=" & formula
    ElseIf (cmpSign = EXPRESSION_GT_STR) Then
        ActiveCell.value = "=" & formula & "+" & toleranceCellAddress
    Else
        ActiveCell.value = "=" & formula & "-" & toleranceCellAddress
    End If
End Sub

Function Expression_GetInputNames(ByVal expression As String, _
                    ByRef inputs() As Expression_Input, _
                    ByRef nInputs As Integer, _
                    ByRef log As String) As Integer
    '***********************************************************
    '
    'AUTHOR:            Le Thai (RBVH\EMB3)
    'DATE CREATED:      12/Feb/2014
    'FUNCTION NAME:     Expression_GetInputNames
    'DESCRIPTION:
    '       INPUT:
    '               expression   :   expression
    '
    '       OUTPUT:
    '               inputs()        :   array contains all the inputs names and value
    '               nInputs         :   array size
    '
    '***********************************************************
    Dim trimmedExpression As String
    Dim patterns() As String
    Dim repStr As String
    Dim names() As String
    Dim i As Integer
    
    trimmedExpression = expression
    'Remove " ", "(", ")"
    patterns = Split(" |(|)", "|")
    repStr = ""
    trimmedExpression = ReplaceAll(trimmedExpression, patterns(), repStr)
    'Split to get single conditions
    patterns = Split(EXPRESSION_OPERATORS, ",")
    names = SplitAll(trimmedExpression, patterns())
    'Get inputs
    nInputs = 0
    For i = 0 To UBound(names)
        If (IsNumeric(names(i)) = False) Then
            inputs(nInputs).name = names(i)
            nInputs = nInputs + 1
        End If
     Next i
    Expression_GetInputNames = OK
End Function
Function Expression_GetInputCells(ByVal sheet As Variant, _
                    ByVal activeRow As Integer, _
                    ByVal nameRow As Integer, _
                    ByRef inputs() As Expression_Input, _
                    ByRef nInputs As Integer, _
                    ByRef log As String) As Integer
    '***********************************************************
    '
    'AUTHOR:            Le Thai (RBVH\EMB3)
    'DATE CREATED:      12/Feb/2014
    'FUNCTION NAME:     Expression_GetInputCells
    'DESCRIPTION:
    '       INPUT:
    '               sheet   :   sheet
    '               activeRow   :   expression
    '               nameRow   :   expression
    '               inputs()        :   array contains all the inputs names
    '               nInputs         :   array size
    '
    '       OUTPUT:
    '               inputs()        :   array contains all the inputs names and cells
    '
    '***********************************************************
    Dim i As Integer
    Dim col As Integer
    Dim patterns() As String
    For i = 0 To nInputs - 1
        'Find input column
        col = FindColAll(sheet, nameRow, EXPRESSION_NAME_PREFIX_1 & inputs(i).name & EXPRESSION_NAME_SUFFIX, 1)
        If (col < 0) Then
            'Try with no PREFIX and SUFFIX
            col = FindColAll(sheet, nameRow, inputs(i).name, 1)
            If (col < 0) Then
                'Try with PREFIX 2
                col = FindColAll(sheet, nameRow, EXPRESSION_NAME_PREFIX_2 & inputs(i).name, 1)
                If (col < 0) Then
                    log = "Expression_GetInputCells! unable to find column number for input '" & inputs(i).name & "'!"
                    Expression_GetInputCells = ERROR
                    Exit Function
                End If
            End If
        End If
        inputs(i).cell = Worksheets(sheet).Cells(activeRow, col).Address
        'Remove "$"
        patterns = Split("$|$", "|")
        inputs(i).cell = ReplaceAll(inputs(i).cell, patterns(), "")
    Next i
    Expression_GetInputCells = OK
End Function
Function Expression_GetFormula(ByVal expression As String, _
                    ByRef inputs() As Expression_Input, _
                    ByVal nInputs As Integer, _
                    ByRef formula As String, _
                    ByRef log As String) As Integer
    '***********************************************************
    '
    'AUTHOR:            Le Thai (RBVH\EMB3)
    'DATE CREATED:      12/Feb/2014
    'FUNCTION NAME:     Expression_GetFormula
    'DESCRIPTION:
    '       INPUT:
    '               expression   :   expression
    '               inputs()        :   array contains all the inputs names and value
    '               nInputs         :   array size
    '       OUTPUT:
    '               formula         :   formula
    '
    '***********************************************************
    Dim trimmedExpression As String
    Dim patterns() As String
    Dim repStr As String
    Dim names() As String
    Dim i As Integer
    
    trimmedExpression = expression
    'Remove " ", tab
    'Note: chr(9) = tab charater
    patterns = Split(" ," & Chr(9), ",")
    repStr = ""
    trimmedExpression = ReplaceAll(trimmedExpression, patterns(), repStr)
    formula = trimmedExpression
    For i = 0 To nInputs - 1
        formula = Replace(formula, inputs(i).name, inputs(i).cell, 1, 1)
    Next i
    Expression_GetFormula = OK
End Function





