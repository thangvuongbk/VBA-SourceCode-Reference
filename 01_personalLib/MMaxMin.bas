Attribute VB_Name = "MMaxMin"
Option Explicit
Const OK = 1
Const ERROR = -1
'$$$$$$$$$$$$ DEFINITIONS FOR "Testcases" SHEET $$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
Const TC_TC_NO_COLUMN = 1
Sub InsertMaxMinFormula()
    '***********************************************************
    '
    'AUTHOR:            Le Thai (RBVH\EMB3)
    'DATE CREATED:      04/Jul/2013
    'FUNCTION NAME:     InsertMaxMinFormula
    'DESCRIPTION:
    '       INPUT: Active cell
    '
    '       OUTPUT: Insert Max Min Formula for active cell
    '
    '***********************************************************
    Dim toleranceRow As Integer
    Dim MaxRow As Integer
    Dim MinRow As Integer
    Dim TCSheet As String
    Dim formula As String
    Dim finalFormula As String
    'Back up
    Backup
    'TCSheet
    TCSheet = "Testcases"
    'Find "Tolerance"
    toleranceRow = FindRowAll(TCSheet, TC_TC_NO_COLUMN, "Tolerance", 1)
    If (toleranceRow < 0) Then
        MsgBox "InsertMaxMinFormula! Unable to to find cell 'Tolerance' in '" & TCSheet & "' sheet"
        Exit Sub
    End If
    'Find "Max"
    MaxRow = FindRowAll(TCSheet, TC_TC_NO_COLUMN, "Max", 1)
    If (MaxRow < 0) Then
        MsgBox "InsertMaxMinFormula! Unable to to find cell 'Max' in '" & TCSheet & "' sheet"
        Exit Sub
    End If
    'Find "Min"
    MinRow = FindRowAll(TCSheet, TC_TC_NO_COLUMN, "Min", 1)
    If (MinRow < 0) Then
        MsgBox "InsertMaxMinFormula! Unable to to find cell 'Min' in '" & TCSheet & "' sheet"
        Exit Sub
    End If
    
    'Check empty active cell
    If (ActiveCell = vbNullString) Then
        MsgBox "InsertMaxMinFormula! Active cell is empty!"
        Exit Sub
    End If
    'Check Tolerance cell
    If (Worksheets(TCSheet).Cells(toleranceRow, ActiveCell.Column) = vbNullString) Then
        MsgBox "InsertMaxMinFormula! 'Tolerance' cell is empty!"
        Exit Sub
    ElseIf (Not IsNumeric(Worksheets(TCSheet).Cells(toleranceRow, ActiveCell.Column).Value2)) Then
        MsgBox "InsertMaxMinFormula! 'Tolerance' cell is not a number!"
        Exit Sub
    End If
    'Check Max cell
    If (Worksheets(TCSheet).Cells(MaxRow, ActiveCell.Column) = vbNullString) Then
        MsgBox "InsertMaxMinFormula! 'Max' cell is empty!"
        Exit Sub
    ElseIf (Not IsNumeric(Worksheets(TCSheet).Cells(MaxRow, ActiveCell.Column).Value2)) Then
        MsgBox "InsertMaxMinFormula! 'Max' cell is not a number!"
        Exit Sub
    End If
    'Check Min cell
    If (Worksheets(TCSheet).Cells(MinRow, ActiveCell.Column) = vbNullString) Then
        MsgBox "InsertMaxMinFormula! 'Min' cell is empty!"
        Exit Sub
    ElseIf (Not IsNumeric(Worksheets(TCSheet).Cells(MinRow, ActiveCell.Column).Value2)) Then
        MsgBox "InsertMaxMinFormula! 'Min' cell is not a number!"
        Exit Sub
    End If
    'Formula
    formula = ActiveCell.formula
    If (InStr(formula, "=") > 0) Then
        formula = Mid(formula, InStr(formula, "=") + 1)
    End If
    'Formula: recalculate with Q value
    'formula = "RoundDown((" & formula & ")/" & Worksheets(TCSheet).Cells(toleranceRow, ActiveCell.Column).Address & ",0)*" & Worksheets(TCSheet).Cells(toleranceRow, ActiveCell.Column).Address
    formula = "Round((" & formula & ")/" & Worksheets(TCSheet).Cells(toleranceRow, ActiveCell.Column).Address & ",0)*" & Worksheets(TCSheet).Cells(toleranceRow, ActiveCell.Column).Address
    'finalFormular
    finalFormula = "MIN(MAX(" & formula & "," & Worksheets(TCSheet).Cells(MinRow, ActiveCell.Column).Address & ")," & Worksheets(TCSheet).Cells(MaxRow, ActiveCell.Column).Address & ")"
    
    'MsgBox "Cell: " & ActiveCell.Address & vbNewLine & _
            "Row: " & ActiveCell.row & vbNewLine & _
            "Column: " & ActiveCell.Column & vbNewLine & _
            "Value: " & ActiveCell.Value & vbNewLine & _
            "Formula: " & ActiveCell.formula & vbNewLine & _
            "Updated Formula: " & finalFormula
    'Update active cell formula
    ActiveCell.value = "=" & finalFormula
End Sub
