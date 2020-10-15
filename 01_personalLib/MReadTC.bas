Attribute VB_Name = "MReadTC"
Option Explicit
'$$$$$$$$$$$$ DEFINITIONS FOR "Testcases" SHEET $$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
Const TC_TC_NO_COLUMN = 1
Sub ReadTC()
    '***********************************************************
    '
    'AUTHOR:            Le Thai (RBVH\EMB3)
    'DATE CREATED:      19/August/2013
    'FUNCTION NAME:     ReadTC
    'DESCRIPTION:
    '       INPUT: Testcase files in CSV format
    '
    '       OUTPUT: Testcase design in Testcases sheet
    '
    '***********************************************************
    Dim filepath() As String
    Dim lines() As String
    Dim signals() As String
    Dim signalValues() As String
    Dim log As String
    Dim condition As String
    Dim i, j As Integer
    Dim oRow As Integer
    Dim temp As Integer
    Dim remarksCol As Integer
    Dim startingORow As Integer
    Dim TCSheet As String
    
    TCSheet = "Testcases"
    'Get file names
    If (ReadDir(filepath(), log) < 0) Then
        'MsgBox "ReadTC!" & log
        Exit Sub
    End If
    'Find "TC No." in Testcases sheet
    startingORow = FindRowAll(TCSheet, TC_TC_NO_COLUMN, "TC No.", 1)
    If (startingORow < 0) Then
        MsgBox "GenTC! Unable to to find cell '" & TCSheet & "' sheet"
        Exit Sub
    End If
    startingORow = startingORow + 2
    'Loop all files
    For i = 0 To UBound(filepath)
        'MsgBox filePath(i)
        'Get lines
        If (ReadFileCSV(filepath(i), lines(), log) < 0) Then
            Exit Sub
        End If
         'Get Signals
        If (GetSignals(lines(0), signals(), log) < 0) Then
            MsgBox log
            Exit Sub
        End If
         'Get Signal Values
        If (GetSignals(lines(3), signalValues(), log) < 0) Then
            MsgBox log
            Exit Sub
        End If
        
        'DEBUG
        'CSVDisplay1D signals(), 1
        'CSVDisplay1D signalValues(), 2
        
        'Create condition
        condition = ""
        For j = 2 To UBound(signals())
            If (j = 2) Then
                condition = signals(j) & "=" & signalValues(j)
            Else
                condition = condition & " && " & signals(j) & "=" & signalValues(j)
            End If
        Next j
        
        'DEBUG
        'MsgBox condition
        'Find oRow
        oRow = FindRowAll(TCSheet, 1, " ", startingORow)
        If (oRow < 0) Then
            MsgBox "ReadTC!Unable to find row number to write TC!"
        End If
        'Write TCNo
        Worksheets(TCSheet).Cells(oRow, 1).value = "TC" & (i + 1)
        'Write REMARKS = filePath
        'Find "REMARKS" column number
        remarksCol = FindColAll(TCSheet, startingORow - 2, "REMARKS", 1)
        If (remarksCol < 0) Then
            MsgBox "ReadTC! Unable to find cell 'REMARKS' to fill file path."
        Else
            Worksheets(TCSheet).Cells(oRow, remarksCol).value = Dir(filepath(i), vbNormal)
        End If
        'Write TC
        temp = WriteTCAll(TCSheet, condition, oRow, 0, log, startingORow - 1)
        If (temp < 0) Then
            MsgBox log
            Exit Sub
        End If
    Next i
    'Activate Testcases sheet
    Worksheets(TCSheet).Activate
    Cells(2, 2).Select
End Sub
Function ReadDir(ByRef filepath() As String, _
                ByRef log As String) As Integer
    '***********************************************************
    '
    'AUTHOR:            Le Thai (RBVH\EMB3)
    'DATE CREATED:      19/August/2013
    'FUNCTION NAME:     ReadCSV
    'DESCRIPTION:
    '       INPUT: Testcase files in CSV format
    '
    '       OUTPUT:
    '               lines        :   array of lines in CSV file
    '               log          :   log
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
        log = "ReadDir! File not found "
        ReadDir = -1
        Exit Function
    End If
    ReDim Preserve filepath(0 To (UBound(fileToOpen) - 1))
    For i = 1 To UBound(fileToOpen)
        filepath(i - 1) = fileToOpen(i)
    Next i
    ReadDir = 1
End Function
Function ReadFileCSV(ByVal filepath As String, _
                    ByRef lines() As String, _
                    ByRef log As String) As Integer
    '***********************************************************
    '
    'AUTHOR:            Le Thai (RBVH\EMB3)
    'DATE CREATED:      19/August/2013
    'FUNCTION NAME:     ReadFileCSV
    'DESCRIPTION:
    '       INPUT: Testcase files in CSV format
    '
    '       OUTPUT:
    '               lines        :   array of lines in CSV file
    '               log          :   log
    '
    '***********************************************************
    Dim writefile As Integer
    Dim i As Integer
    If (Dir(filepath, vbNormal) = vbNullString) Then
        log = "File not found: '" & filepath & "'"
        ReadFileCSV = -1
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
    
    ReadFileCSV = 1
End Function
Function GetSignals(ByVal line, _
                    ByRef signals() As String, _
                    ByRef log As String) As Integer
    '***********************************************************
    '
    'AUTHOR:            Le Thai (RBVH\EMB3)
    'DATE CREATED:      19/August/2013
    'FUNCTION NAME:     GetSignals
    'DESCRIPTION:
    '       INPUT:
    '               line                :   content of 1 line in csv file
    '       OUTPUT:
    '               signals             :   array of signals
    '
    '***********************************************************
    Dim i As Integer
    Dim semiPos, dotPos As Integer
    Dim tmpLine As String
    If (line = "") Then
        log = "GetSignals! Line is null!"
        GetSignals = -1
        Exit Function
    End If
    i = 0
    tmpLine = line
    Do While (Len(tmpLine) > 0)
        ReDim Preserve signals(0 To i)
        'Check ';' position for signal name
        semiPos = InStr(tmpLine, ";")
        If (semiPos > 0) Then
            'Get signal
            signals(i) = Mid(tmpLine, 1, semiPos - 1)
            'Remove characters "'" at head and tail
            If (InStr(signals(i), "'") > 0) Then
                signals(i) = Mid(signals(i), 2)
                signals(i) = Mid(signals(i), 1, Len(signals(i)) - 1)
                'remove TM, keep signal name only
                dotPos = InStr(signals(i), ".")
                If (dotPos > 0) Then
                    signals(i) = Mid(signals(i), 1, dotPos - 1)
                End If
            End If
            tmpLine = Mid(tmpLine, semiPos + 1)
        Else
            signals(i) = tmpLine
            'Remove characters "'" at head and tail
            If (InStr(signals(i), "'") > 0) Then
                signals(i) = Mid(signals(i), 2)
                signals(i) = Mid(signals(i), 1, Len(signals(i)) - 1)
                'remove TM, keep signal name only
                dotPos = InStr(signals(i), ".")
                If (dotPos > 0) Then
                    signals(i) = Mid(signals(i), 1, dotPos - 1)
                End If
            End If
            tmpLine = ""
        End If
        i = i + 1
    Loop
    'DEBUG
    'For i = 0 To UBound(signals())
    '    MsgBox signals(i)
    'Next i
    GetSignals = 1
End Function
Sub opendfiles()

Dim myfile As Variant

Dim counter As Integer

Dim path As String

Dim myfolder As String

myfolder = "D:\temp\"

ChDir myfolder

myfile = Application.GetOpenFilename("Signal xml Files (*.csv), *.csv", , , , True)

counter = 1

If IsNumeric(myfile) = True Then

MsgBox "No files selected"

End If

While counter <= UBound(myfile)

path = myfile(counter)

MsgBox path

counter = counter + 1

Wend

End Sub

'***********************************************NOT USED***************************************
Function ReadCSV_old(ByRef inputs() As String, _
                ByRef inputValues() As String, _
                ByRef outputs() As String, _
                ByRef outputValues() As String, _
                ByRef localOutputs() As String, _
                ByRef localOutputValues() As String, _
                ByRef log As String) As Integer
    '***********************************************************
    '
    'AUTHOR:            Le Thai (RBVH\EMB3)
    'DATE CREATED:      19/August/2013
    'FUNCTION NAME:     ReadCSV
    'DESCRIPTION:
    '       INPUT: Testcase files in CSV format
    '
    '       OUTPUT:
    '               inputs              :   array of input names
    '               inputValues         :   array of inputValues
    '               outputs             :   array of outputs names
    '               outputValues        :   array of outputValues
    '               localOutputs        :   array of localOutputs
    '               localOutputValues   :   array of localOutputValues
    '               log                 :   log
    '
    '***********************************************************
    Dim fileToOpen As String
    Dim lines() As String
    Dim writefile As Integer
    Dim i As Integer
    Dim signals() As String
    Dim signalValues() As String
    Dim nInputs As Integer
    Dim nOutputs As Integer
    Dim nLocalOutputs As Integer
    Dim nValues As Integer
    'Get file path
    fileToOpen = Application.GetOpenFilename("Signal xml Files (*.csv), *.csv")
    If fileToOpen = "False" Then
        MsgBox "File not found "
        ReadCSV = -1
        Exit Function
    End If
    'Initialize writefile
    writefile = FreeFile
    'Open file
    Open fileToOpen For Input As writefile
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
    'Get Signals
    If (GetSignals(lines(0), signals(), log) < 0) Then
        MsgBox log
        ReadCSV = -1
        Exit Function
    End If
     'Get Signal Values
    If (GetSignals(lines(3), signalValues(), log) < 0) Then
        MsgBox log
        ReadCSV = -1
        Exit Function
    End If
    'Classify signals
    nInputs = 0
    nOutputs = 0
    nLocalOutputs = 0
    For i = 0 To UBound(signals())
        If (InStr(signals(i), "expected") > 0) Then
            'outputs
            ReDim Preserve outputs(0 To nOutputs)
            outputs(nOutputs) = signals(i)
            'outputValues
            ReDim Preserve outputValues(0 To nOutputs)
            outputValues(nOutputs) = signalValues(i)
            'increase nOutputs
            nOutputs = nOutputs + 1
        ElseIf (InStr(signals(i), "exp_asp") > 0) Then
            'localOutputs
            ReDim Preserve localOutputs(0 To nLocalOutputs)
            localOutputs(nLocalOutputs) = signals(i)
            'localOutputValues
            ReDim Preserve localOutputValues(0 To nLocalOutputs)
            localOutputValues(nLocalOutputs) = signalValues(i)
            'increase nLocalOutputs
            nLocalOutputs = nLocalOutputs + 1
        Else
            'inputs
            ReDim Preserve inputs(0 To nInputs)
            inputs(nInputs) = signals(i)
            'inputValues
            ReDim Preserve inputValues(0 To nInputs)
            inputValues(nInputs) = signalValues(i)
            'increase nInputs
            nInputs = nInputs + 1
        End If
    Next i
    'DEBUG
    'CSVDisplay1D inputs(), 1
    'CSVDisplay1D inputValues(), 2
    'CSVDisplay1D outputs(), 4
    'CSVDisplay1D outputValues(), 5
    'CSVDisplay1D localOutputs(), 7
    'CSVDisplay1D localOutputValues(), 8
    
    ReadCSV = 1
End Function
