Attribute VB_Name = "MSupport"
Option Explicit
Const OK = 1
Const ERROR = -1
Const MAX_ROW = 1000
Const MAX_COL = 255
' Convert to the PDF naming
Function ConvertPDFName(ByVal inString As Variant)
    Dim pattern() As String
    Dim convertString As String
    pattern = Split(" | ", "|")
    
    inString = StrConv(inString, vbLowerCase)
    inString = StrConv(inString, vbProperCase)
    ConvertPDFName = ReplaceAll(inString, pattern(), "")
End Function
' Generate the password based on DOB and ID
Function GenPwd(ByVal DOB As Variant, ByVal id As Variant)
    Dim m_formatDOB
    Dim m_customID
    m_formatDOB = Format(DOB, "ddmmmyyyy")
    m_customID = Mid(id, Len(id) - 2, 3)
    GenPwd = m_formatDOB & m_customID
End Function
Function ReplaceAll(ByVal inString As Variant, _
                        ByRef patterns() As String, _
                        ByVal repStr As String) As String
    '***********************************************************
    '
    'AUTHOR:            Thang Vuong
    'DATE CREATED:      13/Oct/2020
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



