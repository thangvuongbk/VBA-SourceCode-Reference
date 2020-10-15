Attribute VB_Name = "MSupportFunctions"
'Public Function Param2(Tbl As Range, _
'    dColVal As Double, dRowVal As Double) As Variant
'     ' Returns the bilinear interpolation of Tbl using dColVal and dRowVal
'
'     ' Tbl can be in any order by the first row,
'     '                     and by the first column
'
'    Dim nRow As Long, nCol As Long ' dimensions of table
'    Dim iRow1 As Long, iRow2 As Long ' bounding rows
'    Dim iCol1 As Long, iCol2 As Long ' bounding columns
'    Dim rf As Double, cf As Double ' row and column fractions, 0..1
'    Dim sL As Double, sR As Double '
'    Dim sT As Double, sB As Double '
'    Dim i As Integer, j As Integer
'    Dim MyTbl(), MyTbl_Sav()
'    Dim MinRow As Double, MaxRow As Double
'    Dim MinCol As Double, MaxCol As Double
'     ' four corner table values flanking dColVal, dRowVal
'    Dim sTL As Double, sTR As Double, sBR As Double, sBL As Double
'
'    nRow = Tbl.Rows.Count
'    nCol = Tbl.Columns.Count
'
''======================================================
'' value to be interpolated must lie within row and column headers
'    MinRow = IIf(Tbl(1, 2) < Tbl(1, nCol), Tbl(1, 2), Tbl(1, nCol))
'    MaxRow = IIf(Tbl(1, 2) > Tbl(1, nCol), Tbl(1, 2), Tbl(1, nCol))
'    MinCol = IIf(Tbl(2, 1) < Tbl(nRow, 1), Tbl(2, 1), Tbl(nRow, 1))
'    MaxCol = IIf(Tbl(2, 1) > Tbl(nRow, 1), Tbl(2, 1), Tbl(nRow, 1))
'
'
'      If dColVal < MinRow Then
'
'       dColVal = MinRow
'
'       ElseIf dColVal > MaxRow Then
'
'       dColVal = MaxRow
'
'       End If
'
'
'
'    MyTbl = Tbl
'    If (MyTbl(1, 2) > MyTbl(1, nCol)) Then
''----   Reverse column order
'        For j = 2 To nCol
'            For i = 1 To nRow
'                MyTbl(i, j) = Tbl(i, nCol - j + 2)
'            Next i
'        Next j
'    End If
'
'      If dRowVal < MinCol Then
'
'       dRowVal = MinCol
'
'       ElseIf dRowVal > Tbl(nRow, 1) Then
'
'       dRowVal = MaxCol
'
'       End If
'
'    If (MyTbl(2, 1) > MyTbl(nRow, 1)) Then
''----   Reverse Row order
'        MyTbl_Sav = MyTbl
'        For i = 2 To nRow
'            For j = 1 To nCol
'                MyTbl(i, j) = MyTbl_Sav(nRow - i + 2, j)
'            Next j
'        Next i
'    End If
''----    Prepare  interpolation
'    For i = 2 To nCol
'        If (dColVal < MyTbl(1, i)) Then Exit For
'    Next i
'    iCol1 = i - 1
'    sL = MyTbl(1, iCol1)
'    If dColVal = sL Then
'        iCol2 = iCol1
'        sR = sL
'    Else
'        iCol2 = iCol1 + 1
'       sR = MyTbl(1, iCol2)
'        cf = (dColVal - sL) / (sR - sL) ' column fraction
'    End If
'
'    For i = 2 To nRow
'        If (dRowVal < MyTbl(i, 1)) Then Exit For
'    Next i
'    iRow1 = i - 1
'
'    sT = MyTbl(iRow1, 1)
'    If dRowVal = sT Then
'        iRow2 = iRow1
'        sT = sB
'    Else
'        iRow2 = iRow1 + 1
'        sB = MyTbl(iRow2, 1)
'        rf = (dRowVal - sT) / (sB - sT)
'    End If
'
'    sTL = MyTbl(iRow1, iCol1)
'    sTR = MyTbl(iRow1, iCol2)
'    sBR = MyTbl(iRow2, iCol2)
'    sBL = MyTbl(iRow2, iCol1)
'
'     ' Compute the weighted  sum of four locations in MyTbl
'    Param2 = sTL * (1 - rf) * (1 - cf) _
'    + sTR * (1 - rf) * cf _
'    + sBR * rf * cf _
'    + sBL * rf * (1 - cf)
'End Function

Function Param1D(x As Double, Tbl As Range) As Variant
    '***********************************************************
    '
    'AUTHOR:            Le Thai (RBVH\EMB3)
    'DATE CREATED:      23/Dec/2013
    'FUNCTION NAME:     Param1D
    'DESCRIPTION: Calculation of y value for 1D table
    '       INPUT:
    '               x:   x value
    '               Tbl: Range
    '                   Ex:
    '                       A   B   C   D   E   F
    '                   1
    '                   2
    '                   3   x   1   3   9   25
    '                   4   y   0.1 0.9 1.2 3
    '                   5
    '               Tbl should be "A3:E4"
    '       OUTPUT: y value for 1D table
    '
    '***********************************************************
    Dim x0 As Double, x1 As Double
    Dim y0 As Double, y1 As Double
    Dim i As Integer, j As Integer
    Dim MyTbl()
    Dim nRows, nCols As Integer
    Dim iCol1 As Integer
    
    On Error Resume Next
    nRow = Tbl.Rows.count
    nCol = Tbl.Columns.count
    
    'The sequence of x must be from small number to big number
    'So reverse the table if the sequence is from big to small
    MyTbl = Tbl
    If (MyTbl(1, 2) > MyTbl(1, nCol)) Then
        For j = 2 To nCol
            For i = 1 To nRow
                MyTbl(i, j) = Tbl(i, nCol - j + 2)
            Next i
        Next j
    End If
    'Check limitation of x
    If x <= MyTbl(1, 2) Then
       Param1D = MyTbl(2, 2)
    ElseIf x >= MyTbl(1, nCol) Then
       Param1D = MyTbl(2, nCol)
    Else
        'Find the column number of x value
        For i = 2 To nCol
            If (x < MyTbl(1, i)) Then Exit For
        Next i
        iCol1 = i - 1
    
        'Calculation of return value
        x0 = MyTbl(1, iCol1)
        x1 = MyTbl(1, iCol1 + 1)
        y0 = MyTbl(2, iCol1)
        y1 = MyTbl(2, iCol1 + 1)
        Param1D = y0 + (x - x0) * (y1 - y0) / (x1 - x0)
    End If
  
End Function

Function Param2D(x As Double, y As Double, Tbl As Range) As Variant
    '***********************************************************
    '
    'AUTHOR:            Le Thai (RBVH\EMB3)
    'DATE CREATED:      13/Feb/2014
    'FUNCTION NAME:     Param1D
    'DESCRIPTION: Calculation of f(x,y) value for 2D table (bilinear interpolation)
    '       INPUT:
    '               x:   x value
    '               y:   y value
    '               Tbl: Range
    '                   Ex:
    '                       A     B   C   D   E   F
    '                   1
    '                   2
    '                   3   y,x   1   3   9   25
    '                   4   10   0.1 0.9 1.2 3.0
    '                   5   20   1.7 3.3 2.5 8.0
    '                   6
    '               Tbl should be "A3:E5"
    '       OUTPUT: f(x,y) value for 2D table
    '
    '***********************************************************
    Dim x0 As Double, x1 As Double
    Dim y0 As Double, y1 As Double
    Dim f00 As Double, f01 As Double
    Dim f10 As Double, f11 As Double
    Dim fR0 As Double, fR1 As Double
    Dim i As Integer, j As Integer
    Dim MyTbl()
    Dim MyTblTemp()
    Dim nRows, nCols As Integer
    Dim iCol0 As Integer, iCol1 As Integer
    Dim iRow0 As Integer, iRow1 As Integer
    
    On Error Resume Next
    nRow = Tbl.Rows.count
    nCol = Tbl.Columns.count
    
    'The sequence of x must be from small number to big number
    'So reverse the table if the sequence is from big to small
    MyTbl = Tbl
    If (MyTbl(1, 2) > MyTbl(1, nCol)) Then
        For j = 2 To nCol
            For i = 1 To nRow
                MyTbl(i, j) = Tbl(i, nCol - j + 2)
            Next i
        Next j
    End If
    'The sequence of y must be from small number to big number
    'So reverse the table if the sequence is from big to small
    MyTblTemp = MyTbl
    If (MyTbl(2, 1) > MyTbl(nRow, 1)) Then
        For i = 2 To nRow
            For j = 1 To nCol
                MyTbl(i, j) = MyTblTemp(nRow - i + 2, j)
            Next j
        Next i
    End If
    
    'Find the column number for x value
    'Check limitation of x
    If x <= MyTbl(1, 2) Then
       iCol0 = 2
       iCol1 = iCol0
    ElseIf x >= MyTbl(1, nCol) Then
       iCol0 = nCol
       iCol1 = iCol0
    Else
        For i = 2 To nCol
            If (x < MyTbl(1, i)) Then Exit For
        Next i
        iCol0 = i - 1
        iCol1 = iCol0 + 1
    End If
    
    'Find the row number for y value
    'Check limitation of y
    If y <= MyTbl(2, 1) Then
       iRow0 = 2
       iRow1 = iRow0
    ElseIf y >= MyTbl(nRow, 1) Then
       iRow0 = nRow
       iRow1 = iRow0
    Else
        'Find the column number for y value
        For i = 2 To nRow
            If (y < MyTbl(i, 1)) Then Exit For
        Next i
        iRow0 = i - 1
        iRow1 = iRow0 + 1
    End If
    
    'Calculation of return value
    x0 = MyTbl(1, iCol0)
    x1 = MyTbl(1, iCol1)
    y0 = MyTbl(iRow0, 1)
    y1 = MyTbl(iRow1, 1)
    f00 = MyTbl(iRow0, iCol0)
    f01 = MyTbl(iRow0, iCol1)
    f10 = MyTbl(iRow1, iCol0)
    f11 = MyTbl(iRow1, iCol1)
    'If x0 = x1 And y0 = y1 --> Return value directly
    If (x0 = x1 And y0 = y1) Then
        Param2D = f00
    'If x0 = x1 --> Calculate in x-direction only
    ElseIf (x0 = x1) Then
        Param2D = (y1 - y) / (y1 - y0) * f00 + (y - y0) / (y1 - y0) * f10
    'If y0 = y1 --> Calculate in y-direction only
    ElseIf (y0 = y1) Then
        Param2D = (x1 - x) / (x1 - x0) * f00 + (x - x0) / (x1 - x0) * f01
    Else
        'Calculate in x-direction
        fR0 = (y1 - y) / (y1 - y0) * f00 + (y - y0) / (y1 - y0) * f10
        fR1 = (y1 - y) / (y1 - y0) * f01 + (y - y0) / (y1 - y0) * f11
        'Calculate in y-direction
        Param2D = (x1 - x) / (x1 - x0) * fR0 + (x - x0) / (x1 - x0) * fR1
    End If

End Function

Function BITAND(in1 As Long, in2 As Long) As Long
    '***********************************************************
    '
    'AUTHOR:            Le Thai (RBVH\EMB3)
    'DATE CREATED:      27/Mar/2014
    'FUNCTION NAME:     BITAND
    'DESCRIPTION: bit and two argrument
    '       INPUT:
    '               in1, in2
    '       OUTPUT:
    '               BITAND(in1, in2) = in1 & in2
    '
    '***********************************************************
    BITAND = in1 And in2
End Function
Function BITOR(in1 As Long, in2 As Long) As Long
    '***********************************************************
    '
    'AUTHOR:            Le Thai (RBVH\EMB3)
    'DATE CREATED:      27/Mar/2014
    'FUNCTION NAME:     BITOR
    'DESCRIPTION: bit or two argrument
    '       INPUT:
    '               in1, in2
    '       OUTPUT:
    '               BITOR(in1, in2) = in1 | in2
    '
    '***********************************************************
    BITOR = in1 Or in2
End Function
Function BITSHIFTR(value As Long, shift As Byte) As Long
    '***********************************************************
    '
    'AUTHOR:            Le Thai (RBVH\EMB3)
    'DATE CREATED:      27/Mar/2014
    'FUNCTION NAME:     BITSHIFTR
    'DESCRIPTION: shift right
    '       INPUT:
    '               value, shift
    '       OUTPUT:
    '               BITSHIFTR(value, shift) = value >> shift
    '
    '***********************************************************
    BITSHIFTR = value
    If shift > 0 Then
        BITSHIFTR = Int(BITSHIFTR / (2 ^ shift))
    End If
End Function
Function BITSHIFTL(value As Long, shift As Long) As Long
    '***********************************************************
    '
    'AUTHOR:            Le Thai (RBVH\EMB3)
    'DATE CREATED:      27/Mar/2014
    'FUNCTION NAME:     BITSHIFTL
    'DESCRIPTION: shift left
    '       INPUT:
    '               value, shift
    '       OUTPUT:
    '               BITSHIFTL(value, shift) = value << shift
    '
    '***********************************************************
    BITSHIFTL = value
    If shift > 0 Then
        BITSHIFTL = BITSHIFTL * (2 ^ shift)
    End If
End Function





