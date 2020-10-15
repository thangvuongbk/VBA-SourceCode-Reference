Attribute VB_Name = "MAttempt"
Option Explicit
Sub attempt()
    Dim log As String
    Dim table As CMCDCTable
    Dim ret As Integer
    Dim req As CRequirement
    Dim mcdc As CMCDC
    
    Set table = New CMCDCTable
    Set req = New CRequirement
    Set mcdc = New CMCDC
    
    'table.Value = "ABC"
    'MsgBox table.Value
    
    'ret = table.GetTCNoInfo("MCDC", 8, 2, log)
    'MsgBox UBound(table.GetTCNo())
    'ret = table.GetSignalDataAll("MCDC", 7, 4, 0, log)
    'MsgBox UBound(table.GetInputValues(), 1) & vbNewLine & _
            UBound(table.GetInputValues(), 2)
    'CSVDisplay table.GetInputValues(), 1
    
    'ret = req.SetReq("MCDC", 1, 1, log)
    'MsgBox req.GetName
    
    ret = mcdc.GetData("MCDC", 1, 1, log)
    For Each req In mcdc.GetReqs()
        'MsgBox "req.name: " & req.GetName & vbNewLine & _
                "req.GetDescription: " & req.GetDescription & vbNewLine & _
                "nInputs: " & UBound(req.GetTable.GetInputValues(), 1)
        MsgBox "req.name: " & req.GetName & vbNewLine & _
                "req.GetDescription: " & req.GetDescription & vbNewLine & _
                "nInputs: " & UBound(req.GetTable.GetInputValues(), 1) & vbNewLine & _
                "nValues: " & UBound(req.GetTable.GetInputValues(), 2)
    Next req
    
End Sub
