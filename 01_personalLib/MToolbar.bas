Attribute VB_Name = "MToolbar"
Const COMMON_USERS = "lth1hc"
Sub AddNewToolBar()
    '*********************************************************
    '
    ' Macro created 06/Jun/2013 by Le Thai (RBVH\EMB3)
    '
    '
    '*********************************************************
     
    ' This procedure creates a new temporary toolbar.
    Dim ComBar, ComBar2 As CommandBar
    Dim ComBarContrl As CommandBarControl
    Dim username As String
    On Error GoTo ErrorHandler
     ' Create a new floating toolbar and make it visible.
    On Error Resume Next
     'Delete the toolbar if it already exists
    CommandBars("NewClassToolbar").Delete
    CommandBars("NewClassToolbar2").Delete
    
    username = Application.username
      
Set ComBar = CommandBars.Add(name:="NewClassToolbar", Position:= _
    msoBarTop, Temporary:=True)
    
    ComBar.Visible = True
      ' Create a button with text on the bar and set some properties.
Set ComBarContrl = ComBar.Controls.Add(Type:=msoControlButton)
    With ComBarContrl
        .Caption = "Clear"
        .FaceId = 1088
        .Style = msoButtonIconAndCaption
        .TooltipText = "Clear all current data"
         'the onaction line tells the button to run a certain macro
        .OnAction = "Clear"
    End With
      ' Create a button with text on the bar and set some properties.
Set ComBarContrl = ComBar.Controls.Add(Type:=msoControlButton)
    With ComBarContrl
        .Caption = "Add Requirement"
        .FaceId = 97
        .Style = msoButtonIconAndCaption
        .TooltipText = "Add Requirement"
         'the onaction line tells the button to run a certain macro
        .OnAction = "AddReq"
    End With
      ' Create a button with text on the bar and set some properties.
Set ComBarContrl = ComBar.Controls.Add(Type:=msoControlButton)
    With ComBarContrl
        .Caption = "Generate TC(s)"
        .FaceId = 99
        .Style = msoButtonIconAndCaption
        .TooltipText = "Generate testcases in Testcases sheet"
         'the onaction line tells the button to run a certain macro
        .OnAction = "GenTC"
    End With
      ' Create a button with text on the bar and set some properties.
'Set ComBarContrl = ComBar.Controls.Add(Type:=msoControlButton)
'    With ComBarContrl
'        .Caption = "Generate CSV"
'        .FaceId = 142
'        .Style = msoButtonIconAndCaption
'        .TooltipText = "Generate testcase files in csv format"
'         'the onaction line tells the button to run a certain macro
'        .OnAction = "GenCSV"
'    End With
'      ' Create a button with text on the bar and set some properties.
Set ComBarContrl = ComBar.Controls.Add(Type:=msoControlButton)
    With ComBarContrl
        .Caption = "Read CSV"
        .FaceId = 23
        .Style = msoButtonIconAndCaption
        .TooltipText = "Read testcase files in csv format"
         'the onaction line tells the button to run a certain macro
        .OnAction = "ReadTC"
    End With
      ' Create a button with text on the bar and set some properties.
Set ComBarContrl = ComBar.Controls.Add(Type:=msoControlButton)
    With ComBarContrl
        .Caption = "Backup"
        .FaceId = 81
        .Style = msoButtonIconAndCaption
        .TooltipText = "Backup MCDC sheet and Testcases sheet before runing the macro"
         'the onaction line tells the button to run a certain macro
        .OnAction = "Backup"
    End With
      ' Create a button with text on the bar and set some properties.
Set ComBarContrl = ComBar.Controls.Add(Type:=msoControlButton)
    With ComBarContrl
        .Caption = "Undo"
        .FaceId = 37
        .Style = msoButtonIconAndCaption
        .TooltipText = "Restore MCDC sheet and Testcases sheet before runing the macro"
         'the onaction line tells the button to run a certain macro
        .OnAction = "Restore"
    End With
    ' Create a button with text on the bar and set some properties.
Set ComBarContrl = ComBar.Controls.Add(Type:=msoControlButton)
    With ComBarContrl
        .Caption = "Autofit"
        .FaceId = 80
        .Style = msoButtonIconAndCaption
        .TooltipText = "Autofit Columns active sheet"
         'the onaction line tells the button to run a certain macro
        .OnAction = "AutofitCellsActivesheet"
    End With
        ' Create a button with text on the bar and set some properties.
Set ComBarContrl = ComBar.Controls.Add(Type:=msoControlButton)
    With ComBarContrl
        .Caption = "Max-Min"
        .FaceId = 732
        .Style = msoButtonIconAndCaption
        .TooltipText = "Insert MaxMin Formula"
         'the onaction line tells the button to run a certain macro
        .OnAction = "InsertMaxMinFormula"
    End With
        ' Create a button with text on the bar and set some properties.
Set ComBarContrl = ComBar.Controls.Add(Type:=msoControlButton)
    With ComBarContrl
        .Caption = "Fill Local Variables"
        .FaceId = 19
        .Style = msoButtonIconAndCaption
        .TooltipText = "GenTDSkeleton"
         'the onaction line tells the button to run a certain macro
        .OnAction = "GenTDSkeleton"
    End With
    If (InStr(COMMON_USERS, username) > 0) Then
                ' Create a button with text on the bar and set some properties.
        Set ComBarContrl = ComBar.Controls.Add(Type:=msoControlButton)
            With ComBarContrl
                .Caption = "Formula"
                .FaceId = 85
                .Style = msoButtonIconAndCaption
                .TooltipText = "Insert Expression Formula"
                 'the onaction line tells the button to run a certain macro
                .OnAction = "InsertExpressionFormula"
            End With
    End If
'********************** Toolbar 2 **********************************************************************
Set ComBar2 = CommandBars.Add(name:="NewClassToolbar2", Position:= _
    msoBarTop, Temporary:=True)
    
    ComBar2.Visible = False
     
      ' Create a button with text on the bar and set some properties.
'Set ComBarContrl = ComBar2.Controls.Add(Type:=msoControlButton)
'    With ComBarContrl
'        .Caption = "Add Requirement"
'        .FaceId = 97
'        .Style = msoButtonIconAndCaption
'        .TooltipText = "Add Requirement"
'         'the onaction line tells the button to run a certain macro
'        .OnAction = "AddReq"
'    End With
    
    
    Exit Sub
ErrorHandler:
    MsgBox "Error " & Err.Number & vbCr & Err.description
    Exit Sub
End Sub
Sub DeleteToolbar()
    '*********************************************************
    '
    ' Macro created 06/Jun/2013 by Le Thai (RBVH\EMB3)
    '
    '
    '*********************************************************
    On Error Resume Next
    CommandBars("NewClassToolbar").Delete
    CommandBars("NewClassToolbar2").Delete
End Sub




