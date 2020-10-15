VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCSV 
   Caption         =   "Generate CSV"
   ClientHeight    =   3210
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4920
   OleObjectBlob   =   "frmCSV.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmCSV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_OK_Click()
    'Check Output Path
    If Me.txtPath.value = "" Then
        MsgBox "Please enter Output Path.", vbExclamation, "Generate CSV"
        Me.txtPath.SetFocus
        Exit Sub
    ElseIf (Dir(Me.txtPath.value, vbDirectory) = "") Then
        MsgBox "Output Path is incorrect!" & vbNewLine & _
                "Please enter output path again!" & vbNewLine & _
                "Hint: CSV folder may not exist!"
        Me.txtPath.SetFocus
        Exit Sub
    End If
    'Check TM Name
    If Me.txtTMName.value = "" Then
            MsgBox "Please enter TM Name.", vbExclamation, "Generate CSV"
            Me.txtTMName.SetFocus
            Exit Sub
    End If
    'Check Module Index
    If Not IsNumeric(Me.txtModuleIndex.value) Then
        MsgBox "The Module Index must contain a number.", vbExclamation, "Generate CSV"
        Me.txtModuleIndex.SetFocus
        Exit Sub
    End If
    'Check Constant Set
    If Me.txtConstantSet.value = "" Then
        MsgBox "Please enter Constant Set.", vbExclamation, "Generate CSV"
        Me.txtConstantSet.SetFocus
        Exit Sub
    'Check whether Constant Set exists in Constants sheet
    ElseIf (FindRowAll("Constants", 1, Me.txtConstantSet.value, 1) < 0) Then
        MsgBox "Unable to find constant set '" & Me.txtConstantSet.value & "' in Constants sheet" & vbNewLine & _
                "Hint: Constant Set name should exist in column 1 of Constants sheet", vbExclamation, "Generat CSV"
        Me.txtConstantSet.SetFocus
        Exit Sub
    End If
    Unload Me
End Sub

Private Sub cmdCancel_Click()
    ' Clear the form
    For Each ctl In Me.Controls
        If TypeName(ctl) = "TextBox" Or TypeName(ctl) = "ComboBox" Then
            ctl.value = ""
        ElseIf TypeName(ctl) = "CheckBox" Then
            ctl.value = False
        End If
    Next ctl
    Unload Me
End Sub
