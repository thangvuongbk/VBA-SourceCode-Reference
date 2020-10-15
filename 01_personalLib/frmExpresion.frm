VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmExpresion 
   Caption         =   "Expression"
   ClientHeight    =   1320
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11640
   OleObjectBlob   =   "frmExpresion.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmExpresion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public cancelButtonClicked As Boolean
Private Sub cmdCancel_Click()
    cancelButtonClicked = True
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
Private Sub cmdOK_Click()
    cancelButtonClicked = False
    'Check File Path
    If Me.txtExpression.value = "" Then
        MsgBox "Please enter File Path.", vbExclamation, "Expression Fomular"
        Me.txtExpression.SetFocus
        Exit Sub
    End If
    Unload Me
End Sub

