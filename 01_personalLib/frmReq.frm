VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmReq 
   Caption         =   "Add Requirement"
   ClientHeight    =   1560
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6870
   OleObjectBlob   =   "frmReq.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmReq"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public cancelButtonClicked As Boolean

Private Sub cmdBrowseFilePath_Click()
    Dim filepath As Variant
    Dim fso As New FileSystemObject

    If (Me.txtPath.value <> "" And Dir(Me.txtPath.value) <> "") Then
        ChDrive (fso.GetDriveName(Me.txtPath.value))
        ChDir (fso.GetParentFolderName(Me.txtPath.value))
    End If
    filepath = Application.GetOpenFilename _
                        (FileFilter:="All Fiels (*.*),*.*,C Source Files (*.c),*.c,Text Files (*.txt),*.txt", _
                        Title:="Open File(s)", MultiSelect:=False)
    If filepath <> False Then
        Me.txtPath.value = filepath
    Else
        Me.txtPath.value = ""
    End If
End Sub

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
    If Me.txtPath.value = "" Then
        MsgBox "Please enter File Path.", vbExclamation, "Add Requirement"
        Me.txtPath.SetFocus
        Exit Sub
    ElseIf (Dir(Me.txtPath.value, vbDirectory) = "") Then
        MsgBox "File Path is incorrect!" & vbNewLine & _
                "Please enter file path again!"
        Me.txtPath.SetFocus
        Exit Sub
    End If
    Unload Me
End Sub

Private Sub UserForm_Click()

End Sub
