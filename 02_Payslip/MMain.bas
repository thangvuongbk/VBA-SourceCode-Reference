Option Explicit
Const OK = 1
Const WARNING = 0
Const ERROR = -1

Const PAYSLIP_SHEET = "Payslip"
Const MAILER_SHEET = "Mailer"
Const PAYROLL_SHEET = "1.Payroll"
Const THIRD_PATH = "3rd\pdftk.exe"
Const OUTPUT_FOLDER = "Protected_Payslip\"
Const START_EMP_COUNT = 12
Const START_EMP_SENDMAIL = 2
' Type define
Type OutputReport
    noTotalRow As Integer
    noValidExported As Integer
    noValidNotExported As Integer
    noOther As Integer
End Type

Sub ExportToPDFOne()
    '***********************************************************
    '
    'AUTHOR:             Thang Vuong (thangvuongbk@gmail.com)
    'DATE CREATED:       12/Oct/2020
    'FUNCTION NAME:      ExportToPDFOne
    'DESCRIPTION:
    '
    '       OUTPUT: Export the current payslip based on Payslip sheet with password provided
    '
    '***********************************************************
    Dim m_inPDFFileName As String
    Dim m_outPDFFileName As String
    Dim m_pwd As String
    Dim m_currentPath As String
    Dim m_nameEmp As String
    Dim m_empCode As Variant
    Dim m_empRowNum As Variant
    Dim m_pdfGenToolPath As String
    Dim cmdStrGenPDF As String
        
    On Error GoTo ErrHandler
    If MsgBox("You are going to generate Payslip for this employee, Are you sure?", vbYesNo, "Confirm") = vbYes Then
    
        m_nameEmp = Sheets(PAYSLIP_SHEET).Range("C7").Value2
        Application.StatusBar = "Generating Payslip for " & m_nameEmp & " , Please Wait ..."
        'Application.StatusBar = "Generating a Payslip, Please Wait ..."
        Application.ScreenUpdating = False
        Application.DisplayAlerts = False
        
        
        m_currentPath = Application.ThisWorkbook.path ' get the current path of this workbook
        m_currentPath = IIf(Right(m_currentPath, 1) = "\", m_currentPath, m_currentPath & "\")  'Ensure "\" at the end.
                
        m_empCode = Sheets(PAYSLIP_SHEET).Range("C7").Value2 ' Getting the employee code at Cell "C7" in current sheet Payslip"/Need to take care this point
        If (m_empCode = "") Then
            MsgBox "ERROR! ExportToPDFOne! Employee Code at Cell C7 of sheet Payslip is empty, Please check and update"
            Exit Sub
        End If
        m_empRowNum = Application.Match(m_empCode, Sheets(MAILER_SHEET).Columns("C"), 0)
        If IsError(m_empRowNum) Then
            MsgBox "ERROR! ExportToPDFOne! Not able to find the " & m_empCode & " at column C of sheet " & MAILER_SHEET
            Exit Sub
        End If
        
        m_inPDFFileName = m_currentPath & Sheets(MAILER_SHEET).Range("K" & m_empRowNum).Value2 & ".pdf"       ' set the filename of input file path
        If (m_inPDFFileName = Empty) Then
            MsgBox "ERROR! ExportToPDFOne! The FILENAME is EMPTY for staff code  " & m_empCode & " at sheet " & MAILER_SHEET & " and Cell " & "K" & m_empRowNum
            Exit Sub ' for print all, we exit for instead
        End If
        
        m_pwd = Sheets(MAILER_SHEET).Range("J" & m_empRowNum).Value2
        If (m_pwd = Empty) Then
            MsgBox "ERROR! ExportToPDFOne! The PASSWORD is EMPTY for staff code  " & m_empCode & " at sheet " & MAILER_SHEET & " and Cell " & "K" & m_empRowNum
            Exit Sub ' for print all, we exit for instead
        End If
        
        ' Generate the pdf file
        Worksheets(PAYSLIP_SHEET).Activate
        With ActiveSheet
            .ExportAsFixedFormat Type:=xlTypePDF, _
                                              Filename:=m_inPDFFileName, _
                                              quality:=xlQualityStandard
        End With
        
        ' Including the password
        ' If Directory does not exist, Create Directory under Payslips in the name of Month & Year
        m_outPDFFileName = m_currentPath & OUTPUT_FOLDER
        createADirectory (m_outPDFFileName)
       
        If Len(Dir(m_inPDFFileName)) = 0 Then
            MsgBox "ERROR! ExportToPDFOne! " & m_inPDFFileName & " File does not exist"
            Exit Sub
        End If
        
        m_pdfGenToolPath = m_currentPath & THIRD_PATH
        m_outPDFFileName = m_outPDFFileName & Sheets(MAILER_SHEET).Range("K" & m_empRowNum).Value2 & ".pdf"
        'm_inPDFFileName = """" & m_inPDFFileName & """"
        'm_outPDFFileName = """" & m_outPDFFileName & """"
        'm_pwd = """" & m_pwd & """"
    
        cmdStrGenPDF = m_pdfGenToolPath & " " & m_inPDFFileName _
                                                                    & " Output " & m_outPDFFileName _
                                                                     & " User_pw " & m_pwd _
                                                                     & " Allow AllFeatures"
        
        Shell cmdStrGenPDF, vbHide
        
        Application.Wait DateAdd("s", 5, Now)
        
        Kill Replace(m_inPDFFileName, """", "")
        
        Application.ScreenUpdating = True
        Application.DisplayAlerts = True
        
        Application.StatusBar = "Protected Salary Slip Generated."
        DoEvents
        If MsgBox("The current payslip has been created at: " & vbNewLine & m_currentPath & OUTPUT_FOLDER & vbNewLine & vbNewLine & "Do you want to open?", vbYesNo, "Confirm") = vbYes Then
            Shell "explorer.exe" & " " & m_currentPath & OUTPUT_FOLDER, vbNormalFocus
        End If
    End If ' for message confirmation
    
ErrHandler:
   If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
End Sub
    
Sub ExportToPDFAll()
    '***********************************************************
    '
    'AUTHOR:             Thang Vuong (thangvuongbk@gmail.com)
    'DATE CREATED:       12/Oct/2020
    'FUNCTION NAME:      ExportToPDFAll
    'DESCRIPTION:
    '
    '       OUTPUT: Export all the payslip based on Payslip sheet with password provided
    '
    '***********************************************************
    Dim m_inPDFFileName As String
    Dim m_outPDFFileName As String
    Dim m_pwd As String
    Dim m_currentPath As String
    Dim m_nameEmp As String
    Dim m_empCode As Variant
    Dim m_empRowNum As Variant
    Dim m_pdfGenToolPath As String
    Dim cmdStrGenPDF As String
    Dim m_empIndex As Integer
    Dim m_listValidNotExported() As Integer

    Dim m_outputReport As OutputReport
    
    If MsgBox("You are going to generate ALL the Payslips, Are you sure?", vbYesNo, "Confirm") = vbYes Then
    'On Error GoTo ErrHandler
    Application.StatusBar = "Generating Payslip for all Employees, Please Wait ..."
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    m_currentPath = Application.ThisWorkbook.path ' get the current path of this workbook
    m_currentPath = IIf(Right(m_currentPath, 1) = "\", m_currentPath, m_currentPath & "\")  'Ensure "\" at the end.
            
    ' Loop to end of the employees need to be generated pdf file
    ' m_empIndex should run from row number = 12
    For m_empIndex = START_EMP_COUNT To Sheets(PAYROLL_SHEET).Cells(Rows.Count, "C").End(xlUp).Row
        m_outputReport.noTotalRow = Sheets(PAYROLL_SHEET).Cells(Rows.Count, "C").End(xlUp).Row - START_EMP_COUNT + 1
        Application.StatusBar = "Progress: " & (m_empIndex - START_EMP_COUNT + 1) & "/" & (Sheets(PAYROLL_SHEET).Cells(Rows.Count, "C").End(xlUp).Row - START_EMP_COUNT + 1) & ", Please Wait ..."
        ' in case the value is not numeric
        If Not IsNumeric(Sheets(PAYROLL_SHEET).Range("C" & m_empIndex).Value2) Then
                'MsgBox "ERROR! ExportToPDFAll! Employee Code at Cell C7 of sheet Payslip is empty, Please check and update"
                m_outputReport.noOther = m_outputReport.noOther + 1
                GoTo MoveToNextEmployee
        End If
        
        Sheets(PAYSLIP_SHEET).Range("C7") = Sheets(PAYROLL_SHEET).Range("C" & m_empIndex).Value2
        m_empCode = Sheets(PAYSLIP_SHEET).Range("C7").Value2 ' Getting the employee code at Cell "C7" in current sheet Payslip"/Need to take care this point, this may be change in the future
        If (m_empCode = "") Then
                'MsgBox "ERROR! ExportToPDFAll! Employee Code at Cell C7 of sheet Payslip is empty, Please check and update"
                GoTo MoveToNextEmployee
        End If
         
         m_empRowNum = Application.Match(m_empCode, Sheets(MAILER_SHEET).Columns("C"), 0)
         If IsError(m_empRowNum) Then
             'MsgBox "ERROR! ExportToPDFAll! Can not to find the " & m_empCode & " at column C of sheet " & MAILER_SHEET
             m_outputReport.noValidNotExported = m_outputReport.noValidNotExported + 1
             Sheets(PAYROLL_SHEET).Range("C" & m_empIndex).Font.Color = vbRed
             GoTo MoveToNextEmployee
         End If
         
         m_inPDFFileName = m_currentPath & Sheets(MAILER_SHEET).Range("K" & m_empRowNum).Value2 & ".pdf"       ' set the filename of input file path
         If (m_inPDFFileName = Empty) Then
             'MsgBox "ERROR! ExportToPDFAll! The FILENAME is EMPTY for staff code  " & m_empCode & " at sheet " & MAILER_SHEET & " and Cell " & "K" & m_empRowNum
             m_outputReport.noValidNotExported = m_outputReport.noValidNotExported + 1
             Sheets(PAYROLL_SHEET).Range("C" & m_empIndex).Font.Color = vbRed
             Sheets(MAILER_SHEET).Range("K" & m_empRowNum).Font.Color = vbRed
             GoTo MoveToNextEmployee
         End If
         
         m_pwd = Sheets(MAILER_SHEET).Range("J" & m_empRowNum).Value2
         If (m_pwd = Empty) Then
             'MsgBox "ERROR! ExportToPDFAll! The PASSWORD is EMPTY for staff code  " & m_empCode & " at sheet " & MAILER_SHEET & " and Cell " & "K" & m_empRowNum
             m_outputReport.noValidNotExported = m_outputReport.noValidNotExported + 1
             Sheets(PAYROLL_SHEET).Range("C" & m_empIndex).Font.Color = vbRed
             Sheets(MAILER_SHEET).Range("J" & m_empRowNum).Font.Color = vbRed
             GoTo MoveToNextEmployee
         End If
         
         '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
         ' Generate the pdf file
         Worksheets(PAYSLIP_SHEET).Activate
         With ActiveSheet
             .ExportAsFixedFormat Type:=xlTypePDF, _
                                               Filename:=m_inPDFFileName, _
                                               quality:=xlQualityStandard
         End With
         
         ' Including the password
         ' If Directory does not exist, Create Directory under Payslips in the name of Month & Year
         m_outPDFFileName = m_currentPath & OUTPUT_FOLDER
         createADirectory (m_outPDFFileName)
        
         If Len(Dir(m_inPDFFileName)) = 0 Then
             'MsgBox "ERROR! ExportToPDFAll! " & m_inPDFFileName & " File does not exist"
             GoTo MoveToNextEmployee
         End If
         
         m_pdfGenToolPath = m_currentPath & THIRD_PATH
         m_outPDFFileName = m_outPDFFileName & Sheets(MAILER_SHEET).Range("K" & m_empRowNum).Value2 & ".pdf"
         
         cmdStrGenPDF = m_pdfGenToolPath & " " & m_inPDFFileName _
                                         & " Output " & m_outPDFFileName _
                                         & " User_pw " & m_pwd _
                                         & " Allow AllFeatures"
         
         Shell cmdStrGenPDF, vbHide
         
         Application.Wait DateAdd("s", 5, Now)
         
         Kill Replace(m_inPDFFileName, """", "")
         
         m_outputReport.noValidExported = m_outputReport.noValidExported + 1
'ErrHandler:
'   If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
MoveToNextEmployee:
    Next m_empIndex
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    Application.StatusBar = "Protected Salary Slip Generated."
    DoEvents
    If MsgBox("Report the result:" & vbNewLine & _
                    "1. Total number: " & m_outputReport.noTotalRow & vbNewLine & _
                    "2. No. valid emp are exported: " & m_outputReport.noValidExported & vbNewLine & _
                    "3. No. valid emp are not exported: " & m_outputReport.noValidNotExported & "!!!,  if this value is not 0, please check the red value at column 'C' in sheet" & PAYROLL_SHEET & vbNewLine & _
                    "4. No. others (not right staff code format): " & m_outputReport.noOther & vbNewLine & vbNewLine & _
                    "All Protected Salary Slip Generated at: " & vbNewLine & m_currentPath & OUTPUT_FOLDER & vbNewLine & vbNewLine & "Do you want to open?", vbYesNo, "Confirm") = vbYes Then
                    Shell "explorer.exe" & " " & m_currentPath & OUTPUT_FOLDER, vbNormalFocus
    End If
   End If
End Sub

Sub NextEmployee()
    '***********************************************************
    '
    'AUTHOR:             Thang Vuong (thangvuongbk@gmail.com)
    'DATE CREATED:       12/Oct/2020
    'FUNCTION NAME:      NextEmployee
    'DESCRIPTION:
    '
    '       OUTPUT: Move to the next employee
    '
    '***********************************************************
    Dim m_empRowNum As Variant
    Dim m_empCode As Variant
    m_empRowNum = START_EMP_COUNT - 1
    
    If (Sheets(PAYSLIP_SHEET).Range("C7").Value2 = Empty) Then
        MsgBox "ERROR! NextEmployee! Value at C7 of " & PAYSLIP_SHEET & " should not be empty. Automatic get the first employee"
        Sheets(PAYSLIP_SHEET).Range("C7") = Sheets(PAYROLL_SHEET).Range("C12").Value2
        Exit Sub
    End If
    m_empCode = Sheets(PAYSLIP_SHEET).Range("C7").Value2
     
    m_empRowNum = Application.Match(m_empCode, Sheets(PAYROLL_SHEET).Columns("C"), 0)
    If m_empRowNum = Sheets(PAYROLL_SHEET).Cells(Rows.Count, "C").End(xlUp).Row Then
        MsgBox "INFO! NextEmployee! End of Employee" & vbNewLine & "Back to the first employee"
        Sheets(PAYSLIP_SHEET).Range("C7") = Sheets(PAYROLL_SHEET).Range("C12").Value2
        Exit Sub
    End If
    
    If IsError(m_empRowNum) Then
        MsgBox "ERROR! NextEmployee! Not able to find the " & m_empCode & " at column C of sheet " & MAILER_SHEET & vbNewLine & "Back to the first employee"
        Sheets(PAYSLIP_SHEET).Range("C7") = Sheets(PAYROLL_SHEET).Range("C12").Value2
        Exit Sub
    End If
    ' update to the next one
    If Not IsNumeric(Sheets(PAYROLL_SHEET).Range("C" & m_empRowNum + 1).Value2) Then
        Sheets(PAYSLIP_SHEET).Range("C7") = Sheets(PAYROLL_SHEET).Range("C" & m_empRowNum + 2).Value2
    Else
        Sheets(PAYSLIP_SHEET).Range("C7") = Sheets(PAYROLL_SHEET).Range("C" & m_empRowNum + 1).Value2
    End If
            
End Sub

Sub PreviousEmployee()
    '***********************************************************
    '
    'AUTHOR:             Thang Vuong (thangvuongbk@gmail.com)
    'DATE CREATED:       12/Oct/2020
    'FUNCTION NAME:      PreviousEmployee
    'DESCRIPTION:
    '
    '       OUTPUT: Move to the previous employee
    '
    '***********************************************************
    Dim m_empRowNum As Variant
    Dim m_empCode As Variant
    m_empRowNum = START_EMP_COUNT + 1
    
    If (Sheets(PAYSLIP_SHEET).Range("C7").Value2 = Empty) Then
        MsgBox "ERROR! PreviousEmployee! Value at C7 of " & PAYSLIP_SHEET & " should not be empty. Automatic get the first employee"
        Sheets(PAYSLIP_SHEET).Range("C7") = Sheets(PAYROLL_SHEET).Range("C12").Value2
        Exit Sub
    End If
    m_empCode = Sheets(PAYSLIP_SHEET).Range("C7").Value2
     
    m_empRowNum = Application.Match(m_empCode, Sheets(PAYROLL_SHEET).Columns("C"), 0)
    
    If m_empRowNum = START_EMP_COUNT Then
        MsgBox "INFO! PreviousEmployee! This is the FIRST of Employee" & vbNewLine & "Move to the last employee"
        Sheets(PAYSLIP_SHEET).Range("C7") = Sheets(PAYROLL_SHEET).Range("C" & Sheets(PAYROLL_SHEET).Cells(Rows.Count, "C").End(xlUp).Row).Value2
        Exit Sub
    End If
        
    If IsError(m_empRowNum) Then
        MsgBox "ERROR! PreviousEmployee! Not able to find the " & m_empCode & " at column C of sheet " & MAILER_SHEET & vbNewLine & "Back to the first employee"
        Sheets(PAYSLIP_SHEET).Range("C7") = Sheets(PAYROLL_SHEET).Range("C12").Value2
        Exit Sub
    End If
    ' update to the next one
       If Not IsNumeric(Sheets(PAYROLL_SHEET).Range("C" & m_empRowNum - 1).Value2) Then
        Sheets(PAYSLIP_SHEET).Range("C7") = Sheets(PAYROLL_SHEET).Range("C" & m_empRowNum - 2).Value2
    Else
        Sheets(PAYSLIP_SHEET).Range("C7") = Sheets(PAYROLL_SHEET).Range("C" & m_empRowNum - 1).Value2
    End If
        
End Sub

Sub SendMailToAll()
    '***********************************************************
    '
    'AUTHOR:             Thang Vuong (thangvuongbk@gmail.com)
    'DATE CREATED:       12/Oct/2020
    'FUNCTION NAME:      PreviousEmployee
    'DESCRIPTION:
    '
    '       OUTPUT: Send the email to all the employees
    '
    '***********************************************************
    Dim objOutlook As Object
    Dim objMail As Object
    
    Dim rgnSubject As Range
    Dim rgnBody As Range
    Dim rgnSignature1 As Range
    Dim rgnSignature2 As Range
    '
    Dim rgnDear As String
    Dim rgnTo As String
    Dim rgnAttach As String ' Read the file and update it on the excel
    '
    Dim m_currentPath As String
    Dim m_outputFolder As String
    Dim m_empIndex As Integer
    Dim m_protectedPDFFileName As String
    Dim tgap As Integer
    
    If MsgBox("You are going to send email to ALL the employees, Are you sure?", vbYesNo, "Confirm") = vbYes Then
            On Error GoTo ErrHandler
            
            m_currentPath = Application.ThisWorkbook.path ' get the current path of this workbook
            m_currentPath = IIf(Right(m_currentPath, 1) = "\", m_currentPath, m_currentPath & "\")  'Ensure "\" at the end.
                    
            ' Start to send mail, starting from row 2
            For m_empIndex = START_EMP_SENDMAIL To Sheets(MAILER_SHEET).Cells(Rows.Count, "A").End(xlUp).Row
                ' and send the mail in case of no sending before or sent failure
                If Sheets(MAILER_SHEET).Range("M" & m_empIndex).Value2 <> "Sent SUCCESS" Then
                Set objOutlook = CreateObject("Outlook.Application")
                Set objMail = objOutlook.CreateItem(0)
                
                ' Common content of an email to be sent out
                Worksheets(MAILER_SHEET).Activate
                With ActiveSheet
                    Set rgnSubject = .Range("U1")
                    'Set rgnDear = .Range("U2")
                    Set rgnBody = .Range("U3")
                    Set rgnSignature1 = .Range("U4")
                    Set rgnSignature2 = .Range("U5")
                End With
                
                ' Checking the pdf file is available or not
                If Sheets(MAILER_SHEET).Range("K" & m_empIndex).Value2 = Empty Then
                    MsgBox "The pdf file name info at cell: K" & m_empIndex & " is empty, Please do a check" & vbNewLine & "Move to next person"
                    GoTo MoveToSendNext
                End If
                ' Check the file
                m_protectedPDFFileName = m_currentPath & OUTPUT_FOLDER & Sheets(MAILER_SHEET).Range("K" & m_empIndex).Value2 & ".pdf"
                If Len(Dir(m_protectedPDFFileName)) = 0 Then
                    Sheets(MAILER_SHEET).Range("L" & m_empIndex).Value2 = "File does not exist"
                    Sheets(MAILER_SHEET).Range("L" & m_empIndex).Font.Color = vbRed
                    Sheets(MAILER_SHEET).Range("M" & m_empIndex).Value2 = "Sent FAIL"
                    Sheets(MAILER_SHEET).Range("M" & m_empIndex).Interior.Color = vbRed
                    GoTo MoveToSendNext
                End If
                Sheets(MAILER_SHEET).Range("L" & m_empIndex).Value2 = "File available"
                Sheets(MAILER_SHEET).Range("L" & m_empIndex).Font.Color = RGB(0, 128, 0)
                rgnAttach = m_protectedPDFFileName
                ' Check the email
                If Sheets(MAILER_SHEET).Range("I" & m_empIndex).Value2 = Empty Then
                    Sheets(MAILER_SHEET).Range("M" & m_empIndex).Value2 = "Sent FAIL"
                    Sheets(MAILER_SHEET).Range("M" & m_empIndex).Interior.Color = vbRed
                    MsgBox "The mail info at cell: I" & m_empIndex & " is empty, Please do a check" & vbNewLine & "Move to next person"
                    GoTo MoveToSendNext
                End If
                rgnTo = Sheets(MAILER_SHEET).Range("I" & m_empIndex).Value2
                rgnDear = Sheets(MAILER_SHEET).Range("U2").Value2 & " " & Sheets(MAILER_SHEET).Range("D" & m_empIndex).Value2
                ' Sending the email
                
                    With objMail
                        .To = rgnTo
                        .Subject = rgnSubject.Value2
                        .htmlbody = "<br>" & rgnDear & "<br><br>" & rgnBody.Value2 & "<br><br>" & rgnSignature1.Value2 & "<br>" & rgnSignature2.Value2
                        .attachments.Add rgnAttach
                        .send
                    End With
                Sheets(MAILER_SHEET).Range("M" & m_empIndex).Value2 = "Sent SUCCESS"
                Sheets(MAILER_SHEET).Range("M" & m_empIndex).Interior.Color = vbGreen
                
    '            rgnDear = Empty
    '            rgnAttach = Empty
    '            rgnTo = Empty
    '
                ' Pause 1 minute
                tgap = tgap + 1
                If tgap = 200 Then
                    Application.StatusBar = "Paused for Time Delay.....  Please wait...."
                    Application.Wait (Now + #12:01:00 AM#)
                    tgap = 1
                End If
            End If ' Check sent SUCCESS
MoveToSendNext:
            Next m_empIndex
ErrHandler:
       If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
    End If  ' Message confirmation
    MsgBox "Send mail to all employees is done. Please check the status!!!"
End Sub

Function createADirectory(ByVal path As String)
   Dim fso As Object
   Set fso = CreateObject("scripting.filesystemobject")
   If Not fso.folderexists(path) Then
      fso.createfolder (path)
   End If
   createADirectory = OK
End Function

