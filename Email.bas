Attribute VB_Name = "Email"
Sub EmailMessage()

Application.ScreenUpdating = False

'Go/No-Go Option - Warning-------------------------------------------
Sheet6.Range("AB1").FormulaR1C1 = "=COUNTIF(Interface!R[6]C[-27]:R156C1,TRUE)"
Sheet6.Range("AB4").FormulaR1C1 = "=""Warning: Please close any dialogue boxes in outlook before continuing. You are about to send ""&R1C28&"" automated emails with the ""&R3C28&"" Notification template. Do you wish to continue?"""
 Application.CutCopyMode = False
ContinueForm.NumberTemplateLabel.Caption = Sheet6.Range("AB4").Value
ContinueForm.Show
If Sheet6.Range("AB5").Value = False Then
    MsgBox "You are now aborting sending the automated emails."
    Exit Sub
End If
'End of Warning -----------------------------------------------------

row_number = 6
Dim subject_name As String

subject_name = Sheet1.Range("C2").Value
numemailssent = 0
numemailsnotsent = 0
lr = Sheet3.Cells(Rows.Count, 1).End(xlUp).Row
i = lr + 5


Do
DoEvents
    row_number = row_number + 1
    Dim mail_body_message As String
    
    Dim resource_name As String
    Dim resource_manager As String
    
    Dim project_name As String
    Dim task_name As String
    
    Dim requested_by As String
    Dim project_owner As String
    
    Dim request_status As String
    Dim start_date As String
    Dim end_date As String
    
    Dim requested_on As String
    Dim hours_requested As String
    
    Application.ScreenUpdating = False
    
    'Assignming mail body message
    mail_body_message = Sheets("Message").Range("A1")
    Sheets("Interface").Activate
    
    'Assigning interface variables
    resource_name = Application.Range("B" & row_number)
    resource_manager = Application.Range("C" & row_number)
    
    project_name = Application.Range("D" & row_number)
    task_name = Application.Range("E" & row_number)
    
    requested_by = Application.Range("F" & row_number)
    project_owner = Application.Range("G" & row_number)
    
    request_status = Application.Range("J" & row_number)
    start_date = Application.Range("K" & row_number)
    end_date = Application.Range("L" & row_number)
    
    requested_on = Application.Range("M" & row_number)
    hours_requested = Application.Range("P" & row_number)
    
    email_add = Application.Range("H" & row_number)
    
    request_id = "RR" & Left(Range("D" & row_number), 5) & Format(Now(), "mmddyyhhmmss") & "-" & Left(Range("M" & row_number), 6)
              
    'Replacing interface variables per email into body message of the email
    mail_body_message = Replace(mail_body_message, "IIresource_name", resource_name)
    mail_body_message = Replace(mail_body_message, "IIresource_manager", resource_manager)
    
    mail_body_message = Replace(mail_body_message, "IIproject_name", project_name)
    mail_body_message = Replace(mail_body_message, "IItask_name", task_name)
    
    mail_body_message = Replace(mail_body_message, "IIrequested_by", requested_by)
    mail_body_message = Replace(mail_body_message, "IIproject_owner", project_owner)
    
    mail_body_message = Replace(mail_body_message, "IIrequest_status", request_status)
    mail_body_message = Replace(mail_body_message, "IIstart_date", start_date)
    mail_body_message = Replace(mail_body_message, "IIend_date", end_date)
    
    mail_body_message = Replace(mail_body_message, "IIrequested_on", requested_on)
    mail_body_message = Replace(mail_body_message, "IIhours_requested", hours_requested)
    
    mail_body_message = Replace(mail_body_message, "IIrequest_id", request_id)
    
    'MsgBox "EMAIL NOTIFICATION:" & Chr(10) & _
            "Row Number: " & row_number & Chr(10) & _
            resource_manager & ": " & email_add & Chr(10) & _
            "Project: " & project_name & ": " & task_name & Chr(10) & _
            "Requested By: " & requested_by & Chr(10) & _
            "Hours Requested: " & hours_requested & Chr(10) & _
            start_date & " - " & end_date
     
    'If checkbox is checked then send, otherwise skip (https://www.youtube.com/watch?v=dUSQ7wZHM7A)
    If Application.Range("A" & row_number) = True Then
        Call SendEmail(Application.Range("H" & row_number), subject_name & " - " & request_id, mail_body_message)
        'How do we know how many times a email was sent?
        Application.ScreenUpdating = True
        Application.Range("I" & row_number) = "Successfully Sent: " & Now
        Application.Range("I" & row_number).Font.Color = 5287936
        numemailssent = numemailssent + 1
        Application.ScreenUpdating = False
    Else
        Application.ScreenUpdating = True
        Application.Range("I" & row_number) = "Not Sent"
        Application.Range("I" & row_number).Font.Color = 255
        numemailsnotsent = numemailsnotsent + 1
        Application.ScreenUpdating = False
    End If
    
    Application.ScreenUpdating = False
    
Loop Until row_number = i
'Find last row and make that the loop until

    
    MsgBox "The sending procedure was successful!" & Chr(10) & _
           "Successfully Sent: " & numemailssent & " Emails" & Chr(10) & _
           "Not Sent: " & numemailsnotsent & " Emails"
            
    Application.ScreenUpdating = True
    
End Sub
Sub SendEmail(what_address As String, subject_line As String, mail_body As String)
    'Microsoft Outlook 16.0 Object Library
    'MO > File > Options > Trust Center > Trust Center Settings > Programming Access > Do Not Warn
    
    Dim olApp As Outlook.Application
    Set olApp = CreateObject("Outlook.Application")
    Dim olMail As Outlook.MailItem
    Dim myItem As Outlook.MailItem
    Set olMail = olApp.CreateItem(olMailItem)
    Set myItem = olApp.CreateItem(olMailItem)
    
    Dim oAccount As Outlook.Account
    Dim email_from As String
    email_from = Sheet1.Range("C3").Value
    
    Application.ScreenUpdating = False
    
    'Send from Email From: on Interface (https://www.slipstick.com/developer/send-using-default-or-specific-account/)
    For Each oAccount In olApp.Session.Accounts
        If oAccount = email_from Then
            Application.ScreenUpdating = False
            olMail.SendUsingAccount = oAccount
            olMail.Display
        End If
    Next
    
    'Below is the Email Attachment Option
    Dim FilePathAttachment As String
    FilePathAttachment = Sheet1.Range("C4").Value
    
    If FilePathAttachment <> "" Then
        Dim myAttachments As Outlook.Attachments
        Dim strAttachments As String
        Set myAttachments = olMail.Attachments
        strAttachments = FilePathAttachment
        olMail.Attachments.Add strAttachments, olByValue, 1
        myItem.Display
    End If
    
    Application.ScreenUpdating = False
    
    olMail.To = what_address
    olMail.Subject = subject_line
    olMail.BodyFormat = olFormatHTML
    olMail.HTMLBody = mail_body
    olMail.Send

End Sub
Sub Attachment_Button()
    'This script retrieves the file path for the attachment.
    Dim FileSelect As Variant
    Dim wb As Workbook
    
    Application.ScreenUpdating = False
    
    MsgBox "Please select the attachment you would like to attach to ALL emails below."
    
    'Locate the file path
    FileSelect = Application.GetOpenFilename(MultiSelect:=False)

    'Check if a file is selected
    If FileSelect = False Then
        MsgBox "You did not select a file. Please hit the 'Select...' button again to attach a file."
        Sheet1.Range("C4:E4").Value = ""
        Exit Sub
    End If
    
    'Send the path to the worksheet
    Sheet1.Range("C4:E4").Value = FileSelect
    
End Sub
