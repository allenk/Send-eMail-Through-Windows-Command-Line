Attribute VB_Name = "cmdMailModule"
'Command prompt \ Command line mailing executable by Stav Mann. ® Stavmann2@gmail.com
'Open-Source, you may use as you wish.
'Visual Basic 6.0

'Usage:
'Important: You can not just run this through the Visual Basic IDE, you must compile and use the Command-Line to pass parameters !

'To use this, start your Visual Studio IDE and load the .vbp file \ emailFromCommandline.bas file
'If the mail account you wish to use to send the mail is not Gmail, make sure you change settings and credentials on the function.
'Compile to .exe
'
'Shell from vb \ from a command line using this syntax for your Gmail account (use your own credentials to test this if you want):
'<File Path> user=USERNAME pass=PASSWORD mail=Sendto@mail.com from=Sentfrom@mail.com subj=Subject body=This Is The Body of the letter

'P.S HTML tags work flawlessly here, so if you wish to make a new line of text, just type in a <BR> tag.

'Example:
'C:\cmdMail.exe user=myGmailUsername pass=myGmailPassword mail=stavmann2@gmail.com from=mail@mail.com subj=Hello This-Is A Subject body=This Is The Mail Body.<BR><BR>Good-Bye :)


Option Explicit

Private Const cmdUSER As String = "user="       'SMTP Username
Private Const cmdPASS As String = "pass="       'SMTP Password
Private Const cmdMAIL As String = "mail="       'Targeted eMail address (Must have legit email address template (mail@domain.com) )
Private Const cmdFROM As String = "from="       '"Replay To" address    (Must have legit email address template (mail@domain.com) )
Private Const cmdSUBJ As String = "subj="       'eMail Subject
Private Const cmdBODY As String = "body="       'eMail Body
Private Const cmdEND  As String = "=END="       'eMail Body

Public Sub Main()

'The idea is to simply grab the parameters, and split them to text strings, and then implement them straight to the mailing function.
'if went well, Msgbox (Mail Sent), Else Msgbox Error (written in the mailing function itself)

If mailSend(Trim(GetBetween(cmdUSER, cmdPASS)), _
            Trim(GetBetween(cmdPASS, cmdMAIL)), _
            Trim(GetBetween(cmdMAIL, cmdFROM)), _
            GetBetween(cmdFROM, cmdSUBJ), _
            GetBetween(cmdSUBJ, cmdBODY), _
            GetBetween(cmdBODY, cmdEND) _
            ) = 0 Then Call MsgBox("Mail Sent!", vbInformation)
       
End Sub



Private Function mailSend(xUsername, xPassword, xMailTo, xFrom, xSubject, xMainText) As Integer

Dim msgA As Object 'declare the CDO
Set msgA = CreateObject("CDO.Message") 'set the CDO to reffer as CDO.Message (microsoft default object that can be found on almost all windows versions since vista by default)
    
    msgA.To = xMailTo 'get targeted mail from command
    msgA.Subject = xSubject 'get subject from command
    msgA.HTMLBody = xMainText 'Main Text - You may use HTML tags here, for example <BR> to immitate "VBCRLF" (start new line) etc.
    msgA.From = xFrom 'The from part, make sure its syntax template is a valid mail one, user@domain.com, or something.
    
    'Notice, i simplified it, however, you may use more values depending on your needs, such as:
    '.Bcc = "mail@mail.com" ' - BCC..
    '.Cc = "mail@mail.com" ' - CC..
    '.CreateMHTMLBody ("www.mywebsite.com/index.html) 'send an entire webpage from a site
    '.CreateMHTMLBody ("c:\program files\download.htm) 'Send an entire webpage from your PC
    '.AddAttachment ("c:\myfile.zip") 'Send a file from your pc (notice uploading may take a while depending on your connection)

    
    'Gmail Username (from which mail will be sent)
    msgA.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = xUsername
    'Gmail Password
    msgA.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = xPassword
    
    'Mail Server address.
    msgA.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "smtp.gmail.com"
    
    'To set SMTP over the network = 2
    'To set Local SMTP = 1
    msgA.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
    
    'Type of Authenthication
    '0 - None
    '1 - Base 64 encoded (Normal)
    '2 - NTLM
    msgA.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
    
    'Outgoing Port
    msgA.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
    
    'Send using SSL True\False
    msgA.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = True
    
    'Update values of the SMTP configuration
    msgA.Configuration.Fields.Update
    
    'Send it.
    msgA.Send
    
    mailSend = Err.Number
        If Err.Number <> 0 Then Call MsgBox("Mail delivery failed: " & Err.Description, vbExclamation)
 
End Function


Private Function GetBetween(strOne As String, strTwo As String) As String

'Grab parameters as a whole, and place the line of text on strBody, in addition to the END-OF-PARAMETERS Flag called cmdEnd.
Dim strBody As String
    strBody = Command$ & cmdEND

'Locate each word's location within strBody, if its not found, don't continue.
Dim lngLocationOne As Long
Dim lngLocationTwo As Long
    
lngLocationOne = InStr(1, strBody, strOne, vbTextCompare)
    If (lngLocationOne = 0) Then GoTo ErrHandle
    
lngLocationTwo = InStr(1, strBody, strTwo, vbTextCompare)
    If (lngLocationTwo = 0) Then GoTo ErrHandle

'Grab a parameter value and return it.
GetBetween = Mid(strBody, lngLocationOne + Len(strOne), (lngLocationTwo - lngLocationOne - Len(strOne)))
        
Exit Function
ErrHandle:
    GetBetween = vbNullString

End Function

