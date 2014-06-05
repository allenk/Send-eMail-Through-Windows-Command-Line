Attribute VB_Name = "cmdMailModule"
'Command prompt \ Command line mailing executable by Stav Mann. ® Stavmann2@gmail.com
'Open-Source, you may use as you wish.
'Visual Basic 6.0

'Usage:
'Important: You can not just run this through the Visual Basic IDE, you must compile and use the Command-Line to pass parameters !

'To use this, start your Visual Studio IDE and load the .prj file \ emailFromCommandline.bas file
'If the mail account you wish to use to send the mail is not Gmail, make sure you change settings and credentials on the function.
'Compile to .exe
'
'Shell from vb \ from a command line using this syntax for your Gmail account (use your own credentials to test this if you want): (without the '<', '>')
'<File Path> user=<username> pass=<password> mail=Sendto@mail.com from=Sentfrom@mail.com subj=Subject body=This-Is-The-Body-of-the-letter (dont use spaces, you may type %20 instead of a space, and <br> instead of new line)

'Example:
'C:\cmdMail.exe user=myGmailUsername pass=myGmailPassword mail=stavmann2@gmail.com from=mail@mail.com subj=Hello-This-Is-A-Subject body=This%20Is%20The%20Mail%20Body.<BR><BR>Good-Bye.


Option Explicit

Dim msgA As Object 'declare the CDO

Public Sub Main()

'It's important to declare each variable as its type otherwise, it gets declared as Variant by default.
Dim tempStr As String, xUsername As String, xPassword As String, xMailTo As String, xFrom As String, xSubject As String, xMainText As String

'Declare the .exe parameters and split them using spaces.
Dim Parameters() As String
    Parameters() = Split(Command, " ")
    
'For each parameter splited
Dim i As Integer
    For i = 0 To UBound(Parameters)
        
        'read the first 5 characters of every parameter to figure out which parameter is it
        tempStr = Right(Parameters(i), (Len(Parameters(i)) - 5))
        
        'spread parameters to variables
        Select Case (Left(Parameters(i), 5))
        
            Case ("user="):
                xUsername = tempStr
            Case ("pass="):
                xPassword = tempStr
            Case ("mail="):
                xMailTo = tempStr
            Case ("subj="):
                xSubject = tempStr
            Case ("body="):
                xMainText = tempStr
            Case ("from="):
                xFrom = tempStr
            Case Else
                 
        End Select
        
    Next i

'if all went well, msgbox (Mail Sent), else msgbox Error (written in the function itself)
If mailSend(xUsername, _
             xPassword, _
             xMailTo, _
             xFrom, _
             xSubject, _
             xMainText _
             ) = 0 Then Call MsgBox("Mail Sent!", vbInformation)
       
End Sub



Private Function mailSend(xUsername, xPassword, xMailTo, xFrom, xSubject, xMainText) As Integer

Set msgA = CreateObject("CDO.Message") 'set the CDO to reffer as.
    
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
