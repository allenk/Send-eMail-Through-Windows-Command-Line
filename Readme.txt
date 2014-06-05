Command Line Emailing using Windows CDO.Message

This will allow developers \ coders \ whoever, to send basic email using a shell \ shellexecute \ batch file.
You may also use this to send attachments, html pages, etc although it will require some additional coding 
(i left very detailed instruction and functions to make it as easy as possible to manipulate)

Command prompt \ Command line mailing executable by Stav Mann. � Stavmann2@gmail.com
Open-Source, you may use as you wish.
Visual Basic 6.0

Usage:
Important: You can not just run this through the Visual Basic IDE, you must compile and use the Command-Line to pass parameters !

To use this, start your Visual Studio IDE and load the .vbp file \ emailFromCommandline.bas file
If the mail account you wish to use to send the mail is not Gmail, make sure you change settings and credentials on the function.
Compile to .exe

Shell from vb \ from a command line using this syntax for your Gmail account (use your own credentials to test this if you want): (without the '<', '>')
<File Path> user=<username> pass=<password> mail=Sendto@mail.com from=Sentfrom@mail.com subj=Subject body=This-Is-The-Body-of-the-letter (dont use spaces, you may type %20 instead of a space, and <br> instead of new line)

Example:
C:\cmdMail.exe user=myGmailUsername pass=myGmailPassword mail=stavmann2@gmail.com from=mail@mail.com subj=Hello-This-Is-A-Subject body=This%20Is%20The%20Mail%20Body.<BR><BR>Good-Bye.


