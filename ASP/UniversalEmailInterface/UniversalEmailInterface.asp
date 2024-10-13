<%
    '**************************************
    ' Name: Universal Email Interface
    ' Description:Lays down a universal inte
    '     rface to sending email. Easier to port t
    '     o other servers that use different COM o
    '     bjects to send email. Even exposes a lis
    '     t of ProgID's that are installed on the 
    '     server. (This is the beginnig. But I am 
    '     sure you can see where the benefit of th
    '     is code is going)
    ' By: Lewis Moten
    '
    '
    ' Inputs:None
    '
    ' Returns:None
    '
    'Assumes:None
    '
    'Side Effects:None
    'This code is copyrighted and has limite
    '     d warranties.
    '**************************************
    
    Class clsEmail
    	
    	Public SenderName
    	Public SenderAddress
    	Public RecipientName
    	Public RecipientAddress
    	Public Subject
    	Public Message
    	Public Host
    	Public ProgID
    	Public LogonUsername
    	Public LogonPassword
    	
    	Private Sub Class_Initialize()
    		' Setup default values
    		SenderAddress	= Application.Value("Settings-Email-Address")
    		SenderName		= Application.Value("Settings-Email-Name")
    		Host			= Application.Value("Settings-Email-Server")
    		ProgID			= Application.Value("Settings-Email-ProgID")
    	End Sub
    	
    	Public Property Get Installed()
    		Dim llngIndex
    		Dim llngMaxIndex
    		Dim lstrProgIDs
    		Dim lstrProgIDAry
    		Dim lstrPairAry
    		Dim lstrInstalled
    		Dim lobjTest
    		
    		On Error Resume Next
    		
    		lstrProgIDs = _
    			"CDONTS.NewMail:Collaborative Data Objects for NT;" & _
    			"SMTPsvg.Mailer:Server Objects - ASPMail/ASPQMail"
    		'	"POPsvb.Mailer:Server Objects - ASP Pop3;" & _
    		'	"SoftArtisans.SMTPMail:Software Artisans - SMTP Mail;" & _
    		'	"Jmail.smtpmail:w3 JMail;" & _
    		'	"Persists.MailSender:Persists - ASPEmail;" & _
    		'	"dkQmail.Qmail:dkQmail;" & _
    		'	"Geocel.Mailer:GeoCel;" & _
    		'	"iismail.iismail.1:IISMail;" & _
    		'	"SmtpMail.SmtpMail.1:SMTP;" & _
    		'	"ocxQmail.ocxQmailCtrl.1:OCXQMail;" & _
    		'	"Dundas.Mailer:Dundas - ASPMailer;" & _
    		'	"EasyMail.SMTP.5:Quicksoft - EasyMail"
    		lstrProgIDAry = Split(lstrProgIDs, ";")
    		
    		llngMaxIndex = UBound(lstrProgIDAry)
    		
    		For llngIndex = 0 To llngMaxIndex
    			lstrPairAry = Split(lstrProgIDAry(llngIndex), ":")
    			Set lobjTest = Server.CreateObject(lstrPairAry(0))
    			If Err Then
    				Err.Clear
    			Else
    				lstrInstalled = lstrInstalled & lstrProgIDAry(llngIndex) & ";"
    			End If
    		Next
    		
    		If Right(lstrInstalled, 1) = ";" Then
    			lstrInstalled = Left(lstrInstalled, Len(lstrInstalled) - 1)
    		End If
    		
    		Installed = lstrInstalled
    		
    	End Property	
    	Public Sub Send()
    		
    		If ProgID = "" Then 
    			Call Err.Raise(vbObjectError + 1, "clsEmail.asp", "The ProgID has not been defined.")
    			Exit Sub
    		End If
    		
    		Dim lobjMailer
    		
    		Set lobjMailer = Server.CreateObject(ProgID)
    		
    		If SenderName = "" Then SenderName = SenderAddress
    		If RecipientName = "" Then RecipientName = RecipientAddress
    		
    		Select Case LCase(ProgID)
    			Case "cdonts.newmail"
    				lobjMailer.From		= SenderName & "<" & SenderAddress & ">"
    				lobjMailer.To		= RecipientName & "<" & RecipientAddress & ">"
    				lobjMailer.Subject	= Subject
    				lobjMailer.Body		= Message
    				lobjMailer.Send()
    			Case "smtpsvg.mailer"
    				Call lobjMailer.AddRecipient(RecipientName, RecipientAddress)
    				lobjMailer.RemoteHost	= Host
    				lobjMailer.FromAddress	= SenderAddress
    				lobjMailer.FromName		= SenderName
    				lobjMailer.Subject		= Subject
    				lobjMailer.BodyText		= Message
    				lobjMailer.SendMail
    			Case Else
    				Call Err.Raise(vbObjectError + 2, "clsEmail.asp", "The ProgID """ & ProgID & """ is not registered.")
    		End Select
    		
    		Set lobjMailer = Nothing
    	End Sub
    End Class
    ' --------------------------------------
    '     ----------------------------------------
    '     
    %>