Imports System.Web.Mail
Imports VMS.Framework

Public Class Smtp
	Inherits VMS.Framework.Services.BaseService

	Private Function Address(ByVal recipient As Config.Alerts.Service.Alert.Recipient) As String
		If recipient.Name = "" Then recipient.Name = recipient.Address
		Return """" & recipient.Name & """ <" & recipient.Address & ">"
	End Function
	Public Overrides Function SendAlert(ByVal alert As Config.Alerts.Service.Alert) As Boolean

		Dim EntryData As EntryData

		' get entries from data provider
		EntryData = Me.DataProvider.FilterEntries(alert)

		' if no entries, return
		If EntryData.Entries.Count = 0 Then Return False

		' prepare message

		Dim Message As New System.Web.Mail.MailMessage

		Message.From = Service("SenderName") & " <" & Service("SenderAddress") & ">"
		Message.Subject = alert.Message.Subject
		Message.Body = Compose(alert.Message, EntryData.Entries)

		For Each Recipient As Config.Alerts.Service.Alert.Recipient In alert.Recipients
			If Message.To = "" Then
				Message.To = Address(Recipient)
			Else
				If Message.Bcc = "" Then
					Message.Bcc = Address(Recipient)
				Else
					Message.Bcc &= "; " & Address(Recipient)
				End If
			End If
		Next

		' Perpare communications

		' if server wasn't specified
		If Me.Service("Server") = "" Then
			' use local smtp service
			System.Web.Mail.SmtpMail.SmtpServer = Environment.MachineName
		Else
			' send message via remote service
			System.Web.Mail.SmtpMail.SmtpServer = Service("Server")
		End If

		' TODO: Add properties for each alert for HTML format/Text format
		If alert("Format").ToLower = "html" Then
			Message.BodyFormat = MailFormat.Html
		Else
			Message.BodyFormat = MailFormat.Text
		End If

		' If it appears that we need to login to the smtp server
		If Not Me.Service("Username") = "" Then

			' Setup properties that allow us to login to
			' smtp service

			' CDONTS has a specified format that it uses in order
			' to log into external servers.
			Dim Schema As String = "http://schemas.microsoft.com/cdo/configuration/"
			Message.Fields(Schema & "smtpauthenticate") = 1
			Message.Fields(Schema & "sendusername") = Service("Username")
			Message.Fields(Schema & "sendpassword") = Service("Password")
		End If

		' Send message
		System.Web.Mail.SmtpMail.Send(Message)

		Return True

	End Function
End Class
