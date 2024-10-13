Imports MessengerAPI
Imports Messenger.MSTATE
Imports VMS.Framework
Imports System.Runtime.InteropServices
Imports System.Web.HttpUtility

Public Class Msim
	Inherits VMS.Framework.Services.BaseService
	<MarshalAs(UnmanagedType.BStr)> Private Header As String
	<MarshalAs(UnmanagedType.BStr)> Private Message As String
	Private WithEvents InstantMessenger As Messenger.MsgrObject
	Private Sent As Boolean

	Public Overrides Function SendAlert(ByVal alert As Config.Alerts.Service.Alert) As Boolean

		Dim EntryData As EntryData

		' get entries from data provider
		EntryData = Me.DataProvider.FilterEntries(alert)

		' if no entries, return
		If EntryData.Entries.Count = 0 Then Return False

		' prepare message
		Message = Compose(alert.Message, EntryData.Entries)

		' enforce limit of 400 characters to prevent any problems.
		If Message.Length > 400 Then Message = Message.Substring(0, 400)

		InstantMessenger = New Messenger.MsgrObject
		Dim User As Messenger.IMsgrUser

		BuildHeader(alert)

		Try

			' Login to messenger service
			Dim LogOff As Boolean = Me.Logon()

			' Loop through recipients
			For Each Recipient As Config.Alerts.Service.Alert.Recipient In alert.Recipients

				' Locate user on contact list
				User = GetUser(Recipient.Address)

				' Send message to user
				SendMessage(User)

				User = Nothing

			Next

			' Logout of messenger service
			If LogOff Then Me.Logoff()

		Finally
			InstantMessenger = Nothing
		End Try

		Return True

	End Function

	Private Sub BuildHeader(ByVal alert As Config.Alerts.Service.Alert)

		Dim Builder As New Text.StringBuilder

		Builder.Append("MIME-Version: 1.0" & vbCrLf)
		Builder.Append("Content-Type: text/plain; charset=UTF-8" & vbCrLf)

		If Not alert("X-MMS-IM-Format") = "" Then
			Builder.Append("X-MMS-IM-Format: " & alert("X-MMS-IM-Format") & vbCrLf)
		Else

			' Font Name
			If Not alert("FontName") = "" Then
				Builder.Append("FN=" & UrlEncode(alert("FontName")) & "; ")
			Else
				Builder.Append("FN=" & UrlEncode("MS Sans Serif") & "; ")
			End If

			' Effects
			Builder.Append("EF=")
			If alert("Bold").ToLower = "true" Then Builder.Append("B")
			If alert("Italic").ToLower = "true" Then Builder.Append("I")
			Builder.Append("; ")

			' Color
			If Not alert("Color").Length = 6 Then
				Builder.Append("CO=0000ff; ")
			Else
				' Reverse colors to simulate html colors
				Dim r As String = alert("color").Substring(0, 2)
				Dim g As String = alert("color").Substring(2, 2)
				Dim b As String = alert("color").Substring(4, 2)
				Builder.Append("CO=" & b & g & r & "; ")
			End If

			' Character Set
			Builder.Append("CS=")
			Select Case alert("CharacterSet").ToLower
				Case "ansi" : Builder.Append("0")
				Case "default" : Builder.Append("1")
				Case "symbol" : Builder.Append("2")
				Case "mac" : Builder.Append("4d")
				Case "shiftjis" : Builder.Append("80")
				Case "hangeul" : Builder.Append("81")
				Case "johab" : Builder.Append("82")
				Case "gb2312" : Builder.Append("86")
				Case "chinesebig5" : Builder.Append("88")
				Case "greek" : Builder.Append("a1")
				Case "turkish" : Builder.Append("a2")
				Case "vietnamese" : Builder.Append("a3")
				Case "hebrew" : Builder.Append("b1")
				Case "arabic" : Builder.Append("b2")
				Case "baltic" : Builder.Append("ba")
				Case "russian" : Builder.Append("cc")
				Case "thai" : Builder.Append("de")
				Case "easteurope" : Builder.Append("ee")
				Case "oem" : Builder.Append("ff")
				Case Else : Builder.Append("1")
			End Select
			Builder.Append("; ")

			' Pitch / Font
			Builder.Append("PF=")
			Select Case alert("FontFamily")
				Case "dontcare" : Builder.Append("0")
				Case "roman" : Builder.Append("1")
				Case "swiss" : Builder.Append("2")
				Case "modern" : Builder.Append("3")
				Case "script" : Builder.Append("4")
				Case "decorative" : Builder.Append("5")
				Case Else : Builder.Append("0")
			End Select

			Select Case alert("FontPitch").ToLower
				Case "default" : Builder.Append("0")
				Case "fixed" : Builder.Append("1")
				Case "variable" : Builder.Append("2")
				Case Else : Builder.Append("0")
			End Select

			If alert("Align").ToLower = "right" Then
				Builder.Append("; RL=1")
			End If
			Builder.Append(vbCrLf)
		End If

		Builder.Append(vbCrLf)

		Header = Builder.ToString
	End Sub
	Private Function GetUser(ByVal emailAddress As String) As Messenger.IMsgrUser

		' find user in contact list
		For Each User As Messenger.IMsgrUser In InstantMessenger.List(Messenger.MLIST.MLIST_CONTACT)
			If User.EmailAddress.ToLower = emailAddress.ToLower Then Return User
		Next

		' couldn't find user - create them.
		Return InstantMessenger.CreateUser(emailAddress, InstantMessenger.Services.PrimaryService)

	End Function

	Private Sub SendMessage(ByVal user As Messenger.IMsgrUser)

		Sent = False

		user.SendText(Header, Message, Messenger.MMSGTYPE.MMSGTYPE_ALL_RESULTS)

		Dim Started As DateTime = Now

		' While we dont' have confirmation, let's wait it out.
		While Not Sent
			System.Threading.Thread.CurrentThread.Sleep(10)
			' Check to see if this process is taking too long.
			If Now.Subtract(Started).Seconds = 5 Then Exit While
		End While
	End Sub

	Private Function Logon() As Boolean
		Dim LogOff As Boolean = False
		If InstantMessenger.Services.PrimaryService.Status = Messenger.MSVCSTATUS.MSS_NOT_LOGGED_ON Then
			InstantMessenger.Logon(Service("Username"), Service("Password"), InstantMessenger.Services.PrimaryService)
			LogOff = True
		End If
		While Not InstantMessenger.Services.PrimaryService.Status = Messenger.MSVCSTATUS.MSS_LOGGED_ON
			System.Threading.Thread.CurrentThread.Sleep(10)
		End While
		Return LogOff
	End Function

	Private Sub Logoff()
		InstantMessenger.Logoff()
		While Not InstantMessenger.Services.PrimaryService.Status = Messenger.MSVCSTATUS.MSS_NOT_LOGGED_ON
			System.Threading.Thread.CurrentThread.Sleep(10)
		End While
	End Sub

	Private Sub InstantMessenger_OnSendResult(ByVal hr As Integer, ByVal lCookie As Integer) Handles InstantMessenger.OnSendResult
		Sent = True
	End Sub
End Class
