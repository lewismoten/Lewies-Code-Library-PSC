'Imports Messenger
'Imports MessengerAddIns
Imports MessengerAPI
Imports System.Runtime.InteropServices
Imports Messenger.MSTATE

Public Class Form1
	Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

	Public Sub New()
		MyBase.New()

		'This call is required by the Windows Form Designer.
		InitializeComponent()

		'Add any initialization after the InitializeComponent() call

	End Sub

	'Form overrides dispose to clean up the component list.
	Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
		If disposing Then
			If Not (components Is Nothing) Then
				components.Dispose()
			End If
		End If
		MyBase.Dispose(disposing)
	End Sub

	'Required by the Windows Form Designer
	Private components As System.ComponentModel.IContainer

	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.  
	'Do not modify it using the code editor.
	Friend WithEvents btnSend As System.Windows.Forms.Button
	Friend WithEvents txtMessage As System.Windows.Forms.TextBox
	Friend WithEvents lblRecipient As System.Windows.Forms.Label
	Friend WithEvents btnLogon As System.Windows.Forms.Button
	Friend WithEvents txtUsername As System.Windows.Forms.TextBox
	Friend WithEvents txtPassword As System.Windows.Forms.TextBox
	Friend WithEvents btnLogoff As System.Windows.Forms.Button
	Friend WithEvents Label1 As System.Windows.Forms.Label
	Friend WithEvents lblPassword As System.Windows.Forms.Label
	Friend WithEvents cboRecipient As System.Windows.Forms.ComboBox
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Me.btnSend = New System.Windows.Forms.Button
		Me.txtMessage = New System.Windows.Forms.TextBox
		Me.lblRecipient = New System.Windows.Forms.Label
		Me.txtUsername = New System.Windows.Forms.TextBox
		Me.txtPassword = New System.Windows.Forms.TextBox
		Me.btnLogon = New System.Windows.Forms.Button
		Me.btnLogoff = New System.Windows.Forms.Button
		Me.Label1 = New System.Windows.Forms.Label
		Me.lblPassword = New System.Windows.Forms.Label
		Me.cboRecipient = New System.Windows.Forms.ComboBox
		Me.SuspendLayout()
		'
		'btnSend
		'
		Me.btnSend.Location = New System.Drawing.Point(180, 230)
		Me.btnSend.Name = "btnSend"
		Me.btnSend.TabIndex = 0
		Me.btnSend.Text = "Send"
		'
		'txtMessage
		'
		Me.txtMessage.Location = New System.Drawing.Point(10, 150)
		Me.txtMessage.MaxLength = 400
		Me.txtMessage.Multiline = True
		Me.txtMessage.Name = "txtMessage"
		Me.txtMessage.Size = New System.Drawing.Size(250, 70)
		Me.txtMessage.TabIndex = 1
		Me.txtMessage.Text = "Hello World."
		'
		'lblRecipient
		'
		Me.lblRecipient.Location = New System.Drawing.Point(10, 120)
		Me.lblRecipient.Name = "lblRecipient"
		Me.lblRecipient.TabIndex = 3
		Me.lblRecipient.Text = "Recipient:"
		'
		'txtUsername
		'
		Me.txtUsername.Location = New System.Drawing.Point(80, 10)
		Me.txtUsername.Name = "txtUsername"
		Me.txtUsername.TabIndex = 4
		Me.txtUsername.Text = ""
		'
		'txtPassword
		'
		Me.txtPassword.Location = New System.Drawing.Point(80, 40)
		Me.txtPassword.Name = "txtPassword"
		Me.txtPassword.PasswordChar = Microsoft.VisualBasic.ChrW(42)
		Me.txtPassword.TabIndex = 5
		Me.txtPassword.Text = ""
		'
		'btnLogon
		'
		Me.btnLogon.Location = New System.Drawing.Point(190, 10)
		Me.btnLogon.Name = "btnLogon"
		Me.btnLogon.TabIndex = 6
		Me.btnLogon.Text = "Logon"
		'
		'btnLogoff
		'
		Me.btnLogoff.Location = New System.Drawing.Point(190, 40)
		Me.btnLogoff.Name = "btnLogoff"
		Me.btnLogoff.TabIndex = 7
		Me.btnLogoff.Text = "Logoff"
		'
		'Label1
		'
		Me.Label1.Location = New System.Drawing.Point(10, 10)
		Me.Label1.Name = "Label1"
		Me.Label1.TabIndex = 8
		Me.Label1.Text = "Username:"
		'
		'lblPassword
		'
		Me.lblPassword.Location = New System.Drawing.Point(10, 40)
		Me.lblPassword.Name = "lblPassword"
		Me.lblPassword.TabIndex = 9
		Me.lblPassword.Text = "Password"
		'
		'cboRecipient
		'
		Me.cboRecipient.Location = New System.Drawing.Point(70, 120)
		Me.cboRecipient.Name = "cboRecipient"
		Me.cboRecipient.Size = New System.Drawing.Size(190, 21)
		Me.cboRecipient.TabIndex = 10
		'
		'Form1
		'
		Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
		Me.ClientSize = New System.Drawing.Size(272, 266)
		Me.Controls.Add(Me.cboRecipient)
		Me.Controls.Add(Me.btnLogoff)
		Me.Controls.Add(Me.btnLogon)
		Me.Controls.Add(Me.txtPassword)
		Me.Controls.Add(Me.txtUsername)
		Me.Controls.Add(Me.lblRecipient)
		Me.Controls.Add(Me.txtMessage)
		Me.Controls.Add(Me.btnSend)
		Me.Controls.Add(Me.lblPassword)
		Me.Controls.Add(Me.Label1)
		Me.Name = "Form1"
		Me.Text = "Windows Messenger Test"
		Me.ResumeLayout(False)

	End Sub
#End Region

	Friend WithEvents WindowsMessenger As New Messenger.MsgrObject
	<MarshalAs(UnmanagedType.BStr)> Private sHeader As String
	<MarshalAs(UnmanagedType.BStr)> Private sText As String


	Private Sub btnSend_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSend.Click

		Dim user As Messenger.IMsgrUser

		' See the following URL for formatting your own messages
		' http://www.hypothetic.org/docs/msn/client/plaintext.php
		sHeader = "MIME-Version: 1.0" & vbCrLf
		sHeader &= "Content-Type: text/plain; charset=UTF-8" & vbCrLf
		sHeader &= "X-MMS-IM-Format: FN=MS%20Sans%20Serif; EF=B; CO=ff; CS=0; PF = 22" & vbCrLf
		sHeader &= vbCrLf

		sText = Me.txtMessage.Text

		For Each user In WindowsMessenger.List(Messenger.MLIST.MLIST_CONTACT)
			If user.EmailAddress = Me.cboRecipient.Text Then
				Select Case user.State
					Case MSTATE_AWAY, MSTATE_BE_RIGHT_BACK, MSTATE_BUSY, _
					MSTATE_IDLE, MSTATE_INVISIBLE, MSTATE_ON_THE_PHONE, _
					MSTATE_ONLINE, MSTATE_OUT_TO_LUNCH
						user.SendText(sHeader, sText, Messenger.MMSGTYPE.MMSGTYPE_ALL_RESULTS)
				End Select
			End If
		Next

		Me.txtMessage.Text = ""

	End Sub

	Private Sub btnLogon_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLogon.Click
		If WindowsMessenger.Services.PrimaryService.Status = Messenger.MSVCSTATUS.MSS_NOT_LOGGED_ON Then
			WindowsMessenger.Logon(Me.txtUsername.Text, Me.txtPassword.Text, WindowsMessenger.Services.PrimaryService)
		End If
		While WindowsMessenger.Services.PrimaryService.Status = Messenger.MSVCSTATUS.MSS_LOGGING_ON
			System.Threading.Thread.CurrentThread.Sleep(10)
		End While
	End Sub

	Private Sub btnLogoff_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLogoff.Click
		If WindowsMessenger.Services.PrimaryService.Status = Messenger.MSVCSTATUS.MSS_LOGGED_ON Then
			WindowsMessenger.Logoff()
		End If
		While WindowsMessenger.Services.PrimaryService.Status = Messenger.MSVCSTATUS.MSS_LOGGING_OFF
			System.Threading.Thread.CurrentThread.Sleep(10)
		End While
	End Sub

	Private Sub SyncList()
		Me.cboRecipient.Items.Clear()
		Dim User As Messenger.IMsgrUser
		For Each User In WindowsMessenger.List(Messenger.MLIST.MLIST_CONTACT)
			Select Case User.State
				Case MSTATE_AWAY, MSTATE_BE_RIGHT_BACK, MSTATE_BUSY, _
				MSTATE_IDLE, MSTATE_INVISIBLE, MSTATE_ON_THE_PHONE, _
				MSTATE_ONLINE, MSTATE_OUT_TO_LUNCH
					Me.cboRecipient.Items.Add(User.EmailAddress)
			End Select
		Next
		If Not Me.cboRecipient.Items.Count = 0 Then
			Me.cboRecipient.SelectedIndex = 0
		End If
		Select Case WindowsMessenger.LocalState
			Case MSTATE_OFFLINE, MSTATE_LOCAL_CONNECTING_TO_SERVER, MSTATE_LOCAL_DISCONNECTING_FROM_SERVER, MSTATE_LOCAL_FINDING_SERVER, MSTATE_LOCAL_SYNCHRONIZING_WITH_SERVER
				Me.txtPassword.Enabled = True
				Me.txtUsername.Enabled = True
				Me.btnLogon.Enabled = True

				Me.btnLogoff.Enabled = False
				Me.cboRecipient.Enabled = False
				Me.txtMessage.Enabled = False
				Me.btnSend.Enabled = False
			Case Else
				Me.txtPassword.Enabled = False
				Me.txtUsername.Enabled = False
				Me.btnLogon.Enabled = False

				Me.btnLogoff.Enabled = True
				Me.cboRecipient.Enabled = True
				Me.txtMessage.Enabled = True
				Me.btnSend.Enabled = True

		End Select
	End Sub
#Region " Events to sync list "
	Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
		SyncList()
	End Sub

	Private Sub WindowsMessenger_OnBuddyPropertyChangeResult(ByVal hr As Integer, ByVal pUser As Messenger.IMsgrUser, ByVal ePropType As Messenger.MUSERPROPERTY, ByVal vPropVal As Object, ByVal pService As Messenger.IMsgrService) Handles WindowsMessenger.OnBuddyPropertyChangeResult
		SyncList()
	End Sub

	Private Sub WindowsMessenger_OnServiceLogoff(ByVal hr As Integer, ByVal pService As Messenger.IMsgrService) Handles WindowsMessenger.OnServiceLogoff
		SyncList()
	End Sub

	Private Sub WindowsMessenger_OnSessionStateChange(ByVal pIMSession As Messenger.IMsgrIMSession, ByVal sPrevState As Messenger.SSTATE) Handles WindowsMessenger.OnSessionStateChange
		SyncList()
	End Sub

	Private Sub WindowsMessenger_OnUserStateChanged(ByVal pUser As Messenger.IMsgrUser, ByVal mPrevState As Messenger.MSTATE, ByRef pfEnableDefault As Boolean) Handles WindowsMessenger.OnUserStateChanged
		SyncList()
	End Sub
	Private Sub WindowsMessenger_OnLogonResult(ByVal hr As Integer, ByVal pService As Messenger.IMsgrService) Handles WindowsMessenger.OnLogonResult
		SyncList()
	End Sub

	Private Sub WindowsMessenger_OnLocalStateChangeResult(ByVal hr As Integer, ByVal mLocalState As Messenger.MSTATE, ByVal pService As Messenger.IMsgrService) Handles WindowsMessenger.OnLocalStateChangeResult
		SyncList()
	End Sub

	Private Sub WindowsMessenger_OnLogoff() Handles WindowsMessenger.OnLogoff
		SyncList()
	End Sub

	Private Sub WindowsMessenger_OnListAddResult(ByVal hr As Integer, ByVal MLIST As Messenger.MLIST, ByVal pUser As Messenger.IMsgrUser) Handles WindowsMessenger.OnListAddResult
		SyncList()
	End Sub

	Private Sub WindowsMessenger_OnListRemoveResult(ByVal hr As Integer, ByVal MLIST As Messenger.MLIST, ByVal pUser As Messenger.IMsgrUser) Handles WindowsMessenger.OnListRemoveResult
		SyncList()
	End Sub
#End Region
End Class
