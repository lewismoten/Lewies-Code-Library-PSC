Imports System.Data.SqlClient

Public Class Install
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
	Friend WithEvents btnSkip As System.Windows.Forms.Button
	Friend WithEvents btnInstall As System.Windows.Forms.Button
	Friend WithEvents lblServer As System.Windows.Forms.Label
	Friend WithEvents txtServer As System.Windows.Forms.TextBox
	Friend WithEvents lblCatalog As System.Windows.Forms.Label
	Friend WithEvents txtCatalog As System.Windows.Forms.TextBox
	Friend WithEvents chkTrusted As System.Windows.Forms.CheckBox
	Friend WithEvents lblUser As System.Windows.Forms.Label
	Friend WithEvents txtUser As System.Windows.Forms.TextBox
	Friend WithEvents lblPassword As System.Windows.Forms.Label
	Friend WithEvents txtPassword As System.Windows.Forms.TextBox
	Friend WithEvents btnTest As System.Windows.Forms.Button
	Friend WithEvents lblNotice As System.Windows.Forms.Label
	Friend WithEvents chkCreate As System.Windows.Forms.CheckBox
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Me.btnSkip = New System.Windows.Forms.Button
		Me.btnInstall = New System.Windows.Forms.Button
		Me.lblServer = New System.Windows.Forms.Label
		Me.txtServer = New System.Windows.Forms.TextBox
		Me.lblCatalog = New System.Windows.Forms.Label
		Me.txtCatalog = New System.Windows.Forms.TextBox
		Me.chkTrusted = New System.Windows.Forms.CheckBox
		Me.lblUser = New System.Windows.Forms.Label
		Me.txtUser = New System.Windows.Forms.TextBox
		Me.lblPassword = New System.Windows.Forms.Label
		Me.txtPassword = New System.Windows.Forms.TextBox
		Me.btnTest = New System.Windows.Forms.Button
		Me.lblNotice = New System.Windows.Forms.Label
		Me.chkCreate = New System.Windows.Forms.CheckBox
		Me.SuspendLayout()
		'
		'btnSkip
		'
		Me.btnSkip.DialogResult = System.Windows.Forms.DialogResult.Cancel
		Me.btnSkip.Location = New System.Drawing.Point(20, 230)
		Me.btnSkip.Name = "btnSkip"
		Me.btnSkip.Size = New System.Drawing.Size(60, 23)
		Me.btnSkip.TabIndex = 0
		Me.btnSkip.Text = "&Skip"
		'
		'btnInstall
		'
		Me.btnInstall.Location = New System.Drawing.Point(170, 230)
		Me.btnInstall.Name = "btnInstall"
		Me.btnInstall.Size = New System.Drawing.Size(100, 23)
		Me.btnInstall.TabIndex = 1
		Me.btnInstall.Text = "&Install Database"
		'
		'lblServer
		'
		Me.lblServer.AutoSize = True
		Me.lblServer.Location = New System.Drawing.Point(10, 90)
		Me.lblServer.Name = "lblServer"
		Me.lblServer.Size = New System.Drawing.Size(41, 16)
		Me.lblServer.TabIndex = 2
		Me.lblServer.Text = "Server:"
		'
		'txtServer
		'
		Me.txtServer.Location = New System.Drawing.Point(60, 90)
		Me.txtServer.Name = "txtServer"
		Me.txtServer.Size = New System.Drawing.Size(140, 20)
		Me.txtServer.TabIndex = 3
		Me.txtServer.Text = "(local)"
		'
		'lblCatalog
		'
		Me.lblCatalog.AutoSize = True
		Me.lblCatalog.Location = New System.Drawing.Point(10, 120)
		Me.lblCatalog.Name = "lblCatalog"
		Me.lblCatalog.Size = New System.Drawing.Size(46, 16)
		Me.lblCatalog.TabIndex = 4
		Me.lblCatalog.Text = "Catalog:"
		'
		'txtCatalog
		'
		Me.txtCatalog.Location = New System.Drawing.Point(60, 120)
		Me.txtCatalog.Name = "txtCatalog"
		Me.txtCatalog.Size = New System.Drawing.Size(140, 20)
		Me.txtCatalog.TabIndex = 5
		Me.txtCatalog.Text = "VMS"
		'
		'chkTrusted
		'
		Me.chkTrusted.Checked = True
		Me.chkTrusted.CheckState = System.Windows.Forms.CheckState.Checked
		Me.chkTrusted.Location = New System.Drawing.Point(40, 150)
		Me.chkTrusted.Name = "chkTrusted"
		Me.chkTrusted.Size = New System.Drawing.Size(130, 20)
		Me.chkTrusted.TabIndex = 6
		Me.chkTrusted.Text = "Trusted Connection"
		'
		'lblUser
		'
		Me.lblUser.AutoSize = True
		Me.lblUser.Location = New System.Drawing.Point(60, 170)
		Me.lblUser.Name = "lblUser"
		Me.lblUser.Size = New System.Drawing.Size(31, 16)
		Me.lblUser.TabIndex = 7
		Me.lblUser.Text = "User:"
		'
		'txtUser
		'
		Me.txtUser.Location = New System.Drawing.Point(60, 190)
		Me.txtUser.Name = "txtUser"
		Me.txtUser.Size = New System.Drawing.Size(90, 20)
		Me.txtUser.TabIndex = 8
		Me.txtUser.Text = ""
		'
		'lblPassword
		'
		Me.lblPassword.AutoSize = True
		Me.lblPassword.Location = New System.Drawing.Point(160, 170)
		Me.lblPassword.Name = "lblPassword"
		Me.lblPassword.Size = New System.Drawing.Size(57, 16)
		Me.lblPassword.TabIndex = 9
		Me.lblPassword.Text = "Password:"
		'
		'txtPassword
		'
		Me.txtPassword.Location = New System.Drawing.Point(160, 190)
		Me.txtPassword.Name = "txtPassword"
		Me.txtPassword.PasswordChar = Microsoft.VisualBasic.ChrW(42)
		Me.txtPassword.Size = New System.Drawing.Size(90, 20)
		Me.txtPassword.TabIndex = 10
		Me.txtPassword.Text = ""
		'
		'btnTest
		'
		Me.btnTest.Location = New System.Drawing.Point(90, 230)
		Me.btnTest.Name = "btnTest"
		Me.btnTest.Size = New System.Drawing.Size(70, 23)
		Me.btnTest.TabIndex = 11
		Me.btnTest.Text = "&Test"
		'
		'lblNotice
		'
		Me.lblNotice.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.lblNotice.Location = New System.Drawing.Point(10, 10)
		Me.lblNotice.Name = "lblNotice"
		Me.lblNotice.Size = New System.Drawing.Size(270, 60)
		Me.lblNotice.TabIndex = 12
		Me.lblNotice.Text = "The Sql Server data store requires a database in order to save and query data fro" & _
		"m.  You may skip this process in favor of another datastore, or if you have prev" & _
		"iously installed this database. "
		'
		'chkCreate
		'
		Me.chkCreate.Location = New System.Drawing.Point(210, 120)
		Me.chkCreate.Name = "chkCreate"
		Me.chkCreate.Size = New System.Drawing.Size(60, 20)
		Me.chkCreate.TabIndex = 13
		Me.chkCreate.Text = "Create"
		'
		'Install
		'
		Me.AcceptButton = Me.btnInstall
		Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
		Me.CancelButton = Me.btnSkip
		Me.ClientSize = New System.Drawing.Size(292, 266)
		Me.Controls.Add(Me.chkCreate)
		Me.Controls.Add(Me.lblNotice)
		Me.Controls.Add(Me.btnTest)
		Me.Controls.Add(Me.txtPassword)
		Me.Controls.Add(Me.lblPassword)
		Me.Controls.Add(Me.txtUser)
		Me.Controls.Add(Me.lblUser)
		Me.Controls.Add(Me.chkTrusted)
		Me.Controls.Add(Me.txtCatalog)
		Me.Controls.Add(Me.lblCatalog)
		Me.Controls.Add(Me.txtServer)
		Me.Controls.Add(Me.lblServer)
		Me.Controls.Add(Me.btnInstall)
		Me.Controls.Add(Me.btnSkip)
		Me.Name = "Install"
		Me.Text = "Sql Server Database Installation"
		Me.ResumeLayout(False)

	End Sub

#End Region

	Private Sub btnTest_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnTest.Click
		Dim Connection As New SqlConnection(ConnectionString(chkCreate.Checked))
		Try
			Connection.Open()
			MsgBox("Success!", MsgBoxStyle.Information, "Connection Test")
		Catch ex As Exception

			MsgBox("Connection Failed." & vbCrLf & ex.Message, MsgBoxStyle.Information, "Connection Test")
		Finally
			If Not Connection Is Nothing Then
				If Not Connection.State = ConnectionState.Closed Then
					Connection.Close()
				End If
				Connection.Dispose()
				Connection = Nothing
			End If
		End Try
	End Sub


	Private Sub btnSkip_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSkip.Click
		Me.Close()
	End Sub

	Private Sub btnInstall_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnInstall.Click

		Dim Sql As String = GetResource("Install.sql")
		If Sql = "" Then Return
		Dim SqlCommands() As String = Split(Sql, vbCrLf & "GO" & vbCrLf)


		Dim Connection As New SqlConnection(ConnectionString(chkCreate.Checked))
		Dim Command As New SqlCommand
		Command.CommandType = CommandType.Text
		Command.Connection = Connection

		Try
			Connection.Open()
			If chkCreate.Checked Then
				Command.CommandText = "CREATE DATABASE [" & txtCatalog.Text & "]"
				Command.ExecuteNonQuery()
				Connection.Close()
				Connection.ConnectionString = ConnectionString(False)
				Connection.Open()
			End If
			For Each SqlCommand As String In SqlCommands
				Command.CommandText = SqlCommand
				Command.ExecuteNonQuery()
			Next
			Me.Close()
		Catch ex As Exception
			MsgBox("An exception has occured." & vbCrLf & ex.Message, MsgBoxStyle.Critical, "Exception")
		Finally
			If Not Command Is Nothing Then
				Command.Dispose()
				Command = Nothing
			End If
			If Not Connection Is Nothing Then
				If Not Connection.State = ConnectionState.Closed Then
					Connection.Close()
				End If
				Connection.Dispose()
				Connection = Nothing
			End If
		End Try
	End Sub

	Public Function GetResource(ByVal filename As String) As String

		' Reads an embedded file from the resource of this assembly
		' and returns it as text

		Dim [Assembly] As Reflection.Assembly
		[Assembly] = MyClass.GetType.Assembly.GetExecutingAssembly
		' Stream is stored within root namespace.
		' each folder also appears in name delimited with dots.
		Dim Prefix As String = "VMS.DataStores."
		Dim Stream As IO.Stream
		Try
			Stream = [Assembly].GetManifestResourceStream(Prefix & filename)
			' convert stream to text
			Dim Buffer(Stream.Length - 1) As Byte
			Stream.Read(Buffer, 0, Buffer.Length)
			Stream.Close()
			Return System.Text.ASCIIEncoding.ASCII.GetString(Buffer)
		Finally
			If Not Stream Is Nothing Then
				Stream.Close()
				Stream = Nothing
			End If
		End Try
	End Function
	Private ReadOnly Property ConnectionString(ByVal Create As Boolean) As String
		Get
			' Create a connection string based on properties provided

			Dim Server As String = txtServer.Text.Replace("=", "==")
			Dim Catalog As String = txtCatalog.Text.Replace("=", "==")
			Dim Username As String = txtUser.Text.Replace("=", "==")
			Dim Password As String = txtPassword.Text.Replace("=", "==")
			Dim Trusted As Boolean = chkTrusted.Checked

			' Since we are to create a new database, we must log into
			' the master database
			If Create Then Catalog = "master"


			Dim Text As New System.Text.StringBuilder
			Text.Append("Persist Security Info=False;")
			If Trusted Then
				Text.Append("Integrated Security=SSPI;")
			Else
				Text.Append("Password=" & Password & ";")
				Text.Append("User ID=" & Username & ";")
			End If
			Text.Append("Initial Catalog=" & Catalog & ";")
			Text.Append("server=" & Server & ";")
			Text.Append("Application Name=VssMonitorService;")

			Return Text.ToString

		End Get
	End Property
End Class
