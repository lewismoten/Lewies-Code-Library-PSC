Imports System.IO

Public Class frmImageDetails
	Inherits System.Windows.Forms.Form
	Private _DefaultPath

#Region " Windows Form Designer generated code "

	Public Sub New(ByVal Path As String)
		MyBase.New()

		'This call is required by the Windows Form Designer.
		InitializeComponent()

		'Add any initialization after the InitializeComponent() call
		FileName = Path
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
	Friend WithEvents btnOk As System.Windows.Forms.Button
	Friend WithEvents btnCancel As System.Windows.Forms.Button
	Friend WithEvents txtPath As System.Windows.Forms.TextBox
	Friend WithEvents txtName As System.Windows.Forms.TextBox
	Friend WithEvents btnBrowse As System.Windows.Forms.Button
	Friend WithEvents pnlPictureScroll As System.Windows.Forms.Panel
	Friend WithEvents pbxPreview As System.Windows.Forms.PictureBox
	Friend WithEvents OpenFileDialog1 As System.Windows.Forms.OpenFileDialog
	Friend WithEvents Panel1 As System.Windows.Forms.Panel
	Friend WithEvents Panel2 As System.Windows.Forms.Panel
	Friend WithEvents lblDimensions As System.Windows.Forms.Label
	Friend WithEvents lblFileSize As System.Windows.Forms.Label
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Me.txtPath = New System.Windows.Forms.TextBox()
		Me.txtName = New System.Windows.Forms.TextBox()
		Me.btnBrowse = New System.Windows.Forms.Button()
		Me.btnOk = New System.Windows.Forms.Button()
		Me.btnCancel = New System.Windows.Forms.Button()
		Me.pnlPictureScroll = New System.Windows.Forms.Panel()
		Me.pbxPreview = New System.Windows.Forms.PictureBox()
		Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog()
		Me.Panel1 = New System.Windows.Forms.Panel()
		Me.Panel2 = New System.Windows.Forms.Panel()
		Me.lblDimensions = New System.Windows.Forms.Label()
		Me.lblFileSize = New System.Windows.Forms.Label()
		Me.pnlPictureScroll.SuspendLayout()
		Me.Panel1.SuspendLayout()
		Me.Panel2.SuspendLayout()
		Me.SuspendLayout()
		'
		'txtPath
		'
		Me.txtPath.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.txtPath.Dock = System.Windows.Forms.DockStyle.Top
		Me.txtPath.Location = New System.Drawing.Point(0, 20)
		Me.txtPath.Multiline = True
		Me.txtPath.Name = "txtPath"
		Me.txtPath.ReadOnly = True
		Me.txtPath.Size = New System.Drawing.Size(150, 30)
		Me.txtPath.TabIndex = 4
		Me.txtPath.Text = "c:\full\path\name.gif"
		'
		'txtName
		'
		Me.txtName.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.txtName.Dock = System.Windows.Forms.DockStyle.Top
		Me.txtName.Multiline = True
		Me.txtName.Name = "txtName"
		Me.txtName.ReadOnly = True
		Me.txtName.Size = New System.Drawing.Size(150, 20)
		Me.txtName.TabIndex = 3
		Me.txtName.Text = "FileName.gif"
		'
		'btnBrowse
		'
		Me.btnBrowse.Location = New System.Drawing.Point(20, 10)
		Me.btnBrowse.Name = "btnBrowse"
		Me.btnBrowse.TabIndex = 0
		Me.btnBrowse.Text = "&Browse ..."
		'
		'btnOk
		'
		Me.btnOk.Location = New System.Drawing.Point(110, 10)
		Me.btnOk.Name = "btnOk"
		Me.btnOk.TabIndex = 1
		Me.btnOk.Text = "&Ok"
		'
		'btnCancel
		'
		Me.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
		Me.btnCancel.Location = New System.Drawing.Point(200, 10)
		Me.btnCancel.Name = "btnCancel"
		Me.btnCancel.TabIndex = 2
		Me.btnCancel.Text = "&Cancel"
		'
		'pnlPictureScroll
		'
		Me.pnlPictureScroll.AutoScroll = True
		Me.pnlPictureScroll.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.pnlPictureScroll.Controls.AddRange(New System.Windows.Forms.Control() {Me.pbxPreview})
		Me.pnlPictureScroll.Location = New System.Drawing.Point(10, 10)
		Me.pnlPictureScroll.Name = "pnlPictureScroll"
		Me.pnlPictureScroll.Size = New System.Drawing.Size(130, 120)
		Me.pnlPictureScroll.TabIndex = 7
		'
		'pbxPreview
		'
		Me.pbxPreview.BackColor = System.Drawing.Color.White
		Me.pbxPreview.Name = "pbxPreview"
		Me.pbxPreview.Size = New System.Drawing.Size(60, 60)
		Me.pbxPreview.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize
		Me.pbxPreview.TabIndex = 1
		Me.pbxPreview.TabStop = False
		'
		'Panel1
		'
		Me.Panel1.Controls.AddRange(New System.Windows.Forms.Control() {Me.btnOk, Me.btnCancel, Me.btnBrowse})
		Me.Panel1.Dock = System.Windows.Forms.DockStyle.Bottom
		Me.Panel1.Location = New System.Drawing.Point(0, 145)
		Me.Panel1.Name = "Panel1"
		Me.Panel1.Size = New System.Drawing.Size(304, 40)
		Me.Panel1.TabIndex = 8
		'
		'Panel2
		'
		Me.Panel2.Controls.AddRange(New System.Windows.Forms.Control() {Me.lblFileSize, Me.lblDimensions, Me.txtPath, Me.txtName})
		Me.Panel2.Location = New System.Drawing.Point(150, 10)
		Me.Panel2.Name = "Panel2"
		Me.Panel2.Size = New System.Drawing.Size(150, 110)
		Me.Panel2.TabIndex = 9
		'
		'lblDimensions
		'
		Me.lblDimensions.Dock = System.Windows.Forms.DockStyle.Top
		Me.lblDimensions.Location = New System.Drawing.Point(0, 50)
		Me.lblDimensions.Name = "lblDimensions"
		Me.lblDimensions.Size = New System.Drawing.Size(150, 23)
		Me.lblDimensions.TabIndex = 7
		Me.lblDimensions.Text = "32x32 pixels"
		'
		'lblFileSize
		'
		Me.lblFileSize.Dock = System.Windows.Forms.DockStyle.Top
		Me.lblFileSize.Location = New System.Drawing.Point(0, 73)
		Me.lblFileSize.Name = "lblFileSize"
		Me.lblFileSize.Size = New System.Drawing.Size(150, 23)
		Me.lblFileSize.TabIndex = 8
		Me.lblFileSize.Text = "1024 bytes"
		'
		'frmImageDetails
		'
		Me.AcceptButton = Me.btnOk
		Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
		Me.CancelButton = Me.btnCancel
		Me.ClientSize = New System.Drawing.Size(304, 185)
		Me.ControlBox = False
		Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.Panel2, Me.Panel1, Me.pnlPictureScroll})
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
		Me.MaximizeBox = False
		Me.MinimizeBox = False
		Me.Name = "frmImageDetails"
		Me.ShowInTaskbar = False
		Me.Text = "Image Details"
		Me.pnlPictureScroll.ResumeLayout(False)
		Me.Panel1.ResumeLayout(False)
		Me.Panel2.ResumeLayout(False)
		Me.ResumeLayout(False)

	End Sub

#End Region
	Public Property DefaultPath() As String
		Get
			Return _DefaultPath
		End Get
		Set(ByVal Value As String)
			_DefaultPath = Value
		End Set
	End Property
	Public Property FileName() As String
		Get
			Return Me.txtPath.Text
		End Get
		Set(ByVal Value As String)
			If Value = "" Then
				Me.pbxPreview.Image = Nothing
				Me.lblDimensions.Text = "-1 x -1 pixels"
				Me.lblFileSize.Text = "0 bytes"
				Me.txtName.Text = "n/a"
				Me.txtPath.Text = ""
				Return
			End If
			Try
				If File.Exists(Value) Then
					Dim oImage As New Bitmap(Value)
					Me.pbxPreview.Image = oImage
					Me.lblDimensions.Text = oImage.Width.ToString & "x" & oImage.Height.ToString & " pixels"
					Me.lblFileSize.Text = FileSystem.FileLen(Value).ToString & " bytes"
					Me.txtName.Text = Path.GetFileName(Value)
					Me.txtPath.Text = Value
				Else
					Throw New Exception("File """ & Value & """ does not exist")
					Return
				End If
			Catch ex As Exception
				Throw New Exception("Error occured while acquiring image properties", ex)
			End Try
		End Set
	End Property

	Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
		Me.DialogResult = DialogResult.Cancel
		Me.Close()
	End Sub

	Private Sub btnOk_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOk.Click
		Me.DialogResult = DialogResult.OK
		Me.Close()
	End Sub

	Private Sub btnBrowse_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnBrowse.Click
		With Me.OpenFileDialog1
			.Title = "Choose an image"
			.InitialDirectory = Me.DefaultPath
			If FileName = "" Then
				.FileName = ""
			Else
				.FileName = Me.FileName
			End If
			.Filter = "Image Files (*.BMP;*.JPG;*.GIF;*.PNG)|*.BMP;*.JPG;*.GIF;*.PNG|All Files (*.*)|*.*"
			If .ShowDialog = DialogResult.Cancel Then Exit Sub
			Me.FileName = .FileName
			Me.DefaultPath = Path.GetDirectoryName(.FileName)
		End With
	End Sub
End Class
