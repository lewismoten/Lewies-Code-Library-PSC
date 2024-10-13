Public Class Form1
	Inherits System.Windows.Forms.Form

	Protected WithEvents Display As AnsiFramework

	Dim Screen As New AnsiFramework
	Private map As AnsiCharacterMap

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
	Friend WithEvents lblCharacter As System.Windows.Forms.Label
	Friend WithEvents picCanvas As System.Windows.Forms.PictureBox
	Friend WithEvents MainMenu1 As System.Windows.Forms.MainMenu
	Friend WithEvents pnlOptions As System.Windows.Forms.Panel
	Friend WithEvents mnuFile As System.Windows.Forms.MenuItem
	Friend WithEvents mnuFile_Open As System.Windows.Forms.MenuItem
	Friend WithEvents OpenFileDialog1 As System.Windows.Forms.OpenFileDialog
	Friend WithEvents mnuFile_Save As System.Windows.Forms.MenuItem
	Friend WithEvents SaveFileDialog1 As System.Windows.Forms.SaveFileDialog
	Friend WithEvents mnuEdit_Clear As System.Windows.Forms.MenuItem
	Friend WithEvents ForegroundPanel As System.Windows.Forms.Panel
	Friend WithEvents BackgroundPanel As System.Windows.Forms.Panel
	Friend WithEvents MenuItem1 As System.Windows.Forms.MenuItem
	Friend WithEvents MenuItem2 As System.Windows.Forms.MenuItem
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Me.pnlOptions = New System.Windows.Forms.Panel
		Me.ForegroundPanel = New System.Windows.Forms.Panel
		Me.lblCharacter = New System.Windows.Forms.Label
		Me.BackgroundPanel = New System.Windows.Forms.Panel
		Me.picCanvas = New System.Windows.Forms.PictureBox
		Me.MainMenu1 = New System.Windows.Forms.MainMenu
		Me.mnuFile = New System.Windows.Forms.MenuItem
		Me.mnuFile_Open = New System.Windows.Forms.MenuItem
		Me.mnuFile_Save = New System.Windows.Forms.MenuItem
		Me.mnuEdit_Clear = New System.Windows.Forms.MenuItem
		Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog
		Me.SaveFileDialog1 = New System.Windows.Forms.SaveFileDialog
		Me.MenuItem1 = New System.Windows.Forms.MenuItem
		Me.MenuItem2 = New System.Windows.Forms.MenuItem
		Me.pnlOptions.SuspendLayout()
		Me.SuspendLayout()
		'
		'pnlOptions
		'
		Me.pnlOptions.BackColor = System.Drawing.SystemColors.Control
		Me.pnlOptions.Controls.Add(Me.ForegroundPanel)
		Me.pnlOptions.Controls.Add(Me.lblCharacter)
		Me.pnlOptions.Controls.Add(Me.BackgroundPanel)
		Me.pnlOptions.Dock = System.Windows.Forms.DockStyle.Bottom
		Me.pnlOptions.Location = New System.Drawing.Point(0, 395)
		Me.pnlOptions.Name = "pnlOptions"
		Me.pnlOptions.Size = New System.Drawing.Size(752, 70)
		Me.pnlOptions.TabIndex = 0
		'
		'ForegroundPanel
		'
		Me.ForegroundPanel.BackColor = System.Drawing.Color.White
		Me.ForegroundPanel.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me.ForegroundPanel.Location = New System.Drawing.Point(70, 10)
		Me.ForegroundPanel.Name = "ForegroundPanel"
		Me.ForegroundPanel.Size = New System.Drawing.Size(30, 30)
		Me.ForegroundPanel.TabIndex = 21
		'
		'lblCharacter
		'
		Me.lblCharacter.BackColor = System.Drawing.Color.Black
		Me.lblCharacter.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me.lblCharacter.Font = New System.Drawing.Font("Courier New", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblCharacter.ForeColor = System.Drawing.Color.White
		Me.lblCharacter.Location = New System.Drawing.Point(10, 10)
		Me.lblCharacter.Name = "lblCharacter"
		Me.lblCharacter.Size = New System.Drawing.Size(50, 50)
		Me.lblCharacter.TabIndex = 20
		Me.lblCharacter.Text = "@"
		Me.lblCharacter.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		'
		'BackgroundPanel
		'
		Me.BackgroundPanel.BackColor = System.Drawing.Color.Black
		Me.BackgroundPanel.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me.BackgroundPanel.Location = New System.Drawing.Point(90, 30)
		Me.BackgroundPanel.Name = "BackgroundPanel"
		Me.BackgroundPanel.Size = New System.Drawing.Size(30, 30)
		Me.BackgroundPanel.TabIndex = 22
		'
		'picCanvas
		'
		Me.picCanvas.Dock = System.Windows.Forms.DockStyle.Fill
		Me.picCanvas.Location = New System.Drawing.Point(0, 0)
		Me.picCanvas.Name = "picCanvas"
		Me.picCanvas.Size = New System.Drawing.Size(752, 395)
		Me.picCanvas.TabIndex = 1
		Me.picCanvas.TabStop = False
		'
		'MainMenu1
		'
		Me.MainMenu1.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuFile})
		'
		'mnuFile
		'
		Me.mnuFile.Index = 0
		Me.mnuFile.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuEdit_Clear, Me.mnuFile_Open, Me.mnuFile_Save, Me.MenuItem1, Me.MenuItem2})
		Me.mnuFile.Text = "File"
		'
		'mnuFile_Open
		'
		Me.mnuFile_Open.Index = 1
		Me.mnuFile_Open.Text = "Open ..."
		'
		'mnuFile_Save
		'
		Me.mnuFile_Save.Index = 2
		Me.mnuFile_Save.Text = "Save ..."
		'
		'mnuEdit_Clear
		'
		Me.mnuEdit_Clear.Index = 0
		Me.mnuEdit_Clear.Text = "Clear"
		'
		'MenuItem1
		'
		Me.MenuItem1.Index = 3
		Me.MenuItem1.Text = "www.lewismoten.com"
		'
		'MenuItem2
		'
		Me.MenuItem2.Index = 4
		Me.MenuItem2.Text = "Exit"
		'
		'Form1
		'
		Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
		Me.BackColor = System.Drawing.Color.Black
		Me.ClientSize = New System.Drawing.Size(752, 465)
		Me.Controls.Add(Me.picCanvas)
		Me.Controls.Add(Me.pnlOptions)
		Me.Menu = Me.MainMenu1
		Me.Name = "Form1"
		Me.Text = "Ascii Painter"
		Me.pnlOptions.ResumeLayout(False)
		Me.ResumeLayout(False)

	End Sub

#End Region

	Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
		Dim frm As New Splash
		frm.ShowDialog(Me)
		frm.Dispose()
		Display = New AnsiFramework
		Me.lblCharacter.Font = AnsiCharacterMap.RegularFont
		Display.EraseDisplay()
	End Sub

	Private Sub picCanvas_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles picCanvas.Click
		Dim Mouse As Point = picCanvas.PointToClient(picCanvas.MousePosition)
		Dim x As Integer = CType(Math.Ceiling(Mouse.X / map.CharacterWidth), Integer)
		Dim y As Integer = CType(Math.Ceiling(Mouse.Y / map.CharacterHeight), Integer)
		Display.MoveTo(y, x)
		Display.WriteCharacter(map.Ascii(Me.lblCharacter.Text))
	End Sub

	Private Sub picCanvas_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles picCanvas.MouseMove
		If MouseButtons = MouseButtons.Left Then picCanvas_Click(sender, e)
	End Sub

	Private Sub mnuEdit_Clear_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuEdit_Clear.Click
		Display.EraseDisplay()
	End Sub

	Private Sub mnuFile_Open_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuFile_Open.Click
		Me.OpenFileDialog1.Filter = "ANSI Graphics (*.ans)|*.ans|All (*.*)|*.*"
		Dim Results As DialogResult = Me.OpenFileDialog1.ShowDialog()
		If Results = DialogResult.Cancel Then Exit Sub
		Display.Open(Me.OpenFileDialog1.FileName)
	End Sub
	Private Sub mnuFile_Save_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuFile_Save.Click
		Me.SaveFileDialog1.Filter = "ANSI Graphics (*.ans)|*.ans|All (*.*)|*.*"
		Dim Results As DialogResult = Me.SaveFileDialog1.ShowDialog()
		If Results = DialogResult.Cancel Then Exit Sub
		Dim Buffer() As Byte = Me.Display.GetBytes()
		Dim Stream As New IO.FileStream(Me.SaveFileDialog1.FileName, IO.FileMode.Create)
		Stream.Write(Buffer, 0, Buffer.Length)
		Stream.Close()
	End Sub

	Private Sub Display_DisplayModified() Handles Display.DisplayModified
		Me.picCanvas.Image = Display.GetImage
		Me.picCanvas.Invalidate()
	End Sub

	Private Sub Display_Initialized(ByVal lines As Integer, ByVal characters As Integer) Handles Display.Initialized
		If Not picCanvas.Image Is Nothing Then picCanvas.Image.Dispose()
		picCanvas.Image = New Bitmap(AnsiCharacterMap.CharacterWidth * characters, AnsiCharacterMap.CharacterHeight * lines)
	End Sub

	Private Sub Display_CharacterModified(ByVal line As Integer, ByVal character As Integer) Handles Display.CharacterModified
		If picCanvas.Image Is Nothing Then
			Me.picCanvas.Image = Display.GetImage
			Me.picCanvas.Invalidate()
			Return
		End If

		Dim g As Graphics
		Dim cWidth As Integer = AnsiCharacterMap.CharacterWidth
		Dim cHeight As Integer = AnsiCharacterMap.CharacterHeight
		Dim cLeft As Integer = AnsiCharacterMap.CharacterLeft
		Dim cTop As Integer = AnsiCharacterMap.CharacterTop
		Dim c As AnsiCharacter

		g = g.FromImage(picCanvas.Image)

		c = Display.Screen(character, line)

		' Fill background
		Dim brush As New SolidBrush(c.BackColor)
		g.FillRectangle(brush, ((character - 1) * cWidth), ((line - 1) * cHeight), cWidth, cHeight)

		' Write character
		brush.Color = c.ForeColor
		g.DrawString(c.Character, AnsiCharacterMap.RegularFont, brush, ((character - 1) * cWidth) + cLeft, ((line - 1) * cHeight) + cTop)

		brush.Dispose()
		g.Dispose()
		Me.picCanvas.Invalidate()
	End Sub

	Private Sub BackgroundPanel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BackgroundPanel.Click
		Dim frm As New frmColors
		frm.IsBackground = True
		frm.ShowDialog(Me)
		lblCharacter.BackColor = frm.Color
		BackgroundPanel.BackColor = frm.Color
		Display.Character.BackColor = frm.Color
		frm.Dispose()
	End Sub

	Private Sub ForegroundPanel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ForegroundPanel.Click
		Dim frm As New frmColors
		frm.IsBackground = False
		frm.ShowDialog(Me)
		lblCharacter.ForeColor = frm.Color
		ForegroundPanel.BackColor = frm.Color
		Display.Character.Bold = frm.IsBold
		Display.Character.ForeColor = frm.Color
		frm.Dispose()
	End Sub

	Private Sub lblCharacter_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lblCharacter.Click
		Dim frm As New frmCharacters
		frm.BackColor = lblCharacter.BackColor
		frm.ForeColor = lblCharacter.ForeColor
		frm.ShowDialog(Me)
		lblCharacter.Text = frm.Character
		frm.Dispose()
	End Sub

	Private Sub MenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem1.Click
		System.Diagnostics.Process.Start("http://www.lewismoten.com")
	End Sub

	Private Sub MenuItem2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem2.Click
		Application.Exit()
	End Sub
End Class
