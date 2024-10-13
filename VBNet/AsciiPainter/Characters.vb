Public Class frmCharacters
	Inherits System.Windows.Forms.Form
	Public Character As String

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
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		'
		'frmCharacters
		'
		Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
		Me.ClientSize = New System.Drawing.Size(154, 186)
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
		Me.Name = "frmCharacters"
		Me.ShowInTaskbar = False
		Me.Text = "Characters"
		Me.TopMost = True

	End Sub

#End Region

	Protected Overrides Sub OnPaint(ByVal e As System.Windows.Forms.PaintEventArgs)
		'9x15
		Dim Map As New AnsiCharacterMap
		Dim Character As New AnsiCharacter

		Dim IsBold As Boolean = False
		Character.ForeColor = ForeColor
		IsBold = Character.Foreground > 7

		Dim g As Graphics = e.Graphics
		Dim b As New SolidBrush(ForeColor)

		g.Clear(BackColor)

		For x As Integer = 0 To 15
			For y As Integer = 2 To 15
				Dim s As String = Map.Unicode((y * 16) + x)
				If IsBold Then
					g.DrawString(s, Map.BoldFont, b, (x * Map.CharacterWidth) + Map.CharacterLeft, ((y - 2) * Map.CharacterHeight) + Map.CharacterTop)
				Else
					g.DrawString(s, Map.RegularFont, b, (x * Map.CharacterWidth) + Map.CharacterLeft, ((y - 2) * Map.CharacterHeight) + Map.CharacterTop)
				End If
			Next
		Next

		b.Dispose()
		g.Dispose()
	End Sub

	Private Sub frmCharacters_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Click
		Dim Map As New AnsiCharacterMap
		Dim Mouse As Point = Me.PointToClient(Me.MousePosition)
		Dim x As Integer = Mouse.X \ Map.CharacterWidth
		Dim y As Integer = Mouse.Y \ Map.CharacterHeight
		y += 2
		Character = Map.Unicode((y * 16) + x)
		Me.Close()
	End Sub

	Private Sub frmCharacters_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
		Dim Map As New AnsiCharacterMap
		Me.Width = Map.CharacterWidth * 16 + (Width - ClientSize.Width)
		Me.Height = Map.CharacterHeight * 14 + (Height - ClientSize.Height)
	End Sub
End Class
