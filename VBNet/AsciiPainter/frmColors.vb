Public Class frmColors
	Inherits System.Windows.Forms.Form
	Private Painter As Form1
	Public Color As Color
	Public IsBackground As Boolean
	Public IsBold As Boolean = False

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
	Friend WithEvents pnlWhite As System.Windows.Forms.Panel
	Friend WithEvents pnlYellow As System.Windows.Forms.Panel
	Friend WithEvents pnlPink As System.Windows.Forms.Panel
	Friend WithEvents pnlOrange As System.Windows.Forms.Panel
	Friend WithEvents pnlBrightCyan As System.Windows.Forms.Panel
	Friend WithEvents pnlBrightGreen As System.Windows.Forms.Panel
	Friend WithEvents pnlBrightBlue As System.Windows.Forms.Panel
	Friend WithEvents pnlCharcoal As System.Windows.Forms.Panel
	Friend WithEvents pnlGray As System.Windows.Forms.Panel
	Friend WithEvents pnlBrown As System.Windows.Forms.Panel
	Friend WithEvents pnlMagenta As System.Windows.Forms.Panel
	Friend WithEvents pnlRed As System.Windows.Forms.Panel
	Friend WithEvents pnlCyan As System.Windows.Forms.Panel
	Friend WithEvents pnlGreen As System.Windows.Forms.Panel
	Friend WithEvents pnlBlue As System.Windows.Forms.Panel
	Friend WithEvents pnlBlack As System.Windows.Forms.Panel
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Me.pnlWhite = New System.Windows.Forms.Panel
		Me.pnlYellow = New System.Windows.Forms.Panel
		Me.pnlPink = New System.Windows.Forms.Panel
		Me.pnlOrange = New System.Windows.Forms.Panel
		Me.pnlBrightCyan = New System.Windows.Forms.Panel
		Me.pnlBrightGreen = New System.Windows.Forms.Panel
		Me.pnlBrightBlue = New System.Windows.Forms.Panel
		Me.pnlCharcoal = New System.Windows.Forms.Panel
		Me.pnlGray = New System.Windows.Forms.Panel
		Me.pnlBrown = New System.Windows.Forms.Panel
		Me.pnlMagenta = New System.Windows.Forms.Panel
		Me.pnlRed = New System.Windows.Forms.Panel
		Me.pnlCyan = New System.Windows.Forms.Panel
		Me.pnlGreen = New System.Windows.Forms.Panel
		Me.pnlBlue = New System.Windows.Forms.Panel
		Me.pnlBlack = New System.Windows.Forms.Panel
		Me.SuspendLayout()
		'
		'pnlWhite
		'
		Me.pnlWhite.BackColor = System.Drawing.Color.White
		Me.pnlWhite.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me.pnlWhite.Location = New System.Drawing.Point(100, 100)
		Me.pnlWhite.Name = "pnlWhite"
		Me.pnlWhite.Size = New System.Drawing.Size(20, 20)
		Me.pnlWhite.TabIndex = 15
		'
		'pnlYellow
		'
		Me.pnlYellow.BackColor = System.Drawing.Color.Yellow
		Me.pnlYellow.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me.pnlYellow.Location = New System.Drawing.Point(70, 100)
		Me.pnlYellow.Name = "pnlYellow"
		Me.pnlYellow.Size = New System.Drawing.Size(20, 20)
		Me.pnlYellow.TabIndex = 14
		'
		'pnlPink
		'
		Me.pnlPink.BackColor = System.Drawing.Color.Fuchsia
		Me.pnlPink.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me.pnlPink.Location = New System.Drawing.Point(40, 100)
		Me.pnlPink.Name = "pnlPink"
		Me.pnlPink.Size = New System.Drawing.Size(20, 20)
		Me.pnlPink.TabIndex = 13
		'
		'pnlOrange
		'
		Me.pnlOrange.BackColor = System.Drawing.Color.Red
		Me.pnlOrange.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me.pnlOrange.Location = New System.Drawing.Point(10, 100)
		Me.pnlOrange.Name = "pnlOrange"
		Me.pnlOrange.Size = New System.Drawing.Size(20, 20)
		Me.pnlOrange.TabIndex = 12
		'
		'pnlBrightCyan
		'
		Me.pnlBrightCyan.BackColor = System.Drawing.Color.Cyan
		Me.pnlBrightCyan.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me.pnlBrightCyan.Location = New System.Drawing.Point(100, 70)
		Me.pnlBrightCyan.Name = "pnlBrightCyan"
		Me.pnlBrightCyan.Size = New System.Drawing.Size(20, 20)
		Me.pnlBrightCyan.TabIndex = 11
		'
		'pnlBrightGreen
		'
		Me.pnlBrightGreen.BackColor = System.Drawing.Color.Lime
		Me.pnlBrightGreen.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me.pnlBrightGreen.Location = New System.Drawing.Point(70, 70)
		Me.pnlBrightGreen.Name = "pnlBrightGreen"
		Me.pnlBrightGreen.Size = New System.Drawing.Size(20, 20)
		Me.pnlBrightGreen.TabIndex = 10
		'
		'pnlBrightBlue
		'
		Me.pnlBrightBlue.BackColor = System.Drawing.Color.Blue
		Me.pnlBrightBlue.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me.pnlBrightBlue.Location = New System.Drawing.Point(40, 70)
		Me.pnlBrightBlue.Name = "pnlBrightBlue"
		Me.pnlBrightBlue.Size = New System.Drawing.Size(20, 20)
		Me.pnlBrightBlue.TabIndex = 9
		'
		'pnlCharcoal
		'
		Me.pnlCharcoal.BackColor = System.Drawing.Color.Gray
		Me.pnlCharcoal.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me.pnlCharcoal.Location = New System.Drawing.Point(10, 70)
		Me.pnlCharcoal.Name = "pnlCharcoal"
		Me.pnlCharcoal.Size = New System.Drawing.Size(20, 20)
		Me.pnlCharcoal.TabIndex = 8
		'
		'pnlGray
		'
		Me.pnlGray.BackColor = System.Drawing.Color.Silver
		Me.pnlGray.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me.pnlGray.Location = New System.Drawing.Point(100, 40)
		Me.pnlGray.Name = "pnlGray"
		Me.pnlGray.Size = New System.Drawing.Size(20, 20)
		Me.pnlGray.TabIndex = 7
		'
		'pnlBrown
		'
		Me.pnlBrown.BackColor = System.Drawing.Color.Olive
		Me.pnlBrown.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me.pnlBrown.Location = New System.Drawing.Point(70, 40)
		Me.pnlBrown.Name = "pnlBrown"
		Me.pnlBrown.Size = New System.Drawing.Size(20, 20)
		Me.pnlBrown.TabIndex = 6
		'
		'pnlMagenta
		'
		Me.pnlMagenta.BackColor = System.Drawing.Color.Purple
		Me.pnlMagenta.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me.pnlMagenta.Location = New System.Drawing.Point(40, 40)
		Me.pnlMagenta.Name = "pnlMagenta"
		Me.pnlMagenta.Size = New System.Drawing.Size(20, 20)
		Me.pnlMagenta.TabIndex = 5
		'
		'pnlRed
		'
		Me.pnlRed.BackColor = System.Drawing.Color.Maroon
		Me.pnlRed.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me.pnlRed.Location = New System.Drawing.Point(10, 40)
		Me.pnlRed.Name = "pnlRed"
		Me.pnlRed.Size = New System.Drawing.Size(20, 20)
		Me.pnlRed.TabIndex = 4
		'
		'pnlCyan
		'
		Me.pnlCyan.BackColor = System.Drawing.Color.Teal
		Me.pnlCyan.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me.pnlCyan.Location = New System.Drawing.Point(100, 10)
		Me.pnlCyan.Name = "pnlCyan"
		Me.pnlCyan.Size = New System.Drawing.Size(20, 20)
		Me.pnlCyan.TabIndex = 3
		'
		'pnlGreen
		'
		Me.pnlGreen.BackColor = System.Drawing.Color.FromArgb(CType(0, Byte), CType(179, Byte), CType(0, Byte))
		Me.pnlGreen.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me.pnlGreen.Location = New System.Drawing.Point(70, 10)
		Me.pnlGreen.Name = "pnlGreen"
		Me.pnlGreen.Size = New System.Drawing.Size(20, 20)
		Me.pnlGreen.TabIndex = 2
		'
		'pnlBlue
		'
		Me.pnlBlue.BackColor = System.Drawing.Color.Navy
		Me.pnlBlue.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me.pnlBlue.Location = New System.Drawing.Point(40, 10)
		Me.pnlBlue.Name = "pnlBlue"
		Me.pnlBlue.Size = New System.Drawing.Size(20, 20)
		Me.pnlBlue.TabIndex = 1
		'
		'pnlBlack
		'
		Me.pnlBlack.BackColor = System.Drawing.Color.Black
		Me.pnlBlack.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me.pnlBlack.Location = New System.Drawing.Point(10, 10)
		Me.pnlBlack.Name = "pnlBlack"
		Me.pnlBlack.Size = New System.Drawing.Size(20, 20)
		Me.pnlBlack.TabIndex = 0
		'
		'frmColors
		'
		Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
		Me.ClientSize = New System.Drawing.Size(132, 126)
		Me.Controls.Add(Me.pnlWhite)
		Me.Controls.Add(Me.pnlYellow)
		Me.Controls.Add(Me.pnlPink)
		Me.Controls.Add(Me.pnlOrange)
		Me.Controls.Add(Me.pnlBrightCyan)
		Me.Controls.Add(Me.pnlBrightGreen)
		Me.Controls.Add(Me.pnlBrightBlue)
		Me.Controls.Add(Me.pnlCharcoal)
		Me.Controls.Add(Me.pnlGray)
		Me.Controls.Add(Me.pnlBrown)
		Me.Controls.Add(Me.pnlMagenta)
		Me.Controls.Add(Me.pnlRed)
		Me.Controls.Add(Me.pnlCyan)
		Me.Controls.Add(Me.pnlGreen)
		Me.Controls.Add(Me.pnlBlue)
		Me.Controls.Add(Me.pnlBlack)
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
		Me.Name = "frmColors"
		Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
		Me.Text = "Color Palette"
		Me.ResumeLayout(False)

	End Sub

#End Region

	Private Sub pnlColor_Click(ByVal sender As Object, ByVal e As System.EventArgs)
		Color = CType(sender, Control).BackColor
		IsBold = CType(sender, Control).TabIndex > 7
		Me.Close()
	End Sub

	Private Sub frmColors_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
		For Each Control As Control In Me.Controls
			AddHandler Control.Click, AddressOf pnlColor_Click
		Next
		If IsBackground Then Height = 90 Else Height = 150
	End Sub
End Class
