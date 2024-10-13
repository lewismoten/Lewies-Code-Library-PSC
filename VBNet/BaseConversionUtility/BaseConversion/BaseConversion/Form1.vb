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
	Friend WithEvents lblFromBase As System.Windows.Forms.Label
	Friend WithEvents lblToBase As System.Windows.Forms.Label
	Friend WithEvents lblValue As System.Windows.Forms.Label
	Friend WithEvents btnConvert As System.Windows.Forms.Button
	Friend WithEvents txtValue As System.Windows.Forms.TextBox
	Friend WithEvents txtConverted As System.Windows.Forms.TextBox
	Friend WithEvents cboFromBase As System.Windows.Forms.ComboBox
	Friend WithEvents cboToBase As System.Windows.Forms.ComboBox
	Friend WithEvents lblConverted As System.Windows.Forms.Label
	Friend WithEvents lblFromBaseCount As System.Windows.Forms.Label
	Friend WithEvents lblToBaseCount As System.Windows.Forms.Label
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Me.lblFromBase = New System.Windows.Forms.Label
		Me.lblToBase = New System.Windows.Forms.Label
		Me.lblValue = New System.Windows.Forms.Label
		Me.btnConvert = New System.Windows.Forms.Button
		Me.txtValue = New System.Windows.Forms.TextBox
		Me.txtConverted = New System.Windows.Forms.TextBox
		Me.cboFromBase = New System.Windows.Forms.ComboBox
		Me.cboToBase = New System.Windows.Forms.ComboBox
		Me.lblConverted = New System.Windows.Forms.Label
		Me.lblFromBaseCount = New System.Windows.Forms.Label
		Me.lblToBaseCount = New System.Windows.Forms.Label
		Me.SuspendLayout()
		'
		'lblFromBase
		'
		Me.lblFromBase.Location = New System.Drawing.Point(10, 20)
		Me.lblFromBase.Name = "lblFromBase"
		Me.lblFromBase.TabIndex = 0
		Me.lblFromBase.Text = "From Base"
		'
		'lblToBase
		'
		Me.lblToBase.Location = New System.Drawing.Point(10, 50)
		Me.lblToBase.Name = "lblToBase"
		Me.lblToBase.TabIndex = 1
		Me.lblToBase.Text = "To Base"
		'
		'lblValue
		'
		Me.lblValue.Location = New System.Drawing.Point(10, 80)
		Me.lblValue.Name = "lblValue"
		Me.lblValue.TabIndex = 2
		Me.lblValue.Text = "Value to convert"
		'
		'btnConvert
		'
		Me.btnConvert.Location = New System.Drawing.Point(110, 110)
		Me.btnConvert.Name = "btnConvert"
		Me.btnConvert.TabIndex = 3
		Me.btnConvert.Text = "Convert"
		'
		'txtValue
		'
		Me.txtValue.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
					Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
		Me.txtValue.Location = New System.Drawing.Point(110, 80)
		Me.txtValue.Name = "txtValue"
		Me.txtValue.Size = New System.Drawing.Size(180, 20)
		Me.txtValue.TabIndex = 5
		Me.txtValue.Text = "123456"
		'
		'txtConverted
		'
		Me.txtConverted.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
					Or System.Windows.Forms.AnchorStyles.Left) _
					Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
		Me.txtConverted.Location = New System.Drawing.Point(110, 140)
		Me.txtConverted.Multiline = True
		Me.txtConverted.Name = "txtConverted"
		Me.txtConverted.Size = New System.Drawing.Size(180, 20)
		Me.txtConverted.TabIndex = 6
		Me.txtConverted.Text = "11110001001000000"
		'
		'cboFromBase
		'
		Me.cboFromBase.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
					Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
		Me.cboFromBase.Items.AddRange(New Object() {"01", "01234567", "0123456789", "0123456789abcdef", " .-", "gatc", "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/", "LEwisMOten.com"})
		Me.cboFromBase.Location = New System.Drawing.Point(110, 20)
		Me.cboFromBase.Name = "cboFromBase"
		Me.cboFromBase.Size = New System.Drawing.Size(140, 21)
		Me.cboFromBase.TabIndex = 7
		Me.cboFromBase.Text = "0123456789"
		'
		'cboToBase
		'
		Me.cboToBase.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
					Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
		Me.cboToBase.Items.AddRange(New Object() {"01", "01234567", "0123456789", "0123456789abcdef", " .-", "gatc", "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/", "LEwisMOten.com"})
		Me.cboToBase.Location = New System.Drawing.Point(110, 50)
		Me.cboToBase.Name = "cboToBase"
		Me.cboToBase.Size = New System.Drawing.Size(140, 21)
		Me.cboToBase.TabIndex = 8
		Me.cboToBase.Text = "01"
		'
		'lblConverted
		'
		Me.lblConverted.Location = New System.Drawing.Point(10, 140)
		Me.lblConverted.Name = "lblConverted"
		Me.lblConverted.TabIndex = 9
		Me.lblConverted.Text = "Converted Value"
		'
		'lblFromBaseCount
		'
		Me.lblFromBaseCount.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
		Me.lblFromBaseCount.Location = New System.Drawing.Point(260, 20)
		Me.lblFromBaseCount.Name = "lblFromBaseCount"
		Me.lblFromBaseCount.Size = New System.Drawing.Size(30, 23)
		Me.lblFromBaseCount.TabIndex = 10
		Me.lblFromBaseCount.Text = "10"
		Me.lblFromBaseCount.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		'
		'lblToBaseCount
		'
		Me.lblToBaseCount.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
		Me.lblToBaseCount.Location = New System.Drawing.Point(260, 50)
		Me.lblToBaseCount.Name = "lblToBaseCount"
		Me.lblToBaseCount.Size = New System.Drawing.Size(30, 23)
		Me.lblToBaseCount.TabIndex = 11
		Me.lblToBaseCount.Text = "2"
		Me.lblToBaseCount.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		'
		'Form1
		'
		Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
		Me.ClientSize = New System.Drawing.Size(302, 176)
		Me.Controls.Add(Me.lblToBaseCount)
		Me.Controls.Add(Me.lblFromBaseCount)
		Me.Controls.Add(Me.cboToBase)
		Me.Controls.Add(Me.cboFromBase)
		Me.Controls.Add(Me.txtConverted)
		Me.Controls.Add(Me.txtValue)
		Me.Controls.Add(Me.lblToBase)
		Me.Controls.Add(Me.lblFromBase)
		Me.Controls.Add(Me.btnConvert)
		Me.Controls.Add(Me.lblConverted)
		Me.Controls.Add(Me.lblValue)
		Me.Name = "Form1"
		Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
		Me.Text = "Base Conversion"
		Me.ResumeLayout(False)

	End Sub

#End Region

	Private Sub cboFromBase_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboFromBase.TextChanged
		Me.lblFromBaseCount.Text = cboFromBase.Text.Length
		Me.txtConverted.Text = ""
	End Sub

	Private Sub cboToBase_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboToBase.TextChanged
		Me.lblToBaseCount.Text = cboToBase.Text.Length
		Me.txtConverted.Text = ""
	End Sub

	Private Sub btnConvert_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnConvert.Click
		Me.txtConverted.Text = Logic.BaseConversion(Me.txtValue.Text, Me.cboFromBase.Text, Me.cboToBase.Text)
	End Sub

	Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
		Dim n As New Splash
		n.ShowDialog(Me)
		n.Dispose()
		n = Nothing
	End Sub
End Class
