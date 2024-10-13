Public Class Splash
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
	Friend WithEvents lblName As System.Windows.Forms.Label
	Friend WithEvents lblVersion As System.Windows.Forms.Label
	Friend WithEvents Timer1 As System.Windows.Forms.Timer
	Friend WithEvents lblCompany As System.Windows.Forms.Label
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Me.components = New System.ComponentModel.Container
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(Splash))
		Me.lblName = New System.Windows.Forms.Label
		Me.lblVersion = New System.Windows.Forms.Label
		Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
		Me.lblCompany = New System.Windows.Forms.Label
		Me.SuspendLayout()
		'
		'lblName
		'
		Me.lblName.BackColor = System.Drawing.Color.Transparent
		Me.lblName.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblName.ForeColor = System.Drawing.Color.LightGray
		Me.lblName.Location = New System.Drawing.Point(10, 20)
		Me.lblName.Name = "lblName"
		Me.lblName.Size = New System.Drawing.Size(150, 90)
		Me.lblName.TabIndex = 0
		Me.lblName.Text = "Base Conversion Utility"
		'
		'lblVersion
		'
		Me.lblVersion.AutoSize = True
		Me.lblVersion.BackColor = System.Drawing.Color.Black
		Me.lblVersion.ForeColor = System.Drawing.SystemColors.GrayText
		Me.lblVersion.Location = New System.Drawing.Point(10, 110)
		Me.lblVersion.Name = "lblVersion"
		Me.lblVersion.Size = New System.Drawing.Size(60, 16)
		Me.lblVersion.TabIndex = 1
		Me.lblVersion.Text = "version 1.0"
		'
		'Timer1
		'
		Me.Timer1.Interval = 3000
		'
		'lblCompany
		'
		Me.lblCompany.BackColor = System.Drawing.Color.Black
		Me.lblCompany.ForeColor = System.Drawing.SystemColors.GrayText
		Me.lblCompany.Location = New System.Drawing.Point(10, 180)
		Me.lblCompany.Name = "lblCompany"
		Me.lblCompany.Size = New System.Drawing.Size(240, 16)
		Me.lblCompany.TabIndex = 2
		Me.lblCompany.Text = "LewisMoten.com"
		Me.lblCompany.TextAlign = System.Drawing.ContentAlignment.MiddleRight
		'
		'Splash
		'
		Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
		Me.BackgroundImage = CType(resources.GetObject("$this.BackgroundImage"), System.Drawing.Image)
		Me.ClientSize = New System.Drawing.Size(261, 210)
		Me.ControlBox = False
		Me.Controls.Add(Me.lblCompany)
		Me.Controls.Add(Me.lblVersion)
		Me.Controls.Add(Me.lblName)
		Me.Cursor = System.Windows.Forms.Cursors.AppStarting
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
		Me.Name = "Splash"
		Me.ShowInTaskbar = False
		Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
		Me.Text = "Splash"
		Me.TopMost = True
		Me.ResumeLayout(False)

	End Sub

#End Region

	Private Sub Splash_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
		Me.lblName.Text = Application.ProductName
		Me.lblVersion.Text = "version " & Application.ProductVersion
		Me.lblCompany.Text = Application.CompanyName
		Timer1.Enabled = True
	End Sub

	Private Sub Timer1_Tick(ByVal sender As Object, ByVal e As System.EventArgs) Handles Timer1.Tick
		Timer1.Enabled = False
		Me.Close()
	End Sub
End Class
