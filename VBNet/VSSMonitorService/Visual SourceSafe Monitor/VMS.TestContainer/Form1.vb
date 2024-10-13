Public Class Form1
	Inherits System.Windows.Forms.Form
	Protected WithEvents Monitor As New JournalMonitor

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
	Friend WithEvents btnStart As System.Windows.Forms.Button
	Friend WithEvents btnStop As System.Windows.Forms.Button
	Friend WithEvents btnPause As System.Windows.Forms.Button
	Friend WithEvents btnContinue As System.Windows.Forms.Button
	Friend WithEvents Label1 As System.Windows.Forms.Label
	Friend WithEvents Label2 As System.Windows.Forms.Label
	Friend WithEvents Label3 As System.Windows.Forms.Label
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Me.btnStart = New System.Windows.Forms.Button
		Me.btnStop = New System.Windows.Forms.Button
		Me.btnPause = New System.Windows.Forms.Button
		Me.btnContinue = New System.Windows.Forms.Button
		Me.Label1 = New System.Windows.Forms.Label
		Me.Label2 = New System.Windows.Forms.Label
		Me.Label3 = New System.Windows.Forms.Label
		Me.SuspendLayout()
		'
		'btnStart
		'
		Me.btnStart.Font = New System.Drawing.Font("Webdings", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(2, Byte))
		Me.btnStart.Location = New System.Drawing.Point(10, 30)
		Me.btnStart.Name = "btnStart"
		Me.btnStart.Size = New System.Drawing.Size(30, 23)
		Me.btnStart.TabIndex = 0
		Me.btnStart.Text = "4"
		Me.btnStart.TextAlign = System.Drawing.ContentAlignment.BottomCenter
		'
		'btnStop
		'
		Me.btnStop.Font = New System.Drawing.Font("Webdings", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(2, Byte))
		Me.btnStop.Location = New System.Drawing.Point(50, 30)
		Me.btnStop.Name = "btnStop"
		Me.btnStop.Size = New System.Drawing.Size(30, 23)
		Me.btnStop.TabIndex = 1
		Me.btnStop.Text = "<"
		Me.btnStop.TextAlign = System.Drawing.ContentAlignment.BottomCenter
		'
		'btnPause
		'
		Me.btnPause.Font = New System.Drawing.Font("Webdings", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(2, Byte))
		Me.btnPause.Location = New System.Drawing.Point(10, 80)
		Me.btnPause.Name = "btnPause"
		Me.btnPause.Size = New System.Drawing.Size(30, 23)
		Me.btnPause.TabIndex = 2
		Me.btnPause.Text = ";"
		Me.btnPause.TextAlign = System.Drawing.ContentAlignment.BottomCenter
		'
		'btnContinue
		'
		Me.btnContinue.Font = New System.Drawing.Font("Webdings", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(2, Byte))
		Me.btnContinue.Location = New System.Drawing.Point(50, 80)
		Me.btnContinue.Name = "btnContinue"
		Me.btnContinue.Size = New System.Drawing.Size(30, 23)
		Me.btnContinue.TabIndex = 3
		Me.btnContinue.Text = "4"
		Me.btnContinue.TextAlign = System.Drawing.ContentAlignment.BottomCenter
		'
		'Label1
		'
		Me.Label1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.Label1.Location = New System.Drawing.Point(110, 10)
		Me.Label1.Name = "Label1"
		Me.Label1.Size = New System.Drawing.Size(170, 100)
		Me.Label1.TabIndex = 4
		Me.Label1.Text = "Test and debug your service in this small application before running it as a true" & _
		" service. (www.lewismoten.com)"
		'
		'Label2
		'
		Me.Label2.Location = New System.Drawing.Point(10, 60)
		Me.Label2.Name = "Label2"
		Me.Label2.TabIndex = 5
		Me.Label2.Text = "Pause / Continue"
		'
		'Label3
		'
		Me.Label3.Location = New System.Drawing.Point(10, 10)
		Me.Label3.Name = "Label3"
		Me.Label3.TabIndex = 6
		Me.Label3.Text = "Start / Stop"
		'
		'Form1
		'
		Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
		Me.ClientSize = New System.Drawing.Size(292, 116)
		Me.Controls.Add(Me.Label1)
		Me.Controls.Add(Me.btnContinue)
		Me.Controls.Add(Me.btnPause)
		Me.Controls.Add(Me.btnStop)
		Me.Controls.Add(Me.btnStart)
		Me.Controls.Add(Me.Label2)
		Me.Controls.Add(Me.Label3)
		Me.Name = "Form1"
		Me.Text = "VMS Test Container"
		Me.ResumeLayout(False)

	End Sub

#End Region

	Private Sub btnStart_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnStart.Click
		Monitor.TestEvent(JournalMonitor.eventTypes.OnStart)
		Me.btnStart.Enabled = False
		Me.btnStop.Enabled = True
		Me.btnPause.Enabled = True
		Me.btnContinue.Enabled = False
	End Sub

	Private Sub btnStop_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnStop.Click
		Monitor.TestEvent(JournalMonitor.eventTypes.OnStop)
		Me.btnStart.Enabled = True
		Me.btnStop.Enabled = False
		Me.btnPause.Enabled = False
		Me.btnContinue.Enabled = False
	End Sub

	Private Sub btnPause_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPause.Click
		Monitor.TestEvent(JournalMonitor.eventTypes.OnPause)
		Me.btnStart.Enabled = False
		Me.btnStop.Enabled = True
		Me.btnPause.Enabled = False
		Me.btnContinue.Enabled = True
	End Sub

	Private Sub btnContinue_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnContinue.Click
		Monitor.TestEvent(JournalMonitor.eventTypes.OnContinue)
		Me.btnStart.Enabled = False
		Me.btnStop.Enabled = True
		Me.btnPause.Enabled = True
		Me.btnContinue.Enabled = False
	End Sub

	Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
		Me.btnStart.Enabled = True
		Me.btnStop.Enabled = False
		Me.btnPause.Enabled = False
		Me.btnContinue.Enabled = False
	End Sub
End Class
