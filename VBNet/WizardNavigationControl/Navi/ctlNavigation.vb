Public Class ctlNavigation
	Inherits System.Windows.Forms.UserControl
	Public Event Next_Click()
	Public Event Back_Click()
	Public Event Finish_Click()

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'UserControl overrides dispose to clean up the component list.
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
	Friend WithEvents Panel1 As System.Windows.Forms.Panel
	Friend WithEvents btnCancel As System.Windows.Forms.Button
	Friend WithEvents btnFinish As System.Windows.Forms.Button
	Friend WithEvents btnNext As System.Windows.Forms.Button
	Friend WithEvents btnBack As System.Windows.Forms.Button
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Me.Panel1 = New System.Windows.Forms.Panel()
		Me.btnCancel = New System.Windows.Forms.Button()
		Me.btnFinish = New System.Windows.Forms.Button()
		Me.btnNext = New System.Windows.Forms.Button()
		Me.btnBack = New System.Windows.Forms.Button()
		Me.Panel1.SuspendLayout()
		Me.SuspendLayout()
		'
		'Panel1
		'
		Me.Panel1.Controls.AddRange(New System.Windows.Forms.Control() {Me.btnFinish, Me.btnNext, Me.btnBack})
		Me.Panel1.Dock = System.Windows.Forms.DockStyle.Right
		Me.Panel1.Location = New System.Drawing.Point(140, 0)
		Me.Panel1.Name = "Panel1"
		Me.Panel1.Size = New System.Drawing.Size(270, 40)
		Me.Panel1.TabIndex = 0
		'
		'btnCancel
		'
		Me.btnCancel.BackColor = System.Drawing.SystemColors.Control
		Me.btnCancel.Location = New System.Drawing.Point(10, 10)
		Me.btnCancel.Name = "btnCancel"
		Me.btnCancel.TabIndex = 5
		Me.btnCancel.Text = "Cancel"
		'
		'btnFinish
		'
		Me.btnFinish.BackColor = System.Drawing.SystemColors.Control
		Me.btnFinish.Location = New System.Drawing.Point(180, 10)
		Me.btnFinish.Name = "btnFinish"
		Me.btnFinish.TabIndex = 4
		Me.btnFinish.Text = "Finish"
		'
		'btnNext
		'
		Me.btnNext.BackColor = System.Drawing.SystemColors.Control
		Me.btnNext.Location = New System.Drawing.Point(90, 10)
		Me.btnNext.Name = "btnNext"
		Me.btnNext.TabIndex = 5
		Me.btnNext.Text = "Next >>"
		'
		'btnBack
		'
		Me.btnBack.BackColor = System.Drawing.SystemColors.Control
		Me.btnBack.Location = New System.Drawing.Point(0, 10)
		Me.btnBack.Name = "btnBack"
		Me.btnBack.TabIndex = 6
		Me.btnBack.Text = "<< Back"
		'
		'ctlNavigation
		'
		Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.btnCancel, Me.Panel1})
		Me.Name = "ctlNavigation"
		Me.Size = New System.Drawing.Size(410, 40)
		Me.Panel1.ResumeLayout(False)
		Me.ResumeLayout(False)

	End Sub

#End Region

	Public Property [Next]() As Boolean
		Get
			Return Me.btnNext.Enabled
		End Get
		Set(ByVal Value As Boolean)
			Me.btnNext.Enabled = Value
		End Set
	End Property
	Public Property Back() As Boolean
		Get
			Return Me.btnBack.Enabled
		End Get
		Set(ByVal Value As Boolean)
			Me.btnBack.Enabled = Value
		End Set
	End Property
	Public Property Finish() As Boolean
		Get
			Return Me.btnFinish.Enabled
		End Get
		Set(ByVal Value As Boolean)
			Me.btnFinish.Enabled = Value
		End Set
	End Property
	Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
		Application.Exit()
	End Sub

	Private Sub btnBack_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnBack.Click
		RaiseEvent Back_Click()
	End Sub

	Private Sub btnNext_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNext.Click
		RaiseEvent Next_Click()
	End Sub

	Private Sub btnFinish_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFinish.Click
		RaiseEvent Finish_Click()
	End Sub
End Class
