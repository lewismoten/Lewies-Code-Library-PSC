Imports Microsoft.VisualBasic
Imports System
Imports System.Windows.Forms

Public Class frmDice
	Inherits System.Windows.Forms.Form
	Private Rolls As Integer
	Private Sides As Integer

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
	Friend WithEvents rdoSides4 As System.Windows.Forms.RadioButton
	Friend WithEvents rdoSides6 As System.Windows.Forms.RadioButton
	Friend WithEvents rdoSides8 As System.Windows.Forms.RadioButton
	Friend WithEvents rdoSides10 As System.Windows.Forms.RadioButton
	Friend WithEvents rdoSides12 As System.Windows.Forms.RadioButton
	Friend WithEvents rdoSides20 As System.Windows.Forms.RadioButton
	Friend WithEvents rdoRolls4 As System.Windows.Forms.RadioButton
	Friend WithEvents rdoRolls3 As System.Windows.Forms.RadioButton
	Friend WithEvents rdoRolls5 As System.Windows.Forms.RadioButton
	Friend WithEvents rdoRolls1 As System.Windows.Forms.RadioButton
	Friend WithEvents rdoRolls2 As System.Windows.Forms.RadioButton
	Friend WithEvents rdoRollsOther As System.Windows.Forms.RadioButton
	Friend WithEvents txtRolls As System.Windows.Forms.TextBox
	Friend WithEvents txtSides As System.Windows.Forms.TextBox
	Friend WithEvents rdoSidesOther As System.Windows.Forms.RadioButton
	Friend WithEvents btnRoll As System.Windows.Forms.Button
	Friend WithEvents lstResults As System.Windows.Forms.ListBox
	Friend WithEvents lblTotalResults As System.Windows.Forms.Label
	Friend WithEvents rdoRolls6 As System.Windows.Forms.RadioButton
	Friend WithEvents lnkLewisMoten_com As System.Windows.Forms.LinkLabel
	Friend WithEvents lblSides As System.Windows.Forms.Label
	Friend WithEvents pnlSides As System.Windows.Forms.Panel
	Friend WithEvents pnlRolls As System.Windows.Forms.Panel
	Friend WithEvents lblRolls As System.Windows.Forms.Label
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Me.lblSides = New System.Windows.Forms.Label
		Me.rdoSides4 = New System.Windows.Forms.RadioButton
		Me.rdoSides6 = New System.Windows.Forms.RadioButton
		Me.rdoSides8 = New System.Windows.Forms.RadioButton
		Me.rdoSides10 = New System.Windows.Forms.RadioButton
		Me.rdoSides12 = New System.Windows.Forms.RadioButton
		Me.rdoSides20 = New System.Windows.Forms.RadioButton
		Me.pnlSides = New System.Windows.Forms.Panel
		Me.txtSides = New System.Windows.Forms.TextBox
		Me.rdoSidesOther = New System.Windows.Forms.RadioButton
		Me.pnlRolls = New System.Windows.Forms.Panel
		Me.txtRolls = New System.Windows.Forms.TextBox
		Me.rdoRolls6 = New System.Windows.Forms.RadioButton
		Me.rdoRolls4 = New System.Windows.Forms.RadioButton
		Me.rdoRolls3 = New System.Windows.Forms.RadioButton
		Me.rdoRolls5 = New System.Windows.Forms.RadioButton
		Me.rdoRolls1 = New System.Windows.Forms.RadioButton
		Me.lblRolls = New System.Windows.Forms.Label
		Me.rdoRolls2 = New System.Windows.Forms.RadioButton
		Me.rdoRollsOther = New System.Windows.Forms.RadioButton
		Me.btnRoll = New System.Windows.Forms.Button
		Me.lstResults = New System.Windows.Forms.ListBox
		Me.lblTotalResults = New System.Windows.Forms.Label
		Me.lnkLewisMoten_com = New System.Windows.Forms.LinkLabel
		Me.pnlSides.SuspendLayout()
		Me.pnlRolls.SuspendLayout()
		Me.SuspendLayout()
		'
		'lblSides
		'
		Me.lblSides.AutoSize = True
		Me.lblSides.Location = New System.Drawing.Point(0, 0)
		Me.lblSides.Name = "lblSides"
		Me.lblSides.Size = New System.Drawing.Size(36, 16)
		Me.lblSides.TabIndex = 0
		Me.lblSides.Text = "Sides:"
		'
		'rdoSides4
		'
		Me.rdoSides4.Location = New System.Drawing.Point(10, 20)
		Me.rdoSides4.Name = "rdoSides4"
		Me.rdoSides4.Size = New System.Drawing.Size(40, 24)
		Me.rdoSides4.TabIndex = 1
		Me.rdoSides4.Text = "4"
		'
		'rdoSides6
		'
		Me.rdoSides6.Checked = True
		Me.rdoSides6.Location = New System.Drawing.Point(10, 40)
		Me.rdoSides6.Name = "rdoSides6"
		Me.rdoSides6.Size = New System.Drawing.Size(40, 24)
		Me.rdoSides6.TabIndex = 2
		Me.rdoSides6.TabStop = True
		Me.rdoSides6.Text = "6"
		'
		'rdoSides8
		'
		Me.rdoSides8.Location = New System.Drawing.Point(10, 60)
		Me.rdoSides8.Name = "rdoSides8"
		Me.rdoSides8.Size = New System.Drawing.Size(40, 24)
		Me.rdoSides8.TabIndex = 3
		Me.rdoSides8.Text = "8"
		'
		'rdoSides10
		'
		Me.rdoSides10.Location = New System.Drawing.Point(10, 80)
		Me.rdoSides10.Name = "rdoSides10"
		Me.rdoSides10.Size = New System.Drawing.Size(50, 24)
		Me.rdoSides10.TabIndex = 4
		Me.rdoSides10.Text = "10"
		'
		'rdoSides12
		'
		Me.rdoSides12.Location = New System.Drawing.Point(10, 100)
		Me.rdoSides12.Name = "rdoSides12"
		Me.rdoSides12.Size = New System.Drawing.Size(50, 24)
		Me.rdoSides12.TabIndex = 6
		Me.rdoSides12.Text = "12"
		'
		'rdoSides20
		'
		Me.rdoSides20.Location = New System.Drawing.Point(10, 120)
		Me.rdoSides20.Name = "rdoSides20"
		Me.rdoSides20.Size = New System.Drawing.Size(50, 24)
		Me.rdoSides20.TabIndex = 7
		Me.rdoSides20.Text = "20"
		'
		'pnlSides
		'
		Me.pnlSides.Controls.Add(Me.txtSides)
		Me.pnlSides.Controls.Add(Me.rdoSidesOther)
		Me.pnlSides.Controls.Add(Me.rdoSides10)
		Me.pnlSides.Controls.Add(Me.rdoSides8)
		Me.pnlSides.Controls.Add(Me.rdoSides12)
		Me.pnlSides.Controls.Add(Me.rdoSides4)
		Me.pnlSides.Controls.Add(Me.lblSides)
		Me.pnlSides.Controls.Add(Me.rdoSides6)
		Me.pnlSides.Controls.Add(Me.rdoSides20)
		Me.pnlSides.Location = New System.Drawing.Point(90, 10)
		Me.pnlSides.Name = "pnlSides"
		Me.pnlSides.Size = New System.Drawing.Size(70, 160)
		Me.pnlSides.TabIndex = 9
		'
		'txtSides
		'
		Me.txtSides.Location = New System.Drawing.Point(30, 140)
		Me.txtSides.Name = "txtSides"
		Me.txtSides.Size = New System.Drawing.Size(30, 20)
		Me.txtSides.TabIndex = 10
		Me.txtSides.Text = "50"
		'
		'rdoSidesOther
		'
		Me.rdoSidesOther.Location = New System.Drawing.Point(10, 140)
		Me.rdoSidesOther.Name = "rdoSidesOther"
		Me.rdoSidesOther.Size = New System.Drawing.Size(10, 24)
		Me.rdoSidesOther.TabIndex = 9
		'
		'pnlRolls
		'
		Me.pnlRolls.Controls.Add(Me.txtRolls)
		Me.pnlRolls.Controls.Add(Me.rdoRolls6)
		Me.pnlRolls.Controls.Add(Me.rdoRolls4)
		Me.pnlRolls.Controls.Add(Me.rdoRolls3)
		Me.pnlRolls.Controls.Add(Me.rdoRolls5)
		Me.pnlRolls.Controls.Add(Me.rdoRolls1)
		Me.pnlRolls.Controls.Add(Me.lblRolls)
		Me.pnlRolls.Controls.Add(Me.rdoRolls2)
		Me.pnlRolls.Controls.Add(Me.rdoRollsOther)
		Me.pnlRolls.Location = New System.Drawing.Point(10, 10)
		Me.pnlRolls.Name = "pnlRolls"
		Me.pnlRolls.Size = New System.Drawing.Size(70, 160)
		Me.pnlRolls.TabIndex = 10
		'
		'txtRolls
		'
		Me.txtRolls.Location = New System.Drawing.Point(30, 140)
		Me.txtRolls.Name = "txtRolls"
		Me.txtRolls.Size = New System.Drawing.Size(30, 20)
		Me.txtRolls.TabIndex = 8
		Me.txtRolls.Text = "7"
		'
		'rdoRolls6
		'
		Me.rdoRolls6.Location = New System.Drawing.Point(10, 120)
		Me.rdoRolls6.Name = "rdoRolls6"
		Me.rdoRolls6.Size = New System.Drawing.Size(40, 24)
		Me.rdoRolls6.TabIndex = 9
		Me.rdoRolls6.Text = "6"
		'
		'rdoRolls4
		'
		Me.rdoRolls4.Location = New System.Drawing.Point(10, 80)
		Me.rdoRolls4.Name = "rdoRolls4"
		Me.rdoRolls4.Size = New System.Drawing.Size(40, 24)
		Me.rdoRolls4.TabIndex = 4
		Me.rdoRolls4.Text = "4"
		'
		'rdoRolls3
		'
		Me.rdoRolls3.Checked = True
		Me.rdoRolls3.Location = New System.Drawing.Point(10, 60)
		Me.rdoRolls3.Name = "rdoRolls3"
		Me.rdoRolls3.Size = New System.Drawing.Size(40, 24)
		Me.rdoRolls3.TabIndex = 3
		Me.rdoRolls3.TabStop = True
		Me.rdoRolls3.Text = "3"
		'
		'rdoRolls5
		'
		Me.rdoRolls5.Location = New System.Drawing.Point(10, 100)
		Me.rdoRolls5.Name = "rdoRolls5"
		Me.rdoRolls5.Size = New System.Drawing.Size(40, 24)
		Me.rdoRolls5.TabIndex = 6
		Me.rdoRolls5.Text = "5"
		'
		'rdoRolls1
		'
		Me.rdoRolls1.Location = New System.Drawing.Point(10, 20)
		Me.rdoRolls1.Name = "rdoRolls1"
		Me.rdoRolls1.Size = New System.Drawing.Size(40, 24)
		Me.rdoRolls1.TabIndex = 1
		Me.rdoRolls1.Text = "1"
		'
		'lblRolls
		'
		Me.lblRolls.AutoSize = True
		Me.lblRolls.Location = New System.Drawing.Point(0, 0)
		Me.lblRolls.Name = "lblRolls"
		Me.lblRolls.Size = New System.Drawing.Size(33, 16)
		Me.lblRolls.TabIndex = 0
		Me.lblRolls.Text = "Rolls:"
		'
		'rdoRolls2
		'
		Me.rdoRolls2.Location = New System.Drawing.Point(10, 40)
		Me.rdoRolls2.Name = "rdoRolls2"
		Me.rdoRolls2.Size = New System.Drawing.Size(40, 24)
		Me.rdoRolls2.TabIndex = 2
		Me.rdoRolls2.Text = "2"
		'
		'rdoRollsOther
		'
		Me.rdoRollsOther.Location = New System.Drawing.Point(10, 140)
		Me.rdoRollsOther.Name = "rdoRollsOther"
		Me.rdoRollsOther.Size = New System.Drawing.Size(10, 24)
		Me.rdoRollsOther.TabIndex = 7
		'
		'btnRoll
		'
		Me.btnRoll.Location = New System.Drawing.Point(170, 10)
		Me.btnRoll.Name = "btnRoll"
		Me.btnRoll.TabIndex = 11
		Me.btnRoll.Text = "Roll"
		'
		'lstResults
		'
		Me.lstResults.Location = New System.Drawing.Point(170, 40)
		Me.lstResults.Name = "lstResults"
		Me.lstResults.Size = New System.Drawing.Size(70, 95)
		Me.lstResults.TabIndex = 12
		'
		'lblTotalResults
		'
		Me.lblTotalResults.BackColor = System.Drawing.SystemColors.Control
		Me.lblTotalResults.Location = New System.Drawing.Point(170, 140)
		Me.lblTotalResults.Name = "lblTotalResults"
		Me.lblTotalResults.Size = New System.Drawing.Size(70, 16)
		Me.lblTotalResults.TabIndex = 15
		Me.lblTotalResults.Text = "0"
		Me.lblTotalResults.TextAlign = System.Drawing.ContentAlignment.TopRight
		'
		'lnkLewisMoten_com
		'
		Me.lnkLewisMoten_com.Location = New System.Drawing.Point(170, 160)
		Me.lnkLewisMoten_com.Name = "lnkLewisMoten_com"
		Me.lnkLewisMoten_com.Size = New System.Drawing.Size(90, 23)
		Me.lnkLewisMoten_com.TabIndex = 16
		Me.lnkLewisMoten_com.TabStop = True
		Me.lnkLewisMoten_com.Text = "LewisMoten.com"
		'
		'frmDice
		'
		Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
		Me.ClientSize = New System.Drawing.Size(254, 176)
		Me.Controls.Add(Me.lnkLewisMoten_com)
		Me.Controls.Add(Me.lblTotalResults)
		Me.Controls.Add(Me.lstResults)
		Me.Controls.Add(Me.btnRoll)
		Me.Controls.Add(Me.pnlRolls)
		Me.Controls.Add(Me.pnlSides)
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
		Me.MaximizeBox = False
		Me.Name = "frmDice"
		Me.Text = "Lewies Dice"
		Me.pnlSides.ResumeLayout(False)
		Me.pnlRolls.ResumeLayout(False)
		Me.ResumeLayout(False)

	End Sub

#End Region

	Private Sub SyncForm()
		Me.Text = "Lewies Dice: " & Rolls.ToString & "d" & Sides.ToString
		Me.lstResults.Items.Clear()
		Me.lblTotalResults.Text = "0"
	End Sub
	Private Sub btnRoll_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRoll.Click

		lstResults.Items.Clear()

		Dim Total As Integer = 0

		For roll As Integer = 1 To Rolls

			Dim Result As Integer = CType(Math.Floor(Sides * Rnd()), Integer) + 1

			Total += Result

			lstResults.Items.Add(Suffixed(roll) & " roll: " & Result.ToString)

		Next

		lblTotalResults.Text = Total.ToString

	End Sub

	Private Function Suffixed(ByVal number As Integer) As String
		Dim s As String = number.ToString

		If s.EndsWith("11") Then Return s & "th"
		If s.EndsWith("12") Then Return s & "th"
		If s.EndsWith("13") Then Return s & "th"
		If s.EndsWith("1") Then Return s & "st"
		If s.EndsWith("2") Then Return s & "nd"
		If s.EndsWith("3") Then Return s & "rd"
		Return s & "th"

	End Function
	Private Sub rdoSides4_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rdoSides4.CheckedChanged
		Sides = 4
		SyncForm()
	End Sub

	Private Sub rdoSides6_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rdoSides6.CheckedChanged
		Sides = 6
		SyncForm()
	End Sub

	Private Sub rdoSides8_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rdoSides8.CheckedChanged
		Sides = 8
		SyncForm()
	End Sub

	Private Sub rdoSides10_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rdoSides10.CheckedChanged
		Sides = 10
		SyncForm()
	End Sub

	Private Sub rdoSides12_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rdoSides12.CheckedChanged
		Sides = 12
		SyncForm()
	End Sub

	Private Sub rdoSides20_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rdoSides20.CheckedChanged
		Sides = 20
		SyncForm()
	End Sub

	Private Sub rdoSidesOther_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rdoSidesOther.CheckedChanged
		Sides = Integer.Parse(txtSides.Text)
		SyncForm()
	End Sub

	Private Sub rdoRolls1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rdoRolls1.CheckedChanged
		Rolls = 1
		SyncForm()
	End Sub

	Private Sub rdoRolls2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rdoRolls2.CheckedChanged
		Rolls = 2
		SyncForm()
	End Sub

	Private Sub rdoRolls3_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rdoRolls3.CheckedChanged
		Rolls = 3
		SyncForm()
	End Sub

	Private Sub rdoRolls4_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rdoRolls4.CheckedChanged
		Rolls = 4
		SyncForm()
	End Sub

	Private Sub rdoRolls5_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rdoRolls5.CheckedChanged
		Rolls = 5
		SyncForm()
	End Sub

	Private Sub rdoRollsOther_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rdoRollsOther.CheckedChanged
		Rolls = Integer.Parse(txtRolls.Text)
		SyncForm()
	End Sub

	Private Sub rdoRolls6_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rdoRolls6.CheckedChanged
		Rolls = 6
		SyncForm()
	End Sub

	Private Sub txtRolls_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtRolls.Validating
		If Not IsNumeric(Me.txtRolls.Text) Then
			e.Cancel = True
		End If
		If rdoRollsOther.Checked Then
			Rolls = Integer.Parse(txtRolls.Text)
			SyncForm()
		End If
	End Sub

	Private Sub txtSides_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtSides.Validating
		If Not IsNumeric(Me.txtSides.Text) Then
			e.Cancel = True
		End If
		If rdoSidesOther.Checked Then
			Sides = Integer.Parse(txtSides.Text)
			SyncForm()
		End If
	End Sub

	Private Sub Form1_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load

		Dim Rolls As Object = Application.UserAppDataRegistry.GetValue("Rolls")
		Dim Sides As Object = Application.UserAppDataRegistry.GetValue("Sides")

		If Not Rolls Is Nothing Then
			Me.Rolls = CType(Rolls, Integer)
			Select Case Me.Rolls
				Case 1
					Me.rdoRolls1.Checked = True
				Case 2
					Me.rdoRolls2.Checked = True
				Case 3
					Me.rdoRolls3.Checked = True
				Case 4
					Me.rdoRolls4.Checked = True
				Case 5
					Me.rdoRolls5.Checked = True
				Case 6
					Me.rdoRolls6.Checked = True
				Case Else
					Me.rdoRollsOther.Checked = True
					Me.txtRolls.Text = Rolls.ToString
			End Select
		End If
		If Not Sides Is Nothing Then
			Me.Sides = CType(Sides, Integer)
			Select Case Me.Sides
				Case 4
					Me.rdoSides4.Checked = True
				Case 6
					Me.rdoSides6.Checked = True
				Case 8
					Me.rdoSides8.Checked = True
				Case 10
					Me.rdoSides10.Checked = True
				Case 12
					Me.rdoSides12.Checked = True
				Case 20
					Me.rdoSides20.Checked = True
				Case Else
					Me.rdoSidesOther.Checked = True
					Me.txtSides.Text = Sides.ToString
			End Select
		End If

		SyncForm()

		Dim Top As Object = Application.UserAppDataRegistry.GetValue("Top")
		Dim Left As Object = Application.UserAppDataRegistry.GetValue("Left")

		If Not Top Is Nothing Then Me.Top = CType(Top, Integer)
		If Not Left Is Nothing Then Me.Left = CType(Left, Integer)

		Me.Icon = New Drawing.Icon(Me.GetType.Assembly.GetManifestResourceStream(Me.GetType.Assembly.GetName.Name & ".LewiesDice.ico"))

	End Sub

	Private Sub Form1_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
		Application.UserAppDataRegistry.SetValue("Top", Me.Top)
		Application.UserAppDataRegistry.SetValue("Left", Me.Left)
		Application.UserAppDataRegistry.SetValue("Rolls", Rolls)
		Application.UserAppDataRegistry.SetValue("Sides", Sides)
	End Sub

	Private Sub lnkLewisMoten_com_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles lnkLewisMoten_com.LinkClicked
		Shell("explorer ""http://www.lewismoten.com""")
	End Sub
End Class
