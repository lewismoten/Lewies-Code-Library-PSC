Public Class GamePad
    Inherits System.Windows.Forms.UserControl
    Public Event Clicked(ByVal Index As Integer)

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
    Friend WithEvents pnlYellow As System.Windows.Forms.Panel
    Friend WithEvents pnlBlue As System.Windows.Forms.Panel
    Friend WithEvents pnlGreen As System.Windows.Forms.Panel
    Friend WithEvents pnlRed As System.Windows.Forms.Panel
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.pnlYellow = New System.Windows.Forms.Panel
        Me.pnlBlue = New System.Windows.Forms.Panel
        Me.pnlGreen = New System.Windows.Forms.Panel
        Me.pnlRed = New System.Windows.Forms.Panel
        Me.SuspendLayout()
        '
        'pnlYellow
        '
        Me.pnlYellow.BackColor = System.Drawing.Color.Yellow
        Me.pnlYellow.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.pnlYellow.Location = New System.Drawing.Point(32, 64)
        Me.pnlYellow.Name = "pnlYellow"
        Me.pnlYellow.Size = New System.Drawing.Size(32, 32)
        Me.pnlYellow.TabIndex = 7
        '
        'pnlBlue
        '
        Me.pnlBlue.BackColor = System.Drawing.Color.Blue
        Me.pnlBlue.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.pnlBlue.Location = New System.Drawing.Point(0, 32)
        Me.pnlBlue.Name = "pnlBlue"
        Me.pnlBlue.Size = New System.Drawing.Size(32, 32)
        Me.pnlBlue.TabIndex = 6
        '
        'pnlGreen
        '
        Me.pnlGreen.BackColor = System.Drawing.Color.Lime
        Me.pnlGreen.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.pnlGreen.Location = New System.Drawing.Point(64, 32)
        Me.pnlGreen.Name = "pnlGreen"
        Me.pnlGreen.Size = New System.Drawing.Size(32, 32)
        Me.pnlGreen.TabIndex = 5
        '
        'pnlRed
        '
        Me.pnlRed.BackColor = System.Drawing.Color.Red
        Me.pnlRed.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.pnlRed.Location = New System.Drawing.Point(32, 0)
        Me.pnlRed.Name = "pnlRed"
        Me.pnlRed.Size = New System.Drawing.Size(32, 32)
        Me.pnlRed.TabIndex = 4
        '
        'GamePad
        '
        Me.Controls.Add(Me.pnlYellow)
        Me.Controls.Add(Me.pnlBlue)
        Me.Controls.Add(Me.pnlGreen)
        Me.Controls.Add(Me.pnlRed)
        Me.Name = "GamePad"
        Me.Size = New System.Drawing.Size(96, 96)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private _Pattern() As Byte

    Private Sub pnlRed_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles pnlRed.Click
        BlinkPanel(0)
        RaiseEvent Clicked(0)
    End Sub

    Private Sub pnlBlue_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles pnlBlue.Click
        BlinkPanel(1)
        RaiseEvent Clicked(1)
    End Sub

    Private Sub pnlGreen_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles pnlGreen.Click
        BlinkPanel(2)
        RaiseEvent Clicked(2)
    End Sub

    Private Sub pnlYellow_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles pnlYellow.Click
        BlinkPanel(3)
        RaiseEvent Clicked(3)
    End Sub

    Public Property Pattern() As Byte()
        Get
            Return _Pattern
        End Get
        Set(ByVal Value As Byte())
            ' Do animation
            _Pattern = Value
            For i As Integer = 0 To _Pattern.Length - 1
                BlinkPanel(_Pattern(i))
            Next
        End Set
    End Property
    Private Sub BlinkPanel(ByVal index As Integer)
        Select Case index
            Case 0
                pnlRed.BackColor = Color.White
                Application.DoEvents()
                System.Threading.Thread.Sleep(500)
                pnlRed.BackColor = Color.Red
                Application.DoEvents()
                System.Threading.Thread.Sleep(250)
            Case 1
                pnlBlue.BackColor = Color.White
                Application.DoEvents()
                System.Threading.Thread.Sleep(500)
                pnlBlue.BackColor = Color.Blue
                Application.DoEvents()
                System.Threading.Thread.Sleep(250)
            Case 2
                pnlGreen.BackColor = Color.White
                Application.DoEvents()
                System.Threading.Thread.Sleep(500)
                pnlGreen.BackColor = Color.LightGreen
                Application.DoEvents()
                System.Threading.Thread.Sleep(250)
            Case 3
                pnlYellow.BackColor = Color.White
                Application.DoEvents()
                System.Threading.Thread.Sleep(500)
                pnlYellow.BackColor = Color.Yellow
                Application.DoEvents()
                System.Threading.Thread.Sleep(250)
        End Select

    End Sub

End Class
