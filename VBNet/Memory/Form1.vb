Public Class Form1
    Inherits System.Windows.Forms.Form
    Private _Score As Integer = 0
    Private _HighScore As Integer = 0
    Private _Position As Integer = 0

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
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents lblHighScore As System.Windows.Forms.Label
    Friend WithEvents lblScore As System.Windows.Forms.Label
    Friend WithEvents btnNewGame As System.Windows.Forms.Button
    Friend WithEvents GamePad1 As Memory.GamePad
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.lblHighScore = New System.Windows.Forms.Label
        Me.lblScore = New System.Windows.Forms.Label
        Me.btnNewGame = New System.Windows.Forms.Button
        Me.GamePad1 = New Memory.GamePad
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(128, 8)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(64, 16)
        Me.Label1.TabIndex = 4
        Me.Label1.Text = "High Score:"
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(128, 32)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(64, 23)
        Me.Label2.TabIndex = 5
        Me.Label2.Text = "Score:"
        '
        'lblHighScore
        '
        Me.lblHighScore.Location = New System.Drawing.Point(200, 8)
        Me.lblHighScore.Name = "lblHighScore"
        Me.lblHighScore.Size = New System.Drawing.Size(48, 23)
        Me.lblHighScore.TabIndex = 6
        Me.lblHighScore.Text = "0"
        Me.lblHighScore.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'lblScore
        '
        Me.lblScore.Location = New System.Drawing.Point(200, 32)
        Me.lblScore.Name = "lblScore"
        Me.lblScore.Size = New System.Drawing.Size(48, 23)
        Me.lblScore.TabIndex = 7
        Me.lblScore.Text = "0"
        Me.lblScore.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'btnNewGame
        '
        Me.btnNewGame.Location = New System.Drawing.Point(176, 80)
        Me.btnNewGame.Name = "btnNewGame"
        Me.btnNewGame.TabIndex = 8
        Me.btnNewGame.Text = "New Game"
        '
        'GamePad1
        '
        Me.GamePad1.Location = New System.Drawing.Point(8, 8)
        Me.GamePad1.Name = "GamePad1"
        Me.GamePad1.Size = New System.Drawing.Size(96, 96)
        Me.GamePad1.TabIndex = 9
        '
        'Form1
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(256, 110)
        Me.Controls.Add(Me.GamePad1)
        Me.Controls.Add(Me.btnNewGame)
        Me.Controls.Add(Me.lblScore)
        Me.Controls.Add(Me.lblHighScore)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Name = "Form1"
        Me.Text = "Memory"
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub btnNewGame_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNewGame.Click
        Me.GamePad1.Enabled = True
        Score = 0
        GamePad1.Pattern = New Byte() {}
        NextPattern()
    End Sub

    Private Sub NextPattern()
        _Position = 0
        Dim Rnd As New System.Random
        Dim Pattern As Byte() = GamePad1.Pattern
        ReDim Preserve Pattern(Pattern.Length)
        Dim Chances As Integer = Rnd.Next(1, 400)
        Dim Digit As Integer = 0
        If Chances > 99 Then Digit += 1
        If Chances > 199 Then Digit += 1
        If Chances > 299 Then Digit += 1
        Pattern(Pattern.Length - 1) = Digit
        GamePad1.Pattern = Pattern
    End Sub

    Private Property Score() As Integer
        Get
            Return _Score
        End Get
        Set(ByVal Value As Integer)
            _Score = Value
            Me.lblScore.Text = Value
            If _Score > _HighScore Then
                _HighScore = _Score
                Me.lblHighScore.Text = _HighScore
                ' TODO: Save High Score, Load High Score
            End If
        End Set
    End Property

    Private Sub GamePad1_Clicked(ByVal Index As Integer) Handles GamePad1.Clicked
        If GamePad1.Pattern Is Nothing Then Return
        If _Position >= GamePad1.Pattern.Length Then Return

        If GamePad1.Pattern(_Position) = Index Then
            Score += 1
            If _Position = GamePad1.Pattern.Length - 1 Then
                NextPattern()
            Else
                _Position += 1
            End If
        Else
            Me.GamePad1.Enabled = False
        End If
    End Sub
End Class
