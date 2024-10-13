'**************************************
' Name: Choose File User Control
' Description:Drag and drop this control onto your forms so you can allow users to pick the files on there computer they want to use with your program. Pretty simple stuff.
' By: Lewis Moten
'
'
' Inputs:None
'
' Returns:None
'
'Assumes:None
'
'Side Effects:None
'This code is copyrighted and has limited warranties.
'**************************************
    
    Public Class ChooseFile
    	Inherits System.Windows.Forms.UserControl
    #Region " Windows Form Designer generated code "
    Public Sub New()
    MyBase.New()
    'This call is required by the Windows Fo
    //     rm Designer.
    InitializeComponent()
    'Add any initialization after the Initia
    //     lizeComponent() call
    End Sub
    'UserControl overrides dispose to clean 
    //     up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
    If disposing Then
    If Not (components Is Nothing) Then
    components.Dispose()
    End If
    End If
    MyBase.Dispose(disposing)
    End Sub
    //     'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer
    'NOTE: The following procedure is requir
    //     ed by the Windows Form Designer
    'It can be modified using the Windows Fo
    //     rm Designer. 
    'Do not modify it using the code editor.
    //     
    	Friend WithEvents btnBrowse As System.Windows.Forms.Button
    	Friend WithEvents txtFilePath As System.Windows.Forms.TextBox
    	Friend WithEvents OpenDialog As System.Windows.Forms.OpenFileDialog
    	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
    		Me.btnBrowse = New System.Windows.Forms.Button
    		Me.txtFilePath = New System.Windows.Forms.TextBox
    		Me.OpenDialog = New System.Windows.Forms.OpenFileDialog
    		Me.SuspendLayout()
    		'
    		'btnBrowse
    		'
    		Me.btnBrowse.Dock = System.Windows.Forms.DockStyle.Right
    		Me.btnBrowse.Location = New System.Drawing.Point(240, 0)
    		Me.btnBrowse.Name = "btnBrowse"
    		Me.btnBrowse.Size = New System.Drawing.Size(30, 20)
    		Me.btnBrowse.TabIndex = 0
    		Me.btnBrowse.Text = "..."
    		'
    		'txtFilePath
    		'
    		Me.txtFilePath.Dock = System.Windows.Forms.DockStyle.Fill
    		Me.txtFilePath.Location = New System.Drawing.Point(0, 0)
    		Me.txtFilePath.Name = "txtFilePath"
    		Me.txtFilePath.Size = New System.Drawing.Size(240, 20)
    		Me.txtFilePath.TabIndex = 1
    		Me.txtFilePath.Text = ""
    		'
    		'ChooseFile
    		'
    		Me.Controls.Add(Me.txtFilePath)
    		Me.Controls.Add(Me.btnBrowse)
    		Me.Name = "ChooseFile"
    		Me.Size = New System.Drawing.Size(270, 20)
    		Me.ResumeLayout(False)
    	End Sub
    #End Region
    	Private Sub btnBrowse_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnBrowse.Click
    		If Me.OpenDialog.ShowDialog(Me) = DialogResult.OK Then
    			Me.txtFilePath.Text = Me.OpenDialog.FileName
    		End If
    	End Sub
    	Public Property FileName() As String
    		Get
    			Return Me.txtFilePath.Text
    		End Get
    		Set(ByVal Value As String)
    			Me.txtFilePath.Text = Value
    		End Set
    	End Property
    	Public Property Filter() As String
    		Get
    			Return Me.OpenDialog.Filter
    		End Get
    		Set(ByVal Value As String)
    			Me.OpenDialog.Filter = Value
    		End Set
    	End Property
    End Class