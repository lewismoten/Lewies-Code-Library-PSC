Option Explicit On 
Option Strict On

Public Class DemoForm
	Inherits System.Web.UI.Page
#Region " Web Form Designer Generated Code "

    'This call is required by the Web Form Designer.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

	End Sub
	Protected WithEvents imgValidation As System.Web.UI.WebControls.Image
	Protected WithEvents txtCode As System.Web.UI.WebControls.TextBox
	Protected WithEvents btnOK As System.Web.UI.WebControls.Button
	Protected WithEvents lblMessage As System.Web.UI.WebControls.Label

    'NOTE: The following placeholder declaration is required by the Web Form Designer.
    'Do not delete or move it.
    Private designerPlaceholderDeclaration As System.Object

    Private Sub Page_Init(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Init
        'CODEGEN: This method call is required by the Web Form Designer
        'Do not modify it using the code editor.
        InitializeComponent()
    End Sub

#End Region

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
		'Put user code to initialize the page here
	End Sub

	Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click

		' make sure session variable exists
		If Session("Validation") Is Nothing Then
			lblMessage.Text = "Session ended.  Try again"
			Return
		End If

		Dim UserCode As String = Me.txtCode.Text
		Dim SessionCode As String = Session("Validation").ToString

		' if end-user matched correctly
		If UserCode.ToLower = SessionCode.ToLower Then

			' notify of success
			lblMessage.Text = "You are a human!"

			' hide everything else
			Me.imgValidation.Visible = False
			Me.txtCode.Visible = False
			Me.btnOK.Visible = False

		Else

			' display generic machine message as a joke
			lblMessage.Text = "00101010 11011010 10101010 10101100"

		End If

	End Sub
End Class
