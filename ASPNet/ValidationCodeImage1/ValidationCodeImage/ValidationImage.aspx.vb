Option Explicit On 
Option Strict On

Public Class ValidationImage
	Inherits System.Web.UI.Page
#Region " Web Form Designer Generated Code "

	'This call is required by the Web Form Designer.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

	End Sub

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

	Protected Overrides Sub Render(ByVal writer As System.Web.UI.HtmlTextWriter)
		Dim VCI As New ValidationCodeImage
		Response.BinaryWrite(VCI.Render())
		Session("Validation") = VCI.Text
	End Sub
End Class
