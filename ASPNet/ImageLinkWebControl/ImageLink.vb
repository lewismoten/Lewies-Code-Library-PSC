Imports System.ComponentModel
Imports System.Web.UI
Imports System.Web.UI.design
Imports System.Web.HttpUtility
Imports System.Web
Imports System.Web.UI.HtmlTextWriterAttribute
Imports System.Web.UI.HtmlTextWriterTag

Imports System.Drawing
Imports System.Drawing.Design

< _
DefaultProperty("NavigateUrl"), _
ToolboxData("<{0}:ImageLink runat=server></{0}:ImageLink>") _
> _
Public Class ImageLink
	Inherits System.Web.UI.WebControls.Image
	Private _NavigateUrl As String = ""

	< _
	Bindable(True), _
	Editor(GetType(System.Web.UI.Design.UrlEditor), GetType(System.Drawing.Design.UITypeEditor)), _
	PersistenceMode(PersistenceMode.Attribute), _
	Category("Navigation"), _
	Description("The Url to navigate to."), _
	DefaultValue("") _
	> _
	Public Property NavigateUrl() As String
		Get
			Return Me._NavigateUrl
		End Get
		Set(ByVal Value As String)
			Me._NavigateUrl = Value
		End Set
	End Property

	Protected Overrides Sub Render(ByVal writer As System.Web.UI.HtmlTextWriter)
		If Not Me.NavigateUrl = "" Then
			writer.AddAttribute(Href, Me.NavigateUrl)
			writer.RenderBeginTag(A)
		End If
		MyBase.Render(writer)
		If Not Me.NavigateUrl = "" Then
			writer.RenderEndTag()
		End If
	End Sub
End Class
