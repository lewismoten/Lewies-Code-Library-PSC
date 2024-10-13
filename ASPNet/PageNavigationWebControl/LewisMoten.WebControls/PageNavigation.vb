Imports System.ComponentModel
Imports System.Web.UI
Imports System.Web.UI.design
Imports System.Web.HttpUtility
Imports System.Web

< _
DefaultProperty("PageCount"), _
ToolboxData("<{0}:PageNavigation runat=server></{0}:PageNavigation>")> _
Public Class PageNavigation
	Inherits System.Web.UI.WebControls.WebControl

#Region " Private Member Declarations "

	Private _CurrentPage As Integer = 3
	Private _PageCount As Integer = 6

	Private _Url As String = ""

	Private _LinkDivider As String = " | "
	Private _SelectedPage As String = "[<B>{0}</B>]"
	Private _UnselectedPage As String = "<A href=""{1}"">{0}</A>"
	Private _BackPage As String = "<A href=""{1}"">Back</A>"
	Private _NextPage As String = "<A href=""{1}"">Next</A>"

#End Region

#Region " HTML Properties "

	< _
	Bindable(True), _
	PersistenceMode(PersistenceMode.Attribute), _
	Category("HTML"), _
	Description("HTML that appears betwen each link.  Use {0} for page number and {1} for URL."), _
	DefaultValue(" | ") _
	> _
	Public Property LinkDivider() As String
		Get
			Return Me._LinkDivider
		End Get
		Set(ByVal Value As String)
			Me._LinkDivider = Value
		End Set
	End Property

	< _
	Bindable(True), _
	PersistenceMode(PersistenceMode.Attribute), _
	Category("HTML"), _
	Description("HTML that represents the current selected page.  Use {0} for page number and {1} for URL."), _
	DefaultValue("[<B>{0}</B>]") _
	> _
	Public Property SelectedPage() As String
		Get
			Return Me._SelectedPage
		End Get
		Set(ByVal Value As String)
			Me._SelectedPage = Value
		End Set
	End Property

	< _
	Bindable(True), _
	PersistenceMode(PersistenceMode.Attribute), _
	Category("HTML"), _
	Description("HTML that represents links to pages that have not been selected.  Use {0} for page number and {1} for URL."), _
	DefaultValue("<A href=""{1}"">{0}</A>") _
	> _
	Public Property UnselectedPage() As String
		Get
			Return Me._UnselectedPage
		End Get
		Set(ByVal Value As String)
			Me._UnselectedPage = Value
		End Set
	End Property

	< _
	Bindable(True), _
	PersistenceMode(PersistenceMode.Attribute), _
	Category("HTML"), _
	Description("HTML that represents link to previous page.  Use {0} for page number and {1} for URL."), _
	DefaultValue("<A href=""{1}"">Back</A>") _
	> _
	Public Property BackPage() As String
		Get
			Return Me._BackPage
		End Get
		Set(ByVal Value As String)
			Me._BackPage = Value
		End Set
	End Property

	< _
	Bindable(True), _
	PersistenceMode(PersistenceMode.Attribute), _
	Category("HTML"), _
	Description("HTML that represents link to next page.  Use {0} for page number and {1} for URL."), _
	DefaultValue("<A href=""{1}"">Next</A>") _
	> _
	Public Property NextPage() As String
		Get
			Return Me._NextPage
		End Get
		Set(ByVal Value As String)
			Me._NextPage = Value
		End Set
	End Property

#End Region

#Region " Navigation Properties "

	< _
	Bindable(True), _
	PersistenceMode(PersistenceMode.Attribute), _
	Category("Navigation"), _
	Description("Selected page that is currently being viewed."), _
	DefaultValue(GetType(Integer), "1") _
	> _
	Public Property CurrentPage() As Integer
		Get
			Return Me._CurrentPage
		End Get
		Set(ByVal Value As Integer)
			Me._CurrentPage = Value
		End Set
	End Property

	< _
	Bindable(True), _
	PersistenceMode(PersistenceMode.Attribute), _
	Category("Navigation"), _
	Description("Total number of pages available."), _
	DefaultValue(GetType(Integer), "1") _
	> _
	Public Property PageCount() As Integer
		Get
			Return Me._PageCount
		End Get
		Set(ByVal Value As Integer)
			Me._PageCount = Value
		End Set
	End Property

#End Region

#Region " Private Properties "

	Private ReadOnly Property Url() As String
		Get
			If Not Me._Url = "" Then Return Me._Url

			If context Is Nothing Then
				Me._Url = "/DesignTime.aspx?Page="
				Return Me._Url
			End If

			Dim Request As HttpRequest = Context.Current.Request

			Me._Url = Request.Url.AbsolutePath

			Me._Url &= "?"

			For Each key As String In Request.QueryString.Keys
				If Not key.ToLower = "page" Then
					Me._Url &= UrlEncode(key)
					Me._Url &= "="
					Me._Url &= UrlEncode(Request.QueryString(key))
					Me._Url &= "&"
				End If
			Next

			Me._Url &= "Page="

			Return Me._Url

		End Get
	End Property
#End Region

#Region " Rendering "

	Protected Overrides Sub RenderContents(ByVal writer As System.Web.UI.HtmlTextWriter)
		Me.RenderRuntime(writer)
	End Sub

	Private Sub RenderRuntime(ByVal writer As System.Web.UI.HtmlTextWriter)

		'Try

		If Me.PageCount < 2 Then
			If context Is Nothing Then
				writer.Write("[<B><FONT color=""navy"">" & Me.ID & "</FONT></B>]")
			End If
			Return
		End If

		If Me.CurrentPage > 1 Then
			writer.Write(String.Format(BackPage, page, Url & (CurrentPage - 1).ToString))
			writer.Write(LinkDivider)
		End If

		For page As Integer = 1 To Me.PageCount
			If Not page = 1 Then writer.Write(LinkDivider)
			If page = Me.CurrentPage Then
				writer.Write(String.Format(SelectedPage, page, Url & page.ToString))
			Else
				writer.Write(String.Format(UnselectedPage, page, Url & page.ToString))
			End If
		Next

		If Me.CurrentPage < Me.PageCount Then
			writer.Write(LinkDivider)
			writer.Write(String.Format(NextPage, page, Url & (CurrentPage + 1).ToString))
		End If

		'Catch ex As Exception
		'	writer.Write(HtmlEncode(ex.Message))
		'	writer.Write("<HR>")
		'	writer.Write(HtmlEncode(ex.StackTrace))
		'End Try
	End Sub

#End Region

End Class
