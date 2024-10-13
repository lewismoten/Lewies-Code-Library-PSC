<!--#INCLUDE FILE="clsDatabase.asp"-->
<%If Not Session("IsAdmin") Then Response.Redirect("admin.Login.asp")%>

<%
Dim objDB
Dim strSQL
Dim lngCategoryID
Dim varCategoryAry

lngCategoryID = Request.Form("CategoryID")

Set objDB = New clsDatabase

' Grab Users
strSQL = _
	"SELECT " & _
		"[CategoryName] " & _
	"FROM " & _
		"[Categories] " & _
	"WHERE " & _
		"[CategoryID] = " & lngCategoryID
		
Call objDB.SetData(strSQL, varCategoryAry)

If IsArray(varCategoryAry) Then
	strCategoryName = varCategoryAry(0, 0)
End If

Set objDB = Nothing
%>
<CENTER><H3><A href="http://www.lewismoten.com/">Lewie's</A> Photo Gallery Administration</H3><H5><A href="admin.asp">Home</A>: Edit Category</H5></CENTER>

<FORM method="post" action="admin.Category.Edit.Now.asp" id=form1 name=form1>
	<INPUT type="hidden" name="CategoryID" value="<%=lngCategoryID%>">
	<TABLE>
		<TR>
			<TD>Name</TD>
			<TD><INPUT name="CategoryName" value="<%=Server.HTMLEncode(strCategoryName)%>"></TD>
		</TR>
		<TR>
			<TD>&nbsp;</TD>
			<TD><INPUT type="Submit" value="Save" id=Submit2 name=Submit2></TD>
		</TR>
	</TABLE>
</FORM>
<HR>
Press this button to delete this category.  This can not be undone.
<FORM method="post" action="admin.Category.Delete.Now.asp" id=form2 name=form2>
	<INPUT type="hidden" name="CategoryID" value="<%=lngCategoryID%>">
	<INPUT type="Submit" value="Delete" id=Submit1 name=Submit1>
</FORM>