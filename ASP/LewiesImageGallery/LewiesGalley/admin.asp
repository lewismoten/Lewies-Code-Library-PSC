<!--#INCLUDE FILE="clsDatabase.asp"-->
<%If Not Session("IsAdmin") Then Response.Redirect("admin.Login.asp")%>

<%
Dim objDB
Dim strSQL
Dim strCategoryOptions
Dim strImageOptions
Dim strUserOptions

Set objDB = New clsDatabase

' Grab Categories
strSQL = _
	"SELECT " & _
		"[CategoryID], " & _
		"[CategoryName] " & _
	"FROM " & _
		"[Categories] " & _
	"ORDER BY " & _
		"[CategoryName] ASC"
Call objDB.SetOptions(strSQL, strCategoryOptions, True)

' Grab Images
strSQL = _
	"SELECT " & _
		"[ImageID], " & _
		"[Title] " & _
	"FROM " & _
		"[Images] " & _
	"ORDER BY " & _
		"[Title] ASC"
Call objDB.SetOptions(strSQL, strImageOptions, True)

' Grab Users
strSQL = _
	"SELECT " & _
		"[UserID], " & _
		"[LoginName] " & _
	"FROM " & _
		"[Users] " & _
	"ORDER BY " & _
		"[LoginName] ASC"
Call objDB.SetOptions(strSQL, strUserOptions, True)

Set objDB = Nothing
%>
<CENTER><H3><A href="http://www.lewismoten.com/">Lewie's</A> Photo Gallery Administration</H3><H5>Home</H5></CENTER>

<TABLE width="100%">
	<TR>
		<TD width="33%" align="center"><B>Categories</B></TD>
		<TD width="33%" align="center"><B>Images</B></TD>
		<TD width="33%" align="center"><B>Accounts</B></TD>
	</TR>
	<TR>
		<TD valign="top">
			<%If Not strCategoryOptions = "" Then%>
				<FORM method="post" action="admin.Category.Edit.asp">
					<SELECT name="CategoryID">
						<%=strCategoryOptions%>
					</SELECT>
					<INPUT type="submit" value="Edit">
				</FORM>
			<%Else%>
				No categories are present within the gallery at this time.
			<%End If%>
		</TD>
		<TD valign="top">
			<%If Not strImageOptions = "" Then%>
				<FORM method="post" action="admin.Image.Edit.asp" id=form1 name=form1>
					<SELECT name="ImageID">
						<%=strImageOptions%>
					</SELECT>
					<INPUT type="submit" value="Edit" id=submit1 name=submit1>
				</FORM>
			<%Else%>
				No images are present within the gallery at this time.
			<%End If%>
		</TD>
		<TD valign="top">
			<%If Not strUserOptions = "" Then%>
				<FORM method="post" action="admin.User.Edit.asp" id=form1 name=form1>
					<SELECT name="UserID">
						<%=strUserOptions%>
					</SELECT>
					<INPUT type="submit" value="Edit" id=submit1 name=submit1>
				</FORM>
			<%Else%>
				No users are present within the gallery at this time.
			<%End If%>
		</TD>
	</TR>
	<TR>
		<TD>
			<A href="admin.Category.Add.asp">Add Category</A>
		</TD>
		<TD>
			<A href="admin.Image.Add.asp">Add Image</A>
		</TD>
		<TD>
			<A href="admin.User.Add.asp">Add User</A>
		</TD>
	</TR>
</TABLE>