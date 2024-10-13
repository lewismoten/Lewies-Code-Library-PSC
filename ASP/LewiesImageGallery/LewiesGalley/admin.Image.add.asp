<!--#INCLUDE FILE="clsDatabase.asp"-->
<%
If Not Session("IsAdmin") Then Response.Redirect("admin.Login.asp")
 
Dim objDB
Dim strSQL
Dim strCategoryOptions

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

Set objDB = Nothing
%>

<CENTER><H3><A href="http://www.lewismoten.com/">Lewie's</A> Photo Gallery Administration</H3><H5><A href="admin.asp">Home</A>: Add Image</H5></CENTER>

<FORM method="post" action="admin.Image.Add.Now.asp" encType="multipart/form-data">
	<TABLE>
		<TR>
			<TD>Title</TD>
			<TD><INPUT name="Title"></TD>
		</TR>
		<TR>
			<TD>Description</TD>
			<TD><TEXTAREA name="Description" cols="50" rows="5"></TEXTAREA></TD>
		</TR>
		<TR>
			<TD>Category</TD>
			<TD>
				<SELECT name="CategoryID">
					<OPTION value="">Miscellaneous</OPTION>
					<%=strCategoryOptions%>
				</SELECT>
			</TD>
		</TR>
		<TR>
			<TD>Image</TD>
			<TD>
				<INPUT type="File" name="ImageData">
			</TD>
		</TR>
		<TR>
			<TD>&nbsp;</TD>
			<TD><INPUT type="Submit" value="Add" id=Submit1 name=Submit1></TD>
		</TR>
	</TABLE>
</FORM>
