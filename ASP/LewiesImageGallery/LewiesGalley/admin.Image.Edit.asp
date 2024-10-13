<!--#INCLUDE FILE="clsDatabase.asp"-->
<%
If Not Session("IsAdmin") Then Response.Redirect("admin.Login.asp")
 
Dim objDB
Dim strSQL
Dim strCategoryOptions
Dim lngImageID
Dim strTitle
Dim strDescription
Dim lngCategoryID
Dim varImageAry

lngImageID = Request.Form("ImageID")

Set objDB = New clsDatabase

strSQL = _
	"SELECT " & _
		"[Title], " & _
		"[Description], " & _
		"[CategoryID] " & _
	"FROM " & _
		"[Images] " & _
	"WHERE " & _
		"[ImageID] = " & lngImageID

Call objDB.SetData(strSQL, varImageAry)

If IsArray(varImageAry) Then
	strTitle = varImageAry(0, 0)
	strDescription = varImageAry(1, 0)
	lngCategoryID = varImageAry(2, 0)
	
	If Not Len(strTitle) = 0 Then strTitle = Server.HTMLEncode(strTitle)
	If Not Len(strDescription) = 0 Then strDescription = server.HTMLEncode(strDescription)
	
End If
		
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

strCategoryOptions = Replace(strCategoryOptions, "value=""" & lngCategoryID & """", "value=""" & lngCategoryID & """ selected")

Set objDB = Nothing
%>

<CENTER><H3><A href="http://www.lewismoten.com/">Lewie's</A> Photo Gallery Administration</H3><H5><A href="admin.asp">Home</A>: Edit Image</H5></CENTER>

<CENTER>
	<IMG src="image.asp?ImageID=<%=lngImageID%>" border="1" width="100"><BR>
	<A href="image.asp?ImageID=<%=lngImageID%>">Full Size</A>
</CENTER>

<FORM method="post" action="admin.Image.Edit.Now.asp" encType="multipart/form-data" id=form1 name=form1>
	<INPUT type="hidden" name="ImageID" value="<%=lngImageID%>">
	<TABLE>
		<TR>
			<TD>Title</TD>
			<TD><INPUT name="Title" value="<%=strTitle%>"></TD>
		</TR>
		<TR>
			<TD>Description</TD>
			<TD><TEXTAREA name="Description" cols="50" rows="5"><%=strDescription%></TEXTAREA></TD>
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
				<INPUT type="radio" name="NewImage" value="False" checked> Keep existing image<BR>
				<INPUT type="radio" name="NewImage" value="True">Upload New Image<BR>
				&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT type="File" name="ImageData">
			</TD>
		</TR>
		<TR>
			<TD>&nbsp;</TD>
			<TD><INPUT type="Submit" value="Save" id=Submit1 name=Submit1></TD>
		</TR>
	</TABLE>
</FORM>
<HR>
Press this button to delete this image.  This can not be undone.
<FORM method="post" action="admin.Image.Delete.Now.asp" id=form2 name=form2>
	<INPUT type="hidden" name="ImageID" value="<%=lngImageID%>">
	<INPUT type="Submit" value="Delete" id=Submit1 name=Submit1>
</FORM>