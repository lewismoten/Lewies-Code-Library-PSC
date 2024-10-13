<!--#INCLUDE FILE="clsDatabase.asp"-->
<CENTER><H3>Photo Gallery</H3></CENTER>

<%
Dim objDB
Dim strSQL
Dim strCategoryOptions
Dim lngCategoryID
Dim varImageAry
Dim lngIndex
Dim lngMaxIndex
Dim lngImageID
Dim strTitle
Dim strDescription

lngCategoryID = Request.Form("CategoryID")

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

strCategoryOptions = Replace(strCategoryOptions, "value=""" & lngCategoryID & """", "value=""" & lngCategoryID & """ selected")

strSQL = _
	"SELECT " & _
		"[ImageID], " & _
		"[Title], " & _
		"[Description] " & _
	"FROM " & _
		"[Images] " & _
	"WHERE "
	
If lngCategoryID = "" Then
	strSQL = strSQL & "[CategoryID] IS NULL "
Else
	strSQL = strSQL & "[CategoryID] = " & lngCategoryID & " "
End If

strSQL = strSQL & "ORDER BY [Title] ASC"

Call objDB.SetData(strSQL, varImageAry)
		
Set objDB = Nothing
%>

<FORM method="post" action="default.asp">

	Select a category:
	<SELECT name="CategoryID">
		<OPTION value="">Miscellaneous</OPTION>
		<%=strCategoryOptions%>
	</SELECT>

	<INPUT type="Submit" value="View Images">
	
</FORM>
<HR>
<%
If IsArray(varImageAry) Then
	lngMaxIndex = UBound(varImageAry, 2)
	Response.Write("<TABLE>")
	For lngIndex = 0 To lngMaxIndex
		lngImageID = varImageAry(0, lngIndex)
		strTitle = varImageAry(1, lngIndex)
		strDescription = varImageAry(2, lngIndex)
		
		If Not Len(strTitle) = 0 Then strTitle = Server.HTMLEncode(strTitle)
		If Not Len(strDescription) = 0 Then strDescription = Server.HTMLEncode(strDescription)
		Response.Write("<TR>")
		Response.Write("<TD valign=""top"">")
		Response.Write("<IMG src=""image.asp?ImageID=" & lngImageID & """ width=""100"" border=""1"">")
		Response.Write("</TD>")
		Response.Write("<TD valign=""top"">")
		Response.Write("<A href=""image.asp?ImageID=" & lngImageID & """><B>" & strTitle & "</B></A><BR>")
		Response.Write(strDescription)
		Response.Write("</TD>")
		Response.Write("</TR>")
		
	Next
	
	Response.Write("</TABLE>")
Else
	Response.Write("No images are present in this category.")
End If
%>