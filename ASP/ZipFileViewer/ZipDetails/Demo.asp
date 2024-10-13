<!--#INCLUDE FILE="clsZip.asp"-->
<H1>Zip Viewer Demonstration</H1>
<H4>Created by Lewis Moten</H4>
<H5><A href="http://www.lewismoten.com">www.lewismoten.com</A></H5>
<%
FileName = Request.QueryString("File")
If FileName = "" Then FileName = "Bottom Tabs.zip"
%>
<FORM id=form1 name=form1>
	Zip File: <INPUT name="File" value="<%=Server.HTMLEncode(FileName)%>">
	<INPUT type="Submit" value="View" id=Submit1 name=Submit1>
</FORM>
<%
Dim FileName
Dim zip


set zip = new clszip
zip.ZipLoad(filename)
Dim nn
Set ZipFile = New clsZipFile
Response.Write "<TABLE width=""100%"">"
Response.Write "<TR bgcolor=""#cccccc"">"
Response.Write "<TD>Name</TD>"
Response.Write "<TD>Modified</TD>"
Response.Write "<TD>Size</TD>"
Response.Write "<TD>Ratio</TD>"
Response.Write "<TD>Packed</TD>"
Response.Write "<TD>Path</TD>"
Response.Write "</TR>"
For nn = 1 To zip.FileCount
	Set ZipFile = zip.GetFile(nn)
	With ZipFile
		If Not (.IsFolder Or .IsOverall) Then
			Response.Write "<TR>"
			Response.Write "<TD>" & .Name & "</TD>"
			Response.Write "<TD>" & .Modified & "</TD>"
			Response.Write "<TD>" & .Size & "</TD>"
			Response.Write "<TD>" & .Ratio & "</TD>"
			Response.Write "<TD>" & .Packed & "</TD>"
			Response.Write "<TD>" & .Path & "</TD>"
			Response.Write "</TR>"
		End If
	End With
Next
Response.Write "</TABLE>"
set zip = nothing
%>