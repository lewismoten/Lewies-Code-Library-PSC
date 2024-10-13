<!--#INCLUDE FILE="clsDatabase.asp"-->
<!--#INCLUDE FILE="clsUpload.asp"-->
<%
If Not Session("IsAdmin") Then Response.Redirect("admin.Login.asp")

Dim objDB
Dim strSQL
Dim objUpload
Dim objConn
Dim objRs
Dim strTitle
Dim strDescription
Dim lngCategoryID
Dim binImageData

Set objUpload = New clsUpload

strTitle		= objUpload.Field("Title").Value
strDescription	= objUpload.Field("Description").Value
lngCategoryID	= objUpload.Field("CategoryID").Value
binImageData	= objUpload.Field("ImageData").BinaryData

Set objUpload = Nothing


Set objConn = Server.CreateObject("ADODB.Connection")
Set objRs = Server.CreateObject("ADODB.Recordset")
Set objDB = New clsDatabase
objConn.Open objDB.ConnectionString
Set objDB = Nothing

' Select an empty record scheme
strSQL = _
	"SELECT " & _
		"[ImageID], " & _
		"[CategoryID], " & _
		"[Title], " & _
		"[Description], " & _
		"[DateCreated], " & _
		"[ImageData] " & _
	"FROM " & _
		"[Images] " & _
	"WHERE " & _
		"1 = 2"

objRs.Open strSQL, objConn, 3, 3

objRs.AddNew

If Not lngCategoryID = "" Then
	objRs.Fields("CategoryID")		= lngCategoryID
End If

objRs.Fields("Title")			= strTitle
objRs.Fields("Description")		= strDescription
objRs.Fields("ImageData").AppendChunk = binImageData & ChrB(0) ' terminator

objRs.Update

objRs.Close
	
Set objRs = Nothing
Set objConn = Nothing



Response.Redirect("admin.asp")
%>
