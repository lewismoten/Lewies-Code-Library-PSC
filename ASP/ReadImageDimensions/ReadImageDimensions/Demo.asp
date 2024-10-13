<%Option Explicit%>
<!--#INCLUDE FILE="clsUpload.asp"-->
<!--#INCLUDE FILE="clsImage.asp"-->
<HTML>
	<HEAD>
		<TITLE>Read Image Dimensions</TITLE>
	</HEAD>
	<BODY>
		<H1>Read Image Dimensions</H1>
		<P>
			Choose an image (JPG, GIF, BMP, or PNG) to upload and the
			script will attempt to acquire the correct height and width
			for the image.
		</P>
		<FORM encType="multipart/form-data" method="post">
			Read a file on the server: <INPUT name="Path"><BR>
			Or upload a file: <INPUT type="file" name="Image"><BR>
			<INPUT type="submit">
		</FORM>
		<%Call Main()%>
		<P>
			This demonstration does not actually save the uploaded
			file on the servers file system or into a database.  It
			only looks at the data that has been received and parses
			the information about it.  A picture will be linked to
			the file on your own computer.  When you upload a file,
			the physical location of that file is included with the
			file sent to the server.
		</P>
		Created by <A href="http://www.lewismoten.com">Lewis Moten</A>.
	</BODY>
</HTML>
<%

Sub Main()

	Dim lObjUpload
	Dim lObjImage
	
	Dim lStrPath
	
	Dim lBinImage_FileName
	Dim lBinImage_ContentType
	Dim lBinImage_Length
	Dim lBinImage_BinaryData
	Dim lBinImage_Height
	Dim lBinImage_Width
	Dim lBinImage_Type
	
	Set lObjUpload = New clsUpload
	lBinImage_FileName		= lObjUpload.Field("Image").FileName
	lBinImage_ContentType	= lObjUpload.Field("Image").ContentType
	lBinImage_Length		= lObjUpload.Field("Image").Length
	lBinImage_BinaryData	= lObjUpload.Field("Image").BinaryData
	lStrPath				= lObjUpload.Field("Path").Value
	Set lObjUpload = Nothing

	Set lObjImage = New clsImage

	If Not lBinImage_Length = 0 Then
		lObjImage.DataStream	= lBinImage_BinaryData
		lBinImage_Width			= lObjImage.Width
		lBinImage_Height		= lObjImage.Height
		lBinImage_Type			= lObjImage.ImageType
		lBinImage_ContentType	= lObjImage.ContentType
	ElseIf Not Len(lStrPath) = 0 Then
		lObjImage.Read(lStrPath)
		lBinImage_Width			= lObjImage.Width
		lBinImage_Height		= lObjImage.Height
		lBinImage_Type			= lObjImage.ImageType
		lBinImage_ContentType	= lObjImage.ContentType
		lBinImage_Length		= lObjImage.Size
		lBinImage_FileName		= lObjImage.Path
	End If
	
	Set lObjImage = Nothing

	If Not(lBinImage_Length = 0 And lStrPath = "") Then

		
		If lStrPath = "" Then
			Response.Write "Local File: <A href=""file:///" & Server.URLPathEncode(lBinImage_FileName) & """>" & lBinImage_FileName & "</A><BR>"
		Else
			If InStr(1, lStrPath, "\") = 0 Then
				Response.Write "Remote File URL: <A href=""" & lStrPath & """>" & lStrPath & "</A><BR>"
			End If
			Response.Write "Remote Physical Location: " & lBinImage_FileName & "<BR>"
		End If
		Response.Write "Content Type: " & lBinImage_ContentType & "<BR>"
		Response.Write "Type: " & lBinImage_Type & "<BR>"
		Response.Write "Size: " & lBinImage_Length & " bytes<BR>"
		Response.Write "Dimensions: " & lBinImage_Width & "x" & lBinImage_Height & "<BR><BR>" & vbCrLf
		If lStrPath = "" Then
			Response.Write "<IMG border=""1"" src=""file:///" & Server.URLPathEncode(lBinImage_FileName) & """>"
		Else
			Response.Write "<IMG border=""1"" src=""" & lStrPath & """>"
		End If

	End If
	
End Sub
%>