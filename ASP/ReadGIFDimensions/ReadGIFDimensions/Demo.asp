<%
Dim Width
Dim Height
Dim Path

Width = -1
Height = -1

Path = Request.QueryString("Path")
If Path = "" Then Path = "Demo.gif"
Call ReadGIF(Path)

%>
<HTML>
	<HEAD>
		<TITLE>Read GIF Dimensions</TITLE>
	</HEAD>
	<BODY>
		<H1>Read GIF Dimensions</H1>
		<P>
			This demonstration will attempt to
			read a GIF file and return the
			dimensions.
		</P>
		<FORM>
			<INPUT name="Path" value="<%=Path%>"><BR>
			<INPUT type="Submit" value="Get Dimensions">
		</FORM>
		Width: <%=Width%><BR>
		Height: <%=Height%><BR>
		<BR>
		<BR>
		Created By <A href="http://www.lewismoten.com">Lewis Moten</A>.
	</BODY>
</HTML>
<%
Sub ReadGIF(ByRef pStrPath)
	
	Dim lObjFSO
	Dim lStrData
	Dim lStrPath

	If Not UCase(Right(pStrPath, 4)) = ".GIF" Then Exit Sub
	
	If InStr(1, pStrPath, "\") = 0 Then
		lStrPath = Server.MapPath(pStrPath)
	Else
		lStrPath = pStrPath
	End If
		
	Set lObjFSO = Server.CreateObject("Scripting.FileSystemObject")
	If lObjFSO.FileExists(lStrPath) Then
		lStrData = lObjFSO.OpenTextFile(lStrPath).ReadAll()
	End If
	Set lObjFSO = Nothing
	If Len(lStrData) < 10 Then Exit Sub		
	If Not Left(lStrData, 3) = "GIF" Then Exit Sub
	Width = CInt("&h" & _
		Right("0" & Hex(Asc(Mid(lStrData, 8))), 2) & _
		Right("0" & Hex(Asc(Mid(lStrData, 7))), 2) _
		)
			
	Height = CInt("&h" & _
		Right("0" & Hex(Asc(Mid(lStrData, 10))), 2) & _
		Right("0" & Hex(Asc(Mid(lStrData, 9))), 2) _
		)
End Sub
%>