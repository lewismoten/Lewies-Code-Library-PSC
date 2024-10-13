<%

Const TristateUseDefault = -2	' Opens the file using the system default. 
Const TristateTrue = -1			' Opens the file as Unicode. 
Const TristateFalse = 0			' Opens the file as ASCII. 

Const ForReading = 1			' Open a file for reading only. You can't write to this file. 
Const ForWriting = 2			' Open a file for writing only. You can't read from this file. 
Const ForAppending = 8			' Open a file and write to the end of the file. 

FileName = Request.QueryString("File")

Dim RowNumbers
Dim HexData
Dim CharacterData

If Not FileName = "" Then

	If InStr(1, FileName, ":\") = 0 Then
		FileName = Server.MapPath(FileName)
	End If

	Set FSO = Server.CreateObject("Scripting.FileSystemObject")
	If FSO.FileExists(FileName) Then
		Set File = FSO.OpenTextFile(FileName, ForReading, False, TristateFalse)
		i = 0
		RowNumbers = Right("00000000" & Hex(i), 8) & "h:<BR>"
		CharacterData = "; "
		While Not File.AtEndOfStream

			i = i + 1
			
			Ascii = Asc(File.Read(1))
			HexData = HexData & Right("0" & Hex(Ascii), 2) & " "
			If Ascii => 20 Then
				CharacterData = CharacterData & Server.HTMLEncode(Chr(Ascii))
			Else
				CharacterData = CharacterData & "."
			End If
			If i Mod 16 = 0 Then 
				HexData = HexData & "<BR>"
				CharacterData = CharacterData & "<BR>; "
				RowNumbers = RowNumbers & Right("00000000" & Hex(i), 8) & "h:<BR>"
			End If

		Wend
		File.Close
			
		Set File = Nothing
	Else
		Response.Write("File does not exist")
	End If
	Set FSO = Nothing

End If
%>
<HTML>
	<HEAD>
		<TITLE>Hex Viewer Demonstration</TITLE>
	</HEAD>
	<BODY>
		<H1>Hex Viewer Demonstration</H1>
		<P>
			Type the file name in the field below to
			view the binary contents of the data.
		</P>
		<FORM>
			<INPUT name="File" value="<%=Request.QueryString("File")%>">
			<INPUT type="submit" value="View Hex">
		</FORM>
		<HR>
		<TABLE>
			<TR>
				<TD><PRE><%=RowNumbers%></PRE></TD>
				<TD><PRE><%=HexData%></PRE></TD>
				<TD><PRE><%=CharacterData%></PRE></TD>
			</TR>
		</TABLE>
		<HR>
		Created By <A href="http://www.lewismoten.com">Lewis Moten</A>.
	</BODY>
</HTML>