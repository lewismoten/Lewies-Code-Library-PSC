<%
SRC = Request.QueryString("SRC")
FirstHex = Request.QueryString("FirstHex")
LastHex = Request.QueryString("LastHex")
%>
<HTML>
	<HEAD>
		<TITLE>Dynamic GIF Image Demo</TITLE>
	</HEAD>
	<BODY>
		<H1>Dynamic GIF Image Demo</H1>
		<P>
			This demonstration allows you to load a template file and define the
			first and last colors in a palette.  The stepping colors in between
			are calculated for you and the end result is an image that can change
			on the fly!
		</P>
		<FORM name="MyForm">

			<SELECT name="PrePop" onChange="document.MyForm.SRC.value = this[this.selectedIndex].value">
				<OPTION value="">Predefined Images</OPTION>
				<OPTION value="GoButton.txt">Go Button</OPTION>
				<OPTION value="TopLeftCorner16.txt">Top Left Corner</OPTION>
				<OPTION value="TopRightCorner16.txt">Top Right Corner</OPTION>
				<OPTION value="BottomLeftCorner16.txt">Bottom Left Corner</OPTION>
				<OPTION value="BottomRightCorner16.txt">Bottom Right Corner</OPTION>
			</SELECT><BR>

			Template File Source:<BR>
			<INPUT name="SRC" value="<%=SRC%>" size="50"><BR><BR>

			First Hex: <INPUT name="FirstHex" value="<%=FirstHex%>" size="6" maxlength="6"><BR>

			Last Hex: <INPUT name="LastHex" value="<%=LastHex%>" size="6" maxlength="6"><BR>
			
			<INPUT type="Submit" value="Create Image">
		</FORM>

		Image:
		<IMG src="Image.asp?<%=Request.QueryString%>" align="absmiddle"><BR>
		<BR>
		Create by <A href="http://www.lewismoten.com">Lewis Moten</A>.
	</BODY>
</HTML>