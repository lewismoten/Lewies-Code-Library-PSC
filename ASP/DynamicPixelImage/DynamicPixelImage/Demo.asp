<H1>Dynamic Pixel Image</H1>

To create a new image, type the color<BR>
(in hex format) below.  If an invalid<BR>
color is specified, then the result will<BR>
be a transparent image.<BR>

<FORM method="get">
	<INPUT name="Color" value="<%=Request.QueryString("Color")%>">
	<INPUT type="Submit" value="Build Image">
</FORM>
<IMG src="Pixel.asp?Color=<%=Request.QueryString("Color")%>" width="100" height="100" border=1>

<BR>
<BR>
Created by <A href="http://www.lewismoten.com">Lewis Moten</A>.