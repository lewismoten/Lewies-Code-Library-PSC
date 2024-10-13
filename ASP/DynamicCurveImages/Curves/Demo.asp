<%
Color1 = Request.QueryString("Color1")
Color2 = Request.QueryString("Color2")

Color1 = Replace(Color1, "#", "")
Color2 = Replace(Color2, "#", "")

If Color1 = "" Then Color1 = "000000"
If Color2 = "" Then Color2 = "FFFFFF"

' Validate Image Color
For Index = 1 To 6
	' if an invalid character was found
	If InStr(1, "0123456789ABCDEF", Mid(Color1, Index, 1), vbTextCompare) = 0 Then
		' set color to nothing
		Color1 = "FFFFFF"
	ElseIf InStr(1, "0123456789ABCDEF", Mid(Color2, Index, 1), vbTextCompare) = 0 Then
		' set color to nothing
		Color2 = "000000"
	End If
Next

%>
<BODY>

	<H1>Dynamic Corner Images</H1>

	This demonstration shows the dynamic image
	generation in action.  4 rounded corners
	are created with aliasing to prevent that
	"blocky" edge on the corners.  Simply enter
	an HTML color (in hex format) in the form
	below.  If an invalid color is specified,
	then that color will default to either white
	(Color1) or Black (Color2).
	<BR>
	<TABLE width="100%" bgColor="#<%=Color1%>"><TR><TD><CENTER>
	
	<TABLE bgColor="#<%=Color2%>" cellspacing="0" border="0" cellpadding="0">
		<FORM method="get">
			<TR>
				<TD valign="top"><IMG src="Curve.asp?Color1=<%=Color1%>&Color2=<%=Color2%>&Curve=TopLeft"></TD>
				<TD rowspan="2" align="center">
						<INPUT name="Color1" value="<%=Color1%>"><BR>
						<INPUT name="Color2" value="<%=Color2%>"><BR>
						<INPUT type="Submit" value="Build Images">
				</TD>
				<TD valign="top"><IMG src="Curve.asp?Color1=<%=Color1%>&Color2=<%=Color2%>&Curve=TopRight"></TD>
			</TR>
			<TR>
				<TD valign="bottom"><IMG src="Curve.asp?Color1=<%=Color1%>&Color2=<%=Color2%>&Curve=BottomLeft"></TD>
				<TD valign="bottom"><IMG src="Curve.asp?Color1=<%=Color1%>&Color2=<%=Color2%>&Curve=BottomRight"></TD>
			</TR>
		</FORM>
	</TABLE>
	</CENTER></TD></TR></TABLE>
	<BR>
	<BR>
	Created by <A href="http://www.lewismoten.com">Lewis Moten</A>.
</BODY>
