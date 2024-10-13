<%

Dim Color1
Dim Color2

' Grab color user wishes to contrast
Color1 = Request.QueryString("Color1")

' If user didn't supply a color, suggest a default color of white
If Color1 = "" Then Color1 = "FFFFFF"

' Remove any pound signs
Color1 = Replace(Color1, "#", "")

' Grab contrasting color
Color2 = ColorContrast(Color1)

%>
<HTML>
	<HEAD>
		<TITLE>Contrasting Colors Demonstration by Lewis Moten</TITLE>
	</HEAD>
	<BODY>
		<H1>Contrasting Colors Demonstration</H1>
		<FORM method="get">
			Color: <INPUT Name="Color1" value="<%=Color1%>">
			<INPUT type="submit" value="GO">
		</FORM>
		<HR>
		<BR>
		Color 1: #<%=Color1%> <SPAN style="background-color:#<%=Color1%>;width:100px;border:1px solid black;">&nbsp;</SPAN><BR><BR>
		Color 2: #<%=Color2%> <SPAN style="background-color:#<%=Color2%>;width:100px;border:1px solid black;">&nbsp;</SPAN><BR><BR>

		<B>Contrast tests:</B><BR><BR>

		<SPAN style="color:#<%=Color2%>;background-color:#<%=Color1%>;width:100px;border:1px solid black;">Test</SPAN><BR><BR>
		<SPAN style="color:#<%=Color1%>;background-color:#<%=Color2%>;width:100px;border:1px solid black;">Test</SPAN><BR><BR>
			
		<HR>
		This demonstration and code was brought to you by
		<A href="http://www.lewismoten.com/">Lewis Moten</A>.

	</BODY>
</HTML>
<%
Function ColorContrast(ByRef pstrColor)
	
	Dim llngRed
	Dim llngGreen
	Dim llngBlue

	' Parse individual color values and add half a byte
	llngRed = CLng("&h" & mid(pstrColor, 1, 2)) + 128
	llngGreen = CLng("&h" & mid(pstrColor, 3, 2)) + 128
	llngBlue = CLng("&h" & mid(pstrColor, 5, 2)) + 128

	' Loop byte values if they extend past 255
	If llngRed > 255 Then llngRed = llngRed - 256
	If llngGreen > 255 Then llngGreen = llngGreen - 256
	If llngBlue > 255 Then llngBlue = llngBlue - 256

	' Convert new color values to a full hex color string
	ColorContrast = _
		Right("0" & Hex(llngRed), 2) & _
		Right("0" & Hex(llngGreen), 2) & _
		Right("0" & Hex(llngBlue), 2)

End Function
%>
