<%
' Dynamic Curve Script

' Created by Lewis Moten
' on April 29, 2001
' Email lewis@moten.com
' URL http://www.lewismoten.com

' This script creates a 16x16 curve in the colors
' that you specify.  It expects to recieve the
' following fields defined in a querystring:
'
' Color1: The background hex color of the curve (Default = FFFFFF)
' Color2: The hex color of the curve (Default = 000000)
' Curve: The type of curve (Required)

' examples:
' Curve.asp?Color1=FF0000&Color2=FFFFFF&Curve=TopRight
' Curve.asp?Color1=FF0000&Color2=FFFFFF&Curve=TopLeft
' Curve.asp?Color1=FF0000&Color2=FFFFFF&Curve=BottomRight
' Curve.asp?Color1=FF0000&Color2=FFFFFF&Curve=BottomLeft

' If the curve field is not specificed, or is incorrect,
' then a transparent image will be returned.

Dim Image
Dim Clear
Dim Palette(7)
Dim Phase(7)
Dim Color1
Dim Color2
Dim Prefix
Dim TopLeft
Dim TopRight
Dim BottomLeft
Dim BottomRight
Dim Color(2, 7)

Prefix = "47494638396110001000A2FF00[Palette]2C00000000100010000003"

' Image Templates (Hex Value of Binary Data)
TopLeft = _
	"3308BA1C32A418C65E39185310C6CD19" & _
	"251060599125B88CA9C9A16D080CB129" & _
	"7CB546E75980F3021E68273CC0843F9E" & _
	"A1884900003B3B"

TopRight = _
	"2F6845230130CA43696972D67A5FDE20" & _
	"217CE0266A65754269E8B5E0C0C2DC48" & _
	"6FF2CD057A25F49401F0401816868604" & _
	"003B"

BottomLeft = _
	"3368B7DC5E2E3622EBA852E018E066C0" & _
	"F52D80E07D40381E6940A1A9F062693C" & _
	"4BB54DE740309C0D5E4DF07308730122" & _
	"A1604800003B"

BottomRight = _
	"3278BADCD6909548098DE342A165E8CB" & _
	"008047218CA0886A4400ACD5F9B2F21C" & _
	"BF78540C2E6E2FBB9A0F6028100682DE" & _
	"109700003B"

Clear = _
	"4749463839610100010080FF00C63535" & _
	"00000021F90401000000002C00000000" & _
	"010001000002024401003B"

' Grab defined colors and remove #
Color1 = Request.QueryString("Color1")
Color2 = Request.QueryString("Color2")
Color1 = Replace(Color1, "#", "")
Color2 = Replace(Color2, "#", "")


' Define default color palette
Palette(0) = "FFFFFF"
Palette(1) = "DCDCDC"
Palette(2) = "B3B3B3"
Palette(3) = "707070"
Palette(4) = "3F3F3F"
Palette(5) = "1C1C1C"
Palette(6) = "070707"
Palette(7) = "000000"

' Define percentages of color to increase
Phase(0) = 0
For Index = 1 To 6
	Phase(Index) = Index * ((256/8)/256)
Next
Phase(7) = 1

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
	
If Len(Color1) = 6 And Len(Color2) = 6 Then

	Color(0, 0) = CInt("&h" & Left(Color1, 2))
	Color(1, 0) = CInt("&h" & Mid(Color1, 3, 2))
	Color(2, 0) = CInt("&h" & Right(Color1, 2))

	Color(0, 7) = CInt("&h" & Left(Color2, 2))
	Color(1, 7) = CInt("&h" & Mid(Color2, 3, 2))
	Color(2, 7) = CInt("&h" & Right(Color2, 2))

	For Index = 0 To 7
		Color(0, Index) = Color(0, 0) + ((Color(0, 7) - Color(0, 0)) * Phase(Index))
		Color(1, Index) = Color(1, 0) + ((Color(1, 7) - Color(1, 0)) * Phase(Index))
		Color(2, Index) = Color(2, 0) + ((Color(2, 7) - Color(2, 0)) * Phase(Index))
		Palette(Index) = _
			Right("0" & Hex(Color(0, Index)), 2) & _
			Right("0" & Hex(Color(1, Index)), 2) & _
			Right("0" & Hex(Color(2, Index)), 2)
	Next

End If

' Setup the palette to the chosen colors
Prefix = Replace(Prefix, "[Palette]", Join(Palette, ""))

' Determine wich image to return

Select Case LCase(Request.QueryString("Curve"))
	Case "topright"
		Image = Prefix & TopRight
	Case "topleft"
		Image = Prefix & TopLeft
	Case "bottomright"
		Image = Prefix & BottomRight
	Case "bottomleft"
		Image = Prefix & BottomLeft
	Case Else
		' Return a transparent image
		Image = Clear
End Select	

' Start returning the binary data

MaxIndex = Len(Image) \ 2

Response.ContentType = "image/gif"

For Index = 1 To MaxIndex
	Response.BinaryWrite(ChrB("&h" & Mid(Image, ((Index - 1) * 2) + 1, 2)))
Next	
%>