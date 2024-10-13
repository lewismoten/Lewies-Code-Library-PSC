<%
' Dynamic Pixel Script

' Created by Lewis Moten
' on April 29, 2001
' Email lewis@moten.com
' URL http://www.lewismoten.com

' This script creates a 1x1 pixel in the color
' that you specify.  It expects to recieve a
' hex color value in the querystring "Color"
' field

' examples:
' Pixel.asp?Color=FF0000
' Pixel.asp?Color=#FF0000
' Pixel.asp?Color=ff0000

' If querystring is not specified - or is inavlid,
' a transparent image will be returned


Dim Image
Dim Color
Dim Pixel
Dim Clear

' Image Templates (Hex Value of Binary Data)
Pixel = "47494638396101000100800000######0000002C00000000010001000002024401003B"
Clear = "4749463839610100010080FF00C6353500000021F90401000000002C00000000010001000002024401003B"

Color = Request.QueryString("Color")
Color = Replace(Color, "#", "")

If Len(Color) = 6 Then
	
	Image = Replace(Pixel, "######", Color)
	
	' Validate Image Color
	For Index = 1 To 6
		' if an invalid character was found
		If InStr(1, "0123456789ABCDEF", Mid(Color, Index, 1), vbTextCompare) = 0 Then
			' set the image to transparent
			Image = Clear
			Exit For
		End If
	Next

Else

	' Return a transparent image
	Image = Clear
	
End If

' Start returning the binary data

MaxIndex = Len(Image) \ 2

Response.ContentType = "image/gif"

For Index = 1 To MaxIndex
	Response.BinaryWrite(ChrB("&h" & Mid(Image, ((Index - 1) * 2) + 1, 2)))
Next	
%>