Option Explicit On 
Option Strict On

Imports System.Drawing.Imaging
Imports System.Drawing.Drawing2D

Public Class ValidationCodeImage
	Private _Text As String = ""	' verification code to be displayed
	Private _Font As Font = New Font("Arial", 10, FontStyle.Regular, GraphicsUnit.Pixel)
	Private _Background As Color

	Public Sub New()
		_Text = Guid.NewGuid.ToString.Substring(0, 6).ToUpper()
	End Sub

	Public Sub New(ByVal text As String)
		_Text = text
	End Sub

	Public Property Text() As String
		Get
			Return _Text
		End Get
		Set(ByVal Value As String)
			_Text = Value
		End Set
	End Property

	Public Function Render() As Byte()

		' create image
		Dim Canvas As System.Drawing.Image
		Canvas = New System.Drawing.Bitmap(140, 30)

		' setup graphics object to manipulate the canvas
		Dim g As Graphics = g.FromImage(Canvas)
		g.PageUnit = GraphicsUnit.Pixel

		' setup background color
		_Background = RandomColor()
		g.Clear(_Background)

		' determine where the first character should begin
		Dim Position As New PointF
		Position.X = 0
		Position.Y = (g.VisibleClipBounds.Height / 2S) + 2S

		' Loop through each character.
		For Index As Integer = 0 To Text.Length - 1

			' parse out the character
			Dim Character As String = Text.Substring(Index, 1)

			' Draw the character
			DrawCharacter(g, Character, Position)

		Next

		g.Dispose()
		g = Nothing

		Try
			' return a byte-array of the image
			Return Image2Bytes(Canvas)
		Finally
			If Not Canvas Is Nothing Then
				Canvas.Dispose()
				Canvas = Nothing
			End If
		End Try

	End Function

	Private Function Image2Bytes(ByVal image As System.Drawing.Image) As Byte()

		' find the JPEG codec information
		Dim JpegCodec As ImageCodecInfo
		For Each Info As ImageCodecInfo In ImageCodecInfo.GetImageEncoders
			If Info.MimeType = "image/jpeg" Then
				JpegCodec = Info
			End If
		Next

		' if we didnt' find it, throw an exception
		If JpegCodec Is Nothing Then
			Throw New Exception("Jpeg Codec missing")
		End If

		' setup quality of image
		Dim Quality As Integer = 100
		Dim Parameters As New EncoderParameters(1)
		Parameters.Param(0) = New EncoderParameter(Encoder.Quality, Quality)

		' save image to binary stream
		Dim Stream As New IO.MemoryStream
		image.Save(Stream, JpegCodec, Parameters)

		' return those bytes
		Return Stream.ToArray

	End Function
	Private Function RandomFontFamily() As String

		Randomize()

		' define a list of font families
		Dim Families(7) As String

		Families(0) = "Arial"
		Families(1) = "Courier"
		Families(2) = "Courier New"
		Families(3) = "Times New Roman"
		Families(4) = "Verdana"
		Families(5) = "Lucidia Console"
		Families(6) = "Comis Sans MS"
		Families(7) = "System"

		' choose a random font
		Dim Index As Integer
		Index = CType(Fix(Families.Length * Rnd()), Integer)

		' return the name
		Return Families(Index)

	End Function
	Private Function RandomFontStyle() As System.Drawing.FontStyle

		Randomize()

		' Initialize our style (regular)
		Dim Style As New System.Drawing.FontStyle

		' We setup the chance of each style being used.
		Const Chance As Single = 25

		' apply each style.  
		If 100 * Rnd() < Chance Then
			Style = Style Or FontStyle.Bold
		End If
		If 100 * Rnd() < Chance Then
			Style = Style Or FontStyle.Italic
		End If
		If 100 * Rnd() < Chance Then
			Style = Style Or FontStyle.Strikeout
		End If
		If 100 * Rnd() < Chance Then
			Style = Style Or FontStyle.Underline
			Style = Style Xor FontStyle.Underline
		End If

		Return Style

	End Function

	Private Function RandomFont() As Font

		Randomize()

		' Determine font size, family, style
		Dim Size As Single = (24 * Rnd()) + 12
		Dim Family As String = RandomFontFamily()
		Dim Style As System.Drawing.FontStyle = RandomFontStyle()

		' create and return new font
		Return New Font(Family, Size, Style, GraphicsUnit.Pixel)

	End Function
	Private Sub RotateCanvas(ByVal g As Graphics, ByVal degrees As Single, ByVal center As PointF)

		' Rotate way we approach the canvas
		Dim mx As New Matrix
		mx.RotateAt(degrees, center)
		g.Transform = mx

	End Sub
	Private Function RandomOffset(ByVal position As PointF) As PointF
		Dim x As Single = position.X
		Dim y As Single = position.Y
		Dim max As Single = 5

		x = x + ((max * 2) * Rnd()) - max
		y = y + ((max * 2) * Rnd()) - max

		Return New PointF(x, y)
	End Function

	Private Function RandomElement(ByVal bgElement As Integer) As Integer

		' bytes are hard to work with vb.net so we virtually
		' work with them as integers.  We need to make sure that
		' if they go over 255, then they are wrapped around to
		' zero + overflow.  If they go under 0, then they are
		' wrapped around the other direction 256 + underflow
		' bytes are unsigned.

		' set byte to its opposite.
		If bgElement < 128 Then bgElement += 128 Else bgElement -= 128

		' Determine max levels of saturation to offset element
		Dim MaxOffset As Integer = 64

		' Choose random value to offset element
		Dim Offset As Integer
		Offset = CType(((MaxOffset * 2) * Rnd()) - MaxOffset, Integer)

		' offset the element and prevent overflow
		bgElement = (Offset + bgElement) Mod 256

		' prevent underflow
		If bgElement < 0 Then bgElement = 256 + bgElement

		' return the value
		Return bgElement

	End Function
	Private Function RandomColor() As Color
		Dim r As Integer = CType(255 * Rnd(), Integer)
		Dim g As Integer = CType(255 * Rnd(), Integer)
		Dim b As Integer = CType(255 * Rnd(), Integer)
		Return Color.FromArgb(r, g, b)
	End Function
	Private Function RandomBrush(ByVal bgColor As Color) As Brush

		' define Red, Green, Blue elements and get a
		' random color that contrasts well against the
		' background color.

		Dim r As Integer = RandomElement(bgColor.R)
		Dim g As Integer = RandomElement(bgColor.G)
		Dim b As Integer = RandomElement(bgColor.B)

		' create and return a new brush, with the
		' contrasting color
		Return New SolidBrush(Color.FromArgb(r, g, b))

	End Function
	Private Sub DrawCharacter(ByVal g As Graphics, ByVal character As String, ByRef position As PointF)

		Randomize()

		' Acquire a random font
		Dim Font As Font = RandomFont()

		' determine actual height/width of character when drawn on image.
		Dim TextSize As SizeF = g.MeasureString(character, Font)

		Dim offset As PointF = RandomOffset(position)

		' set the Y postion where the top of the chacter is written
		offset.Y = offset.Y - (TextSize.Height / 2)

		' Determine where the center of the character will appear on the canvas
		Dim CenterCharacter As New PointF
		CenterCharacter.X = offset.X + (TextSize.Width / 2)
		CenterCharacter.Y = offset.Y + (TextSize.Height / 2)

		' Determine number of degrees to rotate
		Dim MaxDegrees As Single = 45
		Dim Degrees As Single = ((MaxDegrees * 2) * Rnd()) - MaxDegrees

		' Rotate way we approach the canvas
		RotateCanvas(g, Degrees, CenterCharacter)

		' Draw the character
		g.DrawString(character, Font, RandomBrush(_Background), offset)

		' Rotate the canvas back to its original state
		RotateCanvas(g, -Degrees, CenterCharacter)

		' move virtual cursor to next characters position
		position.X = position.X + TextSize.Width
	End Sub
End Class
