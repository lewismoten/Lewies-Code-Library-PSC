Option Explicit On 
Option Strict On

Imports System.Drawing.Imaging
Imports System.Drawing.Drawing2D

Public Class ValidationCodeImage
	Private _Text As String = ""	' verification code to be displayed

	Public Sub New()
		_Text = Guid.NewGuid.ToString.Substring(0, 6)
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
		g.Clear(Color.DarkBlue)

		' setup font to write characters in
		Dim Font As New Font("Arial", 24, FontStyle.Regular, GraphicsUnit.Pixel)

		' determine where the first character should begin
		Dim Position As New PointF
		Position.X = 0
		Position.Y = 0

		' Loop through each character.
		For Index As Integer = 0 To Text.Length - 1

			' parse out the character
			Dim Character As String = Text.Substring(Index, 1)

			' Draw the character
			DrawCharacter(g, Character, Position, Font)

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
		Dim Quality As Integer = 70
		Dim Parameters As New EncoderParameters(1)
		Parameters.Param(0) = New EncoderParameter(Encoder.Quality, Quality)

		' save image to binary stream
		Dim Stream As New IO.MemoryStream
		image.Save(Stream, JpegCodec, Parameters)

		' return those bytes
		Return Stream.ToArray

	End Function
	Private Sub DrawCharacter(ByVal g As Graphics, ByVal character As String, ByRef position As PointF, ByVal font As Font)

		' determine actual height/width of character when drawn on image.
		Dim TextSize As SizeF = g.MeasureString(character, font)

		' Draw the character
		g.DrawString(character, font, New SolidBrush(Color.White), position)

		' move virtual cursor to next characters position
		position.X = position.X + TextSize.Width

	End Sub
End Class
