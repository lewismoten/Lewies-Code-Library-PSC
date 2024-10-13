Public Class AnsiCharacter
	Public Foreground As Byte = 0
	Public Background As Byte = 0
	Public Bold As Boolean = False
	Public Blink As Boolean = False
	Public Reverse As Boolean = False
	Public Underscore As Boolean = False
	Public Ascii As Byte
	Public Sub ResetAttributes()
		Foreground = 0
		Background = 0
		Bold = False
		Blink = False
		Reverse = False
		Underscore = False
	End Sub
	Public Shadows Function ToString() As String
		Dim Value As String
		Dim Buffer() As Byte = Escape()
		Dim MaxIndex As Integer = Buffer.Length - 1
		For Index As Integer = 0 To MaxIndex
			Value &= AnsiCharacterMap.Unicode(Buffer(Index))
		Next
		Value &= AnsiCharacterMap.Unicode(Ascii)
		Return Value
	End Function
	Public Function Escape() As Byte()
		Return AnsiColors.Escape(Foreground, Bold, Background)
	End Function
	Public Function ToHtml() As String
		Dim Html As New System.Text.StringBuilder("")
		Dim Color As Color = ForeColor
		Html.Append("<FONT face=""Courier New"" size=""3"" color=""#")
		Html.Append(Right("0" & Hex(Color.R), 2))
		Html.Append(Right("0" & Hex(Color.G), 2))
		Html.Append(Right("0" & Hex(Color.B), 2))
		Html.Append(""" style=""background-color:#")
		Color = BackColor
		Html.Append(Right("0" & Hex(Color.R), 2))
		Html.Append(Right("0" & Hex(Color.G), 2))
		Html.Append(Right("0" & Hex(Color.B), 2))
		Html.Append(";"">&#")
		Html.Append(AscW(Character))
		Html.Append(";</FONT>")
		Return Html.ToString
	End Function
	Public Function ToGif() As Byte()
		Dim Map As AnsiCharacterMap
		Dim Image As New Bitmap(Map.CharacterWidth, Map.CharacterHeight)
		Dim g As Graphics
		Dim Brush As New SolidBrush(ForeColor)
		g = g.FromImage(Image)
		g.Clear(BackColor)
		g.DrawString(Character, Map.RegularFont, Brush, Map.CharacterLeft, Map.CharacterTop)
		Brush.Dispose()
		g.Dispose()
		Dim Stream As New IO.MemoryStream
		Try
			Image.Save(Stream, System.Drawing.Imaging.ImageFormat.Gif)
			Return Stream.ToArray
		Finally
			Stream.Close()
			Image.Dispose()
		End Try
	End Function
	Public ReadOnly Property GetShortString()
		Get
			Dim offset As Byte = 0
			If Bold Then offset = 8
			Return Right("0" & Hex(Ascii), 2) & Hex(Foreground + offset) & Hex(Background)
		End Get
	End Property
	Public Property Character() As Char
		Get
			Return AnsiCharacterMap.Unicode(Ascii)
		End Get
		Set(ByVal Value As Char)
			Ascii = AnsiCharacterMap.Ascii(Value)
		End Set
	End Property
	Public Property ForeColor() As Color
		Get
			Return AnsiColors.ForeColor(Foreground, Bold)
		End Get
		Set(ByVal Value As Color)
			Foreground = AnsiColors.Foreground(Value)
		End Set
	End Property
	Public Property BackColor() As Color
		Get
			Return AnsiColors.BackColor(Background)
		End Get
		Set(ByVal Value As Color)
			Background = AnsiColors.Background(Value)
		End Set
	End Property
End Class
