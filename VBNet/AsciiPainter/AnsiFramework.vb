Public Class AnsiFramework

	Public Cursor As New Point(1, 1)
	Public SavedCursor As New Point(1, 1)
	Public Character As New AnsiCharacter
	Public Screen(,) As AnsiCharacter
	Private _Characters As Integer = 80
	Private _Lines As Integer = 26
	Private _MaxSize As New Point(0, 0)
	Public TextWrapping As Boolean = False


	Public Event CursorMoved(ByVal location As Point)
	Public Event Progress(ByVal task As String, ByVal value As Integer)
	Public Event ProgressBegin(ByVal task As String, ByVal maxValue As Integer)
	Public Event ProgressEnd(ByVal task As String)
	Public Event CharacterModified(ByVal line As Integer, ByVal character As Integer)
	Public Event DisplayModified()
	Public Event Initialized(ByVal lines As Integer, ByVal characters As Integer)

	Public Sub New()
		InitializeScreen()
	End Sub

	Public Property Lines() As Integer
		Get
			Return _Lines
		End Get
		Set(ByVal Value As Integer)
			_Lines = Value
			InitializeScreen()
		End Set
	End Property

	Public Property Characters() As Integer
		Get
			Return _Characters
		End Get
		Set(ByVal Value As Integer)
			_Characters = Value
			InitializeScreen()
		End Set
	End Property

	Public Sub ResizeTo(ByVal lines As Integer, ByVal characters As Integer)
		_Lines = lines
		_Characters = characters
		InitializeScreen()
	End Sub

	Private Sub InitializeScreen()
		RaiseEvent ProgressBegin("Initializing", _Lines)
		ReDim Screen(_Characters, _Lines)
		For line As Integer = 1 To _Lines
			For column As Integer = 1 To _Characters
				Screen(column, line) = New AnsiCharacter
			Next
			RaiseEvent Progress("Initializing", line)
		Next
		RaiseEvent ProgressEnd("Initializing")

		RaiseEvent DisplayModified()

		Cursor.X = 1
		Cursor.Y = 1
		RaiseEvent CursorMoved(Cursor)

		RaiseEvent Initialized(_Lines, _Characters)
	End Sub

#Region " Drawing "

	Public Sub MoveUp(ByVal lines As Integer)
		Cursor.Y -= lines
		RaiseEvent CursorMoved(Cursor)
	End Sub

	Public Sub MoveDown(ByVal lines As Integer)
		Cursor.Y += lines
		RaiseEvent CursorMoved(Cursor)
	End Sub

	Public Sub MoveRight(ByVal characters As Integer)
		Cursor.X += characters
		RaiseEvent CursorMoved(Cursor)
	End Sub

	Public Sub MoveLeft(ByVal characters As Integer)
		Cursor.X -= characters
		RaiseEvent CursorMoved(Cursor)
	End Sub

	Public Sub MoveToCharacter(ByVal character As Integer)
		Cursor.X = character
		RaiseEvent CursorMoved(Cursor)
	End Sub

	Public Sub MoveToLine(ByVal line As Integer)
		Cursor.Y = line
		RaiseEvent CursorMoved(Cursor)
	End Sub

	Public Sub MoveToStartOfLine(ByVal lines As Integer)
		Cursor.X = 1
		RaiseEvent CursorMoved(Cursor)
	End Sub

	Public Sub MoveTo(ByVal line As Integer, ByVal character As Integer)
		Cursor.X = character
		Cursor.Y = line
		RaiseEvent CursorMoved(Cursor)
	End Sub

	Public Sub Home()
		Cursor.X = 1
		Cursor.Y = 1
		RaiseEvent CursorMoved(Cursor)
	End Sub

	Public Sub MoveCursorTabStopsBack(ByVal tabs As Integer)
		If Not Cursor.X Mod 4 = 0 Then
			Cursor.X -= Cursor.X Mod 4
			tabs -= 1
		End If
		Cursor.X -= (tabs * 4)
		RaiseEvent CursorMoved(Cursor)
	End Sub

	Public Sub InsertBlankLines(ByVal lines As Integer)

	End Sub

	Public Sub InsertBlankCharacters(ByVal characters As Integer)

	End Sub

	Public Sub DeleteLines(ByVal lines As Integer)

	End Sub

	Public Sub DeleteCharacters(ByVal characters As Integer)

	End Sub

	Public Sub EraseDisplayFromCursorToEnd()

	End Sub
	Public Sub EraseDisplayFromBeginToCursor()

	End Sub
	Public Sub EraseDisplay()
		RaiseEvent ProgressBegin("Erasing Display", _Lines)
		_MaxSize.X = 0
		_MaxSize.Y = 0

		For line As Integer = 1 To _Lines
			For column As Integer = 1 To _Characters
				Screen(column, line).ResetAttributes()
				Screen(column, line).Ascii = 0
			Next
			RaiseEvent Progress("Erasing Display", line)
		Next
		'Character = New AnsiCharacter
		Character.Bold = True
		Character.Background = 0
		Character.Foreground = 7
		Character.Ascii = 32
		Character.Blink = False
		Character.Reverse = False
		Character.Underscore = False

		RaiseEvent ProgressEnd("Erasing Display")
		RaiseEvent DisplayModified()

		Cursor.X = 1
		Cursor.Y = 1
		RaiseEvent CursorMoved(Cursor)
	End Sub
	Public Sub EraseLineFromCursorToEnd()
		While Not Cursor.X = Characters
			SetCharacterAttributes()
			WriteCharacter(32)
		End While
		WriteCharacter(32)
	End Sub
	Public Sub EraseLineFromBeginToCursor()

	End Sub
	Public Sub EraseLine()

	End Sub
	Public Sub EraseCharacters()

	End Sub
	Public Sub ScrollDisplayUp(ByVal lines As Integer)

	End Sub
	Public Sub ScrollDisplayDown(ByVal lines As Integer)

	End Sub
	Public Sub SetCharacterAttributes(ByVal ParamArray attributes() As Integer)
		For Each Attribute As Integer In attributes
			Select Case Attribute
				Case 0
					Character.Bold = False
					Character.Blink = False
					Character.Reverse = False
					Character.Underscore = False
					Character.Background = 0
					Character.Foreground = 7
				Case 1
					Character.Bold = True
				Case 4
					Character.Underscore = True
				Case 5
					Character.Blink = True
				Case 7
					Character.Reverse = True
				Case 30, 31, 32, 33, 34, 35, 36, 37
					Character.Foreground = Attribute - 30
				Case 40, 41, 42, 43, 44, 45, 46, 47
					Character.Background = Attribute - 40
				Case Else
					Debug.WriteLine("Unknown character attribute: " + Attribute.ToString)
			End Select
		Next
	End Sub
	Public Sub SaveCursor()
		SavedCursor = Cursor
	End Sub
	Public Sub RestoreCursor()
		Cursor = SavedCursor
		RaiseEvent CursorMoved(Cursor)
	End Sub

	Private Function WithinScreen()
		If Cursor.X < 1 Then Return False
		If Cursor.X > _Characters Then Return False
		If Cursor.Y < 1 Then Return False
		If Cursor.Y > _Lines Then Return False
		Return True
	End Function

	Public ReadOnly Property MaxSize() As Point
		Get
			Return _MaxSize
		End Get
	End Property

	Public Sub WriteCharacter(ByVal ascii As Integer)

		If Cursor.X > _MaxSize.X Then _MaxSize.X = Cursor.X
		If Cursor.Y > _MaxSize.Y Then _MaxSize.Y = Cursor.Y

		If Not WithinScreen() Then Return

		With Screen(Cursor.X, Cursor.Y)
			.Background = Character.Background
			.Blink = Character.Blink
			.Bold = Character.Bold
			.Foreground = Character.Foreground
			.Reverse = Character.Reverse
			.Underscore = Character.Underscore
			.Ascii = ascii
			'Debug.Write(.Character)
		End With
		RaiseEvent CharacterModified(Cursor.Y, Cursor.X)

		Cursor.X += 1

		If TextWrapping Then
			If Cursor.X > _Characters Then
				Cursor.Y += 1
				Cursor.X = 1
			End If
		End If
		RaiseEvent CursorMoved(Cursor)

	End Sub

#End Region

#Region " I/O "

	Public Function GetHtml() As String

	End Function
	Public Function GetBytes() As Byte()
		' TODO: Optimize saving ansi image
		Dim Memory As New IO.MemoryStream
		Try
			With Memory
				For character As Integer = 1 To Me.Characters
					For line As Integer = 1 To Me.Lines
						Dim block As AnsiCharacter = Screen(character, line)
						.WriteByte(27)
						.WriteByte(91)
						For Each c As Char In line.ToString.ToCharArray
							.WriteByte(Asc(c))
						Next
						.WriteByte(Asc(";"))
						For Each c As Char In character.ToString.ToCharArray
							.WriteByte(Asc(c))
						Next
						.WriteByte(Asc("H"))
						.WriteByte(27)
						.WriteByte(91)

						If block.Bold Then
							.WriteByte(Asc("1"))
							.WriteByte(Asc(";"))
						Else
							.WriteByte(Asc("0"))
							.WriteByte(Asc(";"))
						End If
						.WriteByte(Asc("4"))
						.WriteByte(Asc(block.Background.ToString))
						.WriteByte(Asc(";"))
						.WriteByte(Asc("3"))
						.WriteByte(Asc(block.Foreground.ToString))
						.WriteByte(Asc("m"))
						.WriteByte(Screen(character, line).Ascii)
					Next
				Next

			End With
			Return Memory.ToArray
		Finally
			Memory.Close()
		End Try
	End Function
	Public Function GetImage() As Image
		Dim Acm As AnsiCharacterMap
		Dim Image As Image = New Bitmap(Characters * Acm.CharacterWidth, Lines * Acm.CharacterHeight)
		Dim g As Graphics
		Dim brush As New SolidBrush(Color.Black)
		Try
			g = g.FromImage(Image)
			g.Clear(Color.Black)
			For character As Integer = 1 To Characters
				For line As Integer = 1 To Lines
					Dim block As AnsiCharacter = Screen(character, line)
					Dim Top As Integer = line * Acm.CharacterHeight
					Dim Left As Integer = character * Acm.CharacterWidth
					If Not block.Background = 0 Then
						brush = New SolidBrush(block.BackColor)
						g.FillRectangle(brush, Left, Top, Acm.CharacterWidth, Acm.CharacterHeight)
					End If
					If Not block.Ascii = 0 Then
						brush.Color = block.ForeColor
						g.DrawString(block.Character, Acm.RegularFont, brush, Left + Acm.CharacterLeft, Top + Acm.CharacterTop)
					End If
				Next
			Next
			Return Image
		Catch ex As Exception
			Debug.WriteLine(ex.Message)
		Finally
			brush.Dispose()
			g.Dispose()
			'Image.Dispose()
		End Try
	End Function

	Public Sub Open(ByVal fileName As String)
		TextWrapping = False
		If Not IO.File.Exists(fileName) Then Exit Sub

		Dim Stream As IO.FileStream
		Stream = IO.File.OpenRead(fileName)
		Dim Buffer(Stream.Length) As Byte
		Stream.Read(Buffer, 0, Buffer.Length)
		Stream.Close()

		EraseDisplay()

		Dim maxIndex = Buffer.Length - 1

		For i As Integer = 0 To maxIndex
			Dim keycode As Byte = Buffer(i)
			Dim keycode2 As Byte = 0
			If Not i + 1 > maxIndex Then keycode2 = Buffer(i + 1)

			If keycode = 27 And keycode2 = 91 Then			 ' [ESC] + "]"
				Dim hasMore As Boolean = False
				Dim n As Integer = i + 2
				Dim value As String = ""
				Dim code As AnsiEscapeSequence
				Do While Not ((Buffer(n) >= 97 And Buffer(n) <= 122) Or (Buffer(n) >= 65 And Buffer(n) <= 90))
					value &= Chr(Buffer(n))
					n += 1
				Loop
				code = CType(Buffer(n), AnsiEscapeSequence)

				Select Case code
					Case AnsiEscapeSequence.MoveCursorUp
						If value.Length = 0 Then
							MoveUp(1)
						Else
							MoveUp(Integer.Parse(value))
						End If
					Case AnsiEscapeSequence.MoveCursorDown
						If value.Length = 0 Then
							MoveDown(1)
						Else
							MoveDown(Integer.Parse(value))
						End If
					Case AnsiEscapeSequence.MoveCursorRight
						If value.Length = 0 Then
							MoveRight(1)
						Else
							MoveRight(Integer.Parse(value))
						End If
					Case AnsiEscapeSequence.MoveCursorLeft
						If value.Length = 0 Then
							MoveLeft(1)
						Else
							MoveLeft(Integer.Parse(value))
						End If
					Case AnsiEscapeSequence.MoveCursorTo
						If value.Length = 0 Then
							Home()
						ElseIf value.IndexOf(";") = -1 Then
							MoveToCharacter(Integer.Parse(value))
						Else
							Dim character As Integer = Integer.Parse(Split(value, ";")(1))
							Dim line As Integer = Integer.Parse(Split(value, ";")(0))
							MoveTo(line, character)
						End If
					Case AnsiEscapeSequence.ClearScreen
						If value = "2" Then
							EraseDisplay()
						Else
							Debug.WriteLine("Unknown value """ & value & """ for J")
						End If
					Case AnsiEscapeSequence.ClearToEndOfLine
						If value.Length = 0 Then
							Me.EraseLineFromCursorToEnd()
						Else
							Debug.WriteLine("Unknown value """ & value & """ for K")
						End If
					Case AnsiEscapeSequence.PlayMusic
						If value.Length = 0 Then
							' Music ...
							Debug.WriteLine("Music not supported.")
							n += 1
							Do While Not Buffer(n) = 14
								n += 1
							Loop
						Else

						End If
					Case AnsiEscapeSequence.SetOptions
						Select Case value
							Case "?7"
								TextWrapping = True
							Case Else
								Debug.WriteLine("Unknown value """ & value & """ for h")
						End Select
					Case AnsiEscapeSequence.SetVideoOptions
						If value.Length = 0 Then
							SetCharacterAttributes()
						ElseIf value.IndexOf(";") = -1 Then
							SetCharacterAttributes(Integer.Parse(value))
						Else
							Dim Values() As String = Split(value, ";")
							Dim maxValueIndex As Integer = Values.Length - 1
							Dim attributes(maxValueIndex) As Integer
							For valueIndex As Integer = 0 To maxValueIndex
								attributes(valueIndex) = Integer.Parse(Values(valueIndex))
							Next
							SetCharacterAttributes(attributes)
						End If
					Case AnsiEscapeSequence.SaveCursorPosition
						SaveCursor()
					Case AnsiEscapeSequence.RestoreCursorPosition
						RestoreCursor()
					Case Else
						'Debug.WriteLine("Unknown Code: " & code & " (" & Chr(code) & ")")
						Debug.WriteLine("Unknown Escape Sequence: <-[" & value & Chr(CType(code, Byte)) & " " & code.ToString())
				End Select
				i = n
			ElseIf keycode = 13 Then
				MoveToCharacter(1)
				MoveDown(1)
				If keycode2 = 10 Then i += 1
			ElseIf keycode = 10 Then
				MoveToCharacter(1)
				MoveDown(1)
			Else
				WriteCharacter(keycode)
			End If
		Next
		If _MaxSize.X > _Characters Or _MaxSize.Y > _Lines Then
			Debug.WriteLine("Image too large: " & _MaxSize.ToString)
		End If

	End Sub
#End Region
End Class