Public Class AnsiColors

	Private Shared Regular() As Integer = New Integer() { _
	&HFF000000, _
	&HFF800000, _
	&HFF00B300, _
	&HFF808000, _
	&HFF000080, _
	&HFF800080, _
	&HFF008080, _
	&HFFC0C0C0 _
	}

	Private Shared Bold() As Integer = New Integer() { _
	&HFF808080, _
	&HFFFF0000, _
	&HFF00FF00, _
	&HFFFFFF00, _
	&HFF0000FF, _
	&HFFFF00FF, _
	&HFF00FFFF, _
	&HFFFFFFFF _
	}

	Private Shared Regular1() As Integer = New Integer() { _
	&HFF000000, _
	&HFFBB0000, _
	&HFF00BB00, _
	&HFFBBBB00, _
	&HFF0000BB, _
	&HFFBB00BB, _
	&HFF00BBBB, _
	&HFFBBBBBB _
	}

	Private Shared Bold1() As Integer = New Integer() { _
	&HFF555555, _
	&HFFFF5555, _
	&HFF55FF55, _
	&HFFFFFF55, _
	&HFF5555FF, _
	&HFFFF55FF, _
	&HFF55FFFF, _
	&HFFFFFFFF _
	}

	Private Shared Names() As String = New String() {"Black", "Blue", "Green", "Cyan", "Red", "Magenta", "Brown", "Gray", "Charcoal", "Bright Blue", "Bright Green", "Bright Cyan", "Orange", "Pink", "Yellow", "White"}

	Public Shared Function Escape(ByVal foreColor As Color, ByVal backColor As Color) As Byte()
		Return Escape(Foreground(foreColor), IsBold(foreColor), Background(backColor))
	End Function
	Public Shared Function Escape(ByVal foreground As Integer, ByVal isBold As Boolean, ByVal background As Integer) As Byte()
		Dim Buffer(9) As Byte
		Buffer(0) = 27		 ' [ESC]
		Buffer(1) = 91		 ' ]
		Buffer(2) = 48		 ' 0
		If isBold Then Buffer(2) += 1 ' 1
		Buffer(3) = 59		 ' ;
		Buffer(4) = 51		 ' 3
		Buffer(5) = 48 + foreground
		Buffer(6) = 59		  ' ;
		Buffer(7) = 52		  ' 4
		Buffer(8) = 48 + background
		Buffer(9) = 77		  ' m
		Return Buffer
	End Function
	Public Shared Function ForeColor(ByVal index As Byte, ByVal isBold As Boolean) As Color
		If isBold Then Return Color.FromArgb(Bold(index))
		Return Color.FromArgb(Regular(index))
	End Function
	Public Shared Function Foreground(ByVal color As Color) As Byte
		Dim Argb As Integer = color.ToArgb
		For Index As Integer = 0 To 7
			If Regular(Index) = Argb Then Return Index
			If Bold(Index) = Argb Then Return Index
		Next
		Debug.WriteLine("No FG colors match " & color.ToString)
		Return 0
	End Function
	Public Shared Function IsBold(ByVal color As Color) As Boolean
		Dim Argb As Integer = color.ToArgb
		For Index As Integer = 0 To 7
			If Bold(Index) = Argb Then Return True
		Next
		Return False
	End Function
	Public Shared Function BackColor(ByVal index As Byte) As Color
		Return Color.FromArgb(Regular(index))
	End Function
	Public Shared Function Background(ByVal color As Color) As Byte
		Dim Argb As Integer = color.ToArgb
		For Index As Integer = 0 To 7
			If Regular(Index) = Argb Then Return Index
		Next
		Debug.WriteLine("No BG colors match " & color.ToString)
		Return 0
	End Function
End Class
