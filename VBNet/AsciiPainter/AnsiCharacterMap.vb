Public Class AnsiCharacterMap

	Public Const FontName As String = "Courier New"

	Public Const CharacterWidth As Integer = 8
	Public Const CharacterHeight As Integer = 16
	Public Const CharacterTop As Integer = 0
	Public Const CharacterLeft As Integer = -2

	Private Shared Mapping() As Integer = New Integer() { _
	  &H0, &H1, &H2, &H3, &H4, &H5, &H6, &H7, &H8, &H9, &HA, &HB, &HC, &HD, &HE, &HF, _
	  &H10, &H11, &H12, &H13, &H14, &H15, &H16, &H17, &H18, &H19, &H1A, &H1B, &H1C, &H1D, &H1E, &H1F, _
	  &H20, &H21, &H22, &H23, &H24, &H25, &H26, &H27, &H28, &H29, &H2A, &H2B, &H2C, &H2D, &H2E, &H2F, _
	  &H30, &H31, &H32, &H33, &H34, &H35, &H36, &H37, &H38, &H39, &H3A, &H3B, &H3C, &H3D, &H3E, &H3F, _
	  &H40, &H41, &H42, &H43, &H44, &H45, &H46, &H47, &H48, &H49, &H4A, &H4B, &H4C, &H4D, &H4E, &H4F, _
	  &H50, &H51, &H52, &H53, &H54, &H55, &H56, &H57, &H58, &H59, &H5A, &H5B, &H5C, &H5D, &H5E, &H5F, _
	  &H60, &H61, &H62, &H63, &H64, &H65, &H66, &H67, &H68, &H69, &H6A, &H6B, &H6C, &H6D, &H6E, &H6F, _
	  &H70, &H71, &H72, &H73, &H74, &H75, &H76, &H77, &H78, &H79, &H7A, &H7B, &H7C, &H7D, &H7E, &H7F, _
	  &HC7, &HFC, &HE9, &HE2, &HE4, &HE0, &HE5, &HE7, &HEA, &HEB, &HE8, &HEF, &HEE, &HEC, &HC4, &HC5, _
	  &HC9, &HE6, &HC6, &HF4, &HF6, &HF2, &HFB, &HF9, &HFF, &HD6, &HDC, &HA2, &HA3, &HA5, &H20A7, &H192, _
	  &HE1, &HED, &HF3, &HFA, &HF1, &HD1, &HAA, &HBA, &HBF, &H2310, &HAC, &HBD, &HBC, &HA1, &HAB, &HBB, _
	  &H2591, &H2592, &H2593, &H2502, &H2524, &H2561, &H2562, &H2556, &H2555, &H2563, &H2551, &H2557, &H255D, &H255C, &H255B, &H2510, _
	  &H2514, &H2534, &H252C, &H251C, &H2500, &H253C, &H255E, &H255F, &H255A, &H2554, &H2569, &H2566, &H2560, &H2550, &H256C, &H2567, _
	  &H2568, &H2564, &H2565, &H2559, &H2558, &H2552, &H2553, &H256B, &H256A, &H2518, &H250C, &H2588, &H2584, &H258C, &H2590, &H2580, _
	  &H3B1, &HDF, &H393, &H3C0, &H3A3, &H3C3, &HB5, &H3C4, &H3A6, &H398, &H3A9, &H3B4, &H221E, &H3C6, &H3B5, &H2229, _
	  &H2261, &HB1, &H2265, &H2264, &H2320, &H2321, &HF7, &H2248, &HB0, &H2219, &HB7, &H221A, &H207F, &HB2, &H25A0, &HA0 _
	  }
	Private Shared _BoldFont As Font
	Private Shared _RegularFont As Font

	Public Shared ReadOnly Property BoldFont() As Font
		Get
			If _BoldFont Is Nothing Then
				_BoldFont = New Font(FontName, 10, FontStyle.Bold, GraphicsUnit.Point)
			End If
			Return _BoldFont
		End Get
	End Property

	Public Shared ReadOnly Property RegularFont() As Font
		Get
			If _RegularFont Is Nothing Then
				_RegularFont = New Font(FontName, 10, FontStyle.Regular, GraphicsUnit.Point)
			End If
			Return _RegularFont
		End Get
	End Property

	Public Shared Function Ascii(ByVal [unicode] As String) As Byte
		Dim Code As Integer = AscW([unicode])
		For Index As Integer = 0 To 255
			If Code = Mapping(Index) Then Return Index
		Next
		Return 0
	End Function

	Public Shared Function [Unicode](ByVal ascii As Byte) As String
		Return ChrW(Mapping(ascii))
	End Function
End Class
