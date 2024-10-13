Public Class FormatProvider
	Implements ICustomFormatter
	Implements IFormatProvider

	Public Function Format(ByVal formatting As String, ByVal arg As Object, ByVal formatProvider As System.IFormatProvider) As String Implements System.ICustomFormatter.Format

		' no formatting?
		If formatting Is Nothing Then
			' return the normal formatting from base formatter
			' (probably just an arg.tostring)
			Return String.Format("{0}", arg)
		End If

		' Take all standard date formats into consideration
		Select Case formatting.Substring(0, 2)
			Case "d+", "D+", "t+", "T+", "f+", "F+", "g+", "G+", _
			"M+", "m+", "R+", "r+", "s+", "u+", "U+", "Y+", "y+", _
			"d-", "D-", "t-", "T-", "f-", "F-", "g-", "G-", "M-", _
			"m-", "R-", "r-", "s-", "u-", "U-", "Y-", "y-"
				' Parse offset
				Dim offset = CType(formatting.Substring(1), Integer)
				' Apply offset
				Dim NewDate As DateTime = CType(arg, DateTime)
				NewDate = NewDate.AddHours(offset)
				Return NewDate.ToString(formatting.Substring(0, 1))
		End Select

		' If we can format it
		If TypeOf arg Is IFormattable Then
			' Let normal format provider take over.
			Return CType(arg, IFormattable).ToString(Format, formatProvider)
		End If

		' just return the string value
		Return arg.ToString

	End Function

	Public Function GetFormat(ByVal formatType As System.Type) As Object Implements System.IFormatProvider.GetFormat
		If formatType.ToString() = GetType(ICustomFormatter).ToString() Then
			Return Me
		Else
			Return Nothing
		End If
	End Function
End Class
