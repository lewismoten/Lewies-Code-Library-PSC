Public Class Logic
	Public Shared Function BaseConversion(ByVal value As String, ByVal FromBase As String, ByVal ToBase As String) As String

		' Converts between bases with customized character sets.

		' First, get decimal value
		Dim DecimalValue As Double = 0
		For startIndex As Integer = value.Length - 1 To 0 Step -1
			Dim digitChar As String = value.Substring(startIndex, 1)
			Dim digitValue As Integer = FromBase.IndexOf(digitChar)
			If digitValue = -1 Then Return "#Error: Invalid Format"
			Dim position As Integer = ((value.Length - 1) - startIndex) + 1
			DecimalValue += digitValue * (FromBase.Length ^ (position - 1))
			If DecimalValue > Double.MaxValue Then Return "#Error: Overflow"
		Next

		' Now, convert to next base
		Dim FinalValue As String = ""
		While Not DecimalValue = 0
			Dim digitvalue As Integer = DecimalValue - (Math.Floor(DecimalValue / ToBase.Length) * ToBase.Length)
			FinalValue = ToBase.Substring(digitValue, 1) & FinalValue
			DecimalValue = Math.Floor(DecimalValue / ToBase.Length)
		End While

		Return FinalValue
	End Function
End Class
