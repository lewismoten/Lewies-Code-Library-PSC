<%
' This class is used to interpret phone numbers to and from a database
' in a specific format.  The format is 25 characters long and is explained
' below.

' Phone number is formated by the following:
' Country Code	: 3 digits	(1-3)
' Area Code		: 3 digits	(4-6)
' Exchange		: 3 digits	(7-9)
' Suffix		: 4 digits  (10-13)
' Extension		: 5 digits  (14-18)
' Letters		: 7 digits	(19-25)
'
' Spaces are used as padding on the left side.
' Example:
'
' oPhoneNumber.CountryCode = 1
' oPhoneNumber.AreaCode = 800
' oPhoneNumber.Exchange = 837
' oPhoneNumber.Suffix = 8464
' oPhoneNumber.Extension = 123
' oPhoneNumber.Letters = "TESTING"
'
' Response.Write PhoneNumber.Data
'
' Entering the previous information will return a string of
' "  18008378464  123TESTING"
'
' Response.Write oPhoneNumber.PhoneNumber
'
' Entering the above function will return the following
' "1 (800) 837-8464 ext. 123, 1 (800) TES-TING"
'
' More example code may be found at the bottom of this script.
'
' ------------------------------------------------------------------------------
Class clsPhoneNumber
' ------------------------------------------------------------------------------
	Public InternationalCode
	Public AreaCode
	Public Exchange
	Public Suffix
	Public Extension
	Public Letters
' ------------------------------------------------------------------------------
	Public Property Get PhoneNumber()
		
		Dim lStrPhoneNumber
		
		If Not InternationalCode = "" Then
			lStrPhoneNumber = InternationalCode & " "
		End If
		
		If Not AreaCode = "" Then
			lStrPhoneNumber = lStrPhoneNumber & "(" & AreaCode & ") "
		End If
		
		lStrPhoneNumber = lStrPhoneNumber & Exchange & "-" & Suffix
		
		If Not Extension = "" Then
			lStrPhoneNumber = lStrPhoneNumber & " ext. " & Extension
		End If
		
		If Not Letters = "" Then
			
			lStrPhoneNumber = lStrPhoneNumber & ", "
			
			If Not InternationalCode = "" Then
				lStrPhoneNumber = lStrPhoneNumber & InternationalCode & " "
			End If
		
			If Not AreaCode = "" Then
				lStrPhoneNumber = lStrPhoneNumber & "(" & AreaCode & ") "
			End If
			
			lStrPhoneNumber = lStrPhoneNumber & Left(Letters, 3) & "-" & Right(Letters, 4)
			
		End If
		
		PhoneNumber = lStrPhoneNumber
		
	End Property
' ------------------------------------------------------------------------------
	Public Property Get Data()

		Data = _
			String(3 - Len(CStr(InternationalCode)), " ") & CStr(InternationalCode) & _
			String(3 - Len(CStr(AreaCode)), " ") & CStr(AreaCode) & _
			String(3 - Len(CStr(Exchange)), " ") & CStr(Exchange) & _
			String(4 - Len(CStr(Suffix)), " ") & CStr(Suffix) & _
			String(5 - Len(CStr(Extension)), " ") & CStr(Extension) & _
			String(7 - Len(CStr(Letters)), " ") & CStr(Letters)
		
	End Property
' ------------------------------------------------------------------------------
	Public Property Let Data(ByRef pStrData)
		
		Dim lStrData
		
		Dim lStrInternationalCode
		Dim lStrAreaCode
		Dim lStrExchange
		Dim lStrSuffix
		Dim lStrExtension
		Dim lStrLetters

		lStrData = pStrData & String(25 - Len(pStrData), " ")
		
		lStrInternationalCode	= Mid(lStrData, 1, 3)
		lStrAreaCode			= Mid(lStrData, 4, 3)
		lStrExchange			= Mid(lStrData, 7, 3)
		lStrSuffix				= Mid(lStrData, 10, 4)
		lStrExtension			= Mid(lStrData, 14, 5)
		lStrLetters				= Mid(lStrData, 19, 7)
		
		If IsNumeric(lStrInternationalCode) And Not lStrInternationalCode = "" Then
			InternationalCode = lStrInternationalCode
		Else
			InternationalCode = ""
		End If
		
		If IsNumeric(lStrAreaCode) And Not lStrAreaCode = "" Then
			AreaCode = lStrAreaCode
		Else
			AreaCode = ""
		End If
		
		If IsNumeric(lStrExchange) And Not lStrExchange = "" Then
			Exchange = lStrExchange
		Else
			Exchange = ""
		End If

		If IsNumeric(lStrSuffix) And Not lStrSuffix = "" Then
			Suffix = lStrSuffix
		Else
			Suffix = ""
		End If
		
		If Trim(lStrExtension) = "" Then
			Extension = ""
		Else
			Extension = Trim(lStrExtension)
		End If

		If LTrim(lStrLetters) = "" Then
			Letters = ""
		Else
			Letters = LTrim(lStrLetters)
		End If
	End Property
' ------------------------------------------------------------------------------
	Public Function LettersFromDigits(ByRef pStrDigits)
		Dim lLngPosition
		Dim lLngLength
		lLngLength = Len(pStrDigits)
		For lLngPosition = 1 To lLngLength
			Select Case Mid(pStrDigits, lLngPosition, 1)
				Case "2"
					LettersFromDigits = LettersFromDigits & "[ABC]"
				Case "3"
					LettersFromDigits = LettersFromDigits & "[DEF]"
				Case "4"
					LettersFromDigits = LettersFromDigits & "[GHI]"
				Case "5"
					LettersFromDigits = LettersFromDigits & "[JKL]"
				Case "6"
					LettersFromDigits = LettersFromDigits & "[MNO]"
				Case "7"
					LettersFromDigits = LettersFromDigits & "[PQRS]"
				Case "8"
					LettersFromDigits = LettersFromDigits & "[TUV]"
				Case "9"
					LettersFromDigits = LettersFromDigits & "[WXYZ]"
				Case Else
					LettersFromDigits = LettersFromDigits & Mid(pStrDigits, lLngPosition, 1)
			End Select
		Next
	End Function
' ------------------------------------------------------------------------------
	Function LettersToDigits(ByRef pStrLetters)
		Dim lLngPosition
		Dim lLngLength
		lLngLength = Len(pStrLetters)
		For lLngPosition = 1 To lLngLength
			Select Case UCase(Mid(pStrLetters, lLngPosition, 1))
				Case "A", "B", "C"
					LettersToDigits = LettersToDigits & "2"
				Case "D", "E", "F"
					LettersToDigits = LettersToDigits & "3"
				Case "G", "H", "I"
					LettersToDigits = LettersToDigits & "4"
				Case "J", "K", "L"
					LettersToDigits = LettersToDigits & "5"
				Case "M", "N", "O"
					LettersToDigits = LettersToDigits & "6"
				Case "P", "Q", "R", "S"
					LettersToDigits = LettersToDigits & "7"
				Case "T", "U", "V"
					LettersToDigits = LettersToDigits & "8"
				Case "W", "X", "Y", "Z"
					LettersToDigits = LettersToDigits & "9"
				Case Else
					LettersToDigits = LettersToDigits & Mid(pStrLetters, lLngPosition, 1)
			End Select
		Next
		
	End Function
' ------------------------------------------------------------------------------
End Class
' ------------------------------------------------------------------------------
'	Dim oPhoneNumber
'	Const lsData = "  18008378464  123TESTING"
'	Set oPhoneNumber = New clsPhoneNumber
'	oPhoneNumber.Data = lsData
'	Response.Write "Original Data: " & lsData & "<BR><BR>"
'	Response.Write "Data = " & oPhoneNumber.Data & "<BR>"
'	Response.Write "PhoneNumber = " & oPhoneNumber.PhoneNumber & "<BR>"
'	Response.Write "International Code = " & oPhoneNumber.InternationalCode & "<BR>"
'	Response.Write "Area Code = " & oPhoneNumber.AreaCode & "<BR>"
'	Response.Write "Exchange = " & oPhoneNumber.Exchange & "<BR>"
'	Response.Write "Suffix = " & oPhoneNumber.Suffix & "<BR>"
'	Response.Write "Extension = " & oPhoneNumber.Extension & "<BR>"
'	Response.Write "Letters = " & oPhoneNumber.Letters & "<BR>"
'	Response.Write "Letters To Digits = " & oPhoneNumber.LettersToDigits(oPhoneNumber.Letters) & "<BR>"
'	Response.Write "Letters From Digits = " & oPhoneNumber.LettersFromDigits(oPhoneNumber.Exchange & oPhoneNumber.Suffix) & "<BR>"
'	Set oPhoneNumber = Nothing
%>