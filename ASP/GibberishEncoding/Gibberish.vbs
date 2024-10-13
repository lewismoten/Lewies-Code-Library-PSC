Dim strMessage
Dim strGibberish

strMessage = InputBox("Enter a message to encode.", "Gibberish")

If Not (strMessage = vbCancel Or strMessage = "") Then
	strGibberish = ToGibberish(strMessage)
	MsgBox strGibberish, vbOKOnly + vbInformation, "Gibberish"
End If

Function ToGibberish(ByRef pstrMessage)

	Dim lstrGibberish
	Dim lstrLetter
	Dim llngLength
	Dim llngPosition

	llngLength = Len(pstrMessage)

	For llngPosition = 1 To llngLength

		lstrLetter = Mid(pstrMessage, llngPosition, 1)

		Select Case UCase(lstrLetter)
			Case "A", "E", "I", "O", "U"
				lstrGibberish = lstrGibberish & "idig"
		End Select

		lstrGibberish = lstrGibberish & lstrLetter

	Next

	ToGibberish = lstrGibberish

End Function
