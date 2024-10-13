Option Explicit

Randomize(Timer())
		
Dim DropTimer()
Dim DropSpeed()
Dim DropObject()

Dim X_Max
Dim Y_Max

Sub Drop(ImageFile, DropCount)

	Dim iCount

	Redim DropTimer(DropCount)
	ReDim DropSpeed(DropCount)
	ReDim DropObject(DropCount)

	X_Max = Int(window.screen.availWidth)
	Y_Max = Int(window.screen.availHeight)

	For iCount = 0 To DropCount
		Call Window.document.body.insertAdjacentHTML("BeforeEnd", imgHTML(ImageFile, iCount))
		DropTimer(iCount) = window.setInterval("DropAnimation Dropper" & iCount & ", " & iCount, 75)
		DropSpeed(iCount) = Int(Rnd * 10) + 10
	Next

End Sub

Function imgHTML(ImageFile, iCount)

	imgHTML = _
		"<IMG ID=Dropper" & iCount & " " & _
		"SRC=""" & ImageFile & """ " & _
		"STYLE=""POSITION:absolute;Z-INDEX:-1;TOP:9999"">"

End Function

Sub DropAnimation(DropperID, DropperIndex)

	Dim X_Left
	Dim Y_Top

	X_Left = GetNum(DropperID.style.left)
	Y_Top = GetNum(DropperID.style.top)

	If Y_Top > Y_Max Then

		X_Left = Int(Rnd * X_Max)
		DropperID.style.left = X_Left
		Y_Top = - DropperID.height
		DropSpeed(DropperIndex) = Int(Rnd * 10) + 10
		
	End If

	DropperID.style.top = Y_Top + DropSpeed(DropperIndex)

End Sub
		
Function GetNum(objStrNum)

	Dim tmpStr

	tmpStr = Replace(objStrNum, "px", "")

	If Not IsNumeric(tmpStr) Then tmpStr = "0"

	GetNum = Int(tmpStr)

End Function