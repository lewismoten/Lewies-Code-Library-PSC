Sub window_onload()

	Dim llngIndex
	Dim lstrHTML
	
	For llngIndex = 0 To 134 Step 3
		lstrHTML = lstrHTML & "<SPAN ID=""Cell" 
		lstrHTML = lstrHTML & (llngIndex \ 3) 
		lstrHTML = lstrHTML & """ Class=""Cell"" style=""top:" 
		lstrHTML = lstrHTML & (4 + llngIndex) 
		lstrHTML = lstrHTML & "px;""></SPAN>"
	Next
	
	document.body.innerHTML = document.body.innerHTML & lstrHTML & _
		"<SPAN ID=""Arrow"" style=""top:69px;""></SPAN>"

	Public_SetHue 255, 0, 255
End Sub

Sub Public_SetHue(pbytRed, pbytGreen, pbytBlue)
	
	Dim lbytRed
	Dim lbytGreen
	Dim lbytBlue
	Dim llngPercent
	Dim lstrHEX
	Dim llngY
	
	' White to Hue
	For llngIndex = 0 To 22
		llngPercent = 1 - (llngIndex / 22)
		
		lbytRed = pbytRed + ((255 - pbytRed) * llngPercent)
		lbytGreen = pbytGreen + ((255 - pbytGreen) * llngPercent)
		lbytBlue = pbytBlue + ((255 - pbytBlue) * llngPercent)
		
		lstrHEX = _
			Right("0" & Hex(lbytRed), 2) & _
			Right("0" & Hex(lbytGreen), 2) & _
			Right("0" & Hex(lbytBlue), 2)
		
		document.all("Cell" & llngIndex).style.backgroundColor = lstrHEX
	Next
	
	' Hue to Black
	For llngIndex = 0 To 21
		llngPercent = 1 - (llngIndex / 21)
		
		lbytRed = pbytRed * llngPercent
		lbytGreen = pbytGreen * llngPercent
		lbytBlue = pbytBlue * llngPercent
		
		lstrHEX = _
			Right("0" & Hex(lbytRed), 2) & _
			Right("0" & Hex(lbytGreen), 2) & _
			Right("0" & Hex(lbytBlue), 2)

		document.all("Cell" & llngIndex + 23).style.backgroundColor = lstrHEX
	Next
	
	llngY = Replace(Arrow.style.top, "px", "")
	llngY = llngY \ 3
	Call SendEvent("Cell_onChange()", document.all("Cell" & llngY).style.backgroundColor)
	
End Sub

Sub Document_onClick()
	
	If Not window.event.srcElement.className = "Cell" Then Exit Sub
	
	Arrow.style.top = Replace(window.event.srcElement.style.top, "px", "") - 2
	
	Call SendEvent("Cell_onClick()", window.event.srcElement.style.backgroundColor)
	
End Sub
Dim mblnDrag
Sub document_onmousedown
	mblnDrag = True
End Sub
Sub document_onmouseup
	mblnDrag = False
End Sub

Sub Document_onMouseOver()
	If Not mblnDrag Then Exit Sub
	
	If Not window.event.srcElement.className = "Cell" Then Exit Sub
	
	Arrow.style.top = Replace(window.event.srcElement.style.top, "px", "") - 2
	
	Call SendEvent("Cell_onClick()", window.event.srcElement.style.backgroundColor)
End Sub
