Private lngTabStart
Private mstrSelectedText
lngTabStart = 0
Private Sub Document_onClick()
	Dim lobjSource
	Dim lobjDIV
	Dim blnScriptlet
	
	Select Case window.event.srcElement.tagName
		Case "BODY", "HR": Exit Sub
	End Select
	
	Set lobjSource = window.event.srcElement
	While Not lobjSource.tagName = "DIV"
		Set lobjSource = lobjSource.parentElement
	Wend
	
	public_put_SelectedText lobjSource.title

	Call SendEvent("onClick()", lobjSource.title)

End Sub

Sub public_put_SelectedText(ByRef pstrText)
	Dim lobjDIV
	mstrSelectedText = pstrText
	For Each lobjDIV In document.all.tags("DIV")
		If lobjDIV.title = pstrText Then
			lobjDIV.ClassName = "Over"
		Else
			lobjDIV.ClassName = "Out"
		End If
	Next
End Sub
Function public_get_SelectedText()
	public_get_SelectedText = mstrSelectedText
End Function

Public Sub public_Add(ByRef pstrText)
	Dim llngWidth ' estimated width of text on button
	Dim lstrHTML
	Dim llngIndex
	
	' Find out how long text is
	llngWidth = (Len(pstrText) * 7) + 14

	' Create Container	
	lstrHTML = lstrHTML & "<DIV" & _
		" class=""Out"" " & _
		" title=""" & pstrText & """" & _
		" style=""" & _
			"width:" & (llngWidth + 20) & "px;" & _
			"left:" & lngTabStart & "px;" & _
		""">"
	
	' Build Frame (Layer 1)
	lstrHTML = lstrHTML & "<SPAN class=""Frame"">"
	For llngIndex = 1 to 9
		lstrHTML = lstrHTML & "<SPAN style=""" & _
			"width:" & (llngWidth + 2 + ((10 - llngIndex)* 2)) & "px;" & _
			"height:2px;" & _
			"top:" & ((llngIndex-1) * 2) & "px;" & _
			"left:" & (lngTabStart + llngIndex) & "px;" & _
			"""></SPAN>"
	Next
	lstrHTML = lstrHTML & "</SPAN>"
	
	' Build Button BgColor (Layer 2)
	lstrHTML = lstrHTML & "<SPAN class=""Button"">"
	For llngIndex = 1 to 9
		lstrHTML = lstrHTML & "<SPAN style=""" & _
			"width:" & (llngWidth + ((10 - llngIndex)* 2)) & "px;" & _
			"height:2px;" & _
			"top:" & ((llngIndex-1) * 2) & "px;" & _
			"left:" & (1 + lngTabStart + llngIndex) & "px;" & _
			"""></SPAN>"
	Next
	lstrHTML = lstrHTML & "<SPAN class=""TopLine"" style=""" & _
		"width:" & (llngWidth + 18) & "px;" & _
		"height:1px;" & _
		"top:0px;" & _
		"left:" & (2 + lngTabStart) & "px;" & _
		"""></SPAN>"
	lstrHTML = lstrHTML & "<SPAN class=""BottomLine"" style=""" & _
		"width:" & (llngWidth + 2) & "px;" & _
		"height:1px;" & _
		"top:18px;" & _
		"left:" & (10 + lngTabStart) & "px;" & _
		"""></SPAN>"
	lstrHTML = lstrHTML & "</SPAN>"
	
	' Build Text (Layer 3)
	lstrHTML = lstrHTML & "<SPAN class=""Text"" style=""" & _
		"width:" & (llngWidth + 20) & "px;" & _
		"left:" & lngTabStart & "px;" & _
		""">" & pstrText & "</SPAN>"
	
	lstrHTML = lstrHTML & "</DIV>"
	
	document.body.innerHTML = document.body.innerHTML & lstrHTML
	lngTabStart = lngTabStart + llngWidth + 9

End Sub
