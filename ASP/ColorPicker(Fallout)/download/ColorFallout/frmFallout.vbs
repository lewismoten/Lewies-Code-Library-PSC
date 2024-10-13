Option Explicit

Dim Public_Red			' Red Luminance
Dim Public_Green		' Green Luminance
Dim Public_Blue			' Blue Luminance
Dim Public_Saturation	' Percentage of saturation
Dim Public_Step			' Number of pixels/steps to the next Primary & Secondary colors
Dim Public_Section		' Current color fading section

Function Public_Get_Color()
	Public_Get_Color = _
		Right("0" & Hex(Public_Red), 2) & _
		Right("0" & Hex(Public_Green), 2) & _
		Right("0" & Hex(Public_Blue), 2)
End Function

Sub ChangePosition()
	
	Dim llngIncriment
	
	' Incriment determines the number of luminance to add/subtract
	' on the steps between each primary & secondary color.
	
	llngIncriment = 8.7931034482758620689655172413793
	
	Public_Section = Fix(window.event.x / 30)
	Public_Step = window.event.x Mod 30
	Public_Saturation = 1 - (window.event.y / 140)
	
	' Determine max luminance of each color based on section
	Select Case Public_Section
		Case 0 ' Red To Yellow
			Public_Red = 255
			Public_Green = Public_Step * llngIncriment
			Public_Blue = 0
		Case 1 ' Yellow To Green
			Public_Red = 255 - (Public_Step * llngIncriment)
			Public_Green = 255
			Public_Blue = 0
		Case 2 ' Green To Cyan
			Public_Red = 0
			Public_Green = 255
			Public_Blue = Public_Step * llngIncriment
		Case 3 ' Cyan To Blue
			Public_Red = 0
			Public_Green = 255 - (Public_Step * llngIncriment)
			Public_Blue = 255
		Case 4 ' Blue To Purple
			Public_Red = Public_Step * llngIncriment
			Public_Green = 0
			Public_Blue = 255
		Case 5 ' Purple To Red
			Public_Red = 255
			Public_Green = 0
			Public_Blue = 255 - (Public_Step * llngIncriment)
	End Select
		
	' Apply Saturation
	ApplySaturation Public_Red, Public_Saturation
	ApplySaturation Public_Green, Public_Saturation
	ApplySaturation Public_Blue, Public_Saturation
	
End Sub
Sub ApplySaturation(ByRef pbytColor, ByRef pdblSaturation)
	pbytColor = Fix((127 * (1 - pdblSaturation)) + (pbytColor * pdblSaturation))
End Sub

Dim mblnDrag
Sub Document_onMouseDown()
	Cursor.style.display = "none"
	document.body.style.cursor = "crosshair"
	mblnDrag = True
	Call ChangePosition()
End Sub
Sub Document_onClick()
	Cursor.style.display = "block"
	document.body.style.cursor = "default"
	mblnDrag = False
	Cursor.style.left = window.event.x - 9
	Cursor.style.top = window.event.y - 9
	Call ChangePosition()
	Call SendEvent("Fallout_onClick()", "")
End Sub
Sub Document_onMouseUp()
	Cursor.style.display = "block"
	Cursor.style.left = window.event.x - 9
	Cursor.style.top = window.event.y - 9
	document.body.style.cursor = "default"
	mblnDrag = False
	Call ChangePosition()
	Call SendEvent("Fallout_onClick()", "")
End Sub
Sub Document_onMouseMove()
	If mblnDrag Then Call ChangePosition()
End Sub
