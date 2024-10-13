<TITLE>Parsing a users Windows System Colors Demonstration</TITLE>
<IFRAME ID="SystemColor" Name="SystemColor" style="display:none;"></IFRAME>
<H3>Demonstration to Parse a users windows System Colors Hex Values</H3>
<DIV ID="COLORS">Please wait ... </DIV>
<BR>
Created By <A href="http://www.lewismoten.com">Lewis Moten</A>

<SCRIPT language="vbScript">

	Dim SystemColors(27,1)
	
	Const sysColorName = 0
	Const sysColorValue = 1

	' enumerate array
	Const sysActiveBorder			= 0
	Const sysActiveCaption			= 1
	Const sysAppWorkspace			= 2
	Const sysBackground				= 3
	Const sysButtonFace				= 4
	Const sysButtonHighlight		= 5
	Const sysButtonShadow			= 6
	Const sysButtonText				= 7
	Const sysCaptionText			= 8
	Const sysGrayText				= 9
	Const sysHighlight				= 10
	Const sysHighlightText			= 11
	Const sysInactiveBorder			= 12
	Const sysInactiveCaption		= 13
	Const sysInactiveCaptionText	= 14
	Const sysInfoBackground			= 15
	Const sysInfoText				= 16
	Const sysMenu					= 17
	Const sysMenuText				= 18
	Const sysScrollBar				= 19
	Const sysThreeDDarkShadow		= 20
	Const sysThreeDFace				= 21
	Const sysThreeDHighlight		= 22
	Const sysThreeDLightShadow		= 23
	Const sysThreeDShadow			= 24
	Const sysWindow					= 25
	Const sysWindowFrame			= 26
	Const sysWindowText				= 27
	
	' Populate color names
	SystemColors(sysActiveBorder,sysColorName)			= "ActiveBorder"
	SystemColors(sysActiveCaption,sysColorName)			= "ActiveCaption"
	SystemColors(sysAppWorkspace,sysColorName)			= "AppWorkspace"
	SystemColors(sysBackground,sysColorName)			= "Background"
	SystemColors(sysButtonFace,sysColorName)			= "ButtonFace"
	SystemColors(sysButtonHighlight,sysColorName)		= "ButtonHighlight"
	SystemColors(sysButtonShadow,sysColorName)			= "ButtonShadow"
	SystemColors(sysButtonText,sysColorName)			= "ButtonText"
	SystemColors(sysCaptionText,sysColorName)			= "CaptionText"
	SystemColors(sysGrayText,sysColorName)				= "GrayText"
	SystemColors(sysHighlight,sysColorName)				= "Highlight"
	SystemColors(sysHighlightText,sysColorName)			= "HighlightText"
	SystemColors(sysInactiveBorder,sysColorName)		= "InactiveBorder"
	SystemColors(sysInactiveCaption,sysColorName)		= "InactiveCaption"
	SystemColors(sysInactiveCaptionText,sysColorName)	= "InactiveCaptionText"
	SystemColors(sysInfoBackground,sysColorName)		= "InfoBackground"
	SystemColors(sysInfoText,sysColorName)				= "InfoText"
	SystemColors(sysMenu,sysColorName)					= "Menu"
	SystemColors(sysMenuText,sysColorName)				= "MenuText"
	SystemColors(sysScrollBar,sysColorName)				= "ScrollBar"
	SystemColors(sysThreeDDarkShadow,sysColorName)		= "ThreeDDarkShadow"
	SystemColors(sysThreeDFace,sysColorName)			= "ThreeDFace"
	SystemColors(sysThreeDHighlight,sysColorName)		= "ThreeDHighlight"
	SystemColors(sysThreeDLightShadow,sysColorName)		= "ThreeDLightShadow"
	SystemColors(sysThreeDShadow,sysColorName)			= "ThreeDShadow"
	SystemColors(sysWindow,sysColorName)				= "Window"
	SystemColors(sysWindowFrame,sysColorName)			= "WindowFrame"
	SystemColors(sysWindowText,sysColorName)			= "WindowText"

	' Grab the first color
	Call GetColor(0)
	
	
	Sub ReturnColor(ColorName, ColorValue)
		
		ColorValue = Replace(ColorValue, "#", "")
		
		' Save color value within array
		Execute "SystemColors(sys" & ColorName & ", sysColorValue) = """ & ColorValue & """"

		' If we just got the last color
		If ColorName = SystemColors(UBound(SystemColors, 1), sysColorName) Then
		
			' call finishing script
			Call GotAllColors()
			
		Else
		
			' Else get the next color
			Execute "GetColor(sys" & ColorName & " + 1)"

		End If
		
	End Sub

	Sub GetColor(ColorIndex)

		' Get the HEX value of the system color
		SystemColor.location.href = "SystemColor.asp?Color=" & SystemColors(ColorIndex, sysColorName)

	End Sub
	
	Sub GotAllColors()

		' Write out a table of all the system colors and there hex values.

		Dim Index
		Dim MaxIndex
		Dim Table
		MaxIndex = UBound(SystemColors, 1)
		For Index = 0 To MaxIndex
		Table = Table & _
			"<TR>" & _
			"<TD>" & SystemColors(Index, sysColorName) & "</TD>" & _
			"<TD><FONT face=""Courier New"">" & SystemColors(Index, sysColorValue) & "</FONT></TD>" & _
			"<TD bgColor=""" & SystemColors(Index, sysColorValue) & """>&nbsp;&nbsp;&nbsp;&nbsp;</TD>" & _
			"</TR>"
		Next
		COLORS.innerHTML = "<TABLE border=""1"">" & Table & "</TABLE>"
	End Sub
</SCRIPT>
