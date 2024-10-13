Public Function RestrainHTML(ByVal HTML As String)
	Dim asProperties As New ArrayList()
	Dim asElements As New ArrayList()
	Dim oReg As Regex
	Dim oTags As MatchCollection
	Dim oTag As Match
	Dim sElement As String
	Dim sTag As String
	Dim oParameters As MatchCollection
	Dim oParameter As Match
	Dim sProperty As String
	Dim sParameter As String
	Dim asProtocols As New ArrayList()
	Dim sProtocol As String
	' Only allow these properties (Enter as LowerCase)
	With asProperties
		.Add("href")
		.Add("border")
		.Add("align")
		.Add("valign")
	End With
	' Only allow these elements (Enter as UpperCase)
	With asElements
		.Add("BR")
		.Add("B")
		.Add("A")
		.Add("TABLE")
		.Add("TR")
		.Add("TH")
		.Add("TD")
	End With
	' Only allow these protocols (enter as lowercase)
	With asProtocols
		.Add("http")
		.Add("https")
		.Add("mailto")
		.Add("ftp")
		.Add("news")
	End With
	' Pull out all HTML tags
	oReg = New Regex("</?([-\w]+)( [^>]+)?>")
	oTags = oReg.Matches(HTML)
	' Loop through each tag
	For Each oTag In oTags
		' Store FullTag, as well as Element name
		sTag = oTag.Value
		sElement = oTag.Groups(1).Value.ToUpper
		' If element is not valid
		If asElements.IndexOf(sElement) = -1 Then
			' Remove it from HTML
			HTML = Replace(HTML, sTag, "", 1, -1, CompareMethod.Text)
		Else
			' parse perameters within tag
			oReg = New Regex(" ([-\w]+)(=(-\w+|""[^""]*""|'[^']*'))?")
			oParameters = oReg.Matches(sTag)
			' Loop throgh each parameter
			For Each oParameter In oParameters
				' Store Parameter, and property name
				sParameter = oParameter.Value
				sProperty = oParameter.Groups(1).Value.ToLower
				sProtocol = oParameter.Groups(3).Value.ToLower
				' If property is not valid
				If asProperties.IndexOf(sProperty) = -1 Then
					' Remove it from HTML
					HTML = Replace(HTML, sParameter, "", 1, -1, CompareMethod.Text)
				Else
					' Validate protocol
					sProtocol = Replace(sProtocol, """", "")
					sProtocol = Replace(sProtocol, "'", "")
					If Not sProtocol.IndexOf(":") = -1 Then
						sProtocol = sProtocol.Substring(0, sProtocol.IndexOf(":")).ToLower
						If asProtocols.IndexOf(sProtocol) = -1 Then
							' Remove it from HTML
							HTML = Replace(HTML, sParameter, "", 1, -1, CompareMethod.Text)
						End If
					End If
				End If
			Next
		End If
	Next
	' Remove Comments
	oReg = New Regex("<!--[\s\S]*?-->")
	HTML = oReg.Replace(HTML, "")
	' remove extra white space
	HTML = Replace(HTML, vbCr, " ")
	HTML = Replace(HTML, vbLf, " ")
	HTML = Replace(HTML, vbTab, " ")
	While Not HTML.IndexOf(" ") = -1
		HTML = Replace(HTML, " ", " ")
	End While
	Return HTML
End Function