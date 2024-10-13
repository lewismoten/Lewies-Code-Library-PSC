<%
    '**************************************
    ' Name: Unlimited Alternating Colors
    ' Description:Demonstrates how to have a
    '     lternating colors. Most people usually n
    '     eed just two alternating colors, but thi
    '     s method supports a larger list. Both a 
    '     fixed loop and a conditional loop are pr
    '     ovided as examples.
    ' By: Lewis Moten
    '
    '
    ' Inputs:None
    '
    ' Returns:None
    '
    'Assumes:None
    '
    'Side Effects:None
    'This code is copyrighted and has limite
    '     d warranties.
    '**************************************
    
    Dim strColorAry
    Dim lngCounter
    Dim strColor
    strColorAry = Array("#FF0000", "#00FF00", "#0000FF", "#000000")
    lngColorCount = UBound(strColorAry) + 1
    ' ---- [ Demo using fixed loop ]
    For lngCounter = 0 to 100
    strColor = strColorAry(lngCounter Mod lngColorCount)
    Response.Write("<FONT color=""" & strColor & """>test</FONT><BR>")
    Next
    Response.End
    ' ---- [ Demo using Recordset ]
    lngCounter = 0
    While Not rs.EOF
    strColor = strColorAry(lngCounter Mod lngColorCount)
    Response.Write("<FONT color=""" & strColor & """>" & rs("MyField") & "</FONT><BR>")
    lngCounter = lngCounter + 1
    rs.moveNext
    Wend
%>