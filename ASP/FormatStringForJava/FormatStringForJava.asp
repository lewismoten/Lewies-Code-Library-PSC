<%
    '**************************************
    ' Name: Format string for Java
    ' Description:formats a regular string t
    '     o javascript. i.e. - changes ticks (') t
    '     o proceed with a back slash (\').
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
    
    Function JavaString(ByVal pstrValue)
    	pstrValue = pstrvalue & ""
    	pstrValue = Replace(pstrValue, "\", "\\")
    	pstrValue = Replace(pstrValue, "'", "\'")
    	pstrValue = Replace(pstrValue, vbCr, "\n")
    	pstrValue = Replace(pstrValue, vbLf, "\r")
    	JavaString = pstrValue
    End Function
%>