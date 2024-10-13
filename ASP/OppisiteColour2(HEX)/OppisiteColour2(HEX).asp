<%
    '**************************************
    ' Name: Oppisite Colour2 (HEX)
    ' Description:Grab the opposite 6 charac
    '     ter HEX color. 
    'This sums up David Weirs Color invert script with just one line of code.
    'See his original code at http://www.pscode.com/vb/scripts/showcode.asp?txtCodeId=7176&lngWId=4
    'Be careful when comming close to middle gray scale (808080). opposites are same.
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
    
    Invert = Right("000000" & Hex(&hFFFFFF - CLng("&h" & Color)), 6)
%>