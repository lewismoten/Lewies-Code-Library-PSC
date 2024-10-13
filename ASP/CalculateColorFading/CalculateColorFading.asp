<%
    '**************************************
    ' Name: Calculate Color Fading
    ' Description:Shows an example of how to
    '     fade from one color to another. This is 
    '     a bare-bones example to help you get sta
    '     rted with calculating a fade across 16 c
    '     olors.
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
    
    ' Assign default values (if not specifie
    '     d)
    If Color1 = "" Then Color1 = "FAD92E"
    If Color2 = "" Then Color2 = "000000"
    ' Parse Byte Values of each primary colo
    '     r
    R(0) = CByte("&h" & Mid(Color1, 1, 2))
    G(0) = CByte("&h" & Mid(Color1, 3, 2))
    B(0) = CByte("&h" & Mid(Color1, 5, 2))
    R(15) = CByte("&h" & Mid(Color2, 1, 2))
    G(15) = CByte("&h" & Mid(Color2, 3, 2))
    B(15) = CByte("&h" & Mid(Color2, 5, 2))
    ' Calculate colors between two values
    For i = 1 To 14
    	R(i) = R(0) + ((i / 15) * (R(15) - R(0)))
    	G(i) = G(0) + ((i / 15) * (G(15) - G(0)))
    	B(i) = B(0) + ((i / 15) * (B(15) - B(0)))
    Next
    ' Compile palette
    For i = 0 To 15
    	p = p & Right("0" & Hex(R(i)), 2)
    	p = p & Right("0" & Hex(G(i)), 2)
    	p = p & Right("0" & Hex(B(i)), 2)
    Next

		
%>