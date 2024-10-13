<%
    '**************************************
    ' Name: DOS Date/Time Functions
    ' Description:DOS stores date and time w
    '     ith 4 bytes. 2 for date, 2 for time. The
    '     se functions convert those bytes to appr
    '     opriate vbScript date and time values. I
    '     used these within my WinZip file viewer 
    '     scripts because winzip stored dates in t
    '     his format.
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
    
    Private Function GetDate(plngDOSDate)
    	GetDate = DateSerial( _
    		1980 + (plngDOSDatee And &HFE00) \ &H200, _
    		(plngDOSDatee And &H1E0) \ &H20, _
    		plngDOSDate And &H1F)
    End Function
    Private Function GetTime(plngDOSTime)
    	GetTime = TimeSerial( _
    		(plngDOSTime And &HF800) \ &H800, _
    		(plngDOSTime And &H7E0) \ &H20, _
    		(plngDOSTime And &H1F) * 2)
    End Function

%>