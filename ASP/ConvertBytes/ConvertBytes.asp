<%
    '**************************************
    ' for :Convert Bytes
    '**************************************
    'Copyright (c) 1999, Lewis Moten. All rights reserved.
    '**************************************
    ' Name: Convert Bytes
    ' Description:This code convert bytes in
    '     to larger units such as kilobytes, megab
    '     ytes, gigabytes, and terabytes. I use th
    '     is with the FileSytemObject so that my u
    '     sers do not see a very large number that
    '     doesn't make any sense.
    ' By: Lewis Moten
    '
    '
    ' Inputs:anBytes - the number of bytes t
    '     o convert.
    '
    ' Returns:Returns a string representing 
    '     a larger unit of measurement along with 
    '     the value.
    '
    'Assumes:None
    '
    'Side Effects:Formatting of the number r
    '     eturned can be changed. Some people want
    '     1 or 2 digits after the decimal, others 
    '     do not want a decimal at all. This scrip
    '     t will not return bytes if it is less th
    '     en 1 kilobyte, but you can uncomment the
    '     code to turn that feature on.
    'This code is copyrighted and has limite
    '     d warranties.
    '**************************************
    
    Function ConvertBytes(ByRef anBytes)
    	Dim lnSize			' File Size to be returned
    	Dim lsType			' Type of measurement (Bytes, KB, MB, GB, TB)
    	
    	Const lnBYTE = 1
    	Const lnKILO = 1024						' 2^10
    	Const lnMEGA = 1048576					' 2^20
    	Const lnGIGA = 1073741824				' 2^30
    	Const lnTERA = 1099511627776			' 2^40
    	'	Const lnPETA = 1.12589990684262E+15		' 2^50
    	'	Const lnEXA = 1.15292150460685E+18		' 2^60
    	'	Const lnZETTA = 1.18059162071741E+21	' 2^70
    	'	Const lnYOTTA = 1.20892581961463E+24	' 2^80
    	
    	If anBytes = "" Or Not IsNumeric(anBytes) Then Exit Function
    	
    	If anBytes < 0 Then Exit Function	
    '	If anBytes < lnKILO Then
    '		' ByteConversion
    '		lnSize = anBytes
    '		lsType = "bytes"
    '	Else		
    		If anBytes < lnMEGA Then
    			' KiloByte Conversion
    			lnSize = (anBytes / lnKILO)
    			lsType = "kb"
    		ElseIf anBytes < lnGIGA Then
    			' MegaByte Conversion
    			lnSize = (anBytes / lnMEGA)
    			lsType = "mb"
    		ElseIf anBytes < lnTERA Then
    			' GigaByte Conversion
    			lnSize = (anBytes / lnGIGA)
    			lsType = "gb"
    		Else
    			' TeraByte Conversion
    			lnSize = (anBytes / lnTERA)
    			lsType = "tb"
    		End If
    '	End If
    	' Remove fraction
    	'lnSize = CLng(lnSize)
    	lnSize = FormatNumber(lnSize, 2, True, False, True)
    	
    	' Return the results
    	ConvertBytes = lnSize & " " & lsType
    End Function
    %>
