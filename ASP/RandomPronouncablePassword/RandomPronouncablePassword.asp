<%
    '**************************************
    ' Name: Random Pronouncable Password
    ' Description:Create a random password t
    '     hat is pronouncable. This helps users re
    '     member the password easier and can come 
    '     up with some silly words.
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
    
    ' Created by Lewis Moten
    ' email-lewis@moten.com
    ' Please visit http://www.lewismoten.com
    '     
    Dim i
    For i = 1 to 10
    Response.Write RandomPassword(7, 10) & "<BR>"
    Next
    Function RandomPassword(Min, Max)
    	Dim Password
    	Dim AddConsonant
    	Dim Letter
    	Dim Length
    	Dim Index
    	Dim Random
    	
    	Const DoubleConsonants = "cdfglmnprst"
    	Const SingleConsonants = "bcdfghjklmnprstv"
    	Const Vowels = "aeiou"
    	Randomize()
    	Length = Int(Rnd * (Max - (Min - 1))) + Min
    	While Len(Password) < Length
    		Random = Int(Rnd * 100)
    		If Not Password = "" And AddConsonant And Random < 10 Then
    			Index = Int(Rnd * Len(DoubleConsonants)) + 1
    			Letter = Mid(DoubleConsonants, Index, 1)
    			Password = Password & Letter & Letter
    			AddConsonant = False
    		ElseIf Not Password = "" And AddConsonant And Random < 90 Then
    			Index = Int(Rnd * Len(SingleConsonants)) + 1
    			Letter = Mid(SingleConsonants, Index, 1)
    			Password = Password & Letter
    			AddConsonant = False
    		Else
    			Index = Int(Rnd * Len(Vowels)) + 1
    			Letter = Mid(Vowels, Index, 1)
    			Password = Password & Letter
    			AddConsonant = True
    		End If
    	Wend
    	
    	If Len(Password) > Max Then Password = Left(Password, Max)
    	
    	RandomPassword = Password
    End Function
    %>