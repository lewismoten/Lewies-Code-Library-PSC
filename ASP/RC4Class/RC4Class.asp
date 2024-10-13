<%
    '**************************************
    ' Name: RC4 Class
    ' Description:Applys Encryption/Decrypti
    '     on to strings. I think just about everyo
    '     ne who has seen my code knows how I love
    '     classes. This version is more "cleaned u
    '     p" and thrown into a nice little class f
    '     or an object oriented feeling. (If only 
    '     ASP was object oriented, I would be a ha
    '     ppy camper). It is also a little more op
    '     timized to run quicker if you change the
    '     Key/Password often.
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
    
    Option Explicit
    Dim ObjRC4
    Set ObjRC4 = New clsRC4
    ObjRC4.Key = "Joe"
    Response.Write	"""" & ObjRC4.Crypt("hello") & """"
    Set ObjRC4 = Nothing
    ' --------------------------------------
    '     ----------------------------------------
    '     
    Class clsRC4
    	
    	Private mStrKey
    	Private mBytKeyAry(255)
    	Private mBytCypherAry(255)
    	
    	Private Sub InitializeCypher()
    		
    		Dim lBytJump
    		Dim lBytIndex
    		Dim lBytTemp
    	
    		For lBytIndex = 0 To 255
    		mBytCypherAry(lBytIndex) = lBytIndex
    		Next
    		' Switch values of Cypher arround based off of index and Key value
    		lBytJump = 0
    		For lBytIndex = 0 To 255
    	
    			' Figure index to switch
    		lBytJump = (lBytJump + mBytCypherAry(lBytIndex) + mBytKeyAry(lBytIndex)) Mod 256
    		
    		' Do the switch
    		lBytTemp					= mBytCypherAry(lBytIndex)
    		mBytCypherAry(lBytIndex)	= mBytCypherAry(lBytJump)
    		mBytCypherAry(lBytJump)		= lBytTemp
    		
    		Next
    	End Sub
    	
    	Public Property Let Key(ByRef pStrKey)
    		Dim lLngKeyLength
    		Dim lLngIndex
    		
    		If pStrKey = mStrKey Then Exit Property
    		lLngKeyLength = Len(pStrKey)
    		If lLngKeyLength = 0 Then Exit Property
    		mStrKey = pStrKey
    		lLngKeyLength = Len(pStrKey)
    		For lLngIndex = 0 To 255
    		mBytKeyAry(lLngIndex) = Asc(Mid(pStrKey, ((lLngIndex) Mod (lLngKeyLength)) + 1, 1))
    		Next
    	End Property
    	
    	Public Property Get Key()
    		Key = mStrKey
    	End Property
    	Public Function Crypt(ByRef pStrMessage)
    		Dim lBytIndex
    		Dim lBytJump
    		Dim lBytTemp
    		Dim lBytY
    		Dim lLngT
    		Dim lLngX
    		
    		' Validate data
    		If Len(mStrKey) = 0 Then Exit Function
    		If Len(pStrMessage) = 0 Then Exit Function
    		Call InitializeCypher()
    		
    		lBytIndex = 0
    		lBytJump = 0
    		For lLngX = 1 To Len(pStrMessage)
    		lBytIndex = (lBytIndex + 1) Mod 256 ' wrap index
    		lBytJump = (lBytJump + mBytCypherAry(lBytIndex)) Mod 256 ' wrap J+S()
    		
    			' Add/Wrap those two	
    		lLngT = (mBytCypherAry(lBytIndex) + mBytCypherAry(lBytJump)) Mod 256
    		
    		' Switcheroo
    		lBytTemp					= mBytCypherAry(lBytIndex)
    		mBytCypherAry(lBytIndex)	= mBytCypherAry(lBytJump)
    		mBytCypherAry(lBytJump)		= lBytTemp
    		lBytY = mBytCypherAry(lLngT)
    			' Character Encryption ...
    		Crypt = Crypt & Chr(Asc(Mid(pStrMessage, lLngX, 1)) Xor lBytY)
    		Next
    		
    	End Function
    End Class
    %>