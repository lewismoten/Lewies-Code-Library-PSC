<%
    '**************************************
    ' for :Credit Card Mod 10 Validation
    '**************************************
    'Copyright (c) 1999, Lewis Moten
    '**************************************
    ' Name: Credit Card Mod 10 Validation
    ' Description:Based on ANSI X4.13, the L
    '     UHN formula (also known as the modulus 1
    '     0 -- or mod 10 -- algorithm ) is used to
    '     generate and/or validate and verify the 
    '     accuracy of credit-card numbers.
    ' By: Lewis Moten
    '
    '
    ' Inputs:asCardType - Type of credit car
    '     d. (American Express, Discover, Visa, Ma
    '     sterCard)
    'anCardNumber - The number appearing on the card. Dashes and spaces are ok. Numbers are stripped from the data provided.
    '
    ' Returns:Returns a boolean (true/false)
    '     determining if the number appears to be 
    '     valid or not.
    '
    'Assumes:The user needs to be able to lo
    '     ok through the code and determine wich s
    '     trings represent the cards. You can eith
    '     er type out the entire card name (i.e. "
    '     American Express") or type in just a let
    '     ter representing the card name (i.e. "a"
    '     )
    '
    'Side Effects:Just because the function 
    '     returns that the card is valid, there ar
    '     e several other things that are not bein
    '     g validated.
    'Date - make sure the card has not expired
    'Active Account - This script does not communicate with any banks to determine if the account number is active
    'Authorization - again, this script does not communicate with any banks to determine if the card has authorization to purchase a product.
    'This code is copyrighted and has limite
    '     d warranties.
    '**************************************
    
    	Function IsCreditCard(ByRef asCardType, ByRef anCardNumber)
    		' Performs a Mod 10 check to make sure the credit card number
    		' appears valid
    		' Developers may use the following numbers as dummy data:
    		' Visa:					430-00000-00000
    		' American Express:		372-00000-00000
    		' Mastercard:			521-00000-00000
    		' Discover:				620-00000-00000
    			
    		Dim lsNumber		' Credit card number stripped of all spaces, dashes, etc.
    		Dim lsChar			' an individual character
    		Dim lnTotal			' Sum of all calculations
    		Dim lnDigit			' A digit found within a credit card number
    		Dim lnPosition		' identifies a character position in a string
    		Dim lnSum			' Sum of calculations for a specific set
    			
    		' Default result is false
    		IsCreditCard = False
    			
    		' ====
    		' Strip all characters that are not numbers.
    		' ====
    			
    		' Loop through each character inthe card number submited
    		For lnPosition = 1 To Len(anCardNumber)
    			' Grab the current character
    			lsChar = Mid(anCardNumber, lnPosition, 1)
    			' If the character is a number, append it to our new number
    			If IsNumeric(lsChar) Then lsNumber = lsNumber & lsChar
    			
    		Next ' lnPosition
    			
    		' ====
    		' The credit card number must be between 13 and 16 digits.
    		' ====
    		' If the length of the number is less then 13 digits, then exit the routine
    		If Len(lsNumber) < 13 Then Exit Function
    			
    		' If the length of the number is more then 16 digits, then exit the routine
    		If Len(lsNumber) > 16 Then Exit Function
    			
    			
    		' ====
    		' The credit card number must start with: 
    		'	4 for Visa Cards 
    		'	37 for American Express Cards 
    		'	5 for MasterCards 
    		'	6 for Discover Cards 
    		' ====
    			
    		' Choose action based on type of card
    		Select Case LCase(asCardType)
    			' VISA
    			Case "visa", "v"
    				' If first digit not 4, exit function
    				If Not Left(lsNumber, 1) = "4" Then Exit Function
    			' American Express
    			Case "american express", "americanexpress", "american", "ax", "a"
    				' If first 2 digits not 37, exit function
    				If Not Left(lsNumber, 2) = "37" Then Exit Function
    			' Mastercard
    			Case "mastercard", "master card", "master", "m"
    				' If first digit not 5, exit function
    				If Not Left(lsNumber, 1) = "5" Then Exit Function
    			' Discover
    			Case "discover", "discovercard", "discover card", "d"
    				' If first digit not 6, exit function
    				If Not Left(lsNumber, 1) = "6" Then Exit Function
    				
    			Case Else
    		End Select ' LCase(asCardType)
    			
    		' ====
    		' If the credit card number is less then 16 digits add zeros
    		' to the beginning to make it 16 digits.
    		' ====
    		' Continue loop while the length of the number is less then 16 digits
    		While Not Len(lsNumber) = 16
    				
    			' Insert 0 to the beginning of the number
    			lsNumber = "0" & lsNumber
    			
    		Wend ' Not Len(lsNumber) = 16
    			
    		' ====
    		' Multiply each digit of the credit card number by the corresponding digit of
    		' the mask, and sum the results together.
    		' ====
    			
    		' Loop through each digit
    		For lnPosition = 1 To 16
    				
    			' Parse a digit from a specified position in the number
    			lnDigit = Mid(lsNumber, lnPosition, 1)
    				
    			' Determine if we multiply by:
    			'	1 (Even)
    			'	2 (Odd)
    			' based on the position that we are reading the digit from
    			lnMultiplier = 1 + (lnPosition Mod 2)
    				
    			' Calculate the sum by multiplying the digit and the Multiplier
    			lnSum = lnDigit * lnMultiplier
    				
    			' (Single digits roll over to remain single. We manually have to do this.)
    			' If the Sum is 10 or more, subtract 9
    			If lnSum > 9 Then lnSum = lnSum - 9
    				
    			' Add the sum to the total of all sums
    			lnTotal = lnTotal + lnSum
    			
    		Next ' lnPosition
    			
    		' ====
    		' Once all the results are summed divide
    		' by 10, if there is no remainder then the credit card number is valid.
    		' ====
    		IsCreditCard = ((lnTotal Mod 10) = 0)
    			
    	End Function ' IsCreditCard

%>