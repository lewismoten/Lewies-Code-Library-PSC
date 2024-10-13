'**************************************
' Name: clsRandomString
' Description:Creates a variable length random string according to parameters you specify. MinLength, MaxLength, CharacterSet. Uses a Cryptography RandomNumberGenerator object to ensure an accurately random string.
' By: Lewis Moten
' Inputs:None
' Returns:None
'Assumes:None
'Side Effects:None
'This code is copyrighted and has limited warranties.
'**************************************

'Dim R As clsRandomString = New clsRandomString()
'Response.Write(R.Create & "<BR>")
'Response.Write(R.Create & "<BR>")
'Response.Write(R.Create & "<BR>")

Imports System.Security.Cryptography
Public Class clsRandomString
	' Creates a random password
	Public CharacterSet As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz1234567890"
	Public MinLength As Integer = 1
	Public MaxLength As Integer = 50

	Public Function Create() As String
		Dim Generator As RandomNumberGenerator
		Dim Buffer(MaxLength) As Byte
		Dim Length As Integer
		Dim Index As Integer
		Dim Position As Integer
		Dim Password As String = ""
		Dim CharacterCount As Integer = CharacterSet.Length

		Generator = RandomNumberGenerator.Create()
		Generator.GetBytes(Buffer)

		Length = (Buffer(0) Mod (MaxLength - MinLength)) + MinLength

		For Index = 1 To Length
			Position = Buffer(Index)
			Position = Position Mod CharacterCount
			Password &= CharacterSet.Substring(Position, 1)
		Next

		Return Password
	End Function
End Class