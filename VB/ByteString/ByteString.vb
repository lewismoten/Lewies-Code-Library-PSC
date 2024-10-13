'**************************************
' Name: ByteString
' Description:Returns a string of the Bi
'     t Representation of your Bytes or number
'     s composed of more then 1 byte. I use th
'     is to "peer inside" my bytes. It helps m
'     e with debugging when reading binary fil
'     es.
' By: Lewis Moten
' Inputs:Long - A number to be converted
' Returns:String - Padded with zeros to 
'     make the length of the string a multiple
'     of 4
'This code is copyrighted and has limite
'     d warranties.
'**************************************
Private Function ByteString(ByVal pNumber As Long) As String
    Do
        If pNumber Mod 2 = 1 Then ByteString = "1" & ByteString Else ByteString = "0" & ByteString
        pNumber = pNumber \ 2
    Loop Until pNumber = 0
    ByteString = String$(4 - (Len(ByteString) Mod 4), "0") & ByteString
End Function