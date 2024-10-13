'**************************************
' Name: CalcBytes
' Description:converts number of bytes to a lower number, but changes measurement. (can changed kb's to gb's)
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
'This code is copyrighted and has limited warranties.
'**************************************

Private Shared Function CalcBytes(ByVal bytes As Double, ByVal type As String) As String
    While bytes > 768
    	Select Case type
    		Case "KB"
    			type = "MB"
    			bytes = bytes / 1024
    		Case "MB"
    			type = "GB"
    			bytes = bytes / 1024
    		Case "GB"
    			type = "TB"
    			bytes = bytes / 1024
    		Case "TB"
    			type = "PB"
    			bytes = bytes / 1024
    		Case Else
    			Exit While
    	End Select
    End While
    bytes = CType(CType(bytes * 100, Integer), Double) * 0.01
    Return bytes.ToString & " " & type
End Function