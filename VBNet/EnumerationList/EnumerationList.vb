'**************************************
' Name: EnumerationList
' Description:List the string values of any enumerated item that is flagged in your enumeration.
' By: Lewis Moten
'This code is copyrighted and has limited warranties.
'**************************************

Private Function EnumerationList(ByVal Enumeration As [Enum]) As String
    Dim Values As Array
    Dim Item As [Enum]
    Dim list As String = ""
    Values = Enumeration.GetValues(Enumeration.GetType)
    For Each Item In Values
    	Dim ItemValue As Int32 = CType(CType(Item, Object), Int32)
    	Dim EnumerationValue As Int32 = CType(CType(Enumeration, Object), Int32)
    	If CType((ItemValue And EnumerationValue), Boolean) Then list &= ", " & Item.ToString
    Next
    If Not list.Length = 0 Then list = list.Remove(0, 2)
    Return list
End Function