'**************************************
' Name: Remove Item Confirmation
' Description:Prompt your users if they 
'     wish to remove an item.
' By: Lewis Moten
'This code is copyrighted and has limited warranties.
'**************************************

Public Function RemoveItem(Optional ByRef pstrItemName As String = "item") As Boolean
    Dim llngAnswer As VbMsgBoxResult
    Dim lstrMessage As String
    lstrMessage = "Are you sure you wish to remove this "
    lstrMessage = lstrMessage & pstrItemName & "?"
    llngAnswer = MsgBox(lstrMessage, vbYesNo + vbQuestion, "Confirm")
    RemoveItem = llngAnswer = vbYes
End Function