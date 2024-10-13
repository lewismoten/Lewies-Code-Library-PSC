Private Sub SendEvent(ByRef pstrEvent, ByRef pstrData)
	If Not InScriptlet() Then Exit Sub
    If window.external.frozen Then Exit Sub
	window.external.raiseEvent pstrEvent, pstrData
End Sub

Private Function InScriptlet()
	Dim lblnScriptlet
	lblnScriptlet = False
	On Error Resume Next
	lblnScriptlet = TypeName(window.external.version) = "String"
    InScriptlet = lblnScriptlet
End Function