<TITLE>Bottom Tabs Demo</TITLE>

<STYLE>
#BottomTabs
{
	height:19px;
	width: 400px;
}
</STYLE>


Tab clicked: <SPAN id="lblTabClicked">(none)</SPAN><BR><BR>

<object id="BottomTabs" type="text/x-scriptlet">
    <param name=url value="frmBottomTabs.htm">
</object>
<SCRIPT language="vbScript">
Sub Window_onLoad()
	BottomTabs.Add "Design"
	BottomTabs.Add "Source"
	BottomTabs.Add "Quick View"
	BottomTabs.SelectedText = "Source"
End Sub
Sub BottomTabs_onScriptletEvent(ByRef pstrEventName, ByRef pstrEventData)

	Select Case pstrEventName
		Case "onClick()"
			lblTabClicked.innerHTML = pstrEventData
		Case Else
			
	End Select

End Sub
</SCRIPT>
