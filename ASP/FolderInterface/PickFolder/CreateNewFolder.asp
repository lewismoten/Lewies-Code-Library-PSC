<%
Dim PATH

PATH = Request.QueryString("Path")
%>
<LINK rel="stylesheet" type="text/css" href="PickFolder.css">
<BR>
<B>Create New Folder</B>
<FORM name="Form1" method="post" action="CreateNewFolder.Now.asp?<%=Request.QueryString%>">
	Home <%=Replace(PATH, "/", " / ")%> / <BR>
	<INPUT type="text" name="Folder" value="New Folder">
	<INPUT type="button" value="Create" onClick="cmdCreate_onClick()">
</FORM>

<A href="SubFolders.asp?<%=Request.QueryString%>">Go back to list</A>

<SCRIPT>
	function cmdBack_onClick()
	{
		window.location.href = 'SubFolders.asp' + window.location.search
	}
	function cmdCreate_onClick()
	{
		var Folder = document.Form1.Folder.value;
		
		// Empty String?
		if(Folder=='')
		{
			alert('Please enter a name for the new folder.')
			return
		}
		
		// Invalid Characters?
		//		\	/	:	*	?	"	<	>	|

		var re = new RegExp('[\\\\/:\*\?"<>|]')
		if(re.test(Folder))
		{
			alert('A folder can not contain any of the following characters:\n\\ / : * ? " < > |');
			return
		}
		
		// Go for it
		document.Form1.submit()
		
	}
</SCRIPT>