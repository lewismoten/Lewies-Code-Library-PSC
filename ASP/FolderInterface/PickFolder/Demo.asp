<TITLE>Pick Folder Demonstration</TITLE>
<SCRIPT>
function PickFolder()
{
	var path = document.Form1.Path.value;
	
	path = window.showModalDialog('PickFolder.asp', path, 'dialogWidth:300px; dialogHeight:300px; help:no; status: no;');
	
	document.Form1.Path.value = path;
}
</SCRIPT>

<H3>Folder Demonstration</H3>

<FORM name="Form1">
	<B>Location:</B>
	<INPUT name="Path" size="50">
	<INPUT type="button" onClick="PickFolder()" value="  ...  " id=button1 name=button1>
</FORM>

<P>
	This page demonstrates how you can add the "PickFolder" script to your
	own web page and gives you some interaction.  Feel free to integrate these
	scripts into your own websites.  If possible, please link back to my own
	website at <A href="http://www.lewismoten.com">www.lewismoten.com</A>.
</P>

<P>
	The folders are found within a database that is included with the code.
	It is possible to convert the code to interact with an actual file system
	if you must have this functionality.
</P>
<P>
	The only additional options available at this time is to create a sub folder
	in the current directory.  Other options to consider would be renaming, moving
	folders, and deleting them.
</P>