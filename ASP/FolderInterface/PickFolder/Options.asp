<HTML>
	<HEAD>
		<TITLE>Options</TITLE>
		<SCRIPT>
			<!-- // hide
			function cmdCancel_onClick()
			{
				window.top.returnValue = window.top.dialogArguments;
				window.top.close()
			}
			function cmdOK_onClick()
			{
				window.top.close()
			}
			// done hiding -->
		</SCRIPT>
	</HEAD>
	<BODY bgColor="#cccccc">
		<TABLE width="100%" cellspacing="0" cellpadding="0" border="0" bgColor="#cccccc">
			<TR>
				<TD align="center" width="50%">
					<INPUT type="Button" value="    OK    " onClick="cmdOK_onClick()">
				</TD>
				<TD align="center" width="50%">
					<INPUT type="Button" value=" Cancel " onClick="cmdCancel_onClick()">
				</TD>
			</TR>
		</TABLE>
	</BODY>
</HTML>