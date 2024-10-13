<%If Not Session("IsAdmin") Then Response.Redirect("admin.Login.asp")%>

<CENTER><H3><A href="http://www.lewismoten.com/">Lewie's</A> Photo Gallery Administration</H3><H5><A href="admin.asp">Home</A>: Add Category</H5></CENTER>

<FORM method="post" action="admin.Category.Add.Now.asp" id=form1 name=form1>
	<TABLE>
		<TR>
			<TD>Category Name</TD>
			<TD><INPUT name="CategoryName"></TD>
		</TR>
		<TR>
			<TD>&nbsp;</TD>
			<TD><INPUT type="Submit" value="Add" id=Submit1 name=Submit1></TD>
		</TR>
	</TABLE>
</FORM>
