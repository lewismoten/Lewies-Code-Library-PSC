<!--#INCLUDE FILE="PaginationNavigation.asp"-->
<%
LngPage			= Request.QueryString("Page")
LngPageCount	= Request.QueryString("PageCount")
StrURL			= Request.QueryString("URL")
LngMax			= Request.QueryString("Max")

' populate with default values
If LngPage = "" Or Not IsNumeric(LngPage) Then LngPage = 1
If LngPageCount = "" Or Not IsNumeric(LngPageCount) Then LngPageCount = 10
If StrURL = ""  Then StrURL = Request.ServerVariables("SCRIPT_NAME")
If LngMax = "" Or Not IsNumeric(LngMax) Then LngMax = 3
%>
<HTML>
	<HEAD>
		<TITLE>Lewis Moten's pagination Navigation</TITLE>
	</HEAD>
	<BODY>
		<H1>Pagination Navigation</H1>

		<P>
			This demonstration will build a set of links to navigate through
			the many pages of data that your dynamic asp page may create.
		</P>
		<FORM>
			<TABLE>
				<TR>
					<TD>Page URL</TD>
					<TD><INPUT name="URL" value="<%=StrURL%>" size="30"></TD>
				</TR>
				<TR>
					<TD>Page Count</TD>
					<TD><INPUT size="3" name="PageCount" value="<%=LngPageCount%>"></TD>
				</TR>
				<TR>
					<TD>Current Page</TD>
					<TD><INPUT name="Page" size="3" value="<%=LngPage%>"></TD> 
				</TR>
				<TR>
					<TD>Max Span</TD>
					<TD><INPUT name="Max" size="3" value="<%=LngMax%>"></TD>
				</TR>
			</TABLE>
			<INPUT type="submit" value="Build Navigation">
		</FORM>
		<%=PaginationNavigation(LngPage, LngPageCount, StrURL, LngMax)%><BR>
		<BR>
		Created By <A href="http://www.lewismoten.com">Lewis Moten</A>
	</BODY>
</HTML>