<%
Sub TemplateHeader(ByRef asFormAction)

	Dim lsHTTP_REFERER	' Page that referred the user to this application.
	
	' If form data was not received
	If Not Request.Form = "" Then
		
		
		' OriginalURL must be retrieved from Form Data
		lsHTTP_REFERER = Request.Form("HTTP_REFERER")
	
	' If form data was not received
	ElseIf Not Request.QueryString = "" Then
		
		
		' OriginalURL must be retrieved from ServerVariables
		lsHTTP_REFERER = Request.QueryString("HTTP_REFERER")
		
	' Else form data was received
	Else
		
		' OriginalURL must be retrieved from Server Variables
		lsHTTP_REFERER = Request.ServerVariables("HTTP_REFERER")
		
	End If ' Request.Form = ""

	If Not InStr(1, LCase(lsHTTP_REFERER), "/login/") = 0 Then
		lsHTTP_REFERER = ""
	End If
	%>
	<BODY leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0" marginheight="0" marginwidth="0" scroll="no" bgcolor="#888888">
		<TABLE width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
			<TR>
				<TD valign="middle" align="center">
					<FORM method="post" action="<%=asFormAction%>" Name="NovellClient" ID="NovellClient">
						<INPUT type="hidden" name="HTTP_REFERER" id="HTTP_REFERER" value="<%=Server.HTMLEncode(lsHTTP_REFERER)%>">
						<CENTER>
							<!-- Image Map for title -->
							<MAP name="TitleMap">
								<AREA shape="rect" coords="455,5,470,18" href="/" alt="Close Window">
							</MAP>
							<TABLE border="0" cellspacing="0" cellpadding="0" bgcolor="#C0C0C0" align="center" width="476">
								<TR>
									<TD colspan="3"><IMG src=".\title.gif" width="476" height="22" border="0" ismap usemap="#TitleMap" alt="Novell Login"></TD>
								</TR>
								<TR>
									<TD background=".\left.gif" width="9" valign="top"><IMG src=".\left.gif" width="9" height="100%" alt="" border="0"></TD>
									<TD valign="top" bgcolor="#C0C0C0">
										<CENTER>
											<IMG src=".\banner.gif" width="458" height="81" vspace="10" border="0" alt="Novell Client for WEB SERVICES">
										</CENTER>
										<FONT face="Helvetica,Arial,Sans-Serif" size="2">
<%End Sub%>
<%Sub TemplateFooter()%>
										</FONT>
									</TD>
									<TD background=".\right.gif" width="9" valign="top"><IMG src=".\right.gif" width="9" height="100%" alt="" border="0"></TD>
								</TR>
								<TR>
									<TD colspan="3"><IMG src=".\bottom.gif" width="476" height="5" alt="" border="0"></TD>
								</TR>
							</TABLE>
						</CENTER>
					</FORM>
				</TD>
			</TR>
		</TABLE>
	</BODY>
<%End Sub%>