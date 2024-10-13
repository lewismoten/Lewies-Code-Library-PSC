<%@ Page Language="vb" AutoEventWireup="false" Codebehind="DemoForm.aspx.vb" Inherits="ValidationCodeImage.DemoForm"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>DemoForm</title>
		<meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
		<meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
		<meta name="vs_defaultClientScript" content="JavaScript">
		<meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
	</HEAD>
	<body>
		<form id="Form1" method="post" runat="server">
			<P>
				<asp:Image id="imgValidation" runat="server" ImageUrl="ValidationImage.aspx"></asp:Image></P>
			<P>
				<asp:Label id="lblMessage" runat="server">Please enter the code that you see above:</asp:Label></P>
			<P>
				<asp:TextBox id="txtCode" runat="server"></asp:TextBox>
				<asp:Button id="btnOK" runat="server" Text="OK"></asp:Button></P>
			<P>&nbsp;</P>
		</form>
	</body>
</HTML>
