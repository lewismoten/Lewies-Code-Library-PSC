<%
Call Main()
Response.Redirect("transparent.gif")

Sub Main()

	Dim Query
	Dim Language
	Dim Node
	Dim Registrant
	Dim FileName
	Dim Xml
	Dim FSO
	Dim File
	Const ForReading = 1

	Query = Request.QueryString("query")
	Language = Request.QueryString("language")
	Registrant = Request.Cookies("registrant")

	If Query = "" Then Exit Sub

	Node = "<query text=""" & Server.HTMLEncode(Query) & """"
	Node = Node & " language=""" & Server.HTMLEncode(Language) & """"
	Node = Node & " date=""" & Server.HTMLEncode(Now) & """"
	Node = Node & " registrant=""" & Server.HTMLEncode(Registrant) & """ />"
	
	FileName = Server.MapPath("Browse.xml")
	
	Set FSO = Server.CreateObject("Scripting.FileSystemObject")

	If(FSO.FileExists(FileName)) Then
		Set File = FSO.OpenTextFile(FileName, ForReading)
		Xml = File.ReadAll()
		File.Close()
	Else
		Xml = "<?xml version=""1.0"" ?>" & vbCrLf & "<queries>" & vbCrLf & "</queries>"
	End If

	Xml = Replace(Xml, "<queries>", "<queries>" & vbCrLf & vbTab & Node)

	Set File = FSO.CreateTextFile(FileName, True)
	File.Write(Xml)
	File.Close()
	Set File = Nothing
	Set FSO = Nothing

End Sub
%>