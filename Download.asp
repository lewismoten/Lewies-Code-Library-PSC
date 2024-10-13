<%

Dim Folder
Dim FileType ' Source Code / Installation
Dim Node 
Dim Xml
Dim Path
Dim Stream
Dim Contents
Const adTypeBinary = 1
Dim FileName
Dim id
Const ForReading = 1

If Request.Cookies("registrant") = "" Then
	Response.Redirect("Register.htm?" & Request.QueryString)
	Response.End()
End If

id = Request.Cookies("registrant")
Folder = Request.QueryString("Folder")
FileType = Request.QueryString("FileType")
' Log file user is downloading ...

If Folder = "" And FileType = "" Then
	FileName = "LewiesCodeLibrary.zip"
	Path = Server.MapPath(FileName)
Else
	' -------------------------------------
	' Read File Info FROM XML
	' -------------------------------------
	Set Xml = Server.CreateObject("Msxml.DOMDocument")
	Call Xml.setProperty("SelectionLanguage", "XPath")
	Xml.async = false
	Xml.load(Server.MapPath("scripts.xml"))
	Set Node = Xml.selectNodes("//scripts/script[@folder='" & Folder & "']")
	If node.length = 0 Then
		Response.Write("File not found.")
		Response.End()
	End If	

	' -------------------------------------
	' Build Path to File
	' -------------------------------------
	Select Case FileType
		Case "code"
			FileName = "download." & Node(Index).getAttribute("code")
		Case "demo"
			FileName = "demo." & Node(Index).getAttribute("demo")
		Case "install"
			FileName = "install." & Node(Index).getAttribute("install")
		Case Else
			FileName = "download." & Node(Index).getAttribute("code")
	End Select

	Path = Server.MapPath(Folder & "/" & FileName)
End If

' -------------------------------------
' Send File
' -------------------------------------
Response.Clear()
Response.ContentType = "application/octet-stream"
Response.AddHeader "content-disposition", "attachment; filename=" & FileName
Set Stream = server.CreateObject("ADODB.Stream")
Stream.Type = adTypeBinary
Stream.Open
Stream.LoadFromFile Path
While Not Stream.EOS And Response.IsClientConnected
    Response.BinaryWrite Stream.Read(1024 * 64)
    Response.Flush()
Wend
Stream.Close()
Set Stream = Nothing
' -------------------------------------
' Log
' -------------------------------------
Node = "<download"
Node = Node & " date=""" & Server.HTMLEncode(Now) & """"
Node = Node & " folder=""" & Server.HTMLEncode(Folder) & """"
Node = Node & " file=""" & Server.HTMLEncode(FileName) & """"
Node = Node & " registrant=""" & Server.HTMLEncode(id) & """ />"

FileName = Server.MapPath("Download.xml")

Set FSO = Server.CreateObject("Scripting.FileSystemObject")

If(FSO.FileExists(FileName)) Then
	Set File = FSO.OpenTextFile(FileName, ForReading)
	Xml = File.ReadAll()
	File.Close()
Else
	Xml = "<?xml version=""1.0"" ?>" & vbCrLf & "<downloads>" & vbCrLf & "</downloads>"
End If

Xml = Replace(Xml, "<downloads>", "<downloads>" & vbCrLf & vbTab & Node)

Set File = FSO.CreateTextFile(FileName, True)
File.Write(Xml)
File.Close()
Set File = Nothing
Set FSO = Nothing
%>