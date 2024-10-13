<%
Dim Email
Dim Url
Dim FileName
Dim FSO
Dim File
Dim Spider
Dim Unknown

Email = Request.QueryString("email")
FileName = Server.MapPath("stats/") & "\" & Email & ".xml"
Url = "http://setiathome2.ssl.berkeley.edu/fcgi-bin/fcgi?cmd=user_xml&email=" & Email

Unknown = Server.MapPath("UnknownUser.xml")

Set FSO = Server.CreateObject("Scripting.FileSystemObject")

' if stats are cached and are not too old, redirect to them.
If FSO.FileExists(FileName) Then
	If FSO.GetFile(FileName).DateLastModified > DateAdd("h", -8, Now()) Then
		SendFile(FileName)
	End If
End If

' Download stats
Set Spider = Server.CreateObject ("Microsoft.XMLHTTP")
Spider.Open "GET", Url, False, "", ""
Spider.Send
Xml = Spider.ResponseText
Set Spider = Nothing

If InStr(1, Xml, "No user with that name was found.", vbTextCompare) > 0 Then
	SendFile(Unknown)
ElseIf InStr(1, Xml, "Web Server Error.", vbTextCompare) > 0 Then
	If FSO.FileExists(FileName) Then
		SendFile(FileName)
	Else
		SendFile(Unknown)
	End If
End If


Xml = Replace(Xml, "<!DOCTYPE userstats SYSTEM ""http://setiathome.ssl.berkeley.edu/xml/userstats.dtd"">", "")

Xml = Replace(Xml, "<userstats>", "<userstats><cache>" & Now() & " EST</cache>")

Set File = FSO.CreateTextFile(FileName, True)
File.Write(Xml)
File.Close()

SendFile(FileName)

Sub SendFile(path)
	Set File = FSO.OpenTextFile(path)
	Response.Write(File.ReadAll())
	File.Close()
	Set File = Nothing
	Set FSO = Nothing
	Response.End()
End Sub
%>