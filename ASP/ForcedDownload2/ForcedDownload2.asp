    <%
    '**************************************
    ' Name: Forced Download 2
    ' Description:Allows you to force a file
    '     to be downloaded rather then displayed w
    '     ithin the users browser. This can be use
    '     d with Word documents, Excel Spreadsheet
    '     s, Adobe PDF's, and other files. Script 
    '     has been optimized to support large down
    '     loads and be dynamically called to downl
    '     oad any file (except ASP, ASPX, ASA, ASA
    '     X, MDB files). Also clears up some probl
    '     ems with currupt files users had been ha
    '     ving by clearing all previouse content a
    '     nd headers.
    ' By: Lewis Moten
    '
    '
    ' Inputs:Pass the location to the file v
    '     ia. the query string.
    'Download.asp?FileName=/Files/Policy.doc
    '
    ' Returns:None
    '
    'Assumes:None
    '
    'Side Effects:None
    'This code is copyrighted and has limite
    '     d warranties.
    '**************************************
    
    Dim Stream
    Dim Contents
    Dim FileName
    Dim FileExt
    Const adTypeBinary = 1
    FileName = Request.QueryString("FileName")
    If FileName = "" Then
    	Response.Write "Filename not specified."
    	Response.End
    End If
    ' Make sure they are not requesting your
    '     code
    FileExt = Mid(FileName, InStrRev(FileName, ".") + 1)
    Select Case UCase(FileExt)
    	Case "ASP", "ASA", "ASPX", "ASAX", "MDB"
    		Response.Write "Protected file not allowed."
    		Response.End
    End Select
    ' Download the file
    Response.Clear
    Response.ContentType = "application/octet-stream"
    Response.AddHeader "content-disposition", "attachment; filename=" & FileName
    Set Stream = server.CreateObject("ADODB.Stream")
    Stream.Type = adTypeBinary
    Stream.Open
    Stream.LoadFromFile Server.MapPath(FileName)
    While Not Stream.EOS
    	Response.BinaryWrite Stream.Read(1024 * 64)
    Wend
    Stream.Close
    Set Stream = Nothing
    Response.Flush
    Response.End
    %>