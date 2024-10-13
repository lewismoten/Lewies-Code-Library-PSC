    <%
    '**************************************
    ' Name: Forced Download
    ' Description:Forces users to have choic
    '     e of downloading a binary file. This was
    '     becomming a problem with Word documents 
    '     - as they would either open within the b
    '     rowser or open externally. You may need 
    '     to change the FileName variable.
    ' By: Lewis Moten
    '
    '
    ' Inputs:None
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
    FileName = "History12.doc"
    Response.ContentType = "application/octet-stream"
    Response.AddHeader "content-disposition", "attachment; filename=" & FileName
    Set Stream = server.CreateObject("ADODB.Stream")
    Stream.Open
    Stream.LoadFromFile Server.MapPath(FileName)
    Contents = Stream.ReadText
    Response.BinaryWrite Contents
    Stream.Close
    Set Stream = Nothing
    %>