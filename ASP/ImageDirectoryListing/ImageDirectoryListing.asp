<%
    '**************************************
    ' Name: Image Directory Listing
    ' Description:Lists thumbnail images wit
    '     hin the current directory. Very simple b
    '     eginner stuff dealing with file scriptin
    '     g object. Handy for viewing your images 
    '     quickly. Just copy the code and paste it
    '     a file called default.asp. Then view the
    '     directory on your website.
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
%>    
    <H1>Image Directory Listing</H1>
    <%
    Dim FSO
    Dim Files
    Dim File
    Dim Count
    Const Columns = 3
    Const ImageWidth = 100
    Set FSO = Server.CreateObject("Scripting.FileSystemObject")
    Set Files = FSO.GetFolder(Server.MapPath("./")).Files
    Set FSO = Nothing
    Response.Write "<TABLE width=""100%"" border=""1"" cellspacing=""0"">"
    Response.Write "<TR>"
    Count = 0
    For Each File In Files
    	Select Case LCase(Right(File.Name, 3))
    		Case "jpg", "gif", "bmp", "png"
    			Count = Count + 1
    			If Count Mod Columns = 1 Then Response.Write "</TR><TR>"
    			
    			Response.Write "<TD align=""center"" valign=""top"">"
    			
    			Response.Write "<A href=""" & File.Name & """>"
    			Response.Write File.Name
    			Response.Write "<BR><IMG src=""" & File.Name & """ border=""1"" width=""" & ImageWidth & """><BR>"
    			Response.Write "</A>"
    			
    			Response.Write "</TD>"
    	End Select
    Next
    Response.Write "</TR>"
    Response.Write "</TABLE>"
    Set File = Nothing
    Set Files = Nothing
%>