<%
    '**************************************
    ' for :Find Available File Name
    '**************************************
    'Copyright (C) 2001, Lewis Edward Moten III. All rights reserved.
    '**************************************
    ' Name: Find Available File Name
    ' Description:When a user uploads a file
    '     , this function looks for duplicate file
    '     names. It incriments the file until an a
    '     vailable name is provided and returns th
    '     at name.
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
    
    ' This test writs out "autoexec[1].bat" 
    '     as the available file name.
    strFileName = "autoexec.bat"
    Response.Write "The available file name for """ & strFileName & """ is: "
    Response.Write AvailableFileName("C:\", strFileName)
    ' end test.
    Function AvailableFileName(ByVal pstrPath, ByRef pstrFileName)
    	' Looks for duplicate file names within the folder provided.
    	' if the file name exists, it is incrimented until an available
    	' file name is found.
    	
    	Dim lstrName		' Name of file (without extension)
    	Dim lstrExt			' File extension
    	Dim llngIncriment	' Incriment of file name
    	Dim lobjFSO			' File Scripting Object
    	Dim lstrNewName		' New file name to test
    	
    	' Just in case the path isn't ending with a back slash, we add it.
    	If Not Right(pstrPath, 1) = "\" Then
    		pstrPath = pstrPath & "\"
    	End If
    	
    	' File System Object
    	Set lobjFSO = Server.CreateObject("Scripting.FileSystemObject")
    	
    	' If the file doesn't exist, return the original file name
    	If Not lobjFSO.FileExists(pstrPath & pstrFileName) Then
    		Set lobjFSO = Nothing
    		AvailableFileName = pstrFileName
    		Exit Function
    	End If
    	
    	' Parse file name and extension.
    	lstrName = pstrFilename
    	lstrExt = Mid(lstrName, InStrRev(lstrName, "."))
    	lstrName = Replace(lstrName, lstrExt, "")
    	
    	' prepare incriments
    	llngIncriment = 0
    	
    	' Loop unlit we find an available name
    	Do
    		llngIncriment = llngIncriment + 1
    		lstrNewName = pstrPath & lstrName & "[" & llngIncriment & "]" & lstrExt
    	Loop While lobjFSO.FileExists(lstrNewName)
    	' Clean up	
    	Set lobjFSO = Nothing
    	' Return the available name
    	AvailableFileName = lstrName & "[" & llngIncriment & "]" & lstrExt
    	
    End Function
%>