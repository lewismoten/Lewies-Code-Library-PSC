<%
    '**************************************
    ' for :File/Path/Extension Stripping
    '**************************************
    'Copyright (c) 1999, Lewis Moten. All rights reserved.
    '**************************************
    ' Name: File/Path/Extension Stripping
    ' Description:These routines can strip a
    '     file name from a path, a directory from 
    '     a path, and an extension from a path. Th
    '     ey can also determine the parent directo
    '     ry path. They are fairly simple routines
    '     that I use in misc. places.
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
    
    '---------------------------------------
    '     ---------------------------------------
    Function ParentDirectory(ByVal asDirectory)
    	If Len(asDirectory) = 0 Then Exit Function
    	asDirectory = Replace(asDirectory, "/", "\")
    	If Right(asDirectory, 1) = "\" Then
    		asDirectory = Left(asDirectory, Len(asDirectory) - 1)
    	End If
    	If asDirectory = "" Then Exit Function
    	If InStr(1, asDirectory, "\") = 0 Then Exit Function
    	asDirectory = Left(asDirectory, InStrRev(asDirectory, "\"))
    	ParentDirectory = asDirectory
    End Function
    '---------------------------------------
    '     ---------------------------------------
    Function CurrentDirectory()
    	CurrentDirectory = StripDirectory(Request.ServerVariables("PATH_TRANSLATED"))
    '	CurrentDirectory = Server.MapPath("/")
    '     
    End Function
    '---------------------------------------
    '     ---------------------------------------
    Function StripDirectory(ByVal asPath)
    	If asPath = "" Then Exit Function
    	asPath = Replace(asPath, "/", "\")
    	If InStr(1, asPath, "\") = 0 Then Exit Function
    	asPath = Left(asPath, InStrRev(asPath, "\"))
    	StripDirectory = asPath
    End Function
    '---------------------------------------
    '     ---------------------------------------
    Function StripFileName(ByVal asPath)
    	If asPath = "" Then Exit Function
    	asPath = Replace(asPath, "/", "\")
    	If InStr(asPath, "\") = 0 Then Exit Function
    	If Right(asPath, 1) = "\" Then Exit Function
    	
    	StripFileName = Right(asPath, Len(asPath) - InStrRev(asPath, "\"))
    End Function
    '---------------------------------------
    '     ---------------------------------------
    Function StripFileExt(sFileName)
    	If sFileName = "" Then Exit Function
    	If InStr(1, sFileName, ".") = 0 Then Exit Function
    	StripFileExt = Right(sFileName, Len(sFileName) - InStrRev(sFileName, ".") + 1)
    End Function
    '---------------------------------------
    '     ---------------------------------------

%>