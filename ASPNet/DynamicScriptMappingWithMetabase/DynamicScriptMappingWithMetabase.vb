'**************************************
' for :Dynamic ScriptMapping with Metabase
'**************************************
'Copyright (c) 2003, Lewis Moten. All rights reserved.
'**************************************
' Name: Dynamic ScriptMapping with Metabase
' Description:Allows you to assign almost any extension to be processed by the ASPX script processor. You may add, update, and remove the extension progromatically without opening the IIS Manager. Great for those of you who do not have access to the machines desktop (such as hosted at other ISPs).
' By: Lewis Moten
'This code is copyrighted and has limited warranties.
'**************************************

Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
    Dim ext As String = "GIF"
    ' Setup extension to be processed by ASPX script processor
    Response.Write("<BR>Add: " & Mapping(ext, MappingAction.Add))
    ' Update extension to be processed by latest ASPX script processor
    Response.Write("<BR>Update: " & Mapping(ext, MappingAction.Update))
    ' Setup extension so that it is not processed by any script processor
    Response.Write("<BR>Delete: " & Mapping(ext, MappingAction.Delete))
End Sub
Enum MappingAction
    Add
    Update
    Delete
End Enum
Private Function Mapping(ByVal ext As String, ByVal action As MappingAction) As Boolean
    ' Assigns the ASPX processor to the extension specified.
    ' See IIS documentation for information regarding script mapping
    ' http://localhost/iishelp/iis/htm/asp/apro9tkj.htm
    ' Example of Script Mapping:
    ' ".gif,C:\WINDOWS\Microsoft.NET\Framework\v1.1.4322\aspnet_isapi.dll,1,GET,HEAD,POST,DEBUG"
    Dim APPL_MD_PATH As String
    Dim AppPath As String
    Dim App As DirectoryEntry
    Dim ScriptMaps As ArrayList
    Dim ExtentionExists As Boolean = False
    Dim ScriptProcessor As String = ""
    Dim Flags As String = "5"
    Dim IncludedVerbs As String = "GET,POST,HEAD"
    Dim NewMap As String = "{0},{1},{2},{3}"
    ' Make sure ext is prefixed with "."
    If Not ext.IndexOf(".") = 0 Then ext = "." & ext
    ' Make sure ext is lowercase
    ext = ext.ToLower
    ' Do our best to prevent problems with default script processors
    Dim BadExt As String
    BadExt = ".asa.asax.ascx.ashx.asmx.asp.aspx.asd.cdx.cer.config."
    BadExt &= "cs.csproj.idc.java.jsl.licx.rem.resources.resx.shtm."
    BadExt &= "shtml.soap.stm.vb.vbproj.vjsproj.vsdisco.webinfo."
    If Not BadExt.IndexOf(ext & ".") = -1 Then
    	Return False
    End If
    Response.Write("got here!")
    Response.End()
    ' Get application metadata path
    APPL_MD_PATH = Request.ServerVariables("APPL_MD_PATH").ToString
    ' Format path for ADSI
    AppPath = Replace(APPL_MD_PATH, "/LM/", "IIS://localhost/")
    ' Attempt to acquire the object
    Try
    	If DirectoryEntry.Exists(AppPath) Then
    		App = New DirectoryEntry(AppPath)
    	End If
    Catch ex As Exception
    	Return False
    End Try
    ' Get a list of all script mappings
    ScriptMaps = New ArrayList(CType(App.Properties("ScriptMaps").Value, Object()))
    ' If we are not deleting a script map
    If Not action = MappingAction.Delete Then
    	' We need to get the latest processor installed
    	' that handles aspx pages.
    	For Each Map As String In ScriptMaps
    		' if the first value of the string is an ASPX page
    		If Split(Map, ",", 4)(0) = ".aspx" Then
    			' The latest script processor is the second
    			' value in the comma delimited string
    			ScriptProcessor = Split(Map, ",", 4)(1)
    			Exit For
    		End If
    	Next
    	' Processor not found!
    	If ScriptProcessor = "" Then
    		App.Dispose()
    		App = Nothing
    		Return False
    	End If
    End If
    ' Find out if extension exists
    For index As Integer = 0 To ScriptMaps.Count() - 1
    	' get current map
    	Dim map As String = ScriptMaps(index).ToString
    	' if we found the extension
    	If Split(Map, ",", 4)(0) = ext Then
    		' Determine results based on action
    		If action = MappingAction.Add Then
    			' impossible to add if it exists
    			App.Dispose()
    			App = Nothing
    			Return False
    		ElseIf action = MappingAction.Update Then
    			' update the map with latest script processor
    			ScriptMaps(index) = NewMap.Format(NewMap, ext, ScriptProcessor, Flags, IncludedVerbs)
    			Exit For
    		ElseIf action = MappingAction.Delete Then
    			' delete the map
    			ScriptMaps.RemoveAt(index)
    			Exit For
    		End If
    	End If
    Next
    ' Is user attempting to add map?
    If action = MappingAction.Add Then
    	' Build the new map
    	ScriptMaps.Add(NewMap.Format(NewMap, ext, ScriptProcessor, Flags, IncludedVerbs))
    End If
    ' Save the new set of script maps
    App.Properties("ScriptMaps").Value = ScriptMaps.ToArray
    App.CommitChanges()
    App.Dispose()
    App = Nothing
    Return True
End Function