'**************************************
' Name: GetResource
' Description:Reads embeded files from your executable programs (such as xml, images, etc.) Allows you to distribute one program - the exe, but still access your data and images needed with the program.
' By: Lewis Moten
'Assumes:This only allows you to read the data, not update/delete/add to it, except during design time.
'This code is copyrighted and has limited warranties.
'**************************************

Public Function GetResource(ByVal filename As String) As System.IO.Stream
    ' Demo Instructions:
    '
    ' Add an XML document to your project ca
    '     lled "test.xml"
    '
    ' Click the file and under properties, s
    '     et BuildAction to "Embedded Resource"
    '
    ' paste this routine in one of your clas
    '     ses or forms.
    '
    ' Add the import to the top of your clas
    '     s
    ' Imports System.Xml.XmlDocument
    '
    ' Dim the XML Document object
    ' Dim xmldoc As XmlDocument = New XmlDoc
    '     ument()
    '
    ' Load the XML document from the resourc
    '     e
    ' xmldoc.Load(GetResource("test.xml"))
    '
    ' Congratulations. You have learned how 
    '     to embed files within
    ' your program and read them at run-time
    '     
    '
    Dim oAssembly As Reflection.Assembly = MyClass.GetType.Assembly
    Dim sFiles() As String
    Dim sFile As String
    ' prefix filename with namespace
    filename = oAssembly.GetName.Name & "." & filename
    ' get a list of all embeded files
    sFiles = oAssembly.GetManifestResourceNames()
    ' loop through each file
    For Each sFile In sFiles
		' found the file? return the stream
		If sFile = filename Then
		Return oAssembly.GetManifestResourceStream(sFile)
		End If
    Next
End Function

		
