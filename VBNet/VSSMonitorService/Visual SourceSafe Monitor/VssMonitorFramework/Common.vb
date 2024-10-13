Imports System.Reflection

Public Module Common

	Public Function WriteTextFile(ByVal path As String, ByVal text As String)
		' Write text to a file
		Dim Stream As IO.FileStream
		Try
			Stream = IO.File.Create(path)
			Dim Buffer() As Byte
			Buffer = System.Text.ASCIIEncoding.ASCII.GetBytes(text)
			Stream.Write(Buffer, 0, text.Length)
		Finally
			Stream.Close()
		End Try
	End Function
	Public Function ReadTextFile(ByVal path As String) As String
		' read entire text file
		Return ReadTextFile(path, 0)
	End Function
	Public Function ReadTextFile(ByVal path As String, ByVal position As Integer) As String

		' read text file beginning at a speciffic position

		Dim StartUpPath As String = IO.Path.GetDirectoryName([Assembly].GetEntryAssembly.Location)
		If IO.File.Exists(path) Then
			' path ok
		ElseIf IO.File.Exists(StartUpPath & "\" & path) Then
			path = StartUpPath & "\" & path
#If DEBUG Then
		ElseIf Not StartUpPath.IndexOf("\bin") = -1 Then
			' check for "bin" folder
			path = StartUpPath.Replace("\bin", "") & "\" & path
			If Not IO.File.Exists(path) Then Return ""
#End If
		Else
			Return ""
		End If

		Dim Stream As IO.FileStream
		Try
			Stream = IO.File.OpenRead(path)
			Dim Buffer((Stream.Length - position) - 1) As Byte
			Stream.Position = position
			Stream.Read(Buffer, 0, Buffer.Length)
			Return System.Text.ASCIIEncoding.ASCII.GetString(Buffer)
		Finally
			Stream.Close()
		End Try
	End Function
End Module
