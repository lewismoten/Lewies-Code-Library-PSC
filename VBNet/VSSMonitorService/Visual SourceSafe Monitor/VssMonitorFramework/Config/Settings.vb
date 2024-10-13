Imports System.Configuration
Imports System.Xml.Serialization
Imports System.Reflection
Imports System.Xml
Namespace Config
	<Serializable(), XmlRoot("settings")> _
	 Public Class Settings
		<XmlAttributeAttribute("fileExists")> Public FileExists As Boolean = True
		<XmlAttributeAttribute("lastModified")> Public LastModified As DateTime = DateTime.MinValue
		<XmlAttributeAttribute("length")> Public Length As Long = 0
		Public Shared Function Instance() As Settings
			Dim FileName As String = [Assembly].GetEntryAssembly.Location.Replace(".exe", ".xml")
			Dim Reader As Xml.XmlTextReader
			Try
				Reader = New Xml.XmlTextReader(FileName)
				Dim Serializer As New XmlSerializer(GetType(Settings))
				Return Serializer.Deserialize(Reader)
			Catch ex As Exception
				Return New Settings
			Finally
				If Not Reader Is Nothing Then Reader.Close()
			End Try
		End Function
		Public Sub Save()
			Dim FileName As String = [Assembly].GetEntryAssembly.Location.Replace(".exe", ".xml")
			Dim Stream As IO.FileStream
			Try
				Stream = New IO.FileStream(FileName, IO.FileMode.Create, IO.FileAccess.Write, IO.FileShare.None)
				Dim Serializer As New XmlSerializer(GetType(Settings))
				Serializer.Serialize(Stream, Me)
			Catch ex As Exception
				System.Diagnostics.Debug.WriteLine(ex.Message)
			Finally
				If Not Stream Is Nothing Then Stream.Close()
			End Try
		End Sub

	End Class
End Namespace
