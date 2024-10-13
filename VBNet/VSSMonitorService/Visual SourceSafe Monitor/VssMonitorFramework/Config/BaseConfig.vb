Imports System.Xml.Serialization
Imports System.Text.ASCIIEncoding
Imports System.Configuration
Imports System.Xml

' Allows a standard method of loading serialized objects
' from the configuation file as any object that inherits
' this one.

Namespace Config
	<Serializable()> _
	 Public MustInherit Class BaseConfig

		Implements IConfigurationSectionHandler

		Public Function Create(ByVal parent As Object, ByVal configContext As Object, ByVal section As System.Xml.XmlNode) As Object Implements System.Configuration.IConfigurationSectionHandler.Create

			' We use xml serialization ...
			Dim Serializer As New XmlSerializer(MyClass.GetType)
			Dim Reader As XmlNodeReader
			Try

				' use the section of the configuration file and
				' deserialize it
				Reader = New XmlNodeReader(section)
				Return Serializer.Deserialize(Reader)

			Finally
				If Not Reader Is Nothing Then
					Reader.Close()
					Reader = Nothing
				End If
			End Try
		End Function
	End Class
End Namespace