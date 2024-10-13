Imports System.Xml.Serialization
Imports System.Configuration

Namespace Config
	<Serializable(), XmlRoot("datastore")> _
	Public Class DataStore
		Inherits Config.BaseConfig

		' Assembly type that datastore is created from
		<XmlAttributeAttribute("type")> Public Type As String

		' collection of properties that datastore needs in order
		' to communicate with datasource
		<XmlElement("property")> Public Properties() As Config.KeyPair

		' method of reading properties
		<XmlIgnore()> Default Public ReadOnly Property [Item](ByVal name As String) As String
			Get
				For Each Key As Config.KeyPair In Properties
					If Key.Name = name Then Return Key.value
				Next
				Return ""
			End Get
		End Property

		Public Shared Function Instance() As DataStore
			Return ConfigurationSettings.GetConfig("datastore")
		End Function

	End Class
End Namespace
