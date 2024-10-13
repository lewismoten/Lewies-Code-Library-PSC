Imports System.Configuration
Imports System.Xml.Serialization

Namespace Config
	<Serializable(), XmlRoot("users")> _
	Public Class Users
		Inherits Config.BaseConfig

		' Default format of dates for any user not included on this list
		' MM/DD/YY (usa)
		' DD/MM/YY (uk)
		' YY/MM/DD (computer)
		<XmlAttributeAttribute("dateFormat")> Public DefaultDateFormat As String = "MM/DD/YY"

		' default time zone offset (in hours) from universal time
		<XmlAttributeAttribute("timeZone")> Public DefaultTimzeZone As Integer = 0

		' A collection of users within source safe that will have there
		' dates adjusted accordingly
		<XmlElement("user")> Public Users() As User

		Public Shared Function Instance() As Users
			Return ConfigurationSettings.GetConfig("users")
		End Function

		Public Class User

			' VSS username of user (look at journal file if you don't know)
			<XmlAttributeAttribute("name")> Public Name As String

			' format of date that users VSS client uses.
			' MM/DD/YY (usa)
			' DD/MM/YY (uk)
			' YY/MM/DD (computer)
			<XmlAttributeAttribute("dateFormat")> Public DateFormat As String

			' time zone offset (in hours) from universal time
			<XmlAttributeAttribute("timeZone")> Public TimeZone As Integer = 99

		End Class
	End Class
End Namespace