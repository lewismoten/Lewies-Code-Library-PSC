Imports System.Configuration
Imports System.Xml.Serialization
Imports VMS.Framework

Namespace Config
	<Serializable(), XmlRoot("alerts")> _
	Public Class Alerts

		Inherits Config.BaseConfig

		' Alerts must be delivered through a service.  Multiple
		' types of services may be setup to deliver these alerts.
		<XmlElement("service")> Public Services() As Service

		' This is the timezone offset that all fixed times are based on.
		' This value represents the hours.
		' EST = -5 GMT.  So the value would be -5
		<XmlAttributeAttribute("timeZone")> Public Offset As Double

		' Instance lets you get a the alerts from the config file
		Public Shared Function Instance() As Alerts
			Return ConfigurationSettings.GetConfig("alerts")
		End Function

		Public Class Service

			' Type is a string representation of the type to instantiate
			' when crating a new service.  You can create your own service
			' handlers by inheriting the Services.BaseService type
			<XmlAttributeAttribute("type")> Public Type As String

			' each service may send multiple alerts at different times to
			' different recipients.
			<XmlElement("alert")> Public Alerts() As Alert

			' A service may have many properties including login information
			' and how messages are sent through it.
			<XmlElement("property")> Public Properties() As Config.KeyPair

			' This property is used to read properties associated with the service.
			<XmlIgnore()> Default Public ReadOnly Property [Item](ByVal name As String) As String
				Get
					For Each Key As Config.KeyPair In Properties
						If Key.Name = name Then Return Key.value
					Next
					Return ""
				End Get
			End Property

			Public Class Alert

				' Number of minutes that must pass after then last email
				' before you send anotherone.
				<XmlAttributeAttribute("deliveryCycle")> Public DeliveryCycle As Integer = -1

				' Time of day that an email is sent.  If the time of day is specified,
				' then it overrides the delivery cycle.
				<XmlAttributeAttribute("fixedTime")> Public FixedTime As String = ""

				' Filter used to query VSS actions.  Filter can be specified within
				' the element tag, or specify a file to read using the "template" attribute.
				<XmlElement("filter")> Public Filter As MessageFormat.Template

				' Array of individuals who may receive the alert.
				<XmlElement("recipient")> Public Recipients() As Recipient

				' Structure of the message that is to be sent
				<XmlElement("message")> Public Message As MessageFormat

				' Time that the last message was sent for this alert.
				<XmlIgnore()> Public LastSent As DateTime = DateTime.Now.ToUniversalTime

				' Each alert can have properties as well
				<XmlElement("property")> Public Properties() As Config.KeyPair

				' property to read properties
				<XmlIgnore()> Default Public ReadOnly Property [Item](ByVal name As String) As String
					Get
						For Each Key As Config.KeyPair In Properties
							If Key.Name = name Then Return Key.value
						Next
						Return ""
					End Get
				End Property
				Public Class Recipient

					' This will often be an email address when using an Smtp service.
					' Identifies the account that the message is to be sent to.
					<XmlAttributeAttribute("address")> Public Address As String

					' Name of the recipient
					<XmlAttributeAttribute("name")> Public Name As String

				End Class

				Public Class MessageFormat

					' In order to determine where data appears when renduring
					' each item, we use string formatting.  Take the following
					' for example:

					' My name is {4}.

					' formatters are zero-based, so the "{4}" will be replaced
					' with the fith variable passed to the formatter.  Take a 
					' look at the following list to determine what each index 
					' represents:

					' 0 – Date
					' 1 – Path
					' 2 – Version
					' 3 – Action
					' 4 – User
					' 5 – Event
					' 6 – Comment

					' Along with this functionality, you can also format your
					' values with special attributes.  Using {0:D} will format
					' the first element as a date.  A customized formatter is
					' used with this software that allows you to go one step
					' further.  You can specify the number of hours to add or
					' subtract from the universal time when the event occured.
					' For example: {0:t-5} will display a time that is 5 hours
					' behind then the date passed to it, allowing it to appear
					' in Eastern Standard Time (-5 GMT).

					' Subject of the message (email)
					<XmlAttributeAttribute("subject")> Public Subject As String

					' Content that appears before listing each item
					<XmlElement("header")> Public Header As Template

					' odd items
					<XmlElement("item")> Public Item As Template

					' even items
					<XmlElement("alternateItem")> Public AlternateItem As Template

					' content that appears at the end of the item list
					<XmlElement("footer")> Public Footer As Template
					Public Class Template

						' text file where text can be found.
						<XmlAttributeAttribute("template")> Public File As String

						' text to be used in formatting.
						<XmlText()> Public Text As String

						' Content looks at the the text property.  If nothing
						' is specified, then it attempts to load a file and cache
						' the results into the text property.
						<XmlIgnore()> Public ReadOnly Property Content() As String
							Get
								If Not Text = "" Then Return Text
								Text = Common.ReadTextFile(File)
								Return Text
							End Get
						End Property

					End Class
				End Class
			End Class
		End Class
	End Class
End Namespace