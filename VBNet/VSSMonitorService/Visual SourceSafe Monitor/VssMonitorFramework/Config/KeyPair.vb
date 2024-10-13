Imports System.Xml.Serialization

' A name/value pair that can easily be serialized
' within an xml object.  Specialized collections
' such as hybrid dictionaries to not allow for such
' flexability.
Namespace Config
	Public Class KeyPair
		<XmlAttributeAttribute("name")> Public Name As String
		<XmlAttributeAttribute("value")> Public value As String
	End Class
End Namespace
