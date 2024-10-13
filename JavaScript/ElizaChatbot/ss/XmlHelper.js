/*
 * Cross Browser Xml Helper Class
 * Author: Lewis Moten
 * Date: June 15, 2004 11:36 PM EST
 * Url: http://www.lewismoten.com
 *
 * Tested with MSIE 6.0, Netscape 7.1, Mozilla 1.6
 */

var XmlActiveX = new Boolean(false); // Determines if xml was created with ActiveX
var Xml = null; // Current Xml document

// Event called after xml file loads.  Searches for
// a method that YOU create called Xml_Loaded to handle
// what goes on after the xml file loads.
function LoadXml_Loaded()
{
	if(Xml_Loaded){Xml_Loaded();}
}

// load an Xml file Asynchronously
function LoadXml(fileName)
{
	if(document.implementation && document.implementation.createDocument)
	{
		XmlActiveX = false;
		Xml = document.implementation.createDocument("", "doc", null);
		Xml.onload = LoadXml_Loaded;
		Xml.load(fileName);
		return;
	}
	XmlActiveX = true;

	try{Xml = new ActiveXObject("Msxml2.DOMDocument");XmlActiveX = true;}
	catch (e){Xml = new ActiveXObject("Msxml.DOMDocument");XmlActiveX = true;}

	if(Xml == null) return;

	Xml.async = false;
	Xml.load(fileName);
	LoadXml_Loaded();
}

// select a collection of nodes that match xpath specified
function XmlSelectNodes(xpath)
{
	if(XmlActiveX) return Xml.selectNodes(xpath);
	var XPathResult = Xml.evaluate(xpath, Xml, null, 5, null);
	var Nodes = new Array();
	while(node = XPathResult.iterateNext())
		Nodes[Nodes.length] = node;
	return Nodes;
}

// Select attribute value of the node
function XmlNodeAttribute(node, name)
{
	if(XmlActiveX) return node.getAttribute(name);
	if(node.attributes)
		if(node.attributes[name])
			if(node.attributes[name].nodeValue) return node.attributes[name].nodeValue;
}
// Select text value of the node
function XmlNodeText(node)
{
	if(XmlActiveX) return node.text;
	if(node.firstChild)
		return node.firstChild.nodeValue;
	return null;
}
