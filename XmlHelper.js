/*
 * Cross Browser Xml Helper Class
 * Author: Lewis Moten
 * Date: June 15, 2004 11:36 PM EST
 * Url: http://www.lewismoten.com
 *
 * Tested with MSIE 6.0, Netscape 7.1, Mozilla 1.6
 */
/*
 September 7, 2011 - added support for MSIE 8, MSIE 9, and Google Chrome
*/
var XmlDebug = new Boolean(false); // set to true if you are debugging.
var XmlActiveX = new Boolean(false); // Determines if xml was created with ActiveX
var Xml = null; // Current Xml document
var XmlFileName = null;

// Event called after xml file loads.  Searches for
// a method that YOU create called Xml_Loaded to handle
// what goes on after the xml file loads.
function LoadXml_Loaded(e)
{
	if(Xml){}
	else
	{
		LoadXml_Failed();
		return;
	}
	
	if(e)
		if(Xml.documentElement){}
		else
		{
			LoadXml_Failed();
			return;
		}
	// xml did load properly, notify delegate
	if(Xml_Loaded){Xml_Loaded(XmlFileName);}
}
function LoadXml_Failed()
{
	if(Xml_Failed){Xml_Failed(XmlFileName);}
}

// load an Xml file Asynchronously
function LoadXml(fileName)
{
	Xml = null;
	XmlFileName = fileName;
	if(document.implementation && document.implementation.createDocument)
	{
		XmlActiveX = false;
		Xml = document.implementation.createDocument("", "doc", null);
		Xml.onload = LoadXml_Loaded;
		try
		{
			Xml.load(fileName);
		}
catch (e) {
        // try chrome, webkit, firefox, mozilla, opera, etc.
		    try {
		        var XmlHttp = new window.XMLHttpRequest();
		        XmlHttp.open("GET", fileName, false);
		        XmlHttp.send(null);
		        Xml = XmlHttp.responseXML;//.documentElement;
		        LoadXml_Loaded();
		    }
		    catch (e2) {
		        alert(e2.message);
		        LoadXml_Failed();
            }
		}
		return;
	}
	XmlActiveX = true;
	try{Xml = new ActiveXObject("Msxml2.DOMDocument");XmlActiveX = true;}
	catch (e){Xml = new ActiveXObject("Msxml.DOMDocument");XmlActiveX = true;}

	if(Xml == null) return;
	try
	{
		Xml.setProperty("SelectionLanguage", "XPath");
		Xml.async = false;// setting to true has problems on different systems.
		Xml.onreadystatechange = Xml_ReadyStateChange;
		
		Xml.setProperty("ServerHTTPRequest", true);
		
		if(!Xml.load(fileName))
		{
			if(XmlDebug == true) 
				alert("LoadXml: " + Xml.parseError.errorCode + ": " + Xml.parseError.reason + "\r\n" + Xml.parseError.url);
			LoadXml_Failed();
		}
	}
	catch(e)
	{
		if(XmlDebug == true) 
			alert("LoadXml: " + e.message);
		LoadXml_Failed();
	}
}
function Xml_ReadyStateChange()
{
	if(Xml.readyState == 4)
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
// Select text value of the node (or child node - name is optional)
function XmlNodeText(node, name)
{
	if(name != null) // if name provided, then we search for a child element with that name
	{
		for(var i = 0; i < node.childNodes.length; i++)
		{
			if(node.childNodes[i].nodeName == name)
				if(XmlActiveX) return node.childNodes[i].text;
				else if(node.childNodes[i].firstChild) return node.childNodes[i].firstChild.nodeValue;
				else return null;
		}
		return null;
	}
	if(XmlActiveX) return node.text;
	if(node.firstChild)
		return node.firstChild.nodeValue;
	return null;
}
