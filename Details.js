
var folder = Request_QueryString("folder");

window.onload = window_Load;

function window_Load()
{
	window.focus();
	LoadXml("scripts.xml");
}
function Xml_Failed(fileName)
{
	switch(fileName)
	{
		case "scripts.xml":
			break;
		case folder + "/comments.xml":
			break;
		default:
	}
}
function Xml_Loaded(fileName)
{
	switch(fileName)
	{
		case "scripts.xml":
			setMainInfo();
			LoadXml(folder + "/comments.xml");
			break;
		case folder + "/comments.xml":
			setComments();
			break;
		default:
	}
}
function setComments()
{
	var nodes = XmlSelectNodes("//Feedback/Comments/Comment");
	
	var comments = "";
	for(var i = 0; i < nodes.length; i++)
	{
		var node = nodes[i];
		var name = XmlNodeText(node, "name");
		var date = XmlNodeText(node, "date");
		var comment = XmlNodeText(node, "comment").replace(/\n/gi, "<BR>");
		if(i != 0) comments += "<HR>";
		comments += "<P><B>Name:</B> " 
			+ (name == "" ? "none" : name)
			+ "<BR><B>Date:</B> " 
			+ date 
			+ "<BR><BR>" 
			+ comment 
			+ "</P>";
	}
	document.getElementById("lblComments").innerHTML = comments;
}
function setMainInfo()
{
	var node = XmlSelectNodes("//scripts/script[@folder='" + folder + "']");
	if(node.length == 0)
	{
		alert("Script could not be found.");
		window.location.href = "Browse.htm";
		return;
	}
	node = node[0];

	var title = XmlNodeAttribute(node, "title");
	var description = XmlNodeAttribute(node, "description");
	var date = XmlNodeAttribute(node, "date");
	var code = XmlNodeAttribute(node, "code");
	var install = XmlNodeAttribute(node, "install");
	var image = XmlNodeAttribute(node, "image");
	var language = XmlNodeAttribute(node, "language");
	var psccode = XmlNodeAttribute(node, "psccode");
	var pscworld = XmlNodeAttribute(node, "pscworld");
	var demo = XmlNodeAttribute(node, "demo");
	
	document.getElementById("lblTitle").innerHTML = title;
	document.title = title;
	
	document.getElementById("lblLinks").innerHTML = "<A href=\"http://www.planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId="
	+ psccode + "&lngWId=" + pscworld + "\"><IMG src=\"psc.jpg\" border=\"0\" width=\"100\" height=\"30\"></A>"
	
	document.getElementById("lblDate").innerHTML = date;
	document.getElementById("lblLanguage").innerHTML = language;
	document.getElementById("lblSource").innerHTML = 
		"<A href=\"" + folder + "/" + code + "\">"
		+ "<IMG src=\"download.gif\" border=\"0\" hspace=\"5\" width=\"16\" height=\"16\" align=\"absmiddle\"></A>"
		+ "<A href=\"Download.asp?FileType=code&Folder=" + escape(folder) + "\">download." + code + "</A>";

	if(install == null)
		document.getElementById("lblInstall").innerHTML = "&nbsp;";
	else
		document.getElementById("lblInstall").innerHTML = 
			"<A href=\"" + folder + "/install." + install + "\">"
			+ "<IMG src=\"download.gif\" border=\"0\" hspace=\"5\" width=\"16\" height=\"16\" align=\"absmiddle\"></A>"
			+ "<A href=\"" + folder + "/install." + install + "\">install." + install + "</A>";

//	if(demo == null)
//		lblDemo.innerHTML = "none";
//	else
//		lblDemo.innerHTML =
//		"<A href=\"" + folder + "/" + demo + "\">view demo</A>";
	
	document.getElementById("lblDescription").innerHTML = description;
	
	if(image == null)
		document.getElementById("lblScreenshot").style.display = "none";
	else
		document.getElementById("lblScreenshot").innerHTML = "<IMG src=\"" + folder + "/screenshot." + image + "\" alt=\"screenshot." + image + "\">";
}
