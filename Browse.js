var Xml = null;
var Page = new Number(1);
var MaxPage = new Number(1);
var Entries = new Number(1);
var PageSize = new Number(5);
var XmlActiveX = new Boolean(false);
var SearchQuery = new String("");
var SearchLanguage = new String("");
var Nodes = new Object();

function lstPage_Change()
{
	ShowPage(document.getElementById("lstPage").selectedIndex + 1);
}
function btnBack_Click()
{
	ShowPage(Page - 1);
}
function btnNext_Click()
{
	ShowPage(Page + 1);
}
function ShowPage(pageNumber)
{
	if(pageNumber <= 0) pageNumber = 1;
	if(pageNumber > MaxPage) pageNumber = MaxPage;
	
	Page = pageNumber;
	
	document.getElementById("lstPage").selectedIndex = Page - 1;
	document.getElementById("btnBack").disabled = Page <= 1;
	document.getElementById("btnNext").disabled = Page >= MaxPage;
	
	var start = (Page - 1) * PageSize;
	var end = start + PageSize - 1;
	//if(end >= Entries) end = Entries - 1;
	
	for(var i = start; i <= end; i++)
		SetRow((i - start) + 1, i);
}
function HtmlEncode(text)
{
	var html = new String(text);
	html = html.replace(/&/g, "&amp;");
	html = html.replace(/</g, "&lt;");
	html = html.replace(/>/g, "&gt;");
	html = html.replace(/"/g, "&quot;");//"
	return html;
}
function SetRow(rowNumber, entryIndex)
{
	var title = "";
	var description = "";
	var date = "";
	var code = "";
	var image = "";
	var install = "";
	var cell1 = "&nbsp;";
	var cell2 = "&nbsp;";
	var cell3 = "&nbsp;";
	var language = "";
	var folder = "";
	var psccode = "";
	var pscworld = "";


	if(entryIndex > Entries - 1 || entryIndex < 0)
	{
		// do nothing
	}
	else
	{
		var node = Nodes[entryIndex];
		title = XmlNodeAttribute(node, "title");
		description = XmlNodeAttribute(node, "description") + "<BR>";
		date = XmlNodeAttribute(node, "date");
		code = XmlNodeAttribute(node, "code");
		image = XmlNodeAttribute(node, "image");
		install = XmlNodeAttribute(node, "install");
		language = XmlNodeAttribute(node, "language");
		folder = XmlNodeAttribute(node, "folder");
		psccode = XmlNodeAttribute(node, "psccode");
		pscworld = XmlNodeAttribute(node, "pscworld");
		
		title = "<B><A href=\"Details.htm?folder=" + escape(folder) + "\">" + title + "</A></B><BR>";
		//description += "<BR>[<A href=\"" + file + "\">Download</A>]";
		if(image != null)
		{
			description += "<A target=\"screenshot\" href=\"" + folder + "/screenshot." + image + "\">";
			description += "<IMG src=\"screenshot.gif\" border=\"0\" hspace=\"5\" width=\"16\" height=\"16\" align=\"absmiddle\"></A>";
			description += "[<A target=\"screenshot\" href=\"" + folder + "/screenshot." + image + "\">Screenshot</A>]";
		}
		if(code != null)
		{
			description += "<A href=\"download.asp?FileType=code&Folder=" + escape(folder) + "\">";
			description += "<IMG src=\"download.gif\" border=\"0\" hspace=\"5\" width=\"16\" height=\"16\" align=\"absmiddle\"></A>";
			description += "[<A href=\"download.asp?FileType=code&Folder=" + escape(folder) + "\">Source Code</A>]";
		}
		if(install != null)
		{
			description += "<A href=\"" + folder + "/install." + code + "\">";
			description += "<IMG src=\"download.gif\" border=\"0\" hspace=\"5\" width=\"16\" height=\"16\" align=\"absmiddle\"></A>";
			description += "[<A href=\"" + folder + "/install." + code + "\">Install</A>]";
		}
		//if(psccode != null && pscworld != null)
		//{
			//var psc = "http://www.pscode.com/vb/scripts/ShowCode.asp?txtCodeId=" + psccode + "&amp;lngWId=" + pscworld;
			var psc = "Details.htm?folder=" + escape(folder);
			description += "<A href=\"" + psc + "\">";
			description += "<IMG src=\"information.gif\" border=\"0\" hspace=\"5\" width=\"16\" height=\"16\" align=\"absmiddle\"></A>";
			description += "[<A href=\"" + psc + "\">Information</A>]";
			
		//}
		
	}
	var tr = document.getElementById("Row" + rowNumber + "a");
	tr.childNodes[0].innerHTML = title;
	tr.childNodes[1].innerHTML = date;
	tr.childNodes[2].innerHTML = language;
	
	tr = document.getElementById("Row" + rowNumber + "b");
	tr.childNodes[0].innerHTML = description;
	
}
function Xml_Loaded(fileName)
{
	var lstLanguage = document.getElementById("lstLanguage");
	while(lstLanguage.options.length != 0) lstLanguage.remove(0);
		
	var option = document.createElement("OPTION") 
	option.text = "Any";
	lstLanguage.options.add(option)
	
	var languages = XmlSelectNodes("//scripts/script/@language");
	for(var i = 0; i < languages.length; i++)
		if(LanguageExists(languages[i].value) == false)
		{
			option = document.createElement("OPTION") 
			option.text = languages[i].value;
			lstLanguage.options.add(option)
		}

	btnSearch_Click();
}
function LanguageExists(language)
{
	var lstLanguage = document.getElementById("lstLanguage");
	for(var i = 1; i < lstLanguage.options.length; i++)
		if(lstLanguage[i].text == language)
			return true;
	return false;
}
function window_Load()
{
	window.focus();
	document.getElementById("btnSearch").onclick = btnSearch_Click;
	LoadXml("scripts.xml");
}

function btnSearch_Click()
{
	SearchQuery = document.getElementById("txtQuery").value;
	SearchLanguage = document.getElementById("lstLanguage")[document.getElementById("lstLanguage").selectedIndex].text;
	
	if(SearchLanguage == "Any") SearchLanguage = "";
	
	document.getElementById("imgLog").src = "Browse.asp?query=" + escape(SearchQuery) + "&language=" + escape(SearchLanguage);

	if(SearchQuery == "" && SearchLanguage == "")
		Nodes = XmlSelectNodes("//scripts/script");
	else if(SearchQuery == "")
		Nodes = XmlSelectNodes("//scripts/script[@language='" + SearchLanguage + "']");
	else if(SearchLanguage == "")
	{
		var XPath = new String("");
		XPath += "//scripts/script[contains(translate(@title, 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'), '" + SearchQuery.toUpperCase() + "')]";
		XPath += "| //scripts/script[contains(translate(@description, 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'), '" + SearchQuery.toUpperCase() + "')]";
		XPath += "| //scripts/script[contains(translate(@date, 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'), '" + SearchQuery.toUpperCase() + "')]";
		Nodes = XmlSelectNodes(XPath);
	}
	else
	{
		var XPath = new String("");
		XPath += "//scripts/script[contains(translate(@title, 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'), '" + SearchQuery.toUpperCase() + "') and @language = '" + SearchLanguage + "']";
		XPath += "| //scripts/script[contains(translate(@description, 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'), '" + SearchQuery.toUpperCase() + "') and @language = '" + SearchLanguage + "']";
		XPath += "| //scripts/script[contains(translate(@date, 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'), '" + SearchQuery.toUpperCase() + "') and @language = '" + SearchLanguage + "']";
		Nodes = XmlSelectNodes(XPath);
	}
		
	Entries = Nodes.length;
	
	document.getElementById("lblMatchingResults").innerHTML = Entries;
	MaxPage = Math.floor(Entries / PageSize);
	if(Entries % PageSize != 0) MaxPage++;
	var lstPage = document.getElementById("lstPage");
	while(lstPage.options.length < MaxPage)
	{
		var option = document.createElement("OPTION") 
		option.text = "Page " + (lstPage.options.length + 1);
		lstPage.options.add(option)
	}
	while(lstPage.options.length > MaxPage)
		lstPage.remove(lstPage.options.length -1);
	ShowPage(1);
}

window.onload = window_Load;
