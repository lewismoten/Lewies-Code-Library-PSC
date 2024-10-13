var heard = new String("");
var posessions = new Array();
var patience = null;
var AIName = "Eliza";
/*
 * -----------------------------------------
 * Event Handlers
 * -----------------------------------------
*/
function window_load()
{
	LoadXml("eliza.xml");
}
function Xml_Loaded()
{
	appendToLog(AIName + "> " + ELIZA_GetEvent("introduction"));

	document.getElementById("Message").focus();
	document.getElementById("btnSend").onclick = btnSend_Click;
	document.getElementById("Message").onkeyup = Message_KeyUp;
	patience = window.setTimeout("ELIZA_Impatient();", 1000 * 60);
}
function Message_KeyUp()
{
	if(event.keyCode == 13)
	{
		event.keyCode = null;
		btnSend_Click();
	}
}
function btnSend_Click()
{
	window.clearTimeout(patience);
	
	var hear = new String(document.getElementById("Message").value);
	document.getElementById("Message").value = "";
	appendToLog("> " + hear);
	
	if(hear.toLowerCase().indexOf("meet ") != -1)
	{
		var newbot = hear.substr(5, hear.length);
		ELIZA_Introduce(newbot);
		return;
	}
	hear = hear.replace(/[^\w'\s]/gi, " ");
	hear = hear.replace(/\s+/gi, " ");
	hear = hear.replace(/\s*(\S+([\s\S]*\S)?)\s*/gi, "$1");
	if(hear == heard)
		appendToLog(AIName + "> " + ELIZA_GetEvent("repetitive"));
	else if(hear == "")
		appendToLog(AIName + "> " + ELIZA_GetEvent("silent"));
	else	
	{
		appendToLog(AIName + "> " + ELIZA_getResponses(hear));
		heard = hear;
	}
	document.getElementById("Message").focus();
	patience = window.setTimeout("ELIZA_Impatient();", 1000 * 60);
}
/*
 * -----------------------------------------
 * A.I.
 * -----------------------------------------
*/
function ELIZA_Impatient()
{
	appendToLog(AIName + "> " + ELIZA_GetEvent("impatient"));
	patience = window.setTimeout("ELIZA_Impatient();", 1000 * 60 * 15);
}
function ELIZA_Conjugate(message)
{
	var conjugation = message;
	var Conjugations = XmlSelectNodes("//ai/conjugations/conjugation");
	var searchIndex = 0;
	var regexp = null;
	var pattern = null;
	while(true)
	{
		var matchIndex = -1;
		var matchPosition = Number.MAX_VALUE;
		var start = conjugation.substr(0, searchIndex);
		var end = conjugation.substr(searchIndex, conjugation.length);
		for(var i = 0; i < Conjugations.length; i++)
		{
			pattern = XmlNodeAttribute(Conjugations[i], "in");
			regexp = new RegExp("\\b" + pattern + "\\b", "i");
			
			if(end.search(regexp) != -1)
			{
				if(end.search(regexp) < matchPosition)
				{
					matchPosition = end.search(regexp);
					matchIndex = i;
				}
			}
		}
		if(matchIndex == -1) return conjugation;
		pattern = XmlNodeAttribute(Conjugations[matchIndex], "in");
		var replacement = XmlNodeAttribute(Conjugations[matchIndex], "out");
		regexp = new RegExp("\\b" + pattern + "\\b", "i");
		searchIndex = start.length + end.search(regexp) + replacement.length + 1;
		conjugation = start + end.replace(regexp, replacement);
	}
}
function ELIZA_GetEvent(name)
{
	var response = "Oops!  I do not know how to handle the \"" + name + "\" event.";
	try
	{
		var nodes = XmlSelectNodes("//ai/events/event[@name='" + name + "']/say");
	}
	catch(e){return response + e.message;}
	if(nodes == null) return response + '[null]';
	if(nodes.length == 0) return response + '[length=0]';
	return XmlNodeText(nodes[randomNumber(0, nodes.length -1)]);
}
function ELIZA_getResponses(message)
{
	var keywords = XmlSelectNodes("//ai/keywords/keyword");
	var Responses = new Array();
	var r = null;
	var Response = new String("");

	r = new RegExp("\\bmy\\b", "i");
	if(message.search(r) != -1)
	{
		var posession = message.substr(message.search(r)+3, message.length);
		posessions[posessions.length] = ELIZA_Conjugate(posession);
	}
	for(var i = 0; i < keywords.length; i++)
	{
		var keyword = XmlNodeAttribute(keywords[i], "value");
		r = new RegExp("\\b" + keyword + "\\b", "i");
		if(message.search(r) != -1)
		{
			var suffixPosition = message.search(r) + keyword.length;
			
			ELIZA_AppendKeywordResponses(keyword, Responses, suffixPosition);
		}
	}
	
	if(Responses.length == 0)
	{
		if(posessions.length != 0 && randomNumber(1, 1000) % 5 == 2)
		{
			Response = ELIZA_GetEvent("remember");
			if(Response.indexOf("*") != -1)
			{
				posession = posessions[randomNumber(1, posessions.length) -1];
				Response = Response.replace(/\*/gi, posession)
			}
			return Response;
		}
		return ELIZA_GetEvent("confused");
	} 
	var ResponseIndex = randomNumber(1, Responses.length) - 1
	Response = Responses[ResponseIndex][0];
	if(Response.indexOf("*") != -1)
	{
		var conjugation = ELIZA_Conjugate(message.substr(Responses[ResponseIndex][1], message.length));
		Response = Response.replace(/\*/gi, conjugation)
	}
	return Response;
}
// find available responses
function ELIZA_AppendKeywordResponses(keyword, responses, suffixPosition)
{
	if(responses.length > 20) return;
	var NewResponses = XmlSelectNodes("//ai/keywords/keyword[@value=\"" + keyword + "\"]/response");
	for(var i = 0; i < NewResponses.length; i++)
	{
		var n = responses.length;
		responses[n] = new Array();
		responses[n][0] = XmlNodeText(NewResponses[i]);
		responses[n][1] = suffixPosition;
	}
	var SeeKeyword = XmlSelectNodes("//ai/keywords/keyword[@value=\"" + keyword + "\"]");
	if(SeeKeyword != null && SeeKeyword.length != 0)
		if(XmlNodeAttribute(SeeKeyword[0], "see")!= null)
			ELIZA_AppendKeywordResponses(XmlNodeAttribute(SeeKeyword[0], "see"), responses, suffixPosition);
}
/*
 * -----------------------------------------
 * Helpers
 * -----------------------------------------
*/
function randomNumber(minimum, maximum)
{
	var Result = Math.random() * (maximum - minimum)
	return(Math.round(Result) + minimum)
}
function appendToLog(message)
{
	var chatlog = document.getElementById("ChatLog");
	chatlog.innerHTML += HtmlEncode(message) + "<BR>";
	chatlog.scrollTop = chatlog.scrollHeight;
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

window.onload = window_load;
