// tests to see if the page was loaded without a parent document (not as part of frames)
// and pulls page into the frameset

if(top==self) 
{ 

	// the container is your main page that has the frameset defined
	// you must include the filename of the target.  
	var sContainer = "http://www.lewismoten.com/" 			// - this is INVALID
	var sContainer = "/default.asp" 				// - ok
	var sContainer = "http://www.lewismoten.com/default.asp" 	// - ok
	var sContainer = "default.asp"	 				// - ok

	// grab the current URL that the user is trying to view
	var sThisURL = unescape(window.location.pathname);
	
	// append the url to the container
	var sGotoURL = sContainer + "?URL=" + sThisURL;
	
	// determine browser info
	var oAppVer = navigator.appVersion;
	
	var bIsNetscape = (navigator.appName == 'Netscape') && ((oAppVer.indexOf('3') != -1) || (oAppVer.indexOf('4') != -1));
	
	var bIsMicrosoftIE = (oAppVer.indexOf('MSIE 4') != -1);
	
	// determine how to redirect based on browser type and version
	if (bIsNetscape || bIsMicrosoftIE)
	{
		location.replace(sGotoURL);
	}
	else
	{
		location.href = sGotoURL;
	}

}