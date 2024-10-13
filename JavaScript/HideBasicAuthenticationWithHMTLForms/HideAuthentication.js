// hide basic authenticatin if passed within URL
if(window.location.href.indexOf('@') != -1)
{
	// parse protocal
	var StrProtocal = window.location.href.substring(0, window.location.href.indexOf('//') + 2)
	
	// parse URL
	var StrURL = window.location.href.substring(window.location.href.indexOf('@') + 1)
	
	// clean up that location
	window.location.replace(StrProtocal + StrURL);

}