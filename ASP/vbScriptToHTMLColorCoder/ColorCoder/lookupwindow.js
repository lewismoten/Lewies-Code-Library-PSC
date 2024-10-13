// creates a pop-up window that will focus each time this function is called
var lookUpWindow
		
function lookUp(asKeyword)
{
	var lsURL = 'http://www.lewismoten.com/ref/vbscript/' + asKeyword + '.asp'
	if ((!lookUpWindow) || (lookUpWindow.closed))
	{
		lookUpWindow = window.open(lsURL, 'lookUpWindow', 'width=380,height=200,toolbar=no,menubar=no,resizeable=yes,top=100,left=200,screenX=200,screenY=100');
		lookUpWindow.opener = window;
		return true;
	}
	else
	{
		lookUpWindow.location = lsURL;
		lookUpWindow.focus();
		return true;
	}
}
