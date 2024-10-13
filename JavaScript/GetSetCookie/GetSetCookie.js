function SetCookie(name, value)
{
	var expires = new Date();
	expires.setTime(expires.getTime() + 0x9A7EC800);
	document.cookie = name + "=" + escape(value) + "; expires=" + expires.toUTCString();
}
function GetCookie(name)
{
	var i = 0;
	while(document.cookie.indexOf(name + "=", i) != -1)
	{
		var value = document.cookie.substr(document.cookie.indexOf(name + "=", i));
		value = value.substr(name.length + 1);
		if(value.indexOf(";") != -1)
			if(value.indexOf(";") == 0)
				value = "";
			else
				value = value.substr(0, value.indexOf(";"));
		return unescape(value);
	}
	return "";
}