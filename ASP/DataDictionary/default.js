/*
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

'Program: Lewies Data Dictionary Report Generator
'Purpose: Return schema information about database tables.
'Author: Lewis Moten
'URL: http://www.lewismoten.com
'Email: lewis@moten.com

' See Readme.txt if you are having problems

'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

*/
function body_onload()
{
	// Prepopulate the form with cookie data
	document.form1.Trusted.checked = Cookie('Trusted') == 'True';
	document.form1.Server.value = Cookie('Server');
	document.form1.Catalog.value = Cookie('Catalog');
	document.form1.Password.value = Cookie('Password');
	
	// setup default values for missing data
	if(document.form1.Server.value == '')
	{
		document.form1.Server.value = '(local)'
	}
	if(document.form1.Catalog.value == '')
	{
		document.form1.Catalog.value = 'master'
	}
	if(document.form1.UserName.value == '')
	{
		document.form1.UserName.value = 'sa'
	}
	
}
function Cookie(name) 
{
	var cookies = document.cookie;
	var prefix = name + '=';
	var begin = cookies.indexOf("; " + prefix);
	var end;
	
	if(begin == -1)
	{
		begin = cookies.indexOf(prefix);
		if(begin != 0)
		{
			return('');
		}
	} 
	else
	{
		begin += 2;
    }
	end = cookies.indexOf(";", begin);
	if(end==-1)
	{
		end = cookies.length;
    }
	return unescape(cookies.substring(begin + prefix.length, end));
}