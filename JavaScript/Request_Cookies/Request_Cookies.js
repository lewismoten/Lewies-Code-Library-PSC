//**************************************
//HTML for :Request_Cookies
//**************************************
/*

<SCRIPT src="Request.js"></SCRIPT>
<SCRIPT>
document.write(Request.Cookies("Username"))
</SCRIPT>

*/
//**************************************
// Name: Request_Cookies
// Description:Gives ASP programmers a familliar interface to request cookies. document.write(Request.Cookies("UserName"))
// By: Lewis Moten
//
//
// Inputs:None
//
// Returns:None
//
//Assumes:None
//
//Side Effects:None
//This code is copyrighted and has limited warranties.
//**************************************

function Request_Cookies(Field)
{
    var cookies = new String('')
    var start = new Number(-1)
    var end = new Number(-1)
    var value = new String('')
    
    // Convert both cookies and field name to lowercase for text-comparison
    cookies = document.cookie.toLowerCase()
    Field = Field.toLowerCase()
    // start looking for field name prefixed with cookie
    // delimiter "; " and suffixed with equal sign
    start = cookies.indexOf('; ' + Field + '=')
    if(start != -1){start+= Field.length + 3}
    
    // if cookie was not found, we then look to see if it is the
    // first cookie defined (without the prefix)
    if(start == -1)
    {
        // if its not found, return an empty string
        if(cookies.indexOf(Field + '=') != 0){return(value)}
        start = Field.length + 1
    }
        
    // Look for the next cookie delimiter
    end = cookies.indexOf('; ', start)
    
    // if delimiter wasn't found, then we assume cookie is the
    // last one defined and extends to the end of the cookie string
    if(end==-1){end=cookies.length}
    
    
    // pull the cookie value out (with case differences)
    value = document.cookie.substring(start, end)
    
    // unescape the cookie value
    value = unescape(value)
    
    // return the value of the cookie
    return(value)
}
function request(){}
request.prototype.Cookies = Request_Cookies;
var Request = new request;
