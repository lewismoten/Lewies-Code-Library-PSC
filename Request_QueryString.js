//**************************************
//HTML for :Request_QueryString
//**************************************
/*
<FORM>
<INPUT name="Query">
<INPUT type="Submit">
</FORM>
*/
//**************************************
// Name: Request_QueryString
// Description:Allows browser to read querystring variables through javascript.
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

document.write(Request_QueryString('Query'));
function Request_QueryString(FieldName)
{
    var QueryString = ''
    var FieldValue = ''
    var Start = 0
    var End = 0
    
    // Grab the querystring
    QueryString = window.location.search
    
    // Convert field name and querystring to lowercase so that
    // function is not case sensitive.
    
    FieldName = FieldName.toLowerCase()
    QueryString = QueryString.toLowerCase()
    
    // Look for field as first item ...
    Start = QueryString.indexOf(FieldName + '=')
    
    // If field is not the first ...
    if(Start!=1)
    {
    
        // Search appended fields
        Start = QueryString.indexOf('&' + FieldName + '=')
        
        // If field wasn't found at all, return empty string.
        if(Start==-1){return(FieldValue)}
        // Setup start position after equal sign
        Start += FieldName.length + 2
    }
    else
    {
        // Setup start position after equal sign
        Start = FieldName.length + 2
    }
    // Search for beginning of next field
    End = QueryString.indexOf('&', Start + 1)
    
    // if another field was not defined, set end to length of querystring
    if(End==-1){End=QueryString.length}
    // Parse the field value
    FieldValue = window.location.search.substring(Start, End)
    // unescape special characters within the value (such as %20 = space character)
    FieldValue = unescape(FieldValue)
    
    // Return the results
    return(FieldValue)
}