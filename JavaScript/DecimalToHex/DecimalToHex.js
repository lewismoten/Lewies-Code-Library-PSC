//**************************************
//HTML for :Decimal to Hex
//**************************************

/*
DEMO:

	<SCRIPT>
	alert(Hex(255))
	</SCRIPT>

*/

//**************************************
// Name: Decimal to Hex
// Description:Converts a decimal number to a Hex String (i.e. 255 = 'FF')
// By: Lewis Moten
//
//
// Inputs:pLngDecimal - the decimal to convert to hex
//
// Returns:hex value of pLngDecimal
//
//Assumes:None
//
//Side Effects:None
//This code is copyrighted and has limited warranties.
//**************************************

function Hex(pLngDecimal)


    {
    var digits = new Array('0', '1', '2', '3', '4', '5', '6', '7', '8', '9', 'A', 'B', 'C', 'D', 'E', 'F')
    if(pLngDecimal < digits.length)


        {
        return(digits[pLngDecimal])
    }
    var prefix = '' + Math.floor(pLngDecimal / digits.length)
    var suffix = pLngDecimal - prefix * digits.length
    if (prefix > digits.length)


        {
        return(Hex(prefix) + digits[suffix])
    }
    return(digits[prefix] + digits[suffix])
}