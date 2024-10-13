//**************************************
//HTML for :inverse color
//**************************************
/*
<script language="javascript">
alert(inverse("#000000"));
</SCRIPT>
*/
//**************************************
// Name: inverse color
// Description:returns inverse color of hex string passed to it. If you pass white, it will pass black back to you. If you pass red, it passes cyan back.
// By: Lewis Moten
//
//
// Inputs:hex string of color alert(inverse("#000000"))
//
// Returns:returns hex string.
//
//Assumes:tested with MSIE 6.0
//
//Side Effects:None
//This code is copyrighted and has limited warranties.
//**************************************

function inverse(color)
{
	if(color.length == 7)
		color = color.substring(1, 7);
	color = "00000" + (0xffffff - parseInt(color, 16)).toString(16);
	return color.substr(color.length - 6, 6);
}