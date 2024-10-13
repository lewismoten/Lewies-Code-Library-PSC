/*
-----------------------------------------------------
Author: Lewis E. Moten III
Date: May, 16, 2004
Homepage: http://www.lewismoten.com
Email: lewis@moten.com
-----------------------------------------------------
*/

function window_load()
{
	try
	{
		//self.window.dialogHeight = (323 + 40) + "px";
		self.window.dialogHeight = (303 + 40) + "px";
		self.window.dialogWidth = (488 + 6) + "px";

		hex = window.dialogArguments;
		hex = hex.toUpperCase();
		if(hex.length == 7) hex = hex.substr(1,6);
		frmColorPicker.txtHex.value = hex;
		
		pnlOldColor.style.backgroundColor = "#" + hex;
		
		Hex_Changed();
	}
	catch(e)
	{}
	

	frmColorPicker.btnCancel.onclick = btnCancel_Click;
	frmColorPicker.btnOK.onclick = btnOK_Click;

	// Gradient Top
	pnlGradient_Top.onclick = pnlGradient_Top_Click;
	pnlGradient_Top.onmousemove = pnlGradient_Top_MouseMove;
	pnlGradient_Top.onmousedown = pnlGradient_Top_MouseDown;
	pnlGradient_Top.onmouseup = pnlGradient_Top_MouseUp;

	// Vertical Top
	pnlVertical_Top.onclick = pnlVertical_Top_Click;
	pnlVertical_Top.onmousemove = pnlVertical_Top_MouseMove;
	pnlVertical_Top.onmousedown = pnlVertical_Top_MouseDown;
	pnlVertical_Top.onmouseup = pnlVertical_Top_MouseUp;
	
	frmColorPicker.txtHSB_Hue.onkeyup = Hsb_Changed;
	frmColorPicker.txtHSB_Brightness.onkeyup = Hsb_Changed;
	frmColorPicker.txtHSB_Saturation.onkeyup = Hsb_Changed;
	
	frmColorPicker.txtRGB_Red.onkeyup = Rgb_Changed;
	frmColorPicker.txtRGB_Green.onkeyup = Rgb_Changed;
	frmColorPicker.txtRGB_Blue.onkeyup = Rgb_Changed;
	
	frmColorPicker.txtHex.onkeyup = Hex_Changed;
	
	frmColorPicker.txtHSB_Hue.onkeypress = validateNumber;
	frmColorPicker.txtHSB_Saturation.onkeypress = validateNumber;
	frmColorPicker.txtHSB_Brightness.onkeypress = validateNumber;

	frmColorPicker.txtRGB_Red.onkeypress = validateNumber;
	frmColorPicker.txtRGB_Green.onkeypress = validateNumber;
	frmColorPicker.txtRGB_Blue.onkeypress = validateNumber;
	
	frmColorPicker.txtHex.onkeypress = validateHex;

	frmColorPicker.btnWebSafeColor.onclick = btnWebSafeColor_Click;
	pnlWebSafeColor.onclick = btnWebSafeColor_Click;
	pnlWebSafeColorBorder.onclick = btnWebSafeColor_Click;
	
}
function btnWebSafeColor_Click()
{
	var hex = new String(frmColorPicker.txtHex.value);
	var rgb = HexToRgb(hex);
	rgb = RgbToWebSafeRgb(rgb);
	frmColorPicker.txtHex.value = RgbToHex(rgb);
	Hex_Changed();
}
function RgbIsWebSafe(rgb)
{
	var hex = RgbToHex(rgb);
	var isSafe = new Boolean(true);
	for(var i = 0; i < 3; i++)
		if("00336699CCFF".indexOf(hex.substr(i*2, 2)) == -1)
		{
			isSafe = false;
			break;
		}

	return isSafe;
}
function RgbToWebSafeRgb(rgb)
{
	var safeRgb = new Array(rgb[0], rgb[1], rgb[2]);

	if(RgbIsWebSafe(rgb)) return new Array(rgb[0], rgb[1], rgb[2]);
	
	var safeValue = new Array(0x00, 0x33, 0x66, 0x99, 0xCC, 0xFF);	
	
	for(var i = 0; i < safeRgb.length; i++)
		for(var j = 1; j < safeValue.length; j++)
			if(safeRgb[i] > safeValue[j-1] && safeRgb[i] < safeValue[j])
			{
				if(safeRgb[i] - safeValue[j-1] > safeValue[j] - safeRgb[i]) 
					safeRgb[i] = safeValue[j];
				else
					safeRgb[i] = safeValue[j-1];
				break;
			}

	return safeRgb;
}
function checkWebSafe()
{
	var hex = new String(frmColorPicker.txtHex.value);
	var rgb = HexToRgb(hex);
	
	if(RgbIsWebSafe(rgb))
	{
		frmColorPicker.btnWebSafeColor.style.display = "none";
		pnlWebSafeColor.style.display = "none";
		pnlWebSafeColorBorder.style.display = "none";
	}
	else
	{
		rgb = RgbToWebSafeRgb(rgb);
		pnlWebSafeColor.style.backgroundColor = RgbToHex(rgb);
		frmColorPicker.btnWebSafeColor.style.display = "";
		pnlWebSafeColor.style.display = "";
		pnlWebSafeColorBorder.style.display = "";
	}

}

function validateNumber()
{

	var key = String.fromCharCode(event.keyCode);

	var keys = new Array(0, 8, 9, 13, 27);

	if(key == null) return;

	for(var i = 0; i < keys.length; i++)
		if(event.keyCode == keys[i]) return;
	
	if("01234567879".indexOf(key) != -1) return;

	event.keyCode = 0;	
}
function validateHex()
{

	var key = String.fromCharCode(event.keyCode);

	var keys = new Array(0, 8, 9, 13, 27);

	if(key == null) return;

	for(var i = 0; i < keys.length; i++)
		if(event.keyCode == keys[i]) return;
	
	if("abcdef".indexOf(key) != -1)
	{
		event.keyCode = key.toUpperCase().charCodeAt(0);
		return;
	}

	if("01234567879ABCDEF".indexOf(key) != -1) return;

	event.keyCode = 0;	
}
function RgbToXyz(rgb)
{
	var sRgb = new Array(rgb[0], rgb[1], rgb[2]);
	
	for(var i = 0; i < sRgb.length; i++)
	{
		// linier
		sRgb[i] = rgb[i] / 255;

		// gamma curve
		if(sRgb[i] <= .04045) // .03928
			sRgb[i] = sRgb[i] / 12.92;
		else
			sRgb[i] = Math.pow((sRgb[i] + .055) / 1.055, 2.4);

		sRgb[i] *= 100;
	}

	var xyz = new Array(0, 0, 0);

	var matrix = new Array(
		new Array(.4124, .3576, .1805),
		new Array(.2126, .7152, .0722),
		new Array(.0193, .1192, .9505)
		)

	for(var i = 0; i < xyz.length; i++)
		for(var n = 0; n < sRgb.length; n++)
			xyz[i] += sRgb[n] * matrix[i][n];

	return xyz;
}
function RgbToCmy(rgb)
{
	var cmy = new Array(0, 0, 0);
	for(var i = 0; i < cmy.length; i++)
		cmy[i] = 255 - rgb[i];
	return cmy;
}
function CmyToCmyk(cmy)
{
	var cmyk = new Array(0, 0, 0, 255);
	
	for(var i = 0; i < cmy.length; i++)
		if(cmy[i] < cmyk[3]) cmyk[3] = cmy[i];
		
	for(var i = 0; i < cmy.length; i++)
		cmyk[i] = (cmy[i] - cmyk[3]);

	for(var i = 0; i < cmyk.length; i++)
		cmyk[i] = cmyk[i] / 255;

	return cmyk;
}
function RgbToCmyk(rgb)
{
	return CmyToCmyk(RgbToCmy(rgb));
}
function XyzToHunterLab(xyz)
{
	var Lab = new Array(0, 0, 0);
	Lab[0] = 10 * Math.sqrt(xyz[1]);
	Lab[1] = 17.5 * (((1.02 * xyz[0]) - xyz[1]) / Math.sqrt(xyz[1]));
	Lab[2] = 7 * ((xyz[1] - (.847 * xyz[2])) / Math.sqrt(xyz[1]));
	return Lab;
}
function RgbToLab(rgb)
{
	return XyzToHunterLab(RgbToXyz(rgb));
}

function btnCancel_Click()
{
	window.close();
}
function btnOK_Click()
{
	try
	{
		window.returnValue = new String(frmColorPicker.txtHex.value);
	}
	catch(e)
	{
	}
	window.close();
}
function SetGradientPosition(x, y)
{
	
	pnlGradientPosition.style.left = x + "px";
	pnlGradientPosition.style.top = y + "px";
	
	// get true x/y
	x -= 7;
	y -= 28;
	
	var saturation = (x / 255) * 100;
	var brightness = 100 - ((y / 255) * 100);
	var hue = new Number(frmColorPicker.txtHSB_Hue.value);
	
	frmColorPicker.txtHSB_Saturation.value = Math.round(saturation);
	frmColorPicker.txtHSB_Brightness.value = Math.round(brightness);
	
	Hsb_Changed();
}
function Hex_Changed()
{
	var hex = new String(frmColorPicker.txtHex.value);
	for(var i = 0; i < hex.length; i++)
		if("0123456789ABCDEFabcdef".indexOf(hex.substr(i, 1)) == -1)
		{
			hex = "000000";
			rgb = HexToRgb(hex);
			frmColorPicker.txtHex.value = "000000";
			alert("An integer between 000000 and FFFFFF is required. Closest\r\nvalue inserted.");
		}
		
	while(hex.length < 6)
	{
		hex = "0" + hex;
	}
	
	var rgb = HexToRgb(hex);
	var hsb = RgbToHsb(rgb);
	var Lab = RgbToLab(rgb);
	var cmyk = RgbToCmyk(rgb);
	
	frmColorPicker.txtRGB_Red.value = rgb[0];
	frmColorPicker.txtRGB_Green.value = rgb[1];
	frmColorPicker.txtRGB_Blue.value = rgb[2];

	frmColorPicker.txtHSB_Hue.value = Math.round(hsb[0] * 360);
	frmColorPicker.txtHSB_Saturation.value = Math.round(hsb[1] * 100);
	frmColorPicker.txtHSB_Brightness.value = Math.round(hsb[2] * 100);

	frmColorPicker.txtLab_Luminancy.value = Math.round(Lab[0]);
	frmColorPicker.txtLab_A.value = Math.round(Lab[1]);
	frmColorPicker.txtLab_B.value = Math.round(Lab[2]);
	
	frmColorPicker.txtCMYK_Cyan.value = Math.round(cmyk[0] * 100);
	frmColorPicker.txtCMYK_Magenta.value = Math.round(cmyk[1] * 100);
	frmColorPicker.txtCMYK_Yellow.value = Math.round(cmyk[2] * 100);
	frmColorPicker.txtCMYK_Black.value = Math.round(cmyk[3] * 100);

	pnlNewColor.style.backgroundColor = "#" + RgbToHex(rgb);

	if(Lab[0] >= 54)
		SetGradientPositionDark();
	else
		SetGradientPositionLight();
	
	pnlGradientPosition.style.left = Math.round((hsb[1] * 255) + 7) + "px";
	pnlGradientPosition.style.top = Math.round((255 - (hsb[2] * 255)) + 27) + "px";
	
	if(Math.round(hsb[0] * 360) == 360 || Math.round(hsb[0] * 360) == 0)
	{
		pnlVerticalPosition.style.top = "27px";
		pnlGradient_Hue.style.backgroundColor = "#FF0000";
	}
	else
	{
		pnlVerticalPosition.style.top = Math.round((255 - (hsb[0] * 255)) + 27) + "px";
		pnlGradient_Hue.style.backgroundColor = "#" + RgbToHex(HueToRgb(hsb[0]));
	}
	
	checkWebSafe();
}
function Rgb_Changed()
{
	var rgb = new Array(0, 0, 0);
	rgb[0] = new Number(frmColorPicker.txtRGB_Red.value);
	rgb[1] = new Number(frmColorPicker.txtRGB_Green.value);
	rgb[2] = new Number(frmColorPicker.txtRGB_Blue.value);
	
	if(rgb[0] > 255)
	{
		rgb[0] = 255;
		frmColorPicker.txtRGB_Red.value = 255;
		alert("An integer between 0 and 255 is required. Closest\r\nvalue inserted.");
	}
	if(rgb[1] > 255)
	{
		rgb[1] = 255;
		frmColorPicker.txtRGB_Green.value = 255;
		alert("An integer between 0 and 255 is required. Closest\r\nvalue inserted.");
	}
	if(rgb[2] > 255)
	{
		rgb[2] = 255;
		frmColorPicker.txtRGB_Blue.value = 255;
		alert("An integer between 0 and 255 is required. Closest\r\nvalue inserted.");
	}
	
	var hsb = RgbToHsb(rgb);
	var Lab = RgbToLab(rgb);
	var cmyk = RgbToCmyk(rgb);
	
	frmColorPicker.txtHSB_Hue.value = Math.round(hsb[0] * 360);
	frmColorPicker.txtHSB_Saturation.value = Math.round(hsb[1] * 100);
	frmColorPicker.txtHSB_Brightness.value = Math.round(hsb[2] * 100);

	frmColorPicker.txtLab_Luminancy.value = Math.round(Lab[0]);
	frmColorPicker.txtLab_A.value = Math.round(Lab[1]);
	frmColorPicker.txtLab_B.value = Math.round(Lab[2]);
	
	frmColorPicker.txtCMYK_Cyan.value = Math.round(cmyk[0] * 100);
	frmColorPicker.txtCMYK_Magenta.value = Math.round(cmyk[1] * 100);
	frmColorPicker.txtCMYK_Yellow.value = Math.round(cmyk[2] * 100);
	frmColorPicker.txtCMYK_Black.value = Math.round(cmyk[3] * 100);

	frmColorPicker.txtHex.value = RgbToHex(rgb);
	pnlNewColor.style.backgroundColor = "#" + RgbToHex(rgb);
	

	if(Lab[0] >= 54)
		SetGradientPositionDark();
	else
		SetGradientPositionLight();
	
	pnlGradientPosition.style.left = Math.round((hsb[1] * 255) + 7) + "px";
	pnlGradientPosition.style.top = Math.round((255 - (hsb[2] * 255)) + 27) + "px";
	
	if(Math.round(hsb[0] * 360) == 360 || Math.round(hsb[0] * 360) == 0)
	{
		pnlVerticalPosition.style.top = "27px";
		pnlGradient_Hue.style.backgroundColor = "#FF0000";
	}
	else
	{
		pnlVerticalPosition.style.top = Math.round((255 - (hsb[0] * 255)) + 27) + "px";
		pnlGradient_Hue.style.backgroundColor = "#" + RgbToHex(HueToRgb(hsb[0]));
	}
	
	checkWebSafe();
}
function Hsb_Changed()
{
	var saturation = new Number(frmColorPicker.txtHSB_Saturation.value) / 100;
	var brightness = new Number(frmColorPicker.txtHSB_Brightness.value) / 100;
	var hue = new Number(frmColorPicker.txtHSB_Hue.value) / 360;

	if(hue > 1)
	{
		hue = 1;
		frmColorPicker.txtHSB_Hue.value = 360;
		alert("An integer between 0 and 360 is required. Closest\r\nvalue inserted.");
	}
	if(saturation > 1)
	{
		saturation = 1;
		frmColorPicker.txtHSB_Saturation.value = 100;
		alert("An integer between 0 and 100 is required. Closest\r\nvalue inserted.");
	}
	if(brightness > 1)
	{
		brightness = 1;
		frmColorPicker.txtHSB_Brightness.value = 100;
		alert("An integer between 0 and 100 is required. Closest\r\nvalue inserted.");
	}

	var rgb = HsbToRgb(hue, saturation, brightness);
	var Lab = RgbToLab(rgb);
	var cmyk = RgbToCmyk(rgb);
	
	frmColorPicker.txtHex.value = RgbToHex(rgb);
	pnlNewColor.style.backgroundColor = "#" + RgbToHex(rgb);
	
	switch (event.srcElement.id)
	{
		case "pnlVertical_Top":
		case "txtHSB_Hue":
			pnlGradient_Hue.style.backgroundColor = "#" + RgbToHex(HueToRgb(hue));
			break;
		default:
			break;
	}
	
	frmColorPicker.txtRGB_Red.value = rgb[0];
	frmColorPicker.txtRGB_Green.value = rgb[1];
	frmColorPicker.txtRGB_Blue.value = rgb[2];
	
	frmColorPicker.txtLab_Luminancy.value = Math.round(Lab[0]);
	frmColorPicker.txtLab_A.value = Math.round(Lab[1]);
	frmColorPicker.txtLab_B.value = Math.round(Lab[2]);
	
	frmColorPicker.txtCMYK_Cyan.value = Math.round(cmyk[0] * 100);
	frmColorPicker.txtCMYK_Magenta.value = Math.round(cmyk[1] * 100);
	frmColorPicker.txtCMYK_Yellow.value = Math.round(cmyk[2] * 100);
	frmColorPicker.txtCMYK_Black.value = Math.round(cmyk[3] * 100);
	
	if(Lab[0] >= 54)
		SetGradientPositionDark();
	else
		SetGradientPositionLight();
	
	if(event.srcElement.tagName != "DIV")
	{
		pnlGradientPosition.style.left = ((saturation * 255) + 7) + "px";
		pnlGradientPosition.style.top = ((255 - (brightness * 255)) + 27) + "px";
		if(hue == 0)
			pnlVerticalPosition.style.top = "27px";
		else
			pnlVerticalPosition.style.top = ((255 - (hue * 255)) + 27) + "px";
	}
	
	checkWebSafe();
}

var GradientPositionDark = new Boolean(false);
function SetGradientPositionDark()
{
	if(GradientPositionDark) return;
	GradientPositionDark = true;
	pnlGradientPosition.style.backgroundImage = "url(GradientPositionDark.gif)";
}
function SetGradientPositionLight()
{
	if(!GradientPositionDark) return;
	GradientPositionDark = false;
	pnlGradientPosition.style.backgroundImage = "url(GradientPositionLight.gif)";
}
function RgbToHsb(rgb)
{
	var sRgb = new Array(rgb[0], rgb[1], rgb[2]);
	var min = new Number(1);
	var max = new Number(0);
	var delta = new Number(1);
	var hsb = new Array(0, 0, 0);
	var deltaRgb = new Array(0, 0, 0);

	for(var i = 0; i < sRgb.length; i++)
	{
		sRgb[i] = rgb[i] / 255;
		if(sRgb[i] < min) min = sRgb[i];
		if(sRgb[i] > max) max = sRgb[i];
	}

	delta = max - min;
	hsb[2] = max;

	if(delta == 0) return hsb;

	hsb[1] = delta / max;

	for(var i = 0; i < sRgb.length; i++)
		deltaRgb[i] = (((max - sRgb[i]) / 6) + (delta / 2)) / delta;

	if (sRgb[0] == max)
		hsb[0] = deltaRgb[2] - deltaRgb[1];
	else if (sRgb[1] == max)
		hsb[0] = (1 / 3) + deltaRgb[0] - deltaRgb[2];
	else if (sRgb[2] == max)
		hsb[0] = (2 / 3) + deltaRgb[1] - deltaRgb[0];

	if(hsb[0] < 0)
		hsb[0] += 1;
	else if(hsb[0] > 1)
		hsb[0] -= 1;

	return hsb;
}
function HsbToRgb(hue, saturation, brightness)
{
	var rgb = HueToRgb(hue);
	var s = brightness * 255;

	for(var i = 0; i < rgb.length; i++)
	{
		rgb[i] = rgb[i] * brightness;
		rgb[i] = ((rgb[i] - s) * saturation) + s;
		rgb[i] = Math.round(rgb[i]);
	}
	return rgb;
}
function RgbToHex(rgb)
{
	var hex = new String();
	
	for(var i = 0; i < rgb.length; i++)
	{
		rgb[2 - i] = Math.round(rgb[2 - i]);
		hex = rgb[2 - i].toString(16) + hex;
		if(hex.length % 2 == 1) hex = "0" + hex;
	}
	
	return hex.toUpperCase();
}
function HexToRgb(hex)
{
	var rgb = new Array(0, 0, 0);
	for(var i = 0; i < rgb.length; i++)
		rgb[i] = new Number("0x" + hex.substr(i * 2, 2));
	return rgb;	
}
function pnlGradient_Top_Click()
{
	event.cancelBubble = true;
	SetGradientPosition(event.clientX - 5, event.clientY - 5);
}
function pnlGradient_Top_MouseMove()
{
	event.cancelBubble = true;
	if(event.button != 1) return;
	SetGradientPosition(event.clientX - 5, event.clientY - 5);
}
function pnlGradient_Top_MouseDown()
{
	event.cancelBubble = true;
	SetGradientPosition(event.clientX - 5, event.clientY - 5);
}
function pnlGradient_Top_MouseUp()
{
	event.cancelBubble = true;
	SetGradientPosition(event.clientX - 5, event.clientY - 5);
}
function HueToRgb(hue)
{
	var degrees = hue * 360;
	var rgb = new Array(0, 0, 0);
	var percent = (degrees % 60) / 60;
	
	if(degrees < 60)
	{
		rgb[0] = 255;
		rgb[1] = percent * 255;
	}
	else if(degrees < 120)
	{
		rgb[1] = 255;
		rgb[0] = (1 - percent) * 255;
	}
	else if(degrees < 180)
	{
		rgb[1] = 255;
		rgb[2] = percent * 255;
	}
	else if(degrees < 240)
	{
		rgb[2] = 255;
		rgb[1] = (1 - percent) * 255;
	}
	else if(degrees < 300)
	{
		rgb[2] = 255;
		rgb[0] = percent * 255;
	}
	else if(degrees < 360)
	{
		rgb[0] = 255;
		rgb[2] = (1 - percent) * 255;
	}
	
	return rgb;
}
function SetVerticalPosition(y)
{
	pnlVerticalPosition.style.top = y + "px";
	var hue = 0;
	if(y == 27)
		frmColorPicker.txtHSB_Hue.value = 0;
	else
	{
		// Determine hue degree on color wheel
		hue = ((255 - (y - 27)) / 255) * 360;
		frmColorPicker.txtHSB_Hue.value = Math.round(hue);
	}
	Hsb_Changed();
}
function pnlVertical_Top_Click()
{
	SetVerticalPosition(event.clientY - 5);
	event.cancelBubble = true;
}
function pnlVertical_Top_MouseMove()
{
	if(event.button != 1) return;
	SetVerticalPosition(event.clientY - 5);
	event.cancelBubble = true;
}
function pnlVertical_Top_MouseDown()
{
	SetVerticalPosition(event.clientY - 5);
	event.cancelBubble = true;
}
function pnlVertical_Top_MouseUp()
{
	SetVerticalPosition(event.clientY - 5);
	event.cancelBubble = true;
}
// assign event handlers
window.onload = window_load;
