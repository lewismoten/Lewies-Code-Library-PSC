var BlnHighlight = new Boolean(false);
var StrTypeOptions = new String('<OPTION value="">default</OPTION><OPTION value="FILE">file</OPTION><OPTION value="URL">url</OPTION>')

function ErrorGUI(pLngCode, pLngSubCode, pStrReason, pStrDetails, pStrMapType, pStrMapPath)
{
	var lStrHTML = new String();
	
	BlnHighlight = (BlnHighlight == false)
	
	lStrHTML += '<INPUT type="hidden" name="ErrorCode" value="' + pLngCode + ',' + pLngSubCode + '">';
	
	lStrHTML += '<TR class="Highlight' + BlnHighlight + '">';
	
	lStrHTML += '<TD width="1%">' + pLngCode
	if(pLngSubCode != 0)
		{lStrHTML += ':' + pLngSubCode}
	lStrHTML += '</TD>'

	lStrHTML += '<TD width="33%" nowrap>' + pStrReason
	if(pStrDetails != '')
		{lStrHTML += '<BR><I>' + pStrDetails + '</I>'}
	lStrHTML += '</TD>'

	lStrHTML += '<TD width="1%">';
	lStrHTML += '<SELECT name="CustomType">';
	
	var lObjRegExp = new RegExp('value="' + pStrMapType + '"', 'gi');
	
	
	lStrHTML += StrTypeOptions.replace(lObjRegExp, 'value="' + pStrMapType + '" selected')
	
	lStrHTML += '</SELECT>'
	lStrHTML += '</TD>'
	
	lStrHTML += '<TD>'
	lStrHTML += '<INPUT class="CustomPath" name="CustomPath" type="text" size="30" value="' + pStrMapPath + '">'
	lStrHTML += '</TD>'

	lStrHTML += '</TR>'
	
	return(lStrHTML);
}
function HeaderGUI(pStrServerName, pStrMetabasePath, pStrKeyType)
{
	var lStrHTML = new String();
	
	lStrHTML += '<H3>Error Mapping Properties for ' + pStrServerName + '</H3>';
	lStrHTML += '<H4>Metabase Path: ' + pStrMetabasePath + '</H4>';
	lStrHTML += '<H4>Key Type: ' + pStrKeyType + '</H4>';
	
	lStrHTML += '<FORM method="post" action="Update.asp">';
	lStrHTML +=		'<TABLE Class="TableFrame">'
	lStrHTML +=			'<TR>'
	lStrHTML +=				'<TD>';
	lStrHTML +=					'<TABLE border="0" cellspacing="1">';
	lStrHTML +=						'<TR Class="Header">';
	lStrHTML +=							'<TD bgColor="#d0d0d0">Number</TD>';
	lStrHTML +=							'<TD bgColor="#d0d0d0">Default Text</TD>';
	lStrHTML +=							'<TD bgColor="#d0d0d0">Message Type</TD>';
	lStrHTML +=							'<TD bgColor="#d0d0d0">Path</TD>';
	lStrHTML +=						'</TR>';
	
	return(lStrHTML);
}
function FooterGUI()
{
	var lStrHTML = new String();
	
	lStrHTML +=					'</TABLE>'
	lStrHTML +=				'</TD>'
	lStrHTML +=			'</TR>'
	lStrHTML +=		'</TABLE>'
	lStrHTML +=		'<CENTER>'
	lStrHTML +=			'<BR>'
	lStrHTML +=			'<BR>'
	lStrHTML +=			'<INPUT type="Submit" value="Apply">'
	lStrHTML +=			'<BR>'
	lStrHTML +=			'<BR>'
	lStrHTML +=			'<HR width="80%">'
	lStrHTML +=			'<FONT size="1">'
	lStrHTML +=				'For help using this utility, you may contact '
	lStrHTML +=				'<A href="http://www.lewismoten.com">Lewis Moten</A>.'
	lStrHTML +=			'</FONT>'
	lStrHTML +=		'</CENTER>'
	lStrHTML += '</FORM>'
	
	return(lStrHTML);
}