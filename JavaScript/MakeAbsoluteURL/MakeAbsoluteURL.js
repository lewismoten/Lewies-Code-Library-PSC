//**************************************
// Name: Make Absolute URL
// Description:Converts a relevent url to an absolute URL.
//login.asp would become http://www.yourdomain.com/your/directory/structure/login.asp
// By: Lewis Moten
//
//
// Inputs:Relevent URL
//
// Returns:absolute URL.
//
//Assumes:None
//
//Side Effects:None
//This code is copyrighted and has limited warranties.
//**************************************

alert('The absolute path to login.asp is: ' + MakeAbs('login.asp'))
function MakeAbs(pStrRelURL)
{
	var lStrAbsURL = new String('');
	// remove any file name after last dir
	var lStrPath = new String(window.location.pathname);
	lStrPath = lStrPath.substring(1, lStrPath.lastIndexOf('/') + 1)
	if(pStrRelURL.indexOf('//') != -1)
	{
		// all ready an absolute url
		lStrAbsURL = pStrRelURL
	}
	else
	{
		if(pStrRelURL.indexOf('../') == 0)
		{
			// find out how many parents.
			var lLngParents = new Number(0)
			for(
				lLngIndex = pStrRelURL.indexOf('../', 0);
				lLngIndex != -1;
				lLngIndex = pStrRelURL.indexOf('../', lLngIndex + 3)
			)
			{
				lLngParents ++
			}

			// remove '../'
			pStrRelURL = pStrRelURL.substring(pStrRelURL.lastIndexOf('../') + 3)

			// remove sub directories
			for(var lLngCount = 0;lLngCount < lLngParents;lLngCount++)
			{
				// if the path doesn't have any directory delimiters left
				if(lStrPath.indexOf('/') == lStrPath.lastIndexOf('/'))
				{
					lStrPath = ''
				}
				else
				{
					lStrPath = lStrPath.substring(0, lStrPath.lastIndexOf('/') + 1);
				}
			}

			lStrAbsURL = window.location.protocol + '//'
			lStrAbsURL += window.location.hostname + '/'
			lStrAbsURL += lStrPath
			lStrAbsURL += pStrRelURL
		}
		else
		{
			if(pStrRelURL.indexOf('/') == 0)
			{
				lStrAbsURL = window.location.protocol + '//'
				lStrAbsURL += window.location.hostname + '/'
				lStrAbsURL += pStrRelURL
			}
			else
			{
				lStrAbsURL = window.location.protocol + '//'
				lStrAbsURL += window.location.hostname + '/'
				lStrAbsURL += lStrPath
				lStrAbsURL += pStrRelURL
			}
		}
	}
	return(lStrAbsURL)
}