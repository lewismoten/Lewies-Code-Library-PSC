var bRememberIsDirty			// determines if we have all ready sent a message to the client.

bRememberIsDirty = false		// initialize

function RememberMeWarning()
{
	// if the user just marked the checkbox
	if (document.NovellClient.RememberMe.checked==true)
	{
		// if the user has not yet been warned
		if (bRememberIsDirty!=true)
		{
			// warn user
			window.alert(
				"In order to be remembered, this application will\n" +
				"store your user information in a cookie. If other\n" +
				"people share your workstation, it may be best to\n" +
				"leave this option off for security purposes.");

			// remember that we warned the user
			bRememberIsDirty = true
		}
	}
}
function OK_onClick()
{
	if (document.NovellClient.UserName.value == ''||document.NovellClient.Password.value == '')
	{
		window.alert('Please type in both your Username\nand Password before Logging In.')
	}
	else
	{
		document.NovellClient.submit()
	}
}
function Cancel_onClick()
{
	window.location.href = "/"
}
