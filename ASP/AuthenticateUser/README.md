# [Lewie's Code Library PSC](../../README.md)

Open source projects that I had published to Planet Source Code.

## [Classic ASP / vbScript](../README.md)

### AuthenticateUser

*1/12/2001 2:42:27 PM*

Authenticates a user to make sure if they have previously logged into the site. Requires Session("UserID") to be populated. This usually represents the Users ID within a data base. (Users.UserID) If a user is not loged in, they are redirected to a page to attempt a login. This is useful when the ability to "Auto-Login" has been enabled to use previously saved login information in the users cookies. When a user is redirected to the login page, The URL they were attempting to view is passed along in the Query String along with the reason why they need to login. If the user was posting data to the protected page (perhaps a session timed out), then the previous page they were posting from is sent as the URL that the user is redirected to after they have successfully logged in. This is done to help reduce errors when visiting a page that expected posted form data.


