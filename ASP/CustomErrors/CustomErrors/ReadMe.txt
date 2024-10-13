-----
Notes
-----

Unfortunately, this solution does not work for everyone because the service provider may not assign privledges to access the metebase.  This was a "last hope" to be able to get a custom 404 page with my web site so that I could log errors.  Fortunately, this solution worked for me.

This script interacts with COM objects that communicate with the MetaBase for Internet Information Server.  It is not guarenteed that this solution will work for you.  Your service provider must enable your account to modify these settings.  
You may be prompted to enter your username/password when this application starts.  This is because the Internet User Account where your site is hosted does not have permission to access this object.  You are prompted to login with your "Authoring" account.  You can look at the code to verify that I do not recieve this information in any way.

I will worn you however, that the username/password is passed in Base64 encoding to the web server.  It is an easy format to decode and many people consider this to be clear-text.  If someone is snooping the connection between you and the server, your account can easily be compromised.

The common error pages I provided are extreamly simple.  They are there to get you started.  I suggest that you go through each one and customize them to look like they are a part of your site.

------
Instructions
------

You have 3 different settings for each URL.

FILE - Specify a file on your website.  "C:\FileName.htm".

URL - Specify a URL on your website.  This must be prefixed with a forward slash "/"

DEFAULT - Inherits the properties from the parent.  These can not be seen, but usually represent the default physical file when the server was installed.


If the property is not defined, then it will be changed to "DEFAULT".

------

I have included the common error pages with descriptions.  I would suggest to place everything in a sub folder called "CustomErrors".  Then go to that directory on your webserver.

http://www.yourdomain.com/CustomErrors/

You may be prompted for your Web Site Account UserName/Password.  Then, you will have a list of Errors that you can customize.  I suggest trying out the 404 error first - since it is the easiest to test.  Change the message type to "URL" and type "/CustomErrors/404.asp" for the Path. Go to the bottom of the page and hit "Apply".

The settings will save and bring you back to the list again.  You should see that your 404 path is set to the new URL that you defined.  Type in an invalid url in your web browser to test.

"http://www.yourdomain.com/CustomErrors/InvalidURL.asp"

Did it work?  Congradulations!