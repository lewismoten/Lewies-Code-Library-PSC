<%
Call Write_kbpsField()

Sub Write_kbpsField()
	
	' this procedure writes a hidden input field to the client that states
	' the recorded speed that they are currently downloading at.
	
	' This procedure is not completly accurate.  The client may have
	' other programs sharing bandwidth, or may be downloading other images
	' and files from your website while this script executes.  Network
	' traffic and the number of hops to your web server can also slow
	' the recorded speed down.  This script only provides a "ball park figure"
	' to help you identify the speed that your viewers connect at.
	
	' If the viewer refreshes the page - there is a chance that the page
	' will be retrieved from cache.  This would make it appear that the
	' client has an extreamly fast connection.  You may wish to add headers
	' and meta tags to instruct the viewers client to refrain from caching
	' the page.
	
	' Glossary:
	'	kilobit - A data unit equal to 1,000 bits
	'	kilobits per second (Kbps) - Data transfer speed, as on
	'		a network, measured in multiples of 1,000 bits per second.

	' Common Measurements:
	
	' b = bit
	' B = byte
	' k = kilo
	' m = mega
	
	' +-------------------+--------+------+----------+-------+
	' | connection        |   kb   |  mb  |    B     |  kB   |
	' +-------------------+--------+------+----------+-------+
	' | 28.8 modem        |   28.8 |  .03 |   3600   |  3.56 |
	' | 36.6 modem        |   36.6 |  .04 |   4575   |  4.47 |
	' | 56k modem         |   53.3 |  .05 |   6662.5 |  6.51 |
	' | 64k ISDN          |   64   |  .06 |   8000   |  7.81 |
	' | 128k ISDN         |  128   |  .13 |  16000   | 15.63 |
	' | 384k DSL          |  384   |  .38 |  48000   | 46.88 |
	' | 768k DSL          |  768   |  .76 |  96000   | 93.75 |
	' | 1.5Mb DSL/Cable   | 1536   | 1.53 | 192000   | 187.5 |
	' + ------------------+--------+------+----------+-------+

	Const lLngKilobits = 384
	' lLngKilobits - number of kilobits to send to the client for testing
	' bandwidth.  
	'	Less kilobits = faster results
	'	More kilobits = more accurate results
	
	Dim lLngByte
	Dim lLngMaxBytes
	Dim lStrChar

	' Write javascript to determine the time when the client
	' first loaded the page
	Response.Write("<SCRIPT language=""JavaScript""><!--" & vbCrLf)
	Response.Write("var sd;var st;var kbps=new Number(-1);")
	Response.Write("var kb=new Number(" & lLngKilobits & ");")
	Response.Write("var s=new Number();sd=new Date();")
	Response.Write("st=sd.getTime();//--></SCRIPT><!--")

	' Determine the number of bytes to write to the client.
	lLngMaxBytes = lLngKilobits * 125
		' 8			= number of bits in a byte
		' 1000		= number of bits in a kilobit
		' 1000 \ 8	= 125
		' 125		= Number of bytes within a kilobit
	
	For lLngByte = 1 To lLngMaxBytes
		
		' Write random characters to represent compressed
		' data. some modems support compression and would
		' give less accurate results if we wrote only letters
		' or data with simple patterns.
		
		lStrChar = Chr(Int(Rnd * 255) + 1)
		
		' some characters thrown together would create invalid
		' HTML tags.  (ie. "-->").  This line prevents that from
		' happening.
		While Not InStr(1, "<> " & vbTab & vbCrLf, lStrChar) = 0
			lStrChar = Chr(Int(Rnd * 255) + 1)
		Wend
		Response.Write(lStrChar)
		
	Next
	
	' write the script to calculate the kilobytes per second based
	' on the amount of data written to the client and the time it
	' took to receive the data.  (The script is small to let results
	' have the ability to be more accurate)
	Response.Write("--><SCRIPT language=""JavaScript""><!--" & vbCrLf)
	Response.Write("var ed=new Date();var et=ed.getTime();")
	Response.Write("if(et-st>=1){s=(et-st)/1000;")
	Response.Write("kbps=Math.round(kb/s);};//--></SCRIPT>")

	' write a hidden input field to pass the clients speed to the server.
	Response.Write("<SCRIPT language=""JavaScript""><!--" & vbCrLf)
	Response.Write("document.write('<INPUT name=""Kbps"" value=""' + kbps + '"" type=""hidden"">');")
	Response.Write(";//--></SCRIPT>")
End Sub
%>
