<%
    '**************************************
    ' Name: Display Accurate Time from NIST
    ' Description:Display accurate time to y
    '     our visitors! This script requests the U
    '     TC time from a server that is syncronize
    '     d with an atomic clock. It then adjusts 
    '     the time according to timezone and if da
    '     ylight savings time is in effect.
    ' By: Lewis Moten
    '
    '
    ' Inputs:None
    '
    ' Returns:None
    '
    'Assumes:None
    '
    'Side Effects:None
    'This code is copyrighted and has limite
    '     d warranties.
    '**************************************
    
    ' Author: Lewis Moten
    ' Email: lewis@moten.com
    ' URL: http://www.lewismoten.com
    ' Not a requirement, but I do suggest th
    '     at you link to, visit
    ' or promote my website by referring it 
    '     to others in newsgroups,
    ' email, co-workers, etc. If this code a
    '     ppears on a website
    ' that allows interaction, please submit
    '     votes, comments, reviews,
    ' etc.
    '	List of time servers:
    '	http://boulder.nist.gov/timefreq/servi
    '     ce/time-servers.html
    ' Server used for this demonstration
    ' time.nist.gov
    ' Format returned by server - Daytime Pr
    '     otocol (RFC-867)
    ' JJJJJ YR-MO-DA HH:MM:SS TT L H msADV U
    '     ST(NIST) OTM
    ' Example:
    ' 52399 02-05-05 02:11:02 50 0 0 49.4 UT
    '     C(NIST) *
    ' JJJJJ 
    '	Modified Julian Date (MJD). Last 5 dig
    '     its of julian
    '	date (count of days since January 1, 4
    '     713 B.C.). Add
    '	2.4 Million to get actual Julian Date.
    '     
    ' YR-MO-DA 
    '	Year, Month, Day
    ' HH:MM:SS 
    '	Hours, Minutes, Seconds (in Coordinate
    '     d Universal Time - UTC)
    ' TT
    '	00 - Standard Time
    '	50 - Daylight Savings Time
    '	1 to 49 - Days within current month un
    '     til daylight 
    '		savings time adjustement approaches
    ' L
    '	Leap seconds keep UTC time adjusted wi
    '     th earths
    '	rotation (every 1 to 1 1/2 years)
    '	0 - leap second will not occur this mo
    '     nth.
    '	1 - 61 seconds will appear in last min
    '     ute of month
    '	2 - 59 seconds will appear in last min
    '     ute of month
    ' H
    '	0 - server is healthy
    '	1 - time may be in error by up to 5 se
    '     conds
    '	2 - time is known to be wrong by more 
    '     then 5 seconds
    '	4 - Hardware/Software failure - amount
    '     of time error is unknown.
    ' msADV
    '	milliseconds NIST has advanced time to
    '     compensate for
    '	network delays.
    ' UTC(NIST)
    '	Signifies UTC time comes from National
    '     Institute of
    '	Standards and Technology.
    ' OTM
    '	On-Time marker. Signifies that the arr
    '     ival time of the
    '	time recieved from server should be ac
    '     curate.
    ' ######################################
    '     ##########################
    ' Begin Code
    ' ######################################
    '     ##########################
    ' Server to query time from
    Const TimeServer = "http://time.nist.gov:13"
    ' Define your timezone here
    Const TimeZoneOffset = -5
    ' Set to true or false if you observe da
    '     ylight savings time
    Const DST = True
    ' Use XML HTTP object to request web pag
    '     e content
    'Set Spider = Server.CreateObject ("MSXM
    '     L2.XMLHTTP.3.0")
    'Set Spider = Server.CreateObject ("MSXM
    '     L2.ServerXMLHTTP")
    Set Spider = Server.CreateObject ("Microsoft.XMLHTTP")
    Spider.Open "GET", TimeServer, False, "", ""
    Spider.Send
    NIST = Spider.ResponseText
    Set Spider = Nothing
    ' Parse UTC date
    UTC = Mid(NIST, 11, 2) & "/" & Mid(NIST, 14, 2) & "/" & Mid(NIST, 8, 2) & " " & Mid(NIST, 16, 9)
    ' Is daylight savings in effect?
    IsDaylightSavings = CInt(Mid(NIST, 26, 2)) = 50 Or (Month(UTC) > 6 AND CInt(Mid(NIST, 26, 2)) > 0)
    ' Create the local time
    LocalTime = DateAdd("h", TimeZoneOffset, UTC)
    ' Modify for daylight savings
    If DST Then
    	If IsDaylightSavings Then LocalTime = DateAdd("h", 1, LocalTime)
    End If
    ' Write out the results
    Response.Write "Time Server URL: " & TimeServer & "<BR>"
    Response.Write "Server Responded with: " & NIST & "<BR>"
    Response.Write "Parsed UTC Date: " & UTC & "<BR>"
    Response.Write "Your timezone offset: " & TimeZoneOffset & "<BR>"
    Response.Write "You observe daylight savings time: " & DST & "<BR>"
    Response.Write "Your local time is: " & LocalTime & "<BR>"
    %>
