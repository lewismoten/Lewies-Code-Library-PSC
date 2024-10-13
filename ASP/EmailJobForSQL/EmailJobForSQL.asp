<%
    Dim lObjConn
    Dim vbs
    Dim lStrSQL
    Set lObjConn = Server.CreateObject("ADODB.Connection")
    lObjConn.Open _
    	"Provider=SQLOLEDB.1;" & _
    	"Data Source=LOCALHOST;" & _
    	"Initial Catalog=msdb;" & _
    	"User ID=sa;" & _
    	"Password=;"
    lStrSQL = "sp_add_job " & _
    	"@job_name = 'SendMailJob'," & _
    	"@enabled = 1," & _
    	"@description = 'Sends e-mail messages'," & _
    	"@start_step_id = 1," & _
    	"@category_name = '[Uncategorized (Local)]'"
    	
    Set lObjRs = lObjConn.Execute(lStrSQL)
    vbs = GetScript("LOCALHOST", "Master", "")
    vbs = Replace(vbs, "'", "''")
    lStrSQL = "sp_add_jobstep " & _
    	"@job_name = 'SendMailJob', " & _
    	"@step_id = 1, " & _
    	"@step_name = 'Find and Send Mail', " & _
    	"@subsystem = 'ACTIVESCRIPTING', " & _
    	"@command = '" & vbs & "'"
    lObjConn.Execute lStrSQL
    lStrSQL = "sp_add_jobschedule " & _
    	"@job_name = 'SendMailJob', " & _
    	"@name = 'Every 10 Minutes', " & _
    	"@enabled = 1, " & _
    	"@freq_type = 4, " & _
    	"@freq_interval = 1, " & _
    	"@freq_subday_type = 0x4, " & _
    	"@freq_subday_interval = 10"
    lObjConn.Execute lStrSQL
    lStrSQL = ""
    Set lObjConn = Nothing
    Function GetScript(ByRef pStrDataSource, ByRef pStrInitialCatalog, ByRef pStrSAPassword)
    	GetScript = _
    		"Dim lObjConn" & vbCrLf & _
    		"Dim lObjRs" & vbCrLf & _
    		"Dim lStrSQL" & vbCrLf & _
    		"Dim lObjMailer" & vbCrLf & _
    		vbCrLf & _
    		"Const adOpenForwardOnly = 0" & vbCrLf & _
    		"Const adLockPessimistic = 2" & vbCrLf & _
    		"Const adCmdText = 1" & vbCrLf & _
    		vbCrLf & _
    		"lStrSQL = ""SELECT [From], [To], [Subject], [Body] FROM [Email]""" & vbCrLf & _
    		vbCrLf & _
    		"Set lObjConn = CreateObject(""ADODB.Connection"")" & vbCrLf & _
    		"Set lObjRs = CreateObject(""ADODB.Recordset"")" & vbCrLf & _
    		"lObjConn.Open _" & vbCrLf & _
    		"	""Provider=SQLOLEDB.1;"" & _" & vbCrLf & _
    		"	""Data Source=" & pStrDataSource & ";"" & _" & vbCrLf & _
    		"	""Initial Catalog=" & pStrInitialCatalog & ";"" & _" & vbCrLf & _
    		"	""User ID=sa;"" & _" & vbCrLf & _
    		"	""Password=" & pStrSAPassword & ";""" & vbCrLf & _
    		vbCrLf & _
    		"lObjRs.Open lStrSQL, lObjConn, adOpenForwardOnly, adLockPessimistic, adCmdText" & vbCrLf & _
    		vbCrLf & _
    		"While Not lObjRs.EOF" & vbCrLf & _
    		"	Set lObjMailer = CreateObject(""CDONTS.NewMail"")" & vbCrLf & _
    		"	lObjMailer.From		= lObjRs(0) & """"" & vbCrLf & _
    		"	lObjMailer.To		= lObjRs(1) & """"" & vbCrLf & _
    		"	lObjMailer.Subject	= lObjRs(2) & """"" & vbCrLf & _
    		"	lObjMailer.Body		= lObjRs(3) & """"" & vbCrLf & _
    		"	lObjMailer.Send" & vbCrLf & _
    		"	lObjRs.Delete" & vbCrLf & _
    		"	lObjRs.MoveNext" & vbCrLf & _
    		"	Set lObjMailer = Nothing" & vbCrLf & _
    		"Wend" & vbCrLf & _
    		vbCrLf & _
    		"lObjRs.Close" & vbCrLf & _
    		"lObjConn.Close" & vbCrLf & _
    		vbCrLf & _
    		"Set lObjRs = Nothing" & vbCrLf & _
    		"Set lObjConn = Nothing"
    End Function
    %>
    done . . .<BR><BR>
    You may wish to open SQL Enterprise Manger and find the
    "Management" folder under your database. Find "SQL Server Agent" with
    a child node called "Jobs". A new job called "SendMailJob" should be
    present.
