    <%
    '**************************************
    ' Name: HTML2Word
    ' Description:Convert HTML web pages int
    '     o Word documents. This is great for crea
    '     ting word documents on the fly. Can be u
    '     sed to email documents after finnished c
    '     onverting. This code is in various place
    '     s on the web. I'm supprised I didn't fin
    '     d it on PSC. This is my own version for 
    '     you to modify as you wish. Larger files 
    '     take longer to process. My waiting perio
    '     d is not good! System resources will be 
    '     pegged. Seems to have problems with Webs
    '     ite URLs, but files on the same server d
    '     on't give me as many problems.
    ' By: Lewis Moten
    '
    '
    ' Inputs:SourceURL, TargetFileName
    '
    ' Returns:None
    '
    'Assumes:None
    '
    'Side Effects:None
    'This code is copyrighted and has limite
    '     d warranties.
    '**************************************
    
    Option Explicit
    HTML2Word "file://" & Server.MapPath("test.htm"), Server.MapPath("test.doc")
    Response.Write "Finnished at: " & Now()
    Sub HTML2Word(ByRef pstrSourceURL, ByRef pstrTargetFile)
    	Dim lobjWordApp
    	Dim lobjInternetExplorer
    	Set lobjWordApp = Server.CreateObject("Word.Document")
    	Set lobjInternetExplorer = Server.CreateObject("InternetExplorer.Application")
    	lobjInternetExplorer.Navigate pstrSourceURL
    	IEStandBy lobjInternetExplorer
    	
    	lobjInternetExplorer.document.body.createTextRange.execCommand("Copy")
    	IEStandBy lobjInternetExplorer
    	
    	lobjWordApp.Content.Paste
    	lobjWordApp.SaveAs pstrTargetFile
    	lobjWordApp.Close
    	
    	lobjInternetExplorer.Quit
    	
    	Set lobjInternetExplorer = Nothing
    	Set lobjWordApp = Nothing 
    End Sub
    Sub IEStandBy(ByRef pobjInternetExplorer)
    	' Not good (Pegs system resources)
    	' Need to find a way to "sleep"
    	While pobjInternetExplorer.busy:Wend
    	While Not pobjInternetExplorer.Document.readyState = "complete":Wend
    End Sub
    %>