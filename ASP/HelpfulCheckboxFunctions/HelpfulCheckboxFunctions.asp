    <%
    '**************************************
    ' for :Helpful Checkbox Functions
    '**************************************
    '(c)Copyright 2001 Lewis Edward Moten III, All rights reserved.
    '**************************************
    ' Name: Helpful Checkbox Functions
    ' Description:Ok, I'm uploading somethin
    '     g that is very simple this time - but th
    '     is is one of the most used things that I
    '     have. These little functions help me mak
    '     e code smaller and change the wording to
    '     users viewing the text next to a check b
    '     ox. I find that if you change the words 
    '     based on what the user has chosen, it fe
    '     els a little more comfortable to them.
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
    
    Dim lBlnYummy
    lBlnYummy = Request.QueryString("Yummy") = "True"
    %>
    <FORM>
    	<INPUT type="checkbox" name="Yummy" value="True"<%=Checked(lBlnYummy)%>>
    	<%=CheckAction(lBlnYummy)%> if you think apples taste great.<BR>
    	<BR>
    	<INPUT type="submit" value="OK!">
    </FORM>
    <%
    Function CheckAction(ByRef pBlnIsChecked)
    	If pBlnIsChecked Then
    		CheckAction = "Leave this box checked "
    	Else
    		CheckAction = "Check this box "
    	End If
    End Function
    Function Checked(ByRef pBlnIsChecked)
    	If pBlnIsChecked Then Checked = " checked "
    End Function
%>