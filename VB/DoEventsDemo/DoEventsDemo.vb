'**************************************
' Name: DoEvents Demonstration
' Description:The main point here is to 
'     show beginners how to allow your code to
'     "wait" for other events to take place be


'     fore continuing to execute any code that
    '     follows. It also includes a timeout feat
    '     ure in case the event never raises a fla
    '     g that it is finnished.
' By: Lewis Moten
'This code is copyrighted and has limite
'     d warranties.
'**************************************

Dim mBolDone As Boolean


Sub ExecuteCode()
    Dim lLngTimeout As Long
    Dim lDtmStart As Date
    ' Set number of seconds until a timeout 
    '     occurs
    lLngTimeout = 10
    ' Setup starting time
    lDtmStart = Now()
    ' Setup flag to state that event has not
    '     finnished
    mBlnDone = False
    ' Loop while the process has not yet fin
    '     nished
    ' and has not yet timed out
    While Not mBolDone And DateDiff("s", lDtmStart, Now()) < lLngTimeout
        ' Allow other processes to occur
        DoEvents
    Wend
    MsgBox "Finnished Waiting"
End Sub
Sub cmdCancel_onClick()
    mBlnDone = True
    msgbox "Process Interrupted"
End Sub