'**************************************
' Name: PageRange
' Description:I had database searches that
'     were returning way too many pages to 
'     display an option to view each page. What
'     I did was set up this procedure to return
'     a range of pages, and also options to
'     jump ahead, back, first page, last page,
'     previouse page, and/or the next page.
'     This routine tells you what those links are.
' By: Lewis Moten
'
'
' Inputs:
'	anCurrentPage Long- The current page being viewed
'	anMaxPages Long - The maximum page that the user may view
'	anRange Long - the range of pages to display from the current page
'
' Returns:Variant Array (LBound 0, UBound 3)
'	Array(0) - Page Numbers to link to
'	Array(1) - Descriptive Text
'	Array(2) - Short Text
'
'Side Effects:If using this function with HTML, you should apply HTML Encoding to the Short Text Descriptions.
'This code is copyrighted and has limited warranties.
'**************************************
Function PageRange(ByRef anCurrentPage As Long, ByRef anMaxPages As Long, ByRef anRange As Long) As Variant()
    Dim lvPageData(2) As Variant
    Dim lnMin As Long
    Dim lnMax As Long
    Dim lnPage As Long
    Dim lsPageNums As String
    Dim lsLongText As String
    Dim lsShortText As String' don't forget to encode these if using HTML
    If Not anMaxPages > 1 Then Exit Function
    lnMin = anCurrentPage - anRange
    lnMax = anCurrentPage + anRange


    If lnMax > anMaxPages Then
        lnMax = anMaxPages
        lnMin = lnMax - (anRange * 2)
    End If


    If lnMin < 1 Then
        lnMin = 1
        lnMax = lnMin + (anRange * 2)
    End If
    If lnMax > anMaxPages Then lnMax = anMaxPages
    ' Link to first page
    If lnMin > 1 Then
        lsPageNums = lsPageNums & "/1"
        lsLongText = lsLongText & "/First Page"
        lsShortText = lsShortText & "/|<"
    End If
    ' Link to jump back within a range of pages
    If lnMin > ((anRange * 2) + 1) Then
        lsPageNums = lsPageNums & "/" & CStr(anCurrentPage - ((anRange * 2) + 1))
        lsLongText = lsLongText & "/Jump Back"
        lsShortText = lsShortText & "/<<"
    End If
    ' back one page
    If Not (anCurrentPage = 1 Or anMaxPages < 3) Then
        lsPageNums = lsPageNums & "/" & CStr(anCurrentPage - 1)
        lsLongText = lsLongText & "/Previous Page"
        lsShortText = lsShortText & "/<"
    End If
    ' Loop through range of pages
    For lnPage = lnMin To lnMax
        lsPageNums = lsPageNums & "/" & CStr(lnPage)
        lsLongText = lsLongText & "/Page " & CStr(lnPage)
        lsShortText = lsShortText & "/" & CStr(lnPage)
    Next
    ' forward one page
    If Not anCurrentPage = anMaxPages And anMaxPages > 2 Then
        lsPageNums = lsPageNums & "/" & CStr(anCurrentPage + 1)
        lsLongText = lsLongText & "/Next Page"
        lsShortText = lsShortText & "/>"
    End If
    ' Link to jump ahead within a range of pages
    If lnMax + (anRange * 2) < anMaxPages Then
        lsPageNums = lsPageNums & "/" & CStr(anCurrentPage + ((anRange * 2) + 1))
        lsLongText = lsLongText & "/Jump Forward"
        lsShortText = lsShortText & "/>>"
    End If
    ' Link to last page
    If lnMax < anMaxPages Then
        lsPageNums = lsPageNums & "/" & CStr(anMaxPages)
        lsLongText = lsLongText & "/Last Page"
        lsShortText = lsShortText & "/>|"
    End If
    lsPageNums = Mid(lsPageNums, 2)
    lsLongText = Mid(lsLongText, 2)
    lsShortText = Mid(lsShortText, 2)
    lvPageData(0) = Split(lsPageNums, "/")
    lvPageData(1) = Split(lsLongText, "/")
    lvPageData(2) = Split(lsShortText, "/")
    PageRange = lvPageData
End Function