Sub Get_Page_HTML(JobHyperlink As String, JobDescription As String, Remote As String, Variable1 As String, Variable2 As String)
    ' Assumes web page is already open and loaded
    Dim Pinned_Posts As Integer
    Pinned_Posts = 0
    
    ' Hotkey ctrl+shift+i to slide open the element inspector
    Application.SendKeys ("^+i"), True
    Application.Wait (Now + TimeValue("00:00:10"))
    
    ' Copy HTML element into clipboard
    Dim clipboard As MSForms.DataObject
    Set clipboard = New MSForms.DataObject
    clipboard.Clear

    'click on root html element
    Call Mouse_Right_Click_Inspector_Root_Element
    Application.Wait (Now + TimeValue("00:00:01"))
    Call Mouse_Right_Click_Inspector_Root_Element_CopyOUterHTML
    Application.Wait (Now + TimeValue("00:00:01"))
    
    Application.CutCopyMode = False
    Dim PageHTML As String
    PageHTML = ""
    clipboard.GetFromClipboard
    PageHTML = clipboard.GetText
    
    Call Mouse_Click_Chrome_Inspector_Close
    
    If PageHTML = "" Or Len(PageHTML) < 500 Then
        For Repeat = 1 To 100
            MsgBox "Didn't get PageHTML"
        Next Repeat
    End If
    
' Initialize variables
Dim startPos As Long
Dim endPos As Long
Dim questionMarkPos As Long
Dim foundLink As String
Dim linkList As String
Dim dismissedText As String
linkList = ""
dismissedText = "again."
dismissedPreText = "unified"
urlPattern = "/jobs/view/"

' Loop to find occurrences
startPos = 1
Do
    startPos = InStr(startPos, PageHTML, urlPattern, vbTextCompare)
    If startPos > 0 Then
        ' Check if the dismissed text is not within 2000 characters after the start of the found link
        Dim subString As String
        subString = Mid(PageHTML, startPos + 750, 1500)
        'check for "won't show this again"
        Dim checkPos As Long
        checkPos = InStr(10, subString, dismissedText, vbTextCompare)
        If checkPos = 0 Then
            subString = Mid(PageHTML, startPos - 750, 750)
            'check for "won't show this again"
            checkPos = InStr(10, subString, dismissedPreText, vbTextCompare)
            If checkPos = 0 Then
        
                ' Find the end position of the URL
                endPos = InStr(startPos, PageHTML, """", vbTextCompare)
                If endPos > 0 Then
                    ' Check for the position of the "?" character
                    questionMarkPos = InStr(startPos, PageHTML, "?", vbTextCompare)
                    If questionMarkPos > 0 And questionMarkPos < endPos Then
                        endPos = questionMarkPos
                    End If
    
                    ' Parse the URL up to the "?" character
        
                    foundLink = Mid(PageHTML, startPos, endPos - startPos)
                
                    linkList = linkList & foundLink & vbCrLf
    
                    ' Exit the loop once a valid link is found and parsed
                    Exit Do
                End If
                
                
            End If
            
        End If
        startPos = startPos + endPos + 1
    End If
Loop While startPos > 0


    
    ' Output the list of links
    For Repeat = 1 To 100
        'MsgBox linkList
    Next Repeat
    
    Call Mouse_Click_Chrome_Inspector_Close
    
    ' Make JobHyperlink = the first one in the list
    If Len(linkList) > 0 Then
        JobHyperlink = Split(linkList, vbCrLf)(0)
    End If
    
    ' error check
    If LCase(JobHyperlink) Like "*comget*" Then
        For Repeat = 1 To 100
            MsgBox "didn't scrape HTML correctly"
        Next Repeat
    End If
    
    'error check 2
    JobHyperlink = Replace(JobHyperlink, "/?openBottomSheet=verifiedHiringV2", "")
    
    'MsgBox JobHyperlink
End Sub

