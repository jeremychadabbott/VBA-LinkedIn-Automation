Attribute VB_Name = "GetHTML"
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
    'Application.Wait (Now + TimeValue("00:00:02"))
    'Application.SendKeys ("^c"), True
    'Application.Wait (Now + TimeValue("00:00:02"))
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
    Dim foundLink As String
    Dim linkList As String
    linkList = ""
    
    ' Example URL pattern
    Dim urlPattern As String
    urlPattern = "/jobs/view/"
    
    ' Loop to find occurrences
    startPos = 1
    Do
        startPos = InStr(startPos, PageHTML, urlPattern, vbTextCompare)
        If startPos > 0 Then
            endPos = InStr(startPos, PageHTML, "?")
            If endPos > 0 Then
                foundLink = Mid(PageHTML, startPos, endPos - startPos + 1)
                linkList = linkList & foundLink & vbCrLf
                startPos = endPos + 1
            Else
                Exit Do
            End If
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
    
End Sub

