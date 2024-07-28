Attribute VB_Name = "MainBody"
Declare PtrSafe Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Declare PtrSafe Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long
Public Const SW_MAXIMIZE As Long = 3&        'Show window Maximised
Declare PtrSafe Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
Public Declare PtrSafe Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
Public Const MOUSEEVENTF_LEFTDOWN = &H2
Public Const MOUSEEVENTF_LEFTUP = &H4
Public Const MOUSEEVENTF_RIGHTDOWN As Long = &H8
Public Const MOUSEEVENTF_RIGHTUP As Long = &H10
Declare PtrSafe Function GetSystemMetrics32 Lib "user32" Alias "GetSystemMetrics" (ByVal nIndex As Long) As Long
Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Sub Run_Linkedin_Manually(automate As String, comm As String, Only_like As String)

troubleshoot = 0
'Automate = 0 run normal routine, launching chrome and searching suggested jobs
'Automate = 1 run normal routine, search job querry alreaqdy brought up (do not launch chrome or initiate search)
    Dim clipboard As MSForms.DataObject
    Dim URL As String
    Dim Variable1 As String
    Dim Variable2 As String
    Dim Remote As String
    Dim JobDescription As String
    Dim JobHyperlink As String

If automate = "1" Then
    Application.Wait (Now + TimeValue("00:00:01"))
    Application.SendKeys ("%{tab}")
    Application.Wait (Now + TimeValue("00:00:05"))
    GoTo Automate1_JumpPoint
End If

Launch_chrome:
    Application.Wait (Now + TimeValue("00:00:02"))
    file = "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
    If Dir("C:\Program Files\Google\Chrome\Application\chrome.exe") <> "" Then file = "C:\Program Files\Google\Chrome\Application\chrome.exe"
    Shell (file)
    Application.Wait (Now + TimeValue("00:00:03"))
    'maximize
    Application.SendKeys ("%" & Chr(32))
    Sleep 500
    Application.SendKeys ("x")
    Sleep 500

start:

    Call CheckIfChromeIsOPen

    'Auto Comment Setting
    auto_comment = automatic_comment

    If auto_comment <> "" Then
        Randomize
        r = Rnd(1) * 6
        For Repeat = 1 To r
            auto_comment = auto_comment & "!"
        Next Repeat
    End If

    URL = "https://www.linkedin.com/jobs/"

Load_LInkedIN:
    'Load User profile Webpage
    Sleep 150
    Application.SendKeys ("%d")
    Sleep 150
    For Repeat = 1 To 200
       Application.SendKeys ("{BS}")
    Next Repeat
    If URL = "" Then MsgBox "About to enter URL but it's blank!?"
    If URL = Chr(34) Then MsgBox "Freeze"
    'MsgBox URL
    Application.SendKeys (URL) & "      ", True 'Navigate Homepage
    Sleep 250
    Application.SendKeys ("~"), True
    
    'Let Page Load
    Application.Wait (Now + TimeValue("00:00:15"))
    
    'Application.SendKeys (" "), True 'scrolls
    'Pinned Loop
    pinned_post = 0
    try = 0

    ' Click on "Show All" Jobs
    SetCursorPos 593, 732
    Sleep 100
    Call mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0) 'click left mouse
    Sleep 75
    Call mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
    Application.Wait (Now + TimeValue("00:00:06"))
    
    
Automate1_JumpPoint:
    
Get_Job_HTML:

    ' Get Page HTML
    Variable2 = ""
    Variable1 = ""
    Remote = ""
    JobDescription = ""
    JobHyperlink = "get"
    Call Get_Page_HTML(JobHyperlink, JobDescription, Remote, Variable1, Variable2)

    
    ' check if jobhyperlink is already on Db sheet
    
    
    ' Add job hyperlink to Db Sheet
    
    
    


    'open jobhyperlink in second tab
    If JobHyperlink <> "get" Then
        Application.SendKeys ("^t") ' Open a new tab
        Sleep 500
        Application.SendKeys ("https://www.linkedin.com" & JobHyperlink) ' Enter the URL
        Sleep 500
        Application.SendKeys ("~") ' Press Enter
        Application.Wait (Now + TimeValue("00:00:07")) ' Wait for the page to load
    Else
        
    
    End If


    ' Copy page text into string "Page_text"
    Set clipboard = New MSForms.DataObject
    clipboard.Clear
    Application.SendKeys ("^a") ' Select all text
    Sleep 500
    Application.SendKeys ("^c") ' Copy selected text to clipboard
    Sleep 500
    clipboard.GetFromClipboard
    page_text = clipboard.GetText() ' Retrieve text from clipboard
    
    Dim wsPosKeywords As Worksheet
    Dim wsAntikeywords As Worksheet
    Dim keyword As Range
    Dim voffset As Integer

    ' Define and set positive keywords worksheet
    Set wsPosKeywords = ThisWorkbook.Sheets("Keywords")
    
    ' Check Positive Keywords
    For Each keyword In wsPosKeywords.Range("A1:A" & wsPosKeywords.Cells(wsPosKeywords.Rows.Count, "A").End(xlUp).Row)
        If LCase(page_text) Like "*" & LCase(keyword.Value) & "*" Then
            GoTo Job_Candidate_Found
        End If
    Next keyword
    
    
    ' Check Anti Keywords

    Set wsAntikeywords = ThisWorkbook.Sheets("Antikeywords")
    
    ' Loop through each keyword in column A
    For Each keyword In wsAntikeywords.Range("A1:A" & wsAntikeywords.Cells(wsAntikeywords.Rows.Count, "A").End(xlUp).Row)
        If LCase(page_text) Like "*" & LCase(keyword.Value) & "*" Then
            If troubleshoot = 1 Then
                MsgBox "Antikeyword:" & keyword.Value
                Application.Wait (Now + TimeValue("00:00:01"))
                Application.SendKeys ("%{tab}")
                Application.Wait (Now + TimeValue("00:00:02"))
            
            End If
            
            'close 2nd tab
            SetCursorPos 493, 20
            Sleep 100
            Call mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0) 'click left mouse
            Sleep 75
            Call mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
            Application.Wait (Now + TimeValue("00:00:02"))
            
            ' copy page text and check for markers to know where the "X" is in the next step
            Set clipboard = New MSForms.DataObject
            clipboard.Clear
            Application.SendKeys ("^a") ' Select all text
            Sleep 500
            Application.SendKeys ("^c") ' Copy selected text to clipboard
            Sleep 500
            clipboard.GetFromClipboard
            page_text = clipboard.GetText() ' Retrieve text from clipboard
            
            ' Set voffset
            voffset = 220
            If LCase(page_text) Like "*for*you*flex*pto*education*" Then voffset = 288
            If LCase(page_text) Like "*jobs*remote*date*posted*experience*level*" Then voffset = 288
            If LCase(page_text) Like "*set*alert*" Then voffset = 288
            If LCase(page_text) Like "*[0-9] results*" Then voffset = 288
                    
            ' hit "X" / don't show me this job again.
            SetCursorPos 536, voffset
            Sleep 100
            Call mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0) 'click left mouse
            Sleep 75
            Call mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
            Application.Wait (Now + TimeValue("00:00:02"))
            
            ' refresh chrome
            Application.SendKeys ("^r")
            
            'wait
            Application.Wait (Now + TimeValue("00:00:08"))
            
            'click on job (to reset job in right window which can skew parse results later)
            SetCursorPos 451, 306
            Sleep 100
            Call mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0) 'click left mouse
            Sleep 75
            Call mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
            Application.Wait (Now + TimeValue("00:00:02"))
            
            'wait
            Application.Wait (Now + TimeValue("00:00:02"))
            
            
            'restart loop
            GoTo Get_Job_HTML:
            
            
            Exit For ' Exit loop if a keyword is found
        End If
    Next keyword
    

Job_Candidate_Found:


    ' Pause here
    For Repeat = 1 To 100
        MsgBox "Hit <Enter> to resume"
        Application.Wait (Now + TimeValue("00:00:01"))
        Application.SendKeys ("%{tab}")
        Application.Wait (Now + TimeValue("00:00:02"))
        
        ' refresh chrome
        Application.SendKeys ("^r")
        
        'wait
        Application.Wait (Now + TimeValue("00:00:08"))
        
        'restart loop
        GoTo Get_Job_HTML:
    
    Next Repeat



ClickIntoURL:
End Sub

