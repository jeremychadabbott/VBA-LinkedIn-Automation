Attribute VB_Name = "Chrome_CheckIfOPen"
Sub CheckIfChromeIsOPen()



    Dim objWMIService As Object
    Dim colProcessList As Object
    Dim objProcess As Object
    Dim chromeFound As Boolean
    
    chromeFound = False
    
    ' Connect to WMI service
    Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
    
    ' Get a list of all running processes
    Set colProcessList = objWMIService.ExecQuery("Select * from Win32_Process")
    
    ' Iterate through the list of processes
    For Each objProcess In colProcessList
        If objProcess.Name = "chrome.exe" Then
            chromeFound = True
            Exit For
        End If
    Next objProcess
    
    ' Display result
    If chromeFound Then
        ' do Nothing
    Else
        MsgBox "Google Chrome is not running."
    End If
    
    ' Clean up
    Set objWMIService = Nothing
    Set colProcessList = Nothing
    Set objProcess = Nothing
    
End Sub






