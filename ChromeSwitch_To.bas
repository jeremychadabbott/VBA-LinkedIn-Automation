Attribute VB_Name = "ChromeSwitch_To"

Sub SwitchToChrome()


    Dim oApp As Object
    For Each oApp In GetObject("winmgmts:").InstancesOf("Win32_Process")
        If oApp.Name Like "*chrome*" Then
            'MsgBox oApp.Name
            AppActivate "Google chrome"
            Exit For
        End If
        
    Next oApp
End Sub
