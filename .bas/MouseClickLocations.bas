Attribute VB_Name = "MouseClickLocations"
' A NOTE ABOUT MOUSE AUTOMATION
' Mouse click targetting is especially weak code-wise, but works when you don't want to interface with an
' API. the reason it's weak is the resolution of the screen can affect the target so you need to keep the
' resolution set. This was programmed for 1280x720
' another reason it is week is the user can interfere with the code by simply joggling the mouse or hitting the
' keyboard.
' the advantage to this code is almost everyone has excel, VBA is easy to write and use
' Also it doesn't rely on google developer so google chrome does not detect it as automation (for a long time)
' in the end, programs I write with mouse targetting built in often require "babysitting"
' I like to let it run on a second screen and keep an eye on it just in case it runs off track.

Sub Mouse_Click_Chrome_Inspector_Close()

'use Mouse Click to close inspector so we don't accidentally OPEN inspector with hot keys
    SetCursorPos 1285, 100
    Sleep 100
    Call mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0) 'click left mouse
    Sleep 75
    Call mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
    Application.Wait (Now + TimeValue("00:00:03"))


End Sub


Sub Mouse_Right_Click_Inspector_Root_Element()
     
        SetCursorPos 1200, 141
        Sleep 100
        Call mouse_event(MOUSEEVENTF_RIGHTDOWN, 0, 0, 0, 0) 'click left mouse
        Sleep 250
        Call mouse_event(MOUSEEVENTF_RIGHTUP, 0, 0, 0, 0)
        Application.Wait (Now + TimeValue("00:00:01"))
    

End Sub
Sub Mouse_Right_Click_Inspector_Root_Element_CopyOUterHTML()
        
        ' Hover mouse over copy until second menu appears
        SetCursorPos 1075, 349
         Application.Wait (Now + TimeValue("00:00:02"))
        
        'click copy outer HTML
        SetCursorPos 901, 397
        Sleep 100
        Call mouse_event(MOUSEEVENTF_RIGHTDOWN, 0, 0, 0, 0) 'click left mouse
        Sleep 250
        Call mouse_event(MOUSEEVENTF_RIGHTUP, 0, 0, 0, 0)
        Application.Wait (Now + TimeValue("00:00:01"))
    

End Sub

