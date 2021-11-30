Attribute VB_Name = "typeText"
'@Folder("Modules.Misc")

'Note: I tried SendMessageA function first, but Windows kept falsly flagging the compiled program as a virus.
Public Sub TypeTextToWindow(ByVal WindowTitle As String, ByVal Text As String)
On Error Resume Next
Dim objShell: Set objShell = CreateObject("Wscript.Shell")
    
    'Window is selected, activate it and type the text.
    AppActivate WindowTitle
    objShell.SendKeys Text
    objShell.SendKeys "{ENTER}"

End Sub
