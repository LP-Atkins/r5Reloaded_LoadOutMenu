Attribute VB_Name = "getHandles"
'@Folder("Modules.Dll_Calls")

Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function IsWindowVisible Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetCursor Lib "user32" () As Long


'Returns a collection of all window captions and their handles in CSV format.
'Code modified from: http://www.freevbcode.com/ShowCode.asp?ID=8408
'Example Usage:
'Set var = getWindowHandles
Public Function getWindowHandles() As Collection
Dim DeskTophWnd As Long, WindowhWnd As Long
Dim Buff As String * 255
Dim colHandleCaption As New Collection
Dim strTempA, strTempB
    
    DeskTophWnd = GetDesktopWindow
    WindowhWnd = GetWindow(DeskTophWnd, 5)
    Do While (WindowhWnd <> 0)
        GetWindowText WindowhWnd, Buff, 255
        strTempA = Replace(Trim(Buff), Chr(0), "")
        strTempB = Replace(Trim(WindowhWnd), Chr(0), "")
        If strTempA <> "" And (IsWindowVisible(WindowhWnd) > False) Then
            colHandleCaption.Add strTempA & "," & strTempB
        End If
        WindowhWnd = GetWindow(WindowhWnd, 2)
        'Clear the buffer string... otherwise it will have leftovers in it..
        Buff = "                                                                                                                                                                                                                                                               "
    Loop
    
    Set getWindowHandles = colHandleCaption
    
End Function


Public Function CursorVisible() As Boolean
    If GetCursor <> 0 Then CursorVisible = True
End Function
