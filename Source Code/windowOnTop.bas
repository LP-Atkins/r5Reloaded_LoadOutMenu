Attribute VB_Name = "windowOnTop"
'@Folder("Modules.Dll_Calls")
Declare Sub SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Global Const SWP_NOMOVE = 2
Global Const SWP_NOSIZE = 1
Global Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
Global Const HWND_TOPMOST = -1
Global Const HWND_NOTOPMOST = -2

Public Sub FormOnBottom(frm As Form)
    SetWindowPos frm.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, FLAGS
End Sub

Public Sub FormOnTop(frm As Form)
    SetWindowPos frm.hWnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS
End Sub
