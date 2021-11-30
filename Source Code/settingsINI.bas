Attribute VB_Name = "settingsINI"
'@Folder("Modules.Misc")
'Author Source:
'http://tutorpinoy.blogspot.com/2014/03/save-and-retrieve-configuration-files-vb6.html

Option Explicit

'API Function to read information from INI File
Public Declare Function GetPrivateProfileString Lib "kernel32" _
    Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, _
    ByVal lpKeyName As Any, _
    ByVal lpDefault As String, _
    ByVal lpReturnedString As String, _
    ByVal nSize As Long, _
    ByVal lpFileName As String) As Long

'API Function to write information to the INI File
Private Declare Function WritePrivateProfileString Lib "kernel32" _
    Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, _
    ByVal lpKeyName As Any, _
    ByVal lpString As Any, _
    ByVal lpFileName As String) As Long

'Get the INI Setting from the File
Public Function GetINISetting(ByVal sHeading As String, _
    ByVal sKey As String, _
    sINIFileName) As String
    
    Const cparmLen = 500    '<- LongParsnip: "Changed from 50 to 500 because I needed more characters."
    Dim sReturn As String * cparmLen
    Dim sDefault As String * cparmLen
    Dim lLength As Long
    lLength = GetPrivateProfileString(sHeading, _
            sKey, _
            sDefault, _
            sReturn, _
            cparmLen, _
            sINIFileName)
            
    GetINISetting = Mid(sReturn, 1, lLength)
End Function

'Save INI Setting in the File
Public Function PutINISetting(ByVal sHeading As String, _
    ByVal sKey As String, _
    ByVal sSetting As String, _
    sINIFileName) As Boolean
    
    Const cparmLen = 500
    Dim sReturn As String * cparmLen
    Dim sDefault As String * cparmLen
    Dim aLength As Long
    aLength = WritePrivateProfileString(sHeading, _
        sKey, _
        sSetting, _
        sINIFileName)
        
    PutINISetting = True
End Function
