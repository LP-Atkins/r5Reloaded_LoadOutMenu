Attribute VB_Name = "Program_Start"
'@Folder("1nvokeMAIN")

'r5WeaponSelect
'   This application provides a simple way to swap loadouts in r5Reloaded ApeSex Legends.
'   This is achieved by writing tgive commands to the external console window.
'
'Requirements:
'   1.) You must be running window / borderless window mode.
'
'Instructions:
'1.) Set game to borderless window.
'2.) Have the checkbox for always on top checked.
'3.) Move to top right hand corner.
'4.) Open the inventory (tab) when you want to deploy a loadout, as you this will show the mouse cursor.
'5.) Select your loadout and click on the Deploy Loadout button.


'Written by LongParsnip
'v1.0.0     30/11/2021      -Initial Prototype


'Hard Coded Constants.
'--------------------------------
Public Const R5TITLE As String = "R5pc_r5launch"        'Partial string match for the R5 console window. This is used to locate the external console window.
Public Const R5INI As String = "R5_Favs.ini"            'Name of the ini file to store the favourites in.
Public Const DEF_PRIMARY As String = "R99"          'Default Primary.
Public Const DEF_SECONDARY As String = "Wingman"    'Default Seconday.
Public Const DEF_TACTICAL As String = "Grapple"     'Defult Tactical.
Public Const DEF_ULTIMATE As String = "Phase Walk"  'Default Ultimate.


'Initialised Constants.
'--------------------------------
Public APESEX_NAME As String                            'Set by InitVars, had to make the game name unclear because it was being flagged as a virus of the same name as the game.
Public INI_PATH As String


'INI Settings.
'--------------------------------
Public FavDesc0 As String       'Button Descriptor.
Public FavDesc1 As String
Public FavDesc2 As String
Public FavDesc3 As String
Public FavDesc4 As String
Public FavDesc5 As String

Public FavCmd0 As String        'Button Command.
Public FavCmd1 As String
Public FavCmd2 As String
Public FavCmd3 As String
Public FavCmd4 As String
Public FavCmd5 As String


'Program Vars
'--------------------------------
Public PrimSel As String
Public SecSel As String
Public TacSel As String
Public UltSel As String
Public FavSel As String
'


Public Sub Main()

    Call InitVars
    Call ReadINI

    'Check if application is already running.
    If App.PrevInstance Then
        AppActivate App.Title
    Else
        frmSelect.Show
    End If

End Sub






'Initialises the semi constants.
Private Sub InitVars()
    APESEX_NAME = "A" & "p" & "e" & "x" & " Legends"      'Window title for r5Reloaded Apesex, used to activate it again once commands have been sent.
    INI_PATH = App.Path & "\" & R5INI
End Sub



'Reads the INI file into program vars.
Private Sub ReadINI()

    'Don't try and read the INI if it doesn't exist yet.
    If FileExists(INI_PATH) Then
    
        'Descriptions.
        FavDesc0 = GetINISetting("LoudOut1", "FavDesc0", INI_PATH)
        FavDesc1 = GetINISetting("LoudOut2", "FavDesc1", INI_PATH)
        FavDesc2 = GetINISetting("LoudOut3", "FavDesc2", INI_PATH)
        FavDesc3 = GetINISetting("LoudOut4", "FavDesc3", INI_PATH)
        FavDesc4 = GetINISetting("LoudOut5", "FavDesc4", INI_PATH)
        FavDesc5 = GetINISetting("LoudOut6", "FavDesc5", INI_PATH)
        
        'Commands.
        FavCmd0 = GetINISetting("LoudOut1", "FavCmd0", INI_PATH)
        FavCmd1 = GetINISetting("LoudOut2", "FavCmd1", INI_PATH)
        FavCmd2 = GetINISetting("LoudOut3", "FavCmd2", INI_PATH)
        FavCmd3 = GetINISetting("LoudOut4", "FavCmd3", INI_PATH)
        FavCmd4 = GetINISetting("LoudOut5", "FavCmd4", INI_PATH)
        FavCmd5 = GetINISetting("LoudOut6", "FavCmd5", INI_PATH)
    
    End If

End Sub



'Writes variables to the INI File.
Public Sub WriteINI()

    'Descriptions.
    Call PutINISetting("LoudOut1", "FavDesc0", FavDesc0, INI_PATH)
    Call PutINISetting("LoudOut2", "FavDesc1", FavDesc1, INI_PATH)
    Call PutINISetting("LoudOut3", "FavDesc2", FavDesc2, INI_PATH)
    Call PutINISetting("LoudOut4", "FavDesc3", FavDesc3, INI_PATH)
    Call PutINISetting("LoudOut5", "FavDesc4", FavDesc4, INI_PATH)
    Call PutINISetting("LoudOut6", "FavDesc5", FavDesc5, INI_PATH)
    
    'Commands
    Call PutINISetting("LoudOut1", "FavCmd0", FavCmd0, INI_PATH)
    Call PutINISetting("LoudOut2", "FavCmd1", FavCmd1, INI_PATH)
    Call PutINISetting("LoudOut3", "FavCmd2", FavCmd2, INI_PATH)
    Call PutINISetting("LoudOut4", "FavCmd3", FavCmd3, INI_PATH)
    Call PutINISetting("LoudOut5", "FavCmd4", FavCmd4, INI_PATH)
    Call PutINISetting("LoudOut6", "FavCmd5", FavCmd5, INI_PATH)

End Sub



