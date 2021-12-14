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
'v1.0.1     14/12/2021      -Added custom weapons/abilities that can be added to the INI file.
'                           -Screen position is now saved.
'                           -Made some adjustments to the poor INI implementation, added subs to update existing INIs to new format.

Option Explicit

'Hard Coded Constants.
'--------------------------------
Public Const R5TITLE As String = "R5pc_r5launch"        'Partial string match for the R5 console window. This is used to locate the external console window.
Public Const R5INI As String = "R5_Favs.ini"            'Name of the ini file to store the favourites in.
Public Const DEF_PRIMARY As String = "R99"          'Default Primary.
Public Const DEF_SECONDARY As String = "Wingman"    'Default Seconday.
Public Const DEF_TACTICAL As String = "Grapple"     'Defult Tactical.
Public Const DEF_ULTIMATE As String = "Phase Walk"  'Default Ultimate.


'Messages
'--------------------------------
Public Const CUSTOM_WEAPON As String = "Support for 100 customs, change index to add (0-99). Enter the weapon string as the value."
Public Const CUSTOM_ABILITY As String = "Support for 100 customs, change index to add (0-99). Enter the ability string as the value."
Public Const CUSTOM_WEAPON_NAME As String = "Name goes here"
Public Const CUSTOM_ABILITY_NAME As String = "Name goes here"


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
Public colCustomWeapons As New Collection
Public colCustomAbilities As New Collection
Public xPos As Long         'Window X, for position retention.
Public yPos As Long
Public xyPos As Boolean     'If True, means that the settings were found in the INI.

'User Defined Types
'--------------------------------
Public Type customWeapon
    CmdString As String
    Name As String
End Type


Public Sub Main()

    'Check if application is already running.
    If App.PrevInstance Then
        AppActivate App.Title
    Else
        Call InitVars
        Call ReadINI
        Call ReadCustomItems
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
        FavDesc0 = GetINISetting("LoudOut1", "FavDesc", INI_PATH)
        FavDesc1 = GetINISetting("LoudOut2", "FavDesc", INI_PATH)
        FavDesc2 = GetINISetting("LoudOut3", "FavDesc", INI_PATH)
        FavDesc3 = GetINISetting("LoudOut4", "FavDesc", INI_PATH)
        FavDesc4 = GetINISetting("LoudOut5", "FavDesc", INI_PATH)
        FavDesc5 = GetINISetting("LoudOut6", "FavDesc", INI_PATH)
        
        'Commands.
        FavCmd0 = GetINISetting("LoudOut1", "FavCmd", INI_PATH)
        FavCmd1 = GetINISetting("LoudOut2", "FavCmd", INI_PATH)
        FavCmd2 = GetINISetting("LoudOut3", "FavCmd", INI_PATH)
        FavCmd3 = GetINISetting("LoudOut4", "FavCmd", INI_PATH)
        FavCmd4 = GetINISetting("LoudOut5", "FavCmd", INI_PATH)
        FavCmd5 = GetINISetting("LoudOut6", "FavCmd", INI_PATH)
        
        'LegacyRead (read once to convert to newer format).
        If GetINISetting("Settings", "DontReadLegacy", INI_PATH) = vbNullString Then
            Call PutINISetting("Settings", "DontReadLegacy", True, INI_PATH)
            Call LegacyReadINI
            Call LegacyWriteINI
            Call WriteINI   'Save settings back.
        End If
        
        'Get last window positions, if they exist.
        If GetINISetting("Settings", "xPos", INI_PATH) <> vbNullString Then xPos = CLng(GetINISetting("Settings", "xPos", INI_PATH))
        If GetINISetting("Settings", "yPos", INI_PATH) <> vbNullString Then yPos = CLng(GetINISetting("Settings", "yPos", INI_PATH))
        If GetINISetting("Settings", "xPos", INI_PATH) <> vbNullString And GetINISetting("Settings", "yPos", INI_PATH) <> vbNullString Then xyPos = True
    
    End If

End Sub


'For reading original implementation.
Private Sub LegacyReadINI()
    
    'Descriptions.
    If GetINISetting("LoudOut1", "FavDesc0", INI_PATH) <> vbNullString Then FavDesc0 = GetINISetting("LoudOut1", "FavDesc0", INI_PATH)
    If GetINISetting("LoudOut2", "FavDesc1", INI_PATH) <> vbNullString Then FavDesc1 = GetINISetting("LoudOut2", "FavDesc1", INI_PATH)
    If GetINISetting("LoudOut3", "FavDesc2", INI_PATH) <> vbNullString Then FavDesc2 = GetINISetting("LoudOut3", "FavDesc2", INI_PATH)
    If GetINISetting("LoudOut4", "FavDesc3", INI_PATH) <> vbNullString Then FavDesc3 = GetINISetting("LoudOut4", "FavDesc3", INI_PATH)
    If GetINISetting("LoudOut5", "FavDesc4", INI_PATH) <> vbNullString Then FavDesc4 = GetINISetting("LoudOut5", "FavDesc4", INI_PATH)
    If GetINISetting("LoudOut6", "FavDesc5", INI_PATH) <> vbNullString Then FavDesc5 = GetINISetting("LoudOut6", "FavDesc5", INI_PATH)
    
    'Commands.
    If GetINISetting("LoudOut1", "FavCmd0", INI_PATH) <> vbNullString Then FavCmd0 = GetINISetting("LoudOut1", "FavCmd0", INI_PATH)
    If GetINISetting("LoudOut2", "FavCmd1", INI_PATH) <> vbNullString Then FavCmd1 = GetINISetting("LoudOut2", "FavCmd1", INI_PATH)
    If GetINISetting("LoudOut3", "FavCmd2", INI_PATH) <> vbNullString Then FavCmd2 = GetINISetting("LoudOut3", "FavCmd2", INI_PATH)
    If GetINISetting("LoudOut4", "FavCmd3", INI_PATH) <> vbNullString Then FavCmd3 = GetINISetting("LoudOut4", "FavCmd3", INI_PATH)
    If GetINISetting("LoudOut5", "FavCmd4", INI_PATH) <> vbNullString Then FavCmd4 = GetINISetting("LoudOut5", "FavCmd4", INI_PATH)
    If GetINISetting("LoudOut6", "FavCmd5", INI_PATH) <> vbNullString Then FavCmd5 = GetINISetting("LoudOut6", "FavCmd5", INI_PATH)
    
End Sub



'Writes variables to the INI File.
Public Sub WriteINI()

    'Descriptions.
    Call PutINISetting("LoudOut1", "FavDesc", FavDesc0, INI_PATH)
    Call PutINISetting("LoudOut2", "FavDesc", FavDesc1, INI_PATH)
    Call PutINISetting("LoudOut3", "FavDesc", FavDesc2, INI_PATH)
    Call PutINISetting("LoudOut4", "FavDesc", FavDesc3, INI_PATH)
    Call PutINISetting("LoudOut5", "FavDesc", FavDesc4, INI_PATH)
    Call PutINISetting("LoudOut6", "FavDesc", FavDesc5, INI_PATH)
    
    'Commands
    Call PutINISetting("LoudOut1", "FavCmd", FavCmd0, INI_PATH)
    Call PutINISetting("LoudOut2", "FavCmd", FavCmd1, INI_PATH)
    Call PutINISetting("LoudOut3", "FavCmd", FavCmd2, INI_PATH)
    Call PutINISetting("LoudOut4", "FavCmd", FavCmd3, INI_PATH)
    Call PutINISetting("LoudOut5", "FavCmd", FavCmd4, INI_PATH)
    Call PutINISetting("LoudOut6", "FavCmd", FavCmd5, INI_PATH)
    
    'Custom Weapons
    If GetINISetting("Custom Weapon", "weapCmd_0", INI_PATH) = vbNullString Or GetINISetting("Custom Weapon", "weapName_0", INI_PATH) = vbNullString Then
        Call PutINISetting("Custom Weapon", "weapCmd_0", CUSTOM_WEAPON, INI_PATH)
        Call PutINISetting("Custom Weapon", "weapName_0", CUSTOM_WEAPON_NAME, INI_PATH)
    End If
    
    'Custom Abilities
    If GetINISetting("Custom Ability", "abilityCmd_0", INI_PATH) = vbNullString Or GetINISetting("Custom Ability", "abilityName_0", INI_PATH) = vbNullString Then
        Call PutINISetting("Custom Ability", "abilityCmd_0", CUSTOM_ABILITY, INI_PATH)
        Call PutINISetting("Custom Ability", "abilityName_0", CUSTOM_ABILITY_NAME, INI_PATH)
    End If

End Sub


'Clears settings for the legacy INI.
Private Sub LegacyWriteINI()

    'Descriptions.
    Call PutINISetting("LoudOut1", "FavDesc0", vbNullString, INI_PATH)
    Call PutINISetting("LoudOut2", "FavDesc1", vbNullString, INI_PATH)
    Call PutINISetting("LoudOut3", "FavDesc2", vbNullString, INI_PATH)
    Call PutINISetting("LoudOut4", "FavDesc3", vbNullString, INI_PATH)
    Call PutINISetting("LoudOut5", "FavDesc4", vbNullString, INI_PATH)
    Call PutINISetting("LoudOut6", "FavDesc5", vbNullString, INI_PATH)
    
    'Commands
    Call PutINISetting("LoudOut1", "FavCmd0", vbNullString, INI_PATH)
    Call PutINISetting("LoudOut2", "FavCmd1", vbNullString, INI_PATH)
    Call PutINISetting("LoudOut3", "FavCmd2", vbNullString, INI_PATH)
    Call PutINISetting("LoudOut4", "FavCmd3", vbNullString, INI_PATH)
    Call PutINISetting("LoudOut5", "FavCmd4", vbNullString, INI_PATH)
    Call PutINISetting("LoudOut6", "FavCmd5", vbNullString, INI_PATH)

End Sub



'Reads all custom items from the INI file into weapon and ability collections.
Private Sub ReadCustomItems()
Dim i As Long
Dim WeapOrAbil As clsCustomWeaponAbility

    'Loop for custom weapons.
    For i = 0 To 99
    
        'Check that there is a value and not the default.
        If GetINISetting("Custom Weapon", "weapCmd_" & i, INI_PATH) <> vbNullString And GetINISetting("Custom Weapon", "weapCmd_" & i, INI_PATH) <> CUSTOM_WEAPON Then
            
            'Check there is a name and not the default.
            If GetINISetting("Custom Weapon", "weapName_" & i, INI_PATH) <> vbNullString And GetINISetting("Custom Weapon", "weapName_" & i, INI_PATH) <> CUSTOM_WEAPON_NAME Then
                Set WeapOrAbil = New clsCustomWeaponAbility
                WeapOrAbil.Command = GetINISetting("Custom Weapon", "weapCmd_" & i, INI_PATH)
                WeapOrAbil.Name = GetINISetting("Custom Weapon", "weapName_" & i, INI_PATH)
                colCustomWeapons.Add Key:=WeapOrAbil.Name, Item:=WeapOrAbil    'Add item to the collection.
            End If
            
        End If
    
    Next i
    
    
    
    'Loop for custom abilities.
    For i = 0 To 99
    
        'Check that there is a value and not the default.
        If GetINISetting("Custom Ability", "abilityCmd_" & i, INI_PATH) <> vbNullString And GetINISetting("Custom Ability", "abilityCmd_" & i, INI_PATH) <> CUSTOM_ABILITY Then
            
            'Check there is a name and not the default.
            If GetINISetting("Custom Ability", "abilityName_" & i, INI_PATH) <> vbNullString And GetINISetting("Custom Ability", "abilityName_" & i, INI_PATH) <> CUSTOM_ABILITY_NAME Then
                Set WeapOrAbil = New clsCustomWeaponAbility
                WeapOrAbil.Command = GetINISetting("Custom Ability", "abilityCmd_" & i, INI_PATH)
                WeapOrAbil.Name = GetINISetting("Custom Ability", "abilityName_" & i, INI_PATH)
                colCustomAbilities.Add Key:=WeapOrAbil.Name, Item:=WeapOrAbil    'Add item to the collection.
            End If
            
        End If
    
    Next i


End Sub




