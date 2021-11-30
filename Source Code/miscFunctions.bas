Attribute VB_Name = "miscFunctions"
'@Folder("Modules.Misc")
'Contains all the misc functions/subs for use in this application.
Option Explicit



'Returns the console command for a weapon, index denotes whether it is primary or secondary.
Public Function WeaponCommand(ByVal Weapon As String, ByVal Index As Long) As String

    Select Case Index
        Case 0: WeaponCommand = "tgive p " & ResolveWeaponName(Weapon) & " " & GiveAttachments(ResolveWeaponName(Weapon))
        Case 1: WeaponCommand = "tgive s " & ResolveWeaponName(Weapon) & " " & GiveAttachments(ResolveWeaponName(Weapon))
    End Select
    
End Function


'Returns the console command for an ability, index denotes whether it is primary or secondary.
Public Function AbilityCommand(ByVal Ability As String, ByVal Index As Long) As String

    Select Case Index
        Case 0: AbilityCommand = "tgive t " & ResolveAbilityName(Ability)
        Case 1: AbilityCommand = "tgive u " & ResolveAbilityName(Ability)
    End Select

End Function


'Resolves the plain text name for an ability to the console string.
Public Function ResolveAbilityName(ByVal Ability As String) As String

    Select Case Ability
    
        Case "Smoke Grenade": ResolveAbilityName = "mp_weapon_grenade_bangalore"
        Case "Sonar Scan": ResolveAbilityName = "mp_ability_area_sonar_scan"
        Case "Beast of the Hunt": ResolveAbilityName = "mp_ability_hunt_mode"
        Case "Gas Trap": ResolveAbilityName = "mp_weapon_dirty_bomb"
        Case "Gas Grenade": ResolveAbilityName = "mp_weapon_grenade_gas"
        Case "Drop Shield": ResolveAbilityName = "mp_weapon_bubble_bunker"
        Case "Creeping Bombardment": ResolveAbilityName = "mp_weapon_grenade_creeping_bombardment"
        Case "Defensive Bombardment": ResolveAbilityName = "mp_weapon_grenade_defensive_bombardment"
        Case "Bamboozle": ResolveAbilityName = "mp_ability_holopilot"
        Case "Mirage Ult": ResolveAbilityName = "mp_ability_mirage_ultimate"
        Case "Stim Boost": ResolveAbilityName = "mp_ability_heal"
        Case "Jump Pad": ResolveAbilityName = "mp_weapon_jump_pad"
        Case "Grapple": ResolveAbilityName = "mp_ability_grapple"
        Case "Zipline": ResolveAbilityName = "mp_weapon_zipline"
        Case "Phase Walk": ResolveAbilityName = "mp_ability_phase_walk"
        Case "Phase Tunnel": ResolveAbilityName = "mp_weapon_phase_tunnel"
        Case "Tesla Fence": ResolveAbilityName = "mp_weapon_tesla_trap"
        Case "Trophy System": ResolveAbilityName = "mp_weapon_trophy_defense_system"
        Case "Doc Drone": ResolveAbilityName = "mp_weapon_deployable_medic"
        Case "Care Package": ResolveAbilityName = "mp_ability_care_package"
    
    End Select

End Function


'Resolves the plain text name for a weapon to the console string.
Public Function ResolveWeaponName(ByVal Weapon As String) As String

    Select Case Weapon
    
        Case "Alternator":  ResolveWeaponName = "mp_weapon_alternator_smg"
        Case "Charge Rifle":  ResolveWeaponName = "mp_weapon_defender"
        Case "Devotion":  ResolveWeaponName = "mp_weapon_esaw"
        Case "EVA8":  ResolveWeaponName = "mp_weapon_shotgun"
        Case "Flatline":  ResolveWeaponName = "mp_weapon_vinson"
        Case "G7 Scout":  ResolveWeaponName = "mp_weapon_g2"
        Case "Havoc":  ResolveWeaponName = "mp_weapon_energy_ar"
        Case "Hemlock":  ResolveWeaponName = "mp_weapon_hemlok"
        Case "Kraber":  ResolveWeaponName = "mp_weapon_sniper"
        Case "Longbow":  ResolveWeaponName = "mp_weapon_dmr"
        Case "Lstar":  ResolveWeaponName = "mp_weapon_lstar"
        Case "Mastiff":  ResolveWeaponName = "mp_weapon_mastiff"
        Case "Mozambique":  ResolveWeaponName = "mp_weapon_shotgun_pistol"
        Case "P2020":  ResolveWeaponName = "mp_weapon_semipistol"
        Case "Peacekeeper":  ResolveWeaponName = "mp_weapon_energy_shotgun"
        Case "Prowler":  ResolveWeaponName = "mp_weapon_pdw"
        Case "R301":  ResolveWeaponName = "mp_weapon_rspn101"
        Case "R99":  ResolveWeaponName = "mp_weapon_r97"
        Case "RE45":  ResolveWeaponName = "mp_weapon_autopistol"
        Case "Spitfire":  ResolveWeaponName = "mp_weapon_lmg"
        Case "Triple Take":  ResolveWeaponName = "mp_weapon_doubletake"
        Case "Wingman":  ResolveWeaponName = "mp_weapon_wingman"
        
    End Select

End Function


'Returns a string of the highest tier attachements for the input gun.
Public Function GiveAttachments(ByVal Gun As String) As String
Dim strTemp As String

    Select Case Gun
        
        'Light weapons
        Case "mp_weapon_alternator_smg"
                                            strTemp = strTemp & "barrel_stabilizer_l4_flash_hider" & " "
                                            strTemp = strTemp & "bullets_mag_l3" & " "
                                            strTemp = strTemp & "optic_cq_holosight_variable" & " "
                                            strTemp = strTemp & "stock_tactical_l3"
        Case "mp_weapon_g2"
                                            strTemp = strTemp & "barrel_stabilizer_l4_flash_hider" & " "
                                            strTemp = strTemp & "bullets_mag_l3" & " "
                                            strTemp = strTemp & "optic_ranged_aog_variable" & " "
                                            strTemp = strTemp & "stock_sniper_l3"
        Case "mp_weapon_semipistol"
                                            strTemp = strTemp & "bullets_mag_l3" & " "
                                            strTemp = strTemp & "optic_cq_hcog_classic"
        Case "mp_weapon_rspn101"
                                            strTemp = strTemp & "barrel_stabilizer_l4_flash_hider" & " "
                                            strTemp = strTemp & "bullets_mag_l3" & " "
                                            strTemp = strTemp & "optic_cq_holosight_variable" & " "
                                            strTemp = strTemp & "stock_tactical_l3"
                                            
        Case "mp_weapon_r97"
                                            strTemp = strTemp & "barrel_stabilizer_l4_flash_hider" & " "
                                            strTemp = strTemp & "bullets_mag_l3" & " "
                                            strTemp = strTemp & "optic_cq_holosight_variable" & " "
                                            strTemp = strTemp & "stock_tactical_l3"
        Case "mp_weapon_autopistol"
                                            strTemp = strTemp & "barrel_stabilizer_l4_flash_hider" & " "
                                            strTemp = strTemp & "bullets_mag_l3" & " "
                                            strTemp = strTemp & "optic_cq_holosight_variable"
        'Heavy Weapons
        Case "mp_weapon_vinson"
                                            strTemp = strTemp & "highcal_mag_l3" & " "
                                            strTemp = strTemp & "optic_cq_holosight_variable" & " "
                                            strTemp = strTemp & "stock_tactical_l3"
        Case "mp_weapon_hemlok"
                                            strTemp = strTemp & "highcal_mag_l3" & " "
                                            strTemp = strTemp & "optic_cq_holosight_variable" & " "
                                            strTemp = strTemp & "stock_tactical_l3" & " "
                                            strTemp = strTemp & "barrel_stabilizer_l4_flash_hider"
        Case "mp_weapon_pdw"
                                            strTemp = strTemp & "highcal_mag_l3" & " "
                                            strTemp = strTemp & "optic_cq_holosight_variable" & " "
                                            strTemp = strTemp & "stock_tactical_l3" & " "
                                            strTemp = strTemp & "hopup_selectfire"
        Case "mp_weapon_lmg"
                                            strTemp = strTemp & "highcal_mag_l3" & " "
                                            strTemp = strTemp & "optic_cq_holosight_variable" & " "
                                            strTemp = strTemp & "stock_tactical_l3" & " "
                                            strTemp = strTemp & "barrel_stabilizer_l4_flash_hider"
        Case "mp_weapon_wingman"
                                            strTemp = strTemp & "highcal_mag_l3" & " "
                                            strTemp = strTemp & "optic_cq_hcog_classic"
        Case "mp_weapon_dmr"
                                            strTemp = strTemp & "highcal_mag_l3" & " "
                                            strTemp = strTemp & "optic_sniper_variable" & " "
                                            strTemp = strTemp & "stock_sniper_l3" & " "
                                            strTemp = strTemp & "barrel_stabilizer_l4_flash_hider"
        'Energy Weapons
        Case "mp_weapon_defender"
                                            strTemp = strTemp & "optic_sniper_variable" & " "
                                            strTemp = strTemp & "stock_sniper_l3"
        Case "mp_weapon_esaw"
                                            strTemp = strTemp & "energy_mag_l3" & " "
                                            strTemp = strTemp & "optic_cq_holosight_variable" & " "
                                            strTemp = strTemp & "stock_tactical_l3" & " "
                                            strTemp = strTemp & "hopup_turbocharger" & " "
                                            strTemp = strTemp & "barrel_stabilizer_l4_flash_hider"
        Case "mp_weapon_energy_ar"
                                            strTemp = strTemp & "energy_mag_l3" & " "
                                            strTemp = strTemp & "optic_cq_holosight_variable" & " "
                                            strTemp = strTemp & "stock_tactical_l3" & " "
                                            strTemp = strTemp & "hopup_turbocharger"
        Case "mp_weapon_doubletake"
                                            strTemp = strTemp & "optic_ranged_aog_variable" & " "
                                            strTemp = strTemp & "stock_sniper_l3" & " "
                                            strTemp = strTemp & "hopup_energy_choke" & " "
                                            strTemp = strTemp & "energy_mag_l3"
        'Shotgun
        Case "mp_weapon_shotgun"
                                            strTemp = strTemp & "shotgun_bolt_l3" & " "
                                            strTemp = strTemp & "optic_cq_threat"
        Case "mp_weapon_shotgun_pistol"
                                            strTemp = strTemp & "shotgun_bolt_l3" & " "
                                            strTemp = strTemp & "optic_cq_threat"
        Case "mp_weapon_energy_shotgun"
                                            strTemp = strTemp & "shotgun_bolt_l3" & " "
                                            strTemp = strTemp & "optic_cq_threat" & " "
                                            strTemp = strTemp & "hopup_energy_choke"
        'Gold Weapons
        Case "mp_weapon_sniper"
        Case "mp_weapon_lstar"
        Case "mp_weapon_mastiff"
        
    End Select

GiveAttachments = strTemp
End Function


'Sends commands to the external console window.
Public Sub SendCommand(ByVal cmd As String)
On Error Resume Next
Dim colApps As New Collection
Dim arrTemp As Variant
Dim x As Variant
Dim hWnd
Dim WindowTitle As String
    
    
    
    'Set colApps = listApps
    Set colApps = getWindowHandles
    
    For Each x In colApps
        arrTemp = Split(x, ",")
        If arrTemp(0) <> App.Title _
            And arrTemp(0) <> "Start" _
            And arrTemp(0) <> "Program Manager" Then
            
                'Check if window title matches part of the external console window.
                If InStr(1, arrTemp(0), R5TITLE) > 0 Then
                    WindowTitle = arrTemp(0)        'Name of window.
                    hWnd = arrTemp(1)               'hWnd (not used, but could be useful later).
                End If

        End If
    Next

    
    'If the external console window was found, then send commands to it.
    If WindowTitle <> vbNullString Then
        Call TypeTextToWindow(WindowTitle, cmd)     'Write the commands to the external console.
        AppActivate (APESEX_NAME)                   'Bring the game client back to the front.
    End If
    
End Sub


'Returns True or False if a input file exists or not
'Call FileExists("C:\File.xtn")
Public Function FileExists(FilePath As String) As Boolean
    If Dir(FilePath) <> vbNullString Then FileExists = True
End Function
