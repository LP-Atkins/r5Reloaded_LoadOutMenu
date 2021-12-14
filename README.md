# r5-Reloaded Apex Legends, Loadout Menu, v1.0.1
# v1.0.1 will release here once it has been cleard by Microsoft Antivirus Team, for now the source code is available to be compiled.
A small application that can allow you to swap loadouts in modded Apex Legends.
This application has only been tested on Windows 10 & 11.

Written in Visual Basic Classic (VB6).

## Version Summary:
* v1.0.0: Original release.
* v1.0.1:
* 	-Window position is now saved.
* 	-Custom Abilities and Weapons can now be loaded through the INI file.

## Features:
* Allows you to select all available guns for primary & secondary.
* Allows you to select all available abilities for tactical & ultimate.
* Favourites list & menu so that you can quickly setup and deploy classes (saved to an ini file in the same folder as the executable).

## How this works?:
I used the r5 documentation to for the tgive command as well as the game names for the various weapons/attachments/abilities.
When you deploy a loadout this application attempts to locate the external console window for r5Reloaded and types the commands directly into it.

## Instructions:
	1.) Set game to borderless window.
	2.) Launch this app and move it to top right hand corner.
	3.) Open the inventory (tab) when you want to deploy a loadout, as you this will show the mouse cursor.
	4.) Select your loadout and click on the Deploy Loadout button.
	
	You can also save the current loadout (as per the comboboxes) to a favourite, make sure to select the favourite index before doing so.
	To access the favourites menu simply click on the 'Open Fav Menu' button.
	
	**NOTE** This will only work if tgive command is enabled.

## How do I add custom weapons & abilties?
	1.) Once you launch the program an INI file should be generated, (if it doesn't contain the settings click the 'Save Fav' button to force it to.
	2.) Under the "Custom Weapon" and "Custom Ability" Heading in the INI file there should be a few keys that have been generated with instructions.
	
	Should look like this:
		weapCmd_0	<- The 0 in this case is the index, 0-99 are supported.
		weapName_0	<- Both name and command must be valid or it wont be shown.
		
	Example:
	
	[Custom Weapon]
	weapCmd_0=mp_weapon_bfg9000 optic_theopticyouwant mag_themagyouwant attachment_etc
	weapName_0=BFG9000
	[Custom Ability]
	abilityCmd_0=mp_ability_iddqd
	abilityName_0=Degreelessness mode


## The window wont pop up anywhere but the program is running?
	It may have saved a bad location for the x and y position, if this happens close the program and edit the INI file and remove the xPos and yPos keys.
	
# Help My Antivirus is blocking this!:
	Windows AV (machine learning algorithm) originally falsly flagged this, I submitted the exe to Microsoft
	and they have removed the false positive.
	If you are still having trouble with windows AV follow the steps provide by Microsoft:
	
	1. Open command prompt as administrator and change directory to c:\Program Files\Windows Defender
	2. Run “MpCmdRun.exe -removedefinitions -dynamicsignatures”
	3. Run "MpCmdRun.exe -SignatureUpdate"
	
	Alternatively you may want to instead add an exception for this program in your preferred AV software.
	If you are concerned, please review the code and compile for yourself.
	
# Bugs
	If you encounter any bugs (app crashes) report them here, please note that r5Reloaded is a little
	buggy and some of the commands don't work properly / some guns don't spawn with ammo and I cannot do anything about this.
