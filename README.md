# r5-Reloaded Apex Legends, Loadout Menu, v1.00
A small application that can allow you to swap loadouts in modded Apex Legends.
This application has only been tested on Windows 10 & 11.

Written in Visual Basic Classic (VB6).

## Version Summary:
* v1.00: Original release.

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
	
# Help My Antivirus is blocking this!:
	Windows AV (machine learning algorithm) originally falsly flagged this, I submitted the exe to Microsoft and they have removed the false positive.
	If you are still having trouble with windows AV follow the steps provide by Microsoft:
	
	1. Open command prompt as administrator and change directory to c:\Program Files\Windows Defender
	2. Run “MpCmdRun.exe -removedefinitions -dynamicsignatures”
	3. Run "MpCmdRun.exe -SignatureUpdate"
	
	Alternatively you may want to instead add an exception for this program in your preferred AV software.
	If you are concerned, please review the code and compile for yourself.
	
# Bugs
	If you encounter any bugs (app crashes) report them here, please note that r5Reloaded is a little buggy and some of the commands don't work properly / some guns don't spawn with ammo and I cannot do anything about this.
