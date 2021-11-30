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
	
## Bugs & Help My Antivirus is blocking this!:
	Your Antivirus may falsly flag this application as a virus. I am not sure why, and have submitted a false flag ticket to Microsoft.
	In the mean time if you wish to download the .exe file you will need to temporarily disable your antivirus, then download the file, turn your antivirus back on and then in the action centre allow the 'threat'.

	If you are concerned, please review the code and compile for yourself.
	
	If you encounter any bugs (app crashes) report them here, please note that r5Reloaded is a little buggy and some of the commands don't work properly / some guns don't spawn with ammo and I cannot do anything about this.
