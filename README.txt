## Instructions ##
Allow user to pick two languages (i.e, english > < spanish)
* Optional Path
	SKYRIM SE:
		Create CreationKitCustom.ini with language, note this will need to be done for both languages!
		Example: 
			[General]
			sLanguage=english
	SKYRIM:
		Edit SkyrimEditor.ini, or create (NEEDS TESTING)

USER must open creation kit, with their master.esm,
After they must export dialogue, THEN the script can go back on track


* Optional Path
	Script waits for export file to continue,
	Export file will most likely always be within the skyrim folder, so wait in that folder for changes
	* NOTE, must wait until the txt file is COMPLETED writing, as it takes time to creation kit to write fully to the .txt file
	Continue with script
	
Once ExportDialogue file is had (This requires both language exports!), run the script against that,
.txt file will be edited as an excel file


## Notes ##
Skyrim versions MAY have different language supports,
	Skyrim (vanilla):
		English			
		French			
		German			
		Italian			
		Spanish			
		Japanese			
		Czech		-	
		Polish			
		Russian		
	
	Skyrim SE:	
		English			
		French			
		Italian			
		German			
		Spanish			
		Polish			
		Traditional Chinese	-		
		Russian			
		Japanese			
		
## Tests ## add at 3 per
Test finding steamapps folder,
Checking if language is supported for game,
txt to xlsx
xlsx to ALL formats available

## To Do ##
Add openpyxl to requirements
Add where it will check 'dialogue' folder first for txt files, as well as the current __location__ b4 checking actual path