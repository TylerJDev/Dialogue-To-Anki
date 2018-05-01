# Dialogue-To-Anki
This is a small python project to process dialogue from games (I.e, Skyrim, Fallout 4) to Anki friendly formats

#### Requirements ($ = optional)
- B.A.E $ *(Needed if you want audio alongside of text)*
- Unfuz $ *(Needed if you want audio alongside of text)*
- TES5Edit $ *(Needed if language contains special characters (non-latin) i.e, Russian, Japanese, Chinese..))*
- A copy of one of the supported games
- Creation Kit **(Only for Skyrim/Skyrim SE & Fallout 4)**
- G.E.C.K **(Only for Fallout New Vegas)**

As above, I mentioned TES5Edit because I had problems exporting dialogue with languages with non-latin alphabets, though they may work for you, so test before you grab TES5Edit.

# Instructions
## Step 1
Go to your steamapps common folder (I.E, Program Files (x86)\Steam\SteamApps\common\GAME_NAME)

Next you must edit, or create specified file for the selected game below
- **SKYRIM SPECIAL EDITION**
  - Since with the launch of Skyrim SE, Bethesda handled the creation kit differently. You will have to create a file named ```'CreationKitCustom.ini'``` and within add the contents 
  
```'[General]
 sLanguage=PUT SUPPORTED LANGUAGE HERE (i.e, spanish, russian, etc)
```

Replace 'SUPPORTED LANGUAGE HERE' with the language that's supported for your selected game *(see below which languages are supported in each game)*

Now open up Creation Kit or G.E.C.K and load in your main .esm file by going to File > Data > and select the main .esm file
![Image of prompt](https://github.com/TylerJCodes/Dialogue-To-Anki/blob/master/Images/CreationKit_Data.PNG?raw=true)

After it's done loading, proceed to Character > Export Dialogue
Once the dialogue has been exported, take it and move it (preferably to Dialogue-To-Anki/dialogues), be sure to rename it!
Rinse and repeat, do step one over but change the language to the second language you want to pair up. (*I.e, Spanish > English*)

## Step 2
**Skip this step if you don't want audio, or if you already have done it previously**

Just like step 1, you'll have to repeat this twice if you want audio for both languages.
Download [B.A.E](https://www.nexusmods.com/fallout4/mods/78/);

Once installed open it and you should see a screen like this:
![Image of B.A.E](https://github.com/TylerJCodes/Dialogue-To-Anki/blob/master/Images/BAE.PNG?raw=true)

If you haven't already, change your games language in steam by right clicking it and selecting properties and click the 'language' tab and set the language you want to get the audio from
[ADD IMAGE HERE]
 

After its done downloading go to your games folder (I.E, Program Files (x86)\Steam\SteamApps\common\GAME_NAME) and proceed to the Data folder, within you should see .bsa files, locate the 'Voices' .bsa file. Usually it'll have the languages prefix tagged along to it i.e, ('Skyrim- Voices_ru0.bsa').

Open that file up in bae and extract it.
*Note, this may take a while. If you want you can uncheck characters you don't want to extract in bae*

After it's done extracting, place that folder within the dialogues/audio directory, and rename 'sound' to the language you downloaded. (*Be sure to name it correctly!*)

## Step 3
Before we unfuz the files, we must run this script;
```python core.py -g skyrim -l english spanish -t txt -f dialogues/export1.txt dialogues/export2.txt -a```
[ADD PHOTO HERE]

-g = The games title
-l = Languages, be sure to put these in the same order as the -f argument!
-t = Format to give out [Supported formats; txt, csv]
-f = The dialogue export .txt files, like -l put these in the same order as the languages (i.e, '-l english spanish -f export_english.txt export_spanish.txt)
-a = *Optional, if you skipped step 2 don't add this* Processes audio files don't put anything after it as it will result to 'True' if it's tagged along

After the script is done, you should find the result in the main folder 'Dialogue-To-Anki' in whichever format you wanted, if it's a .txt you can import it into Anki (see below).

If you wanted audio, you'll find a folder name 'GAMENAME_AUDIO_LANGUAGE1_LANGUAGE2' (LANGUAGE1 & 2 = your languages respectively),
Open up unfuzer and select this folder to Unfuz,
[ADD PHOTO HERE]

After it's done you're all set! You can now place the audio within your Anki collection.media. Then you can import the .txt file into Anki
[ADD PHOTO HERE]

### Games Supported
- [x] Skyrim Special Edition
- [x] Skyrim (Vanilla)
- [x] Fallout 4
- [x] Fallout New Vegas

### Unsupported (May work)
 - Fallout 3

More supported games coming soon!

### Formats Supported
Formats that are currently supported or in progress
- [x] .txt
- [x] .csv
- [ ] .anki

## Python Requirements
- openpyxl

pip install -r requirements.txt
