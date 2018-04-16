import argparse
import os
import openpyxl
import csv

m_descr = ''' B_Languages is a program to get dialogue from BETHESDA game(s), and convert them into different
formats, (i.e, flashcards, csv, txt, etc). The following games are currently supported:
	- Skyrim
	- Skyrim SE
	- Fallout 4
				
	Games that /may/ work:
		- Fallout NV
		- Fallout 3
		- Oblivion
	
	To see support details, bugs and more, please read the README file.
'''


# If -a is true
# Go through XLSX and remove unused audio, or audio without dialogue
# Put it in its own folder, within this projects directory


sheet_data = [];

def convertFromXLSX(inputFile, game, count):
	global sheet_data;
	wb = openpyxl.load_workbook(inputFile)
	sheet = wb[wb.sheetnames[0]]
	
	# Check if sheet cell carries the proper name (For validation reasons)
	
	responseGames = {
		'Skyrim': ['RESPONSE TEXT', 21, 7],
		'Fallout 4': ['Two', 21, 7]
	}
		
	response_text = responseGames[game][0];
	column_count = responseGames[game][1];
	quest_type = responseGames[game][2];
	ignore_types = ['GenericDialgoueWounded', 'DialogueGiant', 'VoicePowers']
	if sheet.cell(row=1, column=1).value != response_text:
		# Look for the column that carries the specified name
		for i in range (1, sheet.max_column):
			if sheet.cell(row=1, column=i).value == response_text:
				column_count = i;
	
	counted_arr = 0;
	for i in range (1, sheet.max_row + 1):
		try:
			if len(sheet.cell(row=i, column=column_count).value.strip()) != 0:		
				if sheet.cell(row=i, column=quest_type - count).value not in ignore_types:
					
					if count > 0: # if second language...
						sheet_data[counted_arr].append(sheet.cell(row=i, column=column_count).value)
						counted_arr += 1;
					else:
						sheet_data.append([sheet.cell(row=i, column=column_count).value])

		except (AttributeError, IndexError):
			pass;

def convert2(inputFile, game):
	rn = game.gamePath()
	
	input_file = inputFile
	# Example: 'Skyrim_DialogueExport_Spanish_0'
	output_file = game.gamePath() + '_DialogueExport_' + game.language + '.xlsx'
	
	if os.path.exists('dialogues/' + output_file): # If path exists already, will rename with prefix _int
		count = 0;
		choice = input('File already exists! Overwrite %s? (Y/N)' % output_file);
		if choice.lower() == 'y':
			def rename(c, output_file):
				if os.path.exists('dialogues/' + output_file + '_' + str(c)):
					c += 1
					rename(c);
				else:
					output_file = output_file + str(c)
			rename(count, output_file);
		else:
			return 'dialogues/' + output_file;
			
	wb = openpyxl.Workbook()
	ws = wb.worksheets[0]

	try:
		with open(input_file, 'r', encoding='utf-8') as data:
			reader = csv.reader(data, delimiter='\t')
			print('Writing to file...');
			try:
				for row in reader:
					ws.append(row)	
				wb.save('dialogues/' + output_file)
				return 'dialogues/' + output_file;
			except:
				print('error!');
		
			
	except FileNotFoundError:
		print('Couldnt find file at ' + input_file + '!');
		return False;
		
	
	

def support(print_games=False, game='', language=''):
	# * Chinese = Traditional Chinese
	supportedGameDict = {'Fallout 4': [['fallout 4', 'f4', 'fallout4'],
	{'English': {'audio': True}, 'French': {'audio': True}, 'Italian': {'audio': True}, 'German': {'audio': True}, 'Spanish': {'audio': True}, 'Polish': {'audio': False}, 'Portuguese-Brazil': {'audio': False}, 'Russian': {'audio': False}, 'Chinese': {'audio': False}, 'Japanese': {'audio': True}}
	], 
	'Skyrim': [['skyrim', 'skyrim se', 'skyrim special edition'], 
	{'English': {'audio': True}, 'French': {'audio': True}, 'Italian': {'audio': True}, 'German': {'audio': True}, 'Spanish': {'audio': True}, 'Polish': {'audio': True}, 'Chinese':
	{'audio': False, 'support': 'skyrim_se'}, 'Russian': {'audio': True}, 'Japanese': {'audio': True}, 'Czech': {'audio': True, 'support': 'skyrim'}}
	]}
	
	s_arr = [];
	for i in supportedGameDict:
		if print_games == True:
			print(i + ': ' + ', '.join(supportedGameDict[i][0]) + '\n')
		s_arr.append(supportedGameDict[i][0]);
		
		if len(game) > 0 and game in supportedGameDict[i][0]:
			if len(language) > 0:
				check = [];
				# Loop through language list and check if both languages are supported, 1 = True, 0 = False
				for checkLang in language:
					if checkLang in supportedGameDict[i][1]:
						check.append(1);
					else: 
						check.append(0)
				if sum(check) == 2:
					return True;
				else:
					print(language[check.index(0)]);
					return language[check.index(0)]; # Returns the non-supported language
				
			return i;
			
		
	return sum(s_arr, []); # Combines supportGameDict arrays
	
# File location, Creation Kit language,
class games:
	## Store current game, that game folders location, and the language
	def __init__(self, game='', gameFolder='', language=''):
		self.game = game
		self.gameFolder = gameFolder
		self.language = language
		
	def checkGame(self):
		languageSupport = support(game=self.game, language=self.language); # Checks language support, if string one or both langauges isn't supported
		if self.game not in support():
			print('Game not found, please check supported games!\n');
			support(True) # prints supported games, in all of their formats
			return False
		elif languageSupport and type(languageSupport) != str:
			return True;
		else:
			print('Language ' + languageSupport + ' not supported in game selected!');
			
	def gamePath(self):
		return support(game=self.game);
			
def core(game, language, type, file, audio, create=False):
	global sheet_data;
	files = file;
	game_class = games(game=game, language=language)
	# This path is currently for testing purposes
	# Look for game folder to create creationkit settings .txt file, or edit existing .txt
	if create == True:
		drive = os.path.splitdrive(os.getcwd())[0] # Assumes the drive
		print('Checking ' + drive + '\Program Files (x86)\Steam\SteamApps\common \n');
		if os.path.exists(drive + '\Program Files (x86)\Steam\SteamApps\common'):
			print('File found!', end='\r');
			gameClass = games(game=game, language=language);
			mainGame = gameClass.gamePath() # Get the games 'real' name
			special = False # Used to handle unique items (If there's more than one of the same), i.e, 'Skyrim', 'Skyrim Special Edition'
			gamePaths = {'Skyrim': 'Skyrim', 'Skyrim SE': 'Skyrim Special Edition', 'Fallout 4': 'Fallout 4'} # Keys are the filepath after SteamApps\common\
			c_gamepath = gamePaths[mainGame];
			
			if mainGame == 'Skyrim': # Since there's two versions of Skyrim, this checks based on the users previous input
				if game == 'skyrim se' or 'skyrim special edition':
					c_gamepath = gamePaths['Skyrim SE'];
					special = True;
				
			main_path = drive + '\Program Files (x86)\Steam\SteamApps\common\\' + c_gamepath;
			print('Checking ' + main_path + ' for ' + c_gamepath, end='\r\n');
			if os.path.exists(main_path):
				# Path - Skyrim SE 
				print('Found game folder!', end='\r');
				# Check for 'CreationKitCustom.ini' if special is True (This is due to the fact that Skyrim SE makes you create the file, instead of having it already like in Skyrim (vanilla)
				if not os.path.exists(main_path + '/CreationKitCustom.ini'): 
					print('Creating CreationKitCustom.ini...', end='\r');
					with open(main_path + '/CreationKitCustom.ini', "a", encoding='utf-8') as ini: # Creates the ini file with the selected language
						ini.write('[General]\nsLanguage=%s' % gameClass.language);
						
					print('Done!', end='\r');
				else:
					choice_input = input('\rCreationKitCustom.ini already exists! Rewrite? (Y/N)\n');
					if choice_input.lower() == 'y':
						with open(main_path + '/CreationKitCustom.ini', "w", encoding='utf-8') as ini:
							ini.write('[General]\nsLanguage=%s' % gameClass.language);
					else:
						return True;
						
				placeholder = input('Roadblock');
			else:
				print('Cannot find path at %s! Are you sure you entered the correct game?' % main_path);
				# add some way to input here 
				
	# Convert files and place them inside of dialogue folder
	output_files = [];
	for i in range(0, len(files)):
		g_c = games(game=game, language=language[i])
		output_files.append(convert2(files[i], g_c))
		convertFromXLSX(output_files[i], g_c.gamePath(), i);
			
	file_name = game_class.gamePath() + '_Dialogue_' + '_'.join(language)
	
	# .txt to format placeholder
	with open(file_name + '.txt', 'w', encoding='utf-8') as prc:
		for i in sheet_data:
			print(i);
			w4 = input('waiting');
			try:
				if i != ' ' or i != None:
					prc.write(';'.join(i) + '\n')
			except:
				pass
			
def main():
	parser = argparse.ArgumentParser(description=m_descr);
	#python core.py -g skyrim -l english spanish -t txt -f dialogues/export1.txt dialogues/export2.txt

	parser.add_argument('-g', '--game', nargs='+', 
		help="Source game to process")
	
	parser.add_argument('-l', '--language', nargs='+',  
		help="Language to process")
		
	parser.add_argument('-t', '--type',  
		help="Type of format [csv, txt, ")
		
	parser.add_argument('-f', '--file', nargs='+',
		help='The export file')
		
	# Audio isn't grabbed by default, arg is optional
	parser.add_argument('-a', '--audio', action='store_true',
		default=False,
		help='Grab audio *optional')
		
	# FOR TESTING PURPOSES; python core.py -c -g YOUR_GAME_NAME** 
	# parser.add_argument('-c', '--create', action='store_true',
	#	default=False,
	#	help='Creates or edits .ini file for selected game')
	
	args = parser.parse_args()
	
	# See if user stated two files, if not pass, and pass args.file false
	
	try:
		if len(args.file) < 2:
			print('Note: Only %s file(s) stated!' % len(args.file));
	except:
		args.file = False
		pass
	
	if len(args.game) > 1:
		args.game = ' '.join(args.game); # i.e, ['fallout', '4'] => 'fallout 4'
	else:
		args.game = args.game[0]
	
	
	if len(args.language) < 2:
		print('Note: Only %s language(s) stated! Please use more than one language, (i.e English Spanish)' % len(args.file));
		return False
	
	for c in range(0, len(args.language)):
		args.language[c] = args.language[c][:1].upper() + args.language[c][1:].lower() # from 'spanish', to 'Spanish'
	
	if not args.type:
		args.type = 'txt'

	gameClass = games(args.game.lower(), '', args.language);
	
	if gameClass.checkGame(): # Checks to see if 'game' is in the supported games arr
		core(game = args.game.lower(),
		language = args.language,
		type = args.type,
		file = args.file,
		audio = args.audio)
		#create = args.create)
	
if __name__ == '__main__':
	main()

