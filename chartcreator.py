import os
import fnmatch
import numpy as np
import pandas
import matplotlib.pyplot as plt
import sqlite3

# Create SQLite database
def initial_DB_connect(db_file, schemaDataStorage_file, schemaResults_file):
	if(os.path.exists(db_file)):
		print("There is already a database existing with values. Using it ...")
		
	else:
		print("No database found. Creating one ...")
		
		# Establish a connection to the database
		connection = sqlite3.connect(db_file)

		# Create a cursor to the Database
		crsr = connection.cursor()
			
		# Create a connection object
		connection.executescript(open(schemaDataStorage_file, 'r').read())
		connection.executescript(open(schemaResults_file, 'r').read())
			
		# Close connection to database
		connection.close()

# Method to iterate through all files in the subdirectory of this scripts folder
# and save the data from excel sheets coressponding in a SQL database.
def read_data_to_sql(db_file):
	# Get the current working directory (cwd) and current data directory (cdd)
	cwd = os.getcwd()
	cdd = cwd + '/data'

	roomName = ''

	for file in os.listdir(cdd):
		if file.endswith(".ods"):	
			print("Contains data from file:", 'data/' + file)
			df = pandas.read_excel('data/' + file, sheet_name=0)		
			
			df = df.applymap(str)
			if fnmatch.fnmatch(file, 'Infobib_*'):
				print("Processing a Infobib data file")
				roomName = "Infobib"
				
			elif fnmatch.fnmatch(file, 'KleinerLernraum_*'):
		   		print("Processing a KleinerLernraum data file")
		   		roomName = "KleinerLernraum"
		   		
			elif fnmatch.fnmatch(file, 'Einzelarbeitsraum_*'):
		   		print("Processing a Einzelarbeitsraum data file")
		   		roomName = "Einzelarbeitsraum"
		   		
		   	# Insert data from excel files into db
			connection = sqlite3.connect(db_file)
			crsr = connection.cursor()
			
			for x in range(df.shape[0]):	
				crsr.execute("INSERT INTO studyroomUsage VALUES ( ?, ?, ?, ?, ?)", (roomName, str(df._get_value(x, 'KW')), df._get_value(x, 'Termin'), df._get_value(x, 'Wochentag'), df._get_value(x, 'Zeitraum')))
				connection.commit()
			
		connection.close()


# Method to calculate room utilization per room per timeslot and safe the results in table 'results'
def process_data(db_file):
	connection = sqlite3.connect(db_file)

	for w in ['Infobib', 'KleinerLernraum', 'Einzelarbeitsraum']:
		for x in ['2', '3', '4']:
			for y in ['Mo', 'Di', 'Mi', 'Do', 'Fr']:
				for z in ['09:30 - 11:15', '11:20 - 13:50', '13:55 - 15:30', '15:35 - 17:30', '17:35 - 20:00', '20:05 - 22:00']:
				
					crsr = connection.cursor()
					crsr.execute("SELECT COUNT(*) FROM studyroomUsage WHERE Room = ? and CalendarWeek = ? and Weekday = ? and Timeslot = ?", (w, x, y, z))
					count = crsr.fetchall()[0][0]
					crsr.execute("INSERT INTO results VALUES ( ?, ?, ?, ?, ?, ?)", (w, x, 'Termin', y, z, str(count)))
					connection.commit()

	connection.close()




# Method to plot data from a SQL file and visualize it in a GUI
def plot_data(db_file):
	connection = sqlite3.connect(db_file)
	crsr = connection.cursor()
	crsr.execute("SELECT Weekday, Timeslot, NumberOfReservations FROM results WHERE room = 'Infobib' and CalendarWeek = '4'")
	
	# Extract SQL values to lists that can be used by the Matplotlib librarie
	timeslots = []
	reservations = []
	
	for i in crsr:
		timeslots.append(i[0] + i[1])
		#timeslots.append(i[0])
		print(i[2])
		reservations.append(int(i[2]))
	
	connection.close()
	
	# Visualizing data using Matplotlib librarie
	
	#plt.xlim([0, x_max])
	plt.ylim([0, 50])
	plt.xlabel("Timeslot")
	plt.ylabel("Number of reservations")
	plt.title("Study room reservations during the month")
	plt.plot(timeslots, reservations)
	plt.show()

		




db_file = 'database.db'
schemaDataStorage_file = 'schemaDataStorage.sql'
schemaResults_file = 'schemaResults.sql'



#initial_DB_connect(db_file, schemaDataStorage_file, schemaResults_file)
#read_data_to_sql(db_file)
#process_data(db_file)
plot_data(db_file)



# Delete unneeded files to protect user data


# Read data from new excel into SQL db
# Delete unneeded files
# Plot graph

