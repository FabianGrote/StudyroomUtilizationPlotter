import os
import fnmatch
import numpy as np
import pandas
import matplotlib.pyplot as plt
#from numpy import *
import sqlite3

# Create SQLite database
db_file = 'database.db'
schemaDataStorage_file = 'schemaDataStorage.sql'
schemaResults_file = 'schemaResults.sql'

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



# Get the current working directory 
cwd = os.getcwd()

print("Current working directory:" , cwd)


# Iterate through all files in the same folder as this script and save corresponding data in SQL database coressponding
for file in os.listdir(cwd):
	if file.endswith(".ods"):
		if fnmatch.fnmatch(file, 'empty.ods'):
			print("Detected 'empty.ods' ")
			continue		
		print("Contains data from file:", file)
		df = pandas.read_excel(file, sheet_name=0)		
		
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



connection = sqlite3.connect(db_file)

# Calculate room utilization per room per timeslot and safe the results in table 'results'
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
		




# Plot data from db and show in GUI

#plt.plot(x,Infobib_min, 'o')
#plt.plot(x,Infobib_average, 'g')
#plt.plot(x,Infobib_max, 'y')


#plt.plot(x,Stillarbeitsraum_min)
#plt.plot(x,Stillarbeitsraum_average)
#plt.plot(x,Stillarbeitsraum_max)

#plt.show()



# Delete unneeded files to protect user data


# Read data from new excel into SQL db
# Delete unneeded files
# Plot graph

