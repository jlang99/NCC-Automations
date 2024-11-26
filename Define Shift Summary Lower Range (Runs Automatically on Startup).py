import pyodbc


db = r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=G:\Shared drives\Narenco Projects\O&M Projects\NCC\NCC\NCC 039.accdb;'
connect_db = pyodbc.connect(db)
c = connect_db.cursor()

c.execute("SELECT TOP 1 [ActivityLogID] FROM [ActivityLog] ORDER BY [ActivityLogID] DESC")
starting_value = c.fetchone() #This is a list
connect_db.close()

start = starting_value[0]

# Write start variable to a text file
file_path = r"G:\Shared drives\O&M\NCC Automations\Emails\Shift Summary Start.txt"
with open(file_path, 'w+') as file:
    file.write(str(start))
