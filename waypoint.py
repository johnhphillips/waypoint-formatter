
'''
Script that opens Excel workbook, using xlrd module, moves through each worksheet
looking for target data (target name, latitude, and longitude). Saves in list of target
and formats as 'Local waypoints' for use in VIP. Input is file name, provided by user. 
Output is waypoint ini files, one for each worksheet containing target information.
''' 

import xlrd

excelX = ".xlsx"
iniX = ".ini"

class Waypoint:
	locationLabel = '#'
	location = '[Location]'
	label = 'Label='
	position = 'Position='
	offDirection = 'Offset direction=0.0'
	offDistance = 'Offset distance (Meters)='
	offYaxis = 'Offset Y axis (Meters)='

def setLocationLabel(point, label):
	point.locationLabel += label

def setLabel(point, label):
        point.label += label

def setPosition(point, lat, lon):
        point.position = point.position + lat + ' ' + lon

def setOffDirection(point, dir):
	point.offDirection += dir

def setOffDistance(point, dis):
	point.offDistance += dis

def setOffYaxis(point, axis):
	point.offYaxis += axis

print 'Make sure script and spreadsheet are in the same directory (on the desktop)'
print

fname = raw_input('Enter file name: ')
fname = fname + excelx

# Open the Excel workbook based on user input
try:
	book = xlrd.open_workbook(fname, on_demand=True)
except IOError as e:
	print
	print 'File is either not in directory or does not exist!'
	print
	raw_input('Press any key to exit')
	quit()

print
print 'Number of sheets in ', fname, ': ', book.nsheets

# Iterate through sheets in workbook looking for target data 
for sheet in range(book.nsheets):
	dataPresent = False
	# Open sheet
	sh = book.sheet_by_index(sheet)
	print
	print "Opening:", sh.name
	print

	# sh.name -> sheet name, used for target field label #LABEL
	# sh.nrows, sh.ncols -> returns number of rows and columns

	# Initialize list to hold waypoint information
	waypoints = []
	# Initialize list to hold target information
	targets = []

	# Collect target names and add to waypoint list
	for column in range(sh.ncols):
		# Check for right column based on label
		if sh.cell_value(rowx=0, colx=column) == u'Target ID':
			dataPresent = True
			for row in range(sh.nrows):
				# Skip the label row
				if row !=0:
					# Create list to hold target information
					target = []
					# Save target name
					targetName = sh.cell_value(rowx=row, colx=column)
					# Strip off any special characters (degree symbol)
					targetName = targetName.encode('ascii', 'ignore')
					# Set to all CAPS
					targetName = targetName.upper()	
					# Add target name to list
					target.append(targetName)
					# Add target list to waypoints
					targets.append(target)

	
	# Collect target latitudes, reorder, add to target list in waypoint list
	for column in range(sh.ncols):
		if sh.cell_value(rowx=0, colx=column) == u'Latitude':
			for row in range(sh.nrows):
				# Skip the label row
				if row !=0:
					counter = column
					# Reorder latitude information
					a = sh.cell_value(rowx=row, colx=counter)
					while counter < range(sh.ncols):
						counter += 1
						if sh.cell_value(rowx=0, colx=counter) == u'Longitude':
							break
						b = str(sh.cell_value(rowx=row, colx=counter))
						a = a + b
					a = a.encode('ascii', 'ignore')
					b = ''
					b = b + a[0:2]
					b = b + a[len(a)-1]
					b = b + a[3:8]
					if b.endswith("'"):
						b=b[:-1]
					targets[row-1].append(b)

	# Collect target longitudes, reorder, add to target list in waypoint list
	for column in range(sh.ncols):
		if sh.cell_value(rowx=0, colx=column) == u'Longitude':
			for row in range(sh.nrows):
				# Skip the label row
				if row !=0:
					counter = column
					# Reorder latitude information
					a = sh.cell_value(rowx=row, colx=counter)
					while counter < range(sh.ncols):
						counter += 1
						if sh.cell_value(rowx=0, colx=counter) == u'Water Depth (m)':
							break
						b = str(sh.cell_value(rowx=row, colx=counter))
						a = a + b
					a = a.encode('ascii', 'ignore')
					b = ''
					b = b + a[0:3]
					b = b + a[len(a)-1]
					b = b + a[3:9]
					if b.endswith("'"):
						b=b[:-1]
					targets[row-1].append(b)

	# Take target information and format into waypoints
	for target in targets:
		point = Waypoint()
		setLocationLabel(point, sh.name)
		setLabel(point, target[0])
		setPosition(point, target[1], target[2])
		# Add point to waypoint list
		waypoints.append(point)
	
	if dataPresent:
		# Build output file name
		outputName = sh.name + inix
		print 'Building:', outputName
		print
		# Create / open output file in write mode
		fout = open(outputName, 'w')
	
		# Write waypoints to output file
		for point in waypoints:
			fout.write(point.locationLabel + '\n')
			fout.write(point.location + '\n')
			fout.write(point.label + '\n')
			fout.write(point.position + '\n')
			fout.write(point.offDirection + '\n')
			fout.write(point.offDistance + '\n')
			fout.write(point.offYaxis + '\n')
			fout.write('\n')

		# Close output file
		fout.close()
	else:
		print 'No waypoint data present'
		print
	print 'Closing:', sh.name
	print 
# End
raw_input('Press any key to exit')
