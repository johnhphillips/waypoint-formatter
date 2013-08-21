import xlrd

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

# Open the Excel workbook (for now this is hard coded, will add prompt for user input)
book = xlrd.open_workbook('C:\Users\jphilips\Desktop\ATLAS_minefield_configurations_v05.xlsx',on_demand=True)

print
print "Number of sheets in workbook: ", book.nsheets

# Open the first sheet (for now will only open sheet #1, will iterate over workbook in final)
sh = book.sheet_by_index(1)

print
print "Opening: ", sh.name
print
 
# sh.name -> sheet name, used for target field label #LABEL
# sh.nrows, sh.ncols -> give number of rows and columns

# List to hold waypoint information
waypoints = []
# Set up list to hold target information
targets = []

# Collect target names and add to waypoint list
for column in range(sh.ncols):
	# Check for right column 
	if sh.cell_value(rowx=0, colx=column) == u'Target ID':
		for row in range(sh.nrows):
			# Skip the label row
			if row !=0:
				# Create list to hold target information
				target = []
				# Save target name and set to all CAPS
				targetName = sh.cell_value(rowx=row, colx=column)
				targetName = targetName.encode('ascii', 'ignore')
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
				# Reorder latitude information
				a = sh.cell_value(rowx=row, colx=column)
				b = sh.cell_value(rowx=row, colx=column+1)
				a = a.encode('ascii', 'ignore')
				a = a + b[7]
				b=b[:-2]
				c=a+b
				targets[row-1].append(c)


# Collect target longitudes, reorder, add to target list in waypoint list
for column in range(sh.ncols):
	if sh.cell_value(rowx=0, colx=column) == u'Longitude':
		for row in range(sh.nrows):
			# Skip the label row
			if row !=0:
				# Reorder longitude information
				a = sh.cell_value(rowx=row, colx=column)
				b = sh.cell_value(rowx=row, colx=column+1)
				a = a.encode('ascii', 'ignore')
				a = a + b[7]
				b=b[:-2]
				c=a+b
				targets[row-1].append(c)

# Take target information and format into waypoints
for target in targets:
	point = Waypoint()
	setLocationLabel(point, sh.name)
	setLabel(point, target[0])
	setPosition(point, target[1], target[2])
	# Add point to waypoint list
	waypoints.append(point)

outputName = sh.name + '.ini'
# Open output file in write mode
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
