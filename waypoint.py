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
        point.position = position + lat + ' ' + lon

# Open the Excel workbook (for now this is hard coded, will add prompt for user input)
book = xlrd.open_workbook('C:\Users\jphilips\Desktop\ATLAS_minefield_configurations_v05.xlsx',on_demand=True)

print "Number of sheets in workbook: ", book.nsheets

# Open the first sheet (for now will only open sheet #1, will iterate over workbook in final)
sh = book.sheet_by_index(1)
print "Opening: ", sh.name
 
# sh.name -> sheet name, used for target field label #LABEL
# sh.nrows, sh.ncols -> give number of rows and columns

#set up array to hold target name, Lat, and Longs
table = [[0 for rows in range(sh.nrows)] for x in xrange(3)]
counter = 0

#collect and populate target id
for column in range(sh.ncols):
	if sh.cell_value(rowx=0, colx=column) == u'Target ID':
		for row in range(sh.nrows):
			#skip label row
			if row !=0:
				table[counter][row-1]=sh.cell_value(rowx=row, colx=column)
#				print sh.cell_value(rowx=row, colx=column)

counter+=1

#collect lats, reorder, and populate
for column in range(sh.ncols):
	if sh.cell_value(rowx=0, colx=column) == u'Latitude':
		for row in range(sh.nrows):
			#skip label row
			if row !=0:
				a = sh.cell_value(rowx=row, colx=column)
				b = sh.cell_value(rowx=row, colx=column+1)
				a = a.encode('ascii', 'ignore')
				a = a + b[7]
				b=b[:-2]
				c=a+b
				table[counter][row-1]=c
#				print sh.cell_value(rowx=row, colx=column), sh.cell_value(rowx=row, colx=column+1)

counter+=1

#collect longs, reorder, and populate
for column in range(sh.ncols):
	if sh.cell_value(rowx=0, colx=column) == u'Longitude':
		for row in range(sh.nrows):
			#skip label row
			if row !=0:
				a = sh.cell_value(rowx=row, colx=column)
				b = sh.cell_value(rowx=row, colx=column+1)
				a = a.encode('ascii', 'ignore')
				a = a + b[7]
				b=b[:-2]
				c=a+b
				table[counter][row-1]=c
#				print sh.cell_value(rowx=row, colx=column), sh.cell_value(rowx=row, colx=column+1)

print table

