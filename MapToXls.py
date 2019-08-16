#
import sys
# xlwt is a library to export and work in excel sheet
import xlsxwriter

#------------------------------------------------------------------------------------
# 							Create graph for excel
#------------------------------------------------------------------------------------

cities = [];

def create_graphs(workbook, nblines):

	for worksheet in workbook.worksheets():

		# Create a chart object. Type is kind of graph, scatter is line
		chart = workbook.add_chart({'type': 'line'})

		# # Configure the series of the chart from the dataframe data.
		chart.add_series({
			'name': "B",
		 	# Y value
		 	'values': [worksheet.name, 1, 1, nblines, 1],

		 	'marker': {'type': 'automatic'},
		 	'line':   {'width': 1.5},
		})
		chart.add_series({
			'name': "C",
		 	# Y value
		 	'values': [worksheet.name, 1, 2, nblines, 2],

		 	'marker': {'type': 'automatic'},
		 	'line':   {'width': 1.5},
		})
		chart.add_series({
			'name': "D",
		 	# Y value
		 	'values': [worksheet.name, 1, 3, nblines, 3],

		 	'marker': {'type': 'automatic'},
		 	'line':   {'width': 1.5},
		})
		# Add a chart title and some axis labels.
		chart.set_title ({'name': 'CURVAS DE PROBABILIDAD DE EXCEDENCIA '})
		chart.set_x_axis({'name': 'Aceleracion Spectral (gals)',
						  'major_gridlines': {
	        						'visible': True,
	        						'line': {'width': 0.1}
	    				            },
	    				   'minor_gridlines': {
	        						'visible': True,
	        						'line': {'width': 0.01}
	    				            },
 						   'num_format': '0.00'
						})
		chart.set_y_axis({'name': 'Frecuencia Anual de Excedencia (1/anos)',
						  'major_gridlines': {
	        						'visible': True,
	        						'line': {'width': 0.1}
	    				            },
	    				   'minor_gridlines': {
	        						'visible': True,
	        						'line': {'width': 0.01}
	    				            },
	    				   'num_format': '0.00'
						  })

		# Set an Excel chart style. Colors with white outline and shadow.
		chart.set_style(10)

		# # # Insert the chart into the worksheet.
		# print("***********" + worksheet.name)
		worksheet.insert_chart(nblines + 2, 3, chart)

def add_city(city):
	global cities
	if (city not in cities):
		cities.append(city)

def reduce_name(city, footer):
	# sheet name cannot be greater than 30 characters.
	result = city[:30-len(footer)-len("_")]
	result = result +'_'+ footer
	return result;

def parse_map_file(workbook, category, filename, current_column):
	gra_file = open(filename, "r")
	#10 initialization de variable de etat - boolean variable
	# must_print=False and must_print=True 
	must_print=False
	# Initialization of variable
	CITY="Uknown"
	# Initialization of variable object - to read each sheet in the during the loop
	current_sheet=None
	# Initialization of variable to print results in the sheet
	curent_line=0
	rp_map = [];

	#--------------------  LOOP FOR EACH LINE IN THE .GRA FILE----------------------------------
	#loop until end of file
	while True:
		line = gra_file.readline()
		if len(line) == 0:
			break

		split_line=(line.rstrip().split())

		# rp line
		if (len(split_line)>1):
			if (split_line[0] == "RP"):
				for RP in range(0,5):
					print split_line
					rp_map.append(split_line[RP+1])

		# city line
		if (len(split_line)>9):
			read_city = split_line[9]
			must_print = True;
			print read_city
			if (CITY != read_city):
				CITY = read_city
				curent_line = 1;
		else:
			must_print = False

		if must_print==True:
			RP = 1
			if (len(split_line)>5):
				for RP in range(0,5):
					worksheet_name = reduce_name(CITY, "TF_" + str(rp_map[RP]))
					current_sheet = workbook.get_worksheet_by_name(worksheet_name);
					if (current_sheet==None):
						current_sheet = workbook.add_worksheet(worksheet_name)
						current_sheet.write(0, 0 , "Period")
						current_sheet.write(0, 1 , "ROCK_B")
						current_sheet.write(0, 2 , "ROCK_C")
						current_sheet.write(0, 3 , "ROCK_D")
					if (current_column == 1):
						current_sheet.write(curent_line, 0,(int)(split_line[0]))
					current_sheet.write(curent_line, current_column ,(float)(split_line[RP+1]))
				curent_line = curent_line + 1
	return curent_line

#------------------------------------------------------------------------------------
#------------------------------------------------------------------------------------

# this are the arguments passed to the script
# argv[0] is the name of the script
# argv[1] is the name of input file
# argv[2] is the name of output file
print( 'Number of arguments:', len(sys.argv), 'arguments.')
print( 'Argument List:', str(sys.argv))

if (len(sys.argv)<4):
	print "Error - wrong parameters"
	print "Usage: " + sys.argv[0] + " GRA_B.map GRA_C.map GRA_D.map"
	exit(1)

# 11 Create a excel Sheet - WORKBOOK 
workbook = xlsxwriter.Workbook('MAP_AMSV.xlsx')
# 12 Object "sheet_amsv" to create a sheet in the excel workwook called "AMSV1". An object is to do action on it.
#sheet_amsv1 = workbook.add_worksheet('AMSV1')
#sheet_amsv2 = workbook.add_worksheet('AMSV2')

# 13 write the value in your sheet - first colum and line start with 0; (raw,column,value to add)
#sheet_amsv1.write(2,0,"value1")
#sheet_amsv2.write(2,0,"value2")

# open input file
nblines = parse_map_file(workbook, "B", sys.argv[1], 1)
nblines = parse_map_file(workbook, "C", sys.argv[2], 2)
nblines = parse_map_file(workbook, "D", sys.argv[3], 3)

create_graphs(workbook, nblines)

workbook.close()

print ("end..")

print cities
