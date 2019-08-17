#
import sys
# xlwt is a library to export and work in excel sheet
import xlsxwriter

#------------------------------------------------------------------------------------
# 							Create graph for excel
#------------------------------------------------------------------------------------

cities = [];

def create_graph(workbook, city, footer, nblines):

	worksheet = workbook.get_worksheet_by_name(reduce_name(city, footer));

	# Create a chart object. Type is kind of graph, scatter is line
	chart = workbook.add_chart({'type': 'scatter',
								'subtype': 'smooth'})

	# # Configure the series of the chart from the dataframe data.
	chart.add_series({
		'name': "B",
	 	# X value
	 	'categories': [worksheet.name, 1, 2, nblines, 2],
	 	# Y value
	 	'values': [worksheet.name, 1, 1, nblines, 1],

	 	'marker': {'type': 'automatic'},
	 	'line':   {'width': 1.5},
	})
	chart.add_series({
		'name': "C",
	 	# X value
	 	'categories': [worksheet.name, 1, 6, nblines, 6],
	 	# Y value
	 	'values': [worksheet.name, 1, 5, nblines, 5],

	 	'marker': {'type': 'automatic'},
	 	'line':   {'width': 1.5},
	})
	chart.add_series({
		'name': "D",
	 	# X value
	 	'categories': [worksheet.name, 1, 10, nblines, 10],
	 	# Y value
	 	'values': [worksheet.name, 1, 9, nblines, 9],

	 	'marker': {'type': 'automatic'},
	 	'line':   {'width': 1.5},
	})

	# Add a chart title and some axis labels.
	chart.set_title ({'name': 'CURVAS DE PROBABILIDAD DE EXCEDENCIA '})
	chart.set_x_axis({'name': 'Aceleracion Spectral (gals)',
					  'min' : 0.01,
					  'max' : 10,
					  'crossing': 0,
					  'log_base': 10,
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
					  'min' : 0.001,
					  'max' : 10,
					  'log_base': 10,
					  'crossing': 0.001,
					  'major_gridlines': {
        						'visible': True,
        						'line': {'width': 0.1}
    				            },
    				   'minor_gridlines': {
        						'visible': True,
        						'line': {'width': 0.01}
    				            },
					   'num_format': '0.00E+00'
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

def parse_gra_file(workbook, category, filename, current_column):
	gra_file = open(filename, "r")
	#10 initialization de variable de etat - boolean variable
	# must_print=False and must_print=True 
	must_print=False
	# Initialization of variable
	CITY="Uknown"
	# Initialization of variable object - to read each sheet in the during the loop
	current_sheet=None
	# Initialization of variable to print results in the sheet
	current_line=0
	#--------------------  LOOP FOR EACH LINE IN THE .GRA FILE----------------------------------
	#loop until end of file
	while True:
		# read velue from input file
	#-------- JE LIS LA LIGNE ------------------
		line = gra_file.readline()
		# 1 rstrip - remove invisble in the ascci file
		# 2 line is a variable which contain one line on the file NR
		if len(line) == 0:
			break
		# 3 Add .rstrip() to remove the return of each line character (no visible)
		#4 To print line (string) splitted into array of strings (table)
	#print (line.rstrip().split(" "))

		# 5 Assing variable to the split command - Make a line to an Array ["ANA" "ES" "BONITA"]
		split_line=(line.rstrip().split(" "))
	# 6 test
		#print (split_line)


	#JE REGARDE SI JE DOIS IMPRIME LA LIGNE

		#7 to find site in the array
		if split_line[0]=="Site:": 
			print (line.rstrip())
			print (split_line[16])
			# Create a VARIABLE to save CITY
			CITY=split_line[16]
			must_print=False
		#8 to find Intensity in the array
		if split_line[0]=="Intensity":
			must_print=False
			#9 to find T=.... in the array
			if (split_line[4]=="T=0.000"):
				must_print=True
				worksheet_name = reduce_name(CITY, "T000")
				current_sheet = workbook.get_worksheet_by_name(worksheet_name);
				if (current_sheet==None):
					current_sheet = workbook.add_worksheet(worksheet_name)
				current_sheet.write(0,current_column + 0,"X")
				current_sheet.write(0,current_column + 1,"Y (1/anos)")
				# X - X/981
				current_sheet.write(0,current_column + 2,"Aceleracion Spectral (gals)")
				current_line=1
				print (split_line[4])
				add_city(CITY)
			if (split_line[4]=="T=0.200"):
				must_print=True
				worksheet_name = reduce_name(CITY, "T020")
				current_sheet = workbook.get_worksheet_by_name(worksheet_name);
				if (current_sheet==None):
					current_sheet = workbook.add_worksheet(worksheet_name)
				current_sheet.write(0,current_column + 0,"X (gals)")
				current_sheet.write(0,current_column + 1,"Y")
				current_sheet.write(0,current_column + 2,"Aceleracion Spectral (gals)")
				print (split_line[4])
				current_line=1
				add_city(CITY)
			if (split_line[4]=="T=1.000"):
				must_print=True
				worksheet_name = reduce_name(CITY, "T100")
				current_sheet = workbook.get_worksheet_by_name(worksheet_name);
				if (current_sheet==None):
					current_sheet = workbook.add_worksheet(worksheet_name)
				current_sheet.write(0,current_column + 0,"Aceleracion Spectral (gals)")
				current_sheet.write(0,current_column + 1,"Frecuencia Anual de Excedencia (1/anos)")
				current_sheet.write(0,current_column + 2,"Aceleracion Spectral (gals)")
				print (split_line[4])
				current_line=1
				add_city(CITY)
		# Boolean Variable execute when must_print=True and it should be declare as
		# must_print==True beacause I test value 
		# SI JE DOIS IMPRIME LA LIGNE
		if must_print==True:
			#JE IMPRIME LA LIGNE
			print (line.rstrip())
			if len(split_line)<5:
				# First colum and line start with 0; (raw,column,value to add)
				# float to recognize number with decimal 
				current_sheet.write(current_line, current_column + 0,(float)(split_line[0]))
				current_sheet.write(current_line, current_column + 1,(float)(split_line[2]))
				current_sheet.write(current_line, current_column + 2,(float)(split_line[0])/981)
				current_line=current_line + 1
	return current_line

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
	print "Usage: " + sys.argv[0] + " GRA_B.gra GRA_C.gra GRA_D.gra"
	exit(1)

# 11 Create a excel Sheet - WORKBOOK 
workbook = xlsxwriter.Workbook('GRA_AMSV.xlsx')
# 12 Object "sheet_amsv" to create a sheet in the excel workwook called "AMSV1". An object is to do action on it.
#sheet_amsv1 = workbook.add_worksheet('AMSV1')
#sheet_amsv2 = workbook.add_worksheet('AMSV2')

# 13 write the value in your sheet - first colum and line start with 0; (raw,column,value to add)
#sheet_amsv1.write(2,0,"value1")
#sheet_amsv2.write(2,0,"value2")

# open input file
nblines = parse_gra_file(workbook, "B", sys.argv[1], 0)
nblines = parse_gra_file(workbook, "C", sys.argv[2], 4)
nblines = parse_gra_file(workbook, "D", sys.argv[3], 8)

for city in cities:
	create_graph(workbook, city, "T000", nblines)
	create_graph(workbook, city, "T020", nblines)
	create_graph(workbook, city, "T100", nblines)

workbook.close()

print ("end..")

print cities
