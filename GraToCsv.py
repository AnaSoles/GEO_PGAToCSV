#
import sys
# xlwt is a library to export and work in excel sheet
import xlsxwriter

#------------------------------------------------------------------------------------
# 							Create graph for excel
#------------------------------------------------------------------------------------

def create_graph(workbook, worksheet, nblines):
	if worksheet==None:
		return;

	# Create a chart object. Type is kind of graph, scatter is line
	chart = workbook.add_chart({'type': 'line'})

	# # Configure the series of the chart from the dataframe data.
	chart.add_series({
		'name': "test",
	 	# X value
	 	'categories': [worksheet.name, 1, 2, nblines, 2],
	 	# Y value
	 	'values': [worksheet.name, 1, 1, nblines, 1],
	 	
	 	'marker': {'type': 'automatic'},
	 	'line':   {'width': 1.5},
	})

	# Add a chart title and some axis labels.
	chart.set_title ({'name': 'CURVAS DE PROBABILIDAD DE EXCEDENCIA '})
	chart.set_x_axis({'name': 'Aceleracion Spectral (gals)',
					  'min' : 0.01,
					  'max' : 10,
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
    				    'num_format': '0.00'
					  })

	# Set an Excel chart style. Colors with white outline and shadow.
	chart.set_style(10)

	# # # Insert the chart into the worksheet.
	# print("***********" + worksheet.name)
	worksheet.insert_chart(1, 4, chart)

#------------------------------------------------------------------------------------
#------------------------------------------------------------------------------------

# this are the arguments passed to the script
# argv[0] is the name of the script
# argv[1] is the name of input file
# argv[2] is the name of output file
print( 'Number of arguments:', len(sys.argv), 'arguments.')
print( 'Argument List:', str(sys.argv))

# open input file
gra_file = open(sys.argv[1], "r")

#open output file
#---output_file = open(sys.argv[2], "w")

# 11 Create a excel Sheet - WORKBOOK 
workbook = xlsxwriter.Workbook('EXC_AMSV.xlsx')
# 12 Object "sheet_amsv" to create a sheet in the excel workwook called "AMSV1". An object is to do action on it.
#sheet_amsv1 = workbook.add_worksheet('AMSV1')
#sheet_amsv2 = workbook.add_worksheet('AMSV2')

# 13 write the value in your sheet - first colum and line start with 0; (raw,column,value to add)
#sheet_amsv1.write(2,0,"value1")
#sheet_amsv2.write(2,0,"value2")

#10 initialization de variable de etat - boolean variable
# must_print=False and must_print=True 
must_print=False
# Initialization of variable
CITY="Uknown"
# Initialization of variable object - to read each sheet in the during the loop
current_sheet=None
# Initialization of variable to print results in the sheet
curent_line=0
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
			create_graph(workbook, current_sheet, curent_line-1)
			current_sheet = workbook.add_worksheet(CITY+'_'+"T000")
			current_sheet.write(0,0,"X")
			current_sheet.write(0,1,"Y (1/anos)")
			# X - X/981
			current_sheet.write(0,2,"Aceleracion Spectral (gals)")
			curent_line=1
			print (split_line[4])
		if (split_line[4]=="T=0.200"):
			must_print=True
			create_graph(workbook, current_sheet, curent_line-1)
			current_sheet = workbook.add_worksheet(CITY+'_'+"T020")
			current_sheet.write(0,0,"X (gals)")
			current_sheet.write(0,1,"Y")
			current_sheet.write(0,2,"Aceleracion Spectral (gals)")
			print (split_line[4])
			curent_line=1
		if (split_line[4]=="T=1.000"):
			must_print=True
			create_graph(workbook, current_sheet, curent_line-1)
			current_sheet = workbook.add_worksheet(CITY+'_'+"T100")
			current_sheet.write(0,0,"Aceleracion Spectral (gals)")
			current_sheet.write(0,1,"Frecuencia Anual de Excedencia (1/anos)")
			current_sheet.write(0,2,"Aceleracion Spectral (gals)")
			print (split_line[4])
			curent_line=1
	# Boolean Variable execute when must_print=True and it should be declare as
	# must_print==True beacause I test value 
	# SI JE DOIS IMPRIME LA LIGNE
	if must_print==True:
		#JE IMPRIME LA LIGNE
		print (line.rstrip())
		if len(split_line)<5:
			# First colum and line start with 0; (raw,column,value to add)
			# float to recognize number with decimal 
			current_sheet.write(curent_line,0,(float)(split_line[0]))
			current_sheet.write(curent_line,1,(float)(split_line[2]))
			current_sheet.write(curent_line,2,(float)(split_line[0])/981)
			curent_line=curent_line + 1

#write the last graph
create_graph(workbook, current_sheet, curent_line-1)

workbook.close()

print ("end..")
