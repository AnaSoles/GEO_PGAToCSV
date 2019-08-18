#
import sys
# xlwt is a library to export and work in excel sheet
import xlsxwriter

#------------------------------------------------------------------------------------
# 							Create graph for excel
#------------------------------------------------------------------------------------

# global variables defines
intensities = []

# def means function
def create_graphs(workbook, nblines):

	for worksheet in workbook.worksheets():

		# Create a chart object. Type is kind of graph, scatter is line
		chart = workbook.add_chart({'type': 'scatter',
									'subtype': 'smooth'})

		# # Configure the series of the chart from the dataframe data.
		chart.add_series({
			'name': "B",
		 	# X value
		 	'categories': [worksheet.name, 1, 0, nblines, 0],
		 	# Y value
		 	'values': [worksheet.name, 1, 1, nblines, 1],
		 	'marker': {'type': 'automatic'},
		 	'line':   {'width': 1.5},
		})
		chart.add_series({
			'name': "C",
		 	# X value
		 	'categories': [worksheet.name, 1, 0, nblines, 0],
		 	# Y value
		 	'values': [worksheet.name, 1, 2, nblines, 2],
		 	'marker': {'type': 'automatic'},
		 	'line':   {'width': 1.5},
		})
		chart.add_series({
			'name': "D",
		 	# X value
		 	'categories': [worksheet.name, 1, 0, nblines, 0],
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
 						   'num_format': '0.00',
					       'max' : max(intensities),
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
		worksheet.insert_chart(2, 10, chart)

def create_graphs_2(workbook, nblines):

	for worksheet in workbook.worksheets():

		# Create a chart object. Type is kind of graph, scatter is line
		chart = workbook.add_chart({'type': 'scatter',
									'subtype': 'smooth'})
		# # Configure the series of the chart from the dataframe data.
		chart.add_series({
			'name': "B",
		 	# X value
		 	'categories': [worksheet.name, nblines + 3, 0, nblines + 3 + (int)(4.0/0.1), 0],
		 	# Y value
		 	'values': [worksheet.name, nblines + 3, 1, nblines + 3 + (int)(4.0/0.1), 1],
		 	'marker': {'type': 'automatic'},
		 	'line':   {'width': 1.5},
		})
		chart.add_series({
			'name': "C",
		 	# X value
		 	'categories': [worksheet.name, nblines + 3, 3, nblines + 3 + (int)(4.0/0.1), 3],
		 	# Y value
		 	'values': [worksheet.name, nblines + 3, 4, nblines + 3 + (int)(4.0/0.1), 4],
		 	'marker': {'type': 'automatic'},
		 	'line':   {'width': 1.5},
		})
		chart.add_series({
			'name': "D",
		 	# X value
		 	'categories': [worksheet.name, nblines + 3, 6, nblines + 3 + (int)(4.0/0.1), 6],
		 	# Y value
		 	'values': [worksheet.name, nblines + 3, 7, nblines + 3 + (int)(4.0/0.1), 7],
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
 						   'num_format': '0.00',
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
	    				   'num_format': '0.00',
						  })

		# Set an Excel chart style. Colors with white outline and shadow.
		chart.set_style(10)

		# # # Insert the chart into the worksheet.
		worksheet.insert_chart(24, 10, chart)

def reduce_name(city, footer):
	# sheet name cannot be greater than 30 characters.
	result = city[:30-len(footer)-len("_")]
	result = result +'_'+ footer
	return result;

def parse_dat_file(filename):
	gra_file = open(filename, "r")

	# we store the intensities valules into a global variable intensities table
	global intensities;

	#--------------------  LOOP FOR EACH LINE IN THE .MAP FILE TO FIND INTENSITIES VALES FOR UHS ----------------------------------
	#loop until end of file
	while True:
		line = gra_file.readline()
		if len(line) == 0:
			break

		split_line=(line.rstrip().split(","))

		if (len(split_line) == 4) and (split_line[3]=='gals'):
			intensities.append((float)(split_line[0]))

fpga_values = [0.1, 0.2, 0.3, 0.4, 0.5]
fpga_matching = [
	# 0.1
	{
	"B" : 1.0,
	"C" : 1.2,
	"D" : 1.6,
	},
	# 0.2
	{
	"B" : 1.0,
	"C" : 1.2,
	"D" : 1.4,
	},
	# 0.3
	{
	"B" : 1.0,
	"C" : 1.1,
	"D" : 1.2,
	},
	# 0.4
	{
	"B" : 1.0,
	"C" : 1.0,
	"D" : 1.1,
	},
	# 0.5
	{
	"B" : 1.0,
	"C" : 1.0,
	"D" : 1.0,
	}
]

Fa_values = [0.25, 0.50, 0.75, 1.0, 1.25]
Fa_matching = fpga_matching # same table as fpga

Fv_values = [0.10, 0.20, 0.30, 0.40, 0.50]
Fv_matching = [
	# 0.1
	{
	"B" : 1.0,
	"C" : 1.7,
	"D" : 2.4,
	},
	# 0.2
	{
	"B" : 1.0,
	"C" : 1.6,
	"D" : 2.0,
	},
	# 0.3
	{
	"B" : 1.0,
	"C" : 1.5,
	"D" : 1.8,
	},
	# 0.4
	{
	"B" : 1.0,
	"C" : 1.4,
	"D" : 1.6,
	},
	# 0.5
	{
	"B" : 1.0,
	"C" : 1.3,
	"D" : 1.5,
	}
]

def low_value(value, ref_values):
	if (value<=ref_values[0]):
		return ref_values[0]
	for i in range(1, len(ref_values)):
		#print str(i) + " -> " + str(ref_values[i-1])
		if (value < ref_values[i]):
			return ref_values[i-1];
	return ref_values[len(ref_values)-1]

def up_value(value, ref_values):
	if (value>=ref_values[len(ref_values)-1]):
		return ref_values[len(ref_values)-1]
	for i in range(1, len(ref_values) + 1):
		#print str(len(ref_values)-i) + " -> " + str(ref_values[len(ref_values)-i])
		if (value > ref_values[len(ref_values)-i]):
			return ref_values[len(ref_values)-i + 1];
	return ref_values[0]

def compute_value(value, ref_values, matching_values, category):
	
	result = -1

	value_low = low_value(value, ref_values)
	index_value_low = ref_values.index(value_low)
	value_low_result = matching_values[index_value_low][category]

	value_up = up_value(value, ref_values)
	index_value_up = ref_values.index(value_up)
	value_up_result = matching_values[index_value_up][category]

	# print "value " + str(value)
	# print "value_low " + str(value_low) 
	# print "value_up " + str(value_up) 
	# print "value_low_result " + str(value_low_result) 
	# print "value_up_result " + str(value_up_result) 

	if ((value_up-value_low) == 0):
		result = value_low_result
	else:
		result = value_low_result + ((value-value_low) * (value_up_result-value_low_result)) / (value_up-value_low)
	return result

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
					rp_map.append(split_line[RP+1])

		# city line
		if (len(split_line)>9):
			read_city = split_line[9]
			must_print = True;
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
						current_sheet.write(0, 6 , "ROCK B")
						current_sheet.write(0, 7 , "ROCK C")
						current_sheet.write(0, 8 , "ROCK D")
					if (current_column == 1):
						current_sheet.write(curent_line, 0, intensities[curent_line-1])
						current_sheet.write(1, 5, "PGA")
						current_sheet.write(2, 5, "Ss")
						current_sheet.write(3, 5, "Sd")
						current_sheet.write(4, 5, "Fpga")
						current_sheet.write(5, 5, "Fa")
						current_sheet.write(6, 5, "Fv")
						current_sheet.write(7, 5, "As")
						current_sheet.write(8, 5, "Sds")
						current_sheet.write(9, 5, "Sd1")
						# T0 -> formula
						current_sheet.write(10, 5, "T0")
						current_sheet.write(10, 6, "=0.2*G10/G9")
						current_sheet.write(10, 7, "=0.2*H10/H9")
						current_sheet.write(10, 8, "=0.2*I10/I9")
						# Ts -> formula
						current_sheet.write(11, 5, "Ts")
						current_sheet.write(11, 6, "=G10/G9")
						current_sheet.write(11, 7, "=H10/H9")
						current_sheet.write(11, 8, "=I10/I9")
						# Tltl -> constant
						current_sheet.write(12, 5, "Tltl")
						current_sheet.write(12, 6, 4.0)
						current_sheet.write(12, 7, 4.0)
						current_sheet.write(12, 8, 4.0)

					current_sheet.write(curent_line, current_column ,(float)(split_line[RP+1]))

					# compute values
					intensity = intensities[curent_line-1]
					intensities_to_print = [0.0, 0.2, 1.0]
					if (intensity in intensities_to_print):
						intensity_index = intensities_to_print.index(intensity)
						# values / 981
						value981 = ((float)(split_line[RP+1]))/981.0
						current_sheet.write(intensity_index + 1, 5 + current_column, value981)
						# fpga, fa, fv
						scales = [fpga_values, Fa_values, Fv_values]
						matchings = [fpga_matching, Fa_matching, Fv_matching]
						interpolated_value = compute_value(value981, scales[intensity_index], matchings[intensity_index], category)
						current_sheet.write(intensity_index + 4, 5 + current_column, interpolated_value)
						# As / Sds / Sd1
						AsSdsSd1 = interpolated_value * value981;
						current_sheet.write(intensity_index + 7, 5 + current_column, AsSdsSd1)

				curent_line = curent_line + 1
	return curent_line

def create_AASHTO_table(current_sheet, line, column, category):
	# header
	current_sheet.write(line, column, "Rock " + category)
	line  = line + 1
	current_sheet.write(line, column, "Tm (seg)")
	current_sheet.write(line, column + 1, "Cm (-)")
	line  = line + 1
	T = 0.0
	cell_format = workbook.add_format()
	cell_format.num_format = '0.00'
	categoryToColumn = {
		"B":"G",
		"C":"H",
		"D":"I"
	}
	Asline =  "8"
	Sdsline = "9"
	Sd1line = "10"
	T0line  = "11"
	Tsline  = "12"
	while (T <= 4.1):
		current_sheet.write(line, column, T, cell_format)
		T0  = "{}{}".format(categoryToColumn[category], T0line)
		Ts  = "{}{}".format(categoryToColumn[category], Tsline)
		Sds = "{}{}".format(categoryToColumn[category], Sdsline)
		Sd1 = "{}{}".format(categoryToColumn[category], Sd1line)
		As  = "{}{}".format(categoryToColumn[category], Asline)
		formula = "=IF({}<{},({}-{})*({}/{})+{},IF({}>={},{}/{},{}))".format(T,T0,Sds,As,T,T0,As,T,Ts,Sd1,T,Sds)
		current_sheet.write(line, column + 1, formula, cell_format)
		line = line + 1
		T = T + 0.1

#------------------------------------------------------------------------------------
#------------------------------------------------------------------------------------

# this are the arguments passed to the script
# argv[0] is the name of the script
# argv[1] is the name of input file
# argv[2] is the name of output file
print( 'Number of arguments:', len(sys.argv), 'arguments.')
print( 'Argument List:', str(sys.argv))

if (len(sys.argv)<5):
	print "Error - wrong parameters"
	print "Usage (example): " + sys.argv[0] + " GRA_B.map GRA_C.map GRA_D.map GRA_B.dat"
	exit(1)

# 11 Create a excel Sheet - WORKBOOK 
workbook = xlsxwriter.Workbook('MAP_AMSV.xlsx')
# 12 Object "sheet_amsv" to create a sheet in the excel workwook called "AMSV1". An object is to do action on it.
#sheet_amsv1 = workbook.add_worksheet('AMSV1')
#sheet_amsv2 = workbook.add_worksheet('AMSV2')

# 13 write the value in your sheet - first colum and line start with 0; (raw,column,value to add)
#sheet_amsv1.write(2,0,"value1")
#sheet_amsv2.write(2,0,"value2")

# print low_value(0.60, fpga_values)
# print low_value(0.55, fpga_values)
# print low_value(0.50, fpga_values)
# print low_value(0.45, fpga_values)
# print low_value(0.40, fpga_values)
# print low_value(0.35, fpga_values)
# print low_value(0.30, fpga_values)
# print low_value(0.25, fpga_values)
# print low_value(0.20, fpga_values)
# print low_value(0.15, fpga_values)
# print low_value(0.10, fpga_values)
# print low_value(0.05, fpga_values)
#print compute_value(0.35, fpga_values, fpga_matching, "D")
#exit(0)
# open input file
parse_dat_file(sys.argv[4])
if (len(intensities)==0):
	print("ERROR: No gals unit found in .dat file !!!")
	exit(1)
print("Found " + str(len(intensities)) + " in dat files.")
nblines = parse_map_file(workbook, "B", sys.argv[1], 1)
nblines = parse_map_file(workbook, "C", sys.argv[2], 2)
nblines = parse_map_file(workbook, "D", sys.argv[3], 3)


for worksheet in workbook.worksheets():
	create_AASHTO_table(worksheet, nblines + 1, 0 ,"B")
	create_AASHTO_table(worksheet, nblines + 1, 3 ,"C")
	create_AASHTO_table(worksheet, nblines + 1, 6 ,"D")

create_graphs(workbook, nblines)
create_graphs_2(workbook, nblines)

workbook.close()

print ("end..")
