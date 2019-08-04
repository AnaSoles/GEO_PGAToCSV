import sys
# xlwt is a library to export and work in excel sheet
import xlwt
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

# 11 Create a excel Sheet
workbook = xlwt.Workbook(encoding="UTF-8")
# 12 Object "sheet_amsv" to create a sheet in the excel workwook called "AMSV1". An object is to do action on it.
sheet_amsv1 = workbook.add_sheet('AMSV1')
sheet_amsv2 = workbook.add_sheet('AMSV2')

# 13 write the value in your sheet - first colum and line start with 0; (raw,column,value to add)
sheet_amsv1.write(2,0,"value1")
sheet_amsv2.write(2,0,"value2")



#10 Declare variable de etat - initialization - boolean variable
# must_print=False and must_print=True 
must_print=False

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
		must_print=False
	#8 to find Intensity in the array
	if split_line[0]=="Intensity":
		must_print=False
		#9 to find T=.... in the array
		if (split_line[4]=="T=0.000"):
			must_print=True
			print (split_line[4])
		if (split_line[4]=="T=0.200"):
			must_print=True
			print (split_line[4])
		if (split_line[4]=="T=1.000"):
			must_print=True
			print (split_line[4])
	# Boolean Variable execute when must_print=True and it should be declare as
	# must_print==True beacause I test value 
	# SI JE DOIS IMPRIME LA LIGNE
	if must_print==True:
		#JE IMPRIME LA LIGNE
		print (line.rstrip())



		


	 #longitude = input_file.readline().rstrip()
	# if len(longitude) == 0:
	# 	break
	# latitude = input_file.readline().rstrip()
	# if len(latitude) == 0:
	# 	break
	# # print values on screen
	# print(name)
	# print(longitude)
	# print(latitude)
	# # write into file
	# output_file.write(longitude);
	# output_file.write(" ");
	# output_file.write(latitude);
	# output_file.write(" ");
	# output_file.write("\n");
workbook.save('EXC_AMSV.xls')

print ("end..")
