import sys

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

#loop until end of file
while True:
	# read velue from input file
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

	#7 to find site in the array
	if split_line[0]=="Site:": 
		print ("found" )


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

print ("end..")
