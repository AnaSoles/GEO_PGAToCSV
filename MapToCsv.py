import sys

# this are the arguments passed to the script
# argv[0] is the name of the script
# argv[1] is the name of input file
# argv[2] is the name of output file
print 'Number of arguments:', len(sys.argv), 'arguments.'
print 'Argument List:', str(sys.argv)

# open input file
input_file = open(sys.argv[1], "r")

#open output file
output_file = open(sys.argv[2], "w")

#loop until end of file
while True:
	# read velue from input file
	name = input_file.readline().rstrip()
	if len(name) == 0:
		break
	longitude = input_file.readline().rstrip()
	if len(longitude) == 0:
		break
	latitude = input_file.readline().rstrip()
	if len(latitude) == 0:
		break
	# print values on screen
	print(name)
	print(longitude)
	print(latitude)
	# write into file
	output_file.write(longitude);
	output_file.write(" ");
	output_file.write(latitude);
	output_file.write(" ");
	output_file.write("\n");

print "end.."
