#author = Nikhil Gopal
#Last Updated July 15th, 2017

import xlrd         #library used for reading the file

#name of the file that you will create
file_name = raw_input("What would you like the filename to be? Make sure to include .xls at the end!!!")

#opens the input file, and then opens the first sheet for reading
input_file_location = raw_input("What is the location of the input file? Make sure to include .xls at the end!")
wb = xlrd.open_workbook(input_file_location)
sheet = wb.sheet_by_index(0)

#copies the values of the columns we are interested in
file_name_column = sheet.col_values(0)
response_column = sheet.col_values(4)
itsd_response_column = sheet.col_values(5)


import xlwt #library needed for writing a new excel file
workbook = xlwt.Workbook()
sheet1=workbook.add_sheet("Cleaned Data") #creates a new workbook and then adds a new sheet to it

counter = 0 #for loop that pastes the copies columns into the new file
for value in range(0,len(file_name_column)):

    sheet1.write(value,0,str(file_name_column[counter]))
    sheet1.write(value,1,str(response_column[counter]))
    sheet1.write(value,2,str(itsd_response_column[counter]))
    counter += 1

counter2 = 0
#calculates the response ratio cloumn
for item in response_column:
    if counter2 == 0:
        sheet1.write(counter2,3,"Response Ratio")
    elif type(item) == float:
        number = item / float(itsd_response_column[counter2])
        sheet1.write(counter2,3,str(number))
    else:
        pass
    counter2 += 1

workbook.save(file_name)

