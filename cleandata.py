#author = Nikhil Gopal
#Last Updated July 15th, 2017

import xlrd         #library used for reading the file

#name of the file that you will create
file_name = raw_input("What would you like the filename to be? Make sure to include .xls at the end!!!")

#opens the input file, and then opens the first sheet for reading
input_file_location = raw_input("What is the location of the input file? Make sure to include .xls at the end!") #creates a variable that stores the input file location
wb = xlrd.open_workbook(input_file_location, on_demand=True)


list_of_sheet_names = wb.sheet_names() #creates a list to be used later with the names of the sheets in the input file


#copies the values of the columns we are interested in


import xlwt #library needed for writing a new excel file
workbook = xlwt.Workbook()



what_sheet_are_we_on = 0
for x in list_of_sheet_names:
   sheet = wb.sheet_by_index(what_sheet_are_we_on)
   sheet1 = workbook.add_sheet(str(x))
   file_name_column = sheet.col_values(0)
   response_column = sheet.col_values(4)
   itsd_response_column = sheet.col_values(5)
   sample_id_column = sheet.col_values(17)

   counter = 0
   for value in range(0,len(file_name_column)):
    sheet1.write(value,0,str(file_name_column[counter]))
    sheet1.write(value,1,str(response_column[counter]))
    sheet1.write(value,2,str(itsd_response_column[counter]))
    sheet1.write(value,3, str(sample_id_column[counter]))
    counter += 1

   workbook.save(file_name)
   what_sheet_are_we_on += 1



print "Completed. Your file will be saved in the same directory as this script."
