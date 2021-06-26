import datetime
import xlrd
import xlsxwriter
import re

#path = '14decraw.xlsx' #Get filename here

path = input("Enter File Name: ") + ".xlsx"

file = xlrd.open_workbook(path) 
sheet = file.sheet_by_index(0) #First sheet in excel file i.e index 0

information_list = []
detailed_list = []
driver_list = []
price_list = []
return_list = []
count = 0
count2 = 0

for i in range(sheet.nrows):
    information_list.append(sheet.cell_value(i,1)) #column here, (i,0) for timing
    s = information_list[i][:4] #checking only first 4 indexes
    for element in (s): 
        if element.isdigit():
            driver_list.append(element)
            break
    else:
        driver_list.append(" ")
    
    detailed_list.append(information_list[i].split(' '))

for i in range(sheet.nrows):
    try:
        if 'return' in detailed_list[i][0].lower() or "rtn" in detailed_list[i][0].lower() or 'return' in detailed_list[i][1].lower() or "rtn" in detailed_list[i][1].lower():
            return_list.append('yes')
        else:
            return_list.append('no')
    except:
        return_list.append('no')

    #price
    dollar_sign = information_list[i].find('$')
    if "/" in information_list[i] and 'way' in information_list[i] and return_list[i] == 'yes':
        initial_price_index = dollar_sign+1
        price_index = initial_price_index
        while True:
            price_str = information_list[i][price_index]
            if price_str.isdigit():
                price_index+=1
            else:
                break
        if price_index > initial_price_index:
            count+=1
            price_list.append(information_list[i][initial_price_index:price_index]) #string slicing
        else:
            price_list.append(" ")
    elif return_list[i] == 'yes':
        count2+=1
        price_list.append("RETURN")
    else:
        price_list.append(" ")

print("Debug:")        
print("Price entries count:",count)   
print("RETURN entries count:",count2)  
print("Success!")

#Create excel file
workbookOut = xlsxwriter.Workbook("sorted"+path)
worksheetOut = workbookOut.add_worksheet("Sheet1")

for i in range(len(information_list)):
    worksheetOut.write("A"+str(i),driver_list[i]) #add value to i if need more rows
    worksheetOut.write("B"+str(i),price_list[i]) 

workbookOut.close()
