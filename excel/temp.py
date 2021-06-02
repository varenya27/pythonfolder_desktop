import datetime 
import xlsxwriter
from xlsxwriter import workbook

#converts tuples into a string
def tup_str(t):
    n=len(t)
    lst=list()
    for i in range(n):
        lst.append(t[i])
    word=''
    for term in t:
        if type(term)==int:
            term=str(term)
        word+=term
    return word

#function to return the day of the week for an integer
def day(n):
    if n==0:
        return 'MONDAY'
    elif n==1:
        return 'TUESDAY'
    elif n==2:
        return 'WEDNESDAY'
    elif n==3:
        return 'THURSDAY'
    elif n==4:
        return 'FRIDAY'
    elif n==5:
        return 'SATURDAY'
    elif n==6:
        return 'SUNDAY'
    else: return 'invalid'

#inputs the days into excel from a specified row/column vertically
def day_excel(row,column,num_days):
    for i in range(num_days):
        worksheet.write(row+i,column,day(i))    

#to use stuff like 1430+130=1500
def time_converter(start, dur):
    s_h=start//100
    d_h = dur//100
    e_h = s_h+d_h
    sm = start%100
    dm=dur%100
    em1= sm+dm
    em_h=em1//60
    em_m = em1%60
    end= 100*(e_h+em_h)+em_m
    return end

#returns a list with time slots as strings
def time(num_periods,start_time,len_period):
    time_tuple=list()
    for i in range(num_periods):
        t=(time_converter(start_time+len_period*i,0), '-', time_converter(start_time+len_period*i, len_period))
        time_tuple.append(t)
    time_strings=list()
    for tuple in time_tuple:
        time_strings.append(tup_str(tuple))
    return time_strings

#inputs the time slots from specified row/column horizontally
def time_excel(row, column, num_periods, first_hour,len_period):
    
    lst=time(num_periods, first_hour,len_period)
    for slot in lst:
        worksheet.write(row,column,slot) 
        column+=1

#fill one input in one cell
def fill (letter_notation,entry):
    worksheet.write(letter_notation,entry)

#fill one input in 2 cells
def fill2(c1, c2, code):
    worksheet.write(c1, code)
    worksheet.write(c2,code)

#fill one input in 3 cells
def fill3(c1,c2,c3, code):
    worksheet.write(c1, code)
    worksheet.write(c2, code)
    worksheet.write(c3,code)

#fill course code correspnding to slot
def slot_fill(slot, code):
    code=str(code)
    if slot == 'A':
        fill3('B2', 'D4', 'C5',code)
    elif slot == 'B':
        fill3('C2', 'B4', 'D5', code)
    elif slot == 'C':
        fill3('D2','C4','B5',code)
    elif slot == 'D':
        fill3('B3','E2','D6',code)
    elif slot == 'E':
        fill3('C3','E5','B6',code)
    elif slot == 'F':
        fill3('D3','G4','C6',code)
    elif slot == 'G':
        fill3('E3','E4','E6',code)
    elif slot == 'P':
        fill2('G2','H5',code)
    elif slot == 'Q':
        fill2('H2','G5',code)
    elif slot == 'R':
        fill2('G3','H6',code)
    elif slot == 'S':
        fill2('H3','G6',code)
    elif slot == 'W':
        fill2('J2','J5',code)
    elif slot == 'X':
        fill2('K2','K5',code)
    elif slot == 'Y':
        fill2('J3','J6',code)
    elif slot == 'Z':
        fill2('K3','K6',code)


#create an new Excel file and add a worksheet.
workbook = xlsxwriter.Workbook('timetable1.xlsx')
worksheet = workbook.add_worksheet()

#setting a template corresponding to the IITH timetable
worksheet.write(0,0,'DAY\TIME')
day_excel(1,0,5)
time_excel(0,1,4,900,100)
time_excel(0,6,2,1430,130)
time_excel(0,9,2,1800,130)

#taking the courses along with their codes as inputs and filling in the sheet
n=input("Enter the number of courses: ")
n=int(n)
for i in range(n):
    slot=input("Enter slot: ")
    code = input("Enter course code: ")
    slot_fill(slot,code)


workbook.close()