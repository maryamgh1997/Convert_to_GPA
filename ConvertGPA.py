import xlrd

def convert_to_gpa():


    # Give the location of the file
    loc = (r"C:\Users\ahmad\Desktop\Convert_to_GPA/Grades.xlsx")
    # To open Workbook
    wb = xlrd.open_workbook(loc)
    sheet = wb.sheet_by_index(0)

    # For row 0 and column 0
    sheet.cell_value(0, 0)
    multiple_list = []     # a list to store grade*credits in it
    gpa_list =[]           #a list to store grades
    al_grades = []         #a list to store american type grades (including A, B, C, D)


    #to store grades of student in a list
    for i in range(1, sheet.nrows):
        gpa_list.append(sheet.cell_value(i, 2))


    # to transform grades in range 0 to 20 into 1 to 4 and A to D
    for i in range(len(gpa_list)):
        if 16.0<= gpa_list[i] <=20.0:
            gpa_list[i] = 4
            al_grades.append('A')
        elif 14.0<= gpa_list[i] <16.0:
            gpa_list[i] = 3
            al_grades.append('B')
        elif 12.0<= gpa_list[i] <14.0:
            gpa_list[i] = 2
            al_grades.append('C')
        elif 10.0<= gpa_list[i] <12.0:
            gpa_list[i] = 1
            al_grades.append('D')


    # to make a multipication list to store grades*credits of each lesson in it
    for i in range(1, sheet.nrows):
        multiple_list.append(sheet.cell_value(i, 1)*gpa_list[i-1])

    # to calculate the total GPA of student
    total_gpa = sum(multiple_list)/sheet.cell_value(1, 3)

    print(total_gpa)
    print("")
    print(gpa_list , al_grades)


convert_to_gpa()
