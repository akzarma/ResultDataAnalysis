import copy


# this is extra library which must be installed in python3 so that we can create excel files and write on them
# to download this library :


# check if pip3.exe is present in C:/python/scripts     directory
# if yes, add C:\python\Scripts to your environmental variable PATH . so that you use pip3 command

# now type pip3  to check whether it is installed or not
# now to install xlsxwriter type: pip3 install xlsxwriter

import xlsxwriter


# error in result files for two students of entc  , list of them are
# entc
# S150393045
# S150393030
# so these students  result should be deleted from file or code can be optimised to handle these type of errors in result file

# this is the text variable which contains the result file text , skip down to go to coding part
text = '''Entc last result paste it here.'''

#class of particular subject to store different type of marks
class Subject:
    def __init__(self):
        self.sub_nm = ''
        self.oe = ''
        self.oet = ''
        self.th = ''
        self.tht = ''
        self.oe_th = ''
        self.oe_tht = ''
        self.tw = ''
        self.twt = ''
        self.pr = ''
        self.prt = ''

        self.total_percent = ''
        self.crd = ''
        self.grd = ''
        self.grd_pts = ''
        self.crd_pts = ''

#just a function like toString() in c++
    def __str__(self):
        return self.total_percent + " " + self.grd


#class of a student with details like name seat no., and an array of subject
class Student:
    def __init__(self):
        self.name = ''
        self.seat_no = ''
        self.admission_no = ''
        self.clg = ''
        self.sgpa = ''
        self.sgpa2 = 0
        self.credits = ''
        self.div = ''
        self.subject = []

    def __str__(self):
        return self.name + " " + self.seat_no + " " + self.admission_no + " " + self.clg + \
               " " + self.sgpa + " " + self.credits


#variable declaration
students = []

#to analyze data, division wise
studentsA = []
studentsB = []
studentsC = []

#temporary subject array
subjects = []
student_temp = Student()
count = 0
offset = 0

#splits the whole text by new line and stores it in x which gets incremented after each loop
for x in text.split("\n"):

    if x.__contains__("SAVITRIBAI"):
        count = 0
    if x.__contains__("............................"):
        count = count + 1
        if count == 2:
            student_temp.subject = subjects
            # if count == 2 then it means that data of one student is completed , so we should save it in temp variable
            # and append it to our final list of students
            temp = copy.deepcopy(student_temp)
            students.append(temp)

            # clear the subjects array so that we can again use subjects variable to store subject list with marks of next student
            subjects.clear()
            count = 1
            continue

    else:

        # split() function splits the the whole line into words which can accessed by index starting from 0
        word = x.split()

        if word[0].__contains__("S15039"):
            student_temp.seat_no = word[0]
            i = 1
            student_temp.name = ''
            while 1:
                # exception names containing "." or "-" thats why more conditions in if condition
                if word[i].isalpha() or word[i] == "." or word[i].__contains__("-"):
                    #appending the name of student and last name and mother's name.
                    student_temp.name += word[i] + " "
                else:
                    #if word[i] doesn't contain pure alphabets then name traversing is completed
                    break
                i += 1

            # as the name traversing is completed now the next word will be students admission_no. and college name
            student_temp.admission_no = word[i]
            student_temp.clg = word[i + 1]

        # if word[0] doesn't containg "s15039" and if it contains digits then itt means, it is the subject code
        elif word[0].isnumeric():

            #creating temporary object of Subject class
            subject_temp = Subject()

            #as the word[0] contains numeric that means it will be subject code so we'll store it in sub_nm
            subject_temp.sub_nm = word[0]

            # this is some exceptional lines in result file where a "*" is present after subject code
            if word[1] == '*':
                offset = 0
            else:
            # if there is no "*" in subject line... then we will simply subtract one from each index because there is one less word in this case.
                offset = -1

            # word.split("/") will split the the marks combination like 27/50 in two parts
            # first part i.e. 27 will be word.split("/")[0]
            # second part i.e. 50 will be word.split("/")[1]
            subject_temp.oe = word[2 + offset].split("/")[0]
            if word[2 + offset].__contains__("/"):
                subject_temp.oet = word[2 + offset].split("/")[1]

            subject_temp.th = word[3 + offset].split("/")[0]
            if word[3 + offset].__contains__("/"):
                subject_temp.tht = word[3 + offset].split("/")[1]

            subject_temp._oeth = word[4 + offset].split("/")[0]
            if word[4 + offset].__contains__("/"):
                subject_temp.oe_tht = word[4 + offset].split("/")[1]

            subject_temp.tw = word[5 + offset].split("/")[0]
            if word[5 + offset].__contains__("/"):
                subject_temp.twt = word[5 + offset].split("/")[1]

            subject_temp.pr = word[6 + offset].split("/")[0]
            if word[6 + offset].__contains__("/"):
                subject_temp.prt = word[6 + offset].split("/")[1]

            # ignoring oral or

            # now simply storing total_percent etc in our temporary student_temp object
            subject_temp.total_percent = word[8 + offset]

            subject_temp.crd = word[9 + offset]

            subject_temp.grd = word[10 + offset]

            subject_temp.grd_pts = word[11 + offset]

            subject_temp.crd_pts = word[12 + offset]

            # now one subject data is stored subject_temp so we will copy it in new temp variable
            temp = copy.deepcopy(subject_temp)
            # now that temp variable is appended in  our final subjects list
            subjects.append(temp)


        # now in result file if the word[0] i.e. first word is SGPA then we know that next words contains SGPA of student
        # so we will take index according to file and store it in variables of student_temp object
        elif word[0].__contains__("SGPA"):
            student_temp.sgpa = word[2].split(",")[0]
            student_temp.credits = word[6]





# sgpa of sem 2 only calculation
# stu = 0
# while stu < len(students):
#     sub = 12
#     while sub < len(students[0].subject):
#         students[stu].sgpa2 = float(students[stu].sgpa2) + float(students[stu].subject[sub].crd_pts)
#         sub += 1
#     students[stu].sgpa2 = float(students[stu].sgpa2)/25.0
#     stu += 1




# adding div field in Student class

students_without_div = copy.deepcopy(students)








# ============================= ANALYSIS PART ====================================================


# now a class is declared which stores the sum of specific grades of a particular subject
class Grade:
    def __init__(self):
        self.o_grade = 0
        self.ap_grade = 0
        self.a_grade = 0
        self.bp_grade = 0
        self.b_grade = 0
        self.c_grade = 0
        self.p_grade = 0
        self.f_grade = 0
        self.absent = 0

# class of subject which contains the details of a subject
class SubjectAnalyse:
    def __init__(self):
        self.sub_nm = ''
        self.grades = Grade()
        self.pass_percent = 0.0
        self.appeared = 0


# Every file generator. this function does analysis.
def analysis(student_div):
    subjectsAnalysis = []
    c = 0
    while c < len(student_div[0].subject):
        temp = SubjectAnalyse()
        subjectsAnalysis.append(copy.deepcopy(temp))
        c += 1

    sub = 0
    stu = 0

    total_fails = 0
    total_pass = 0

    distinction = 0

    while stu < len(student_div):
        sub = 0
        while sub < len(student_div[0].subject):
            # print(students[stu].seat_no)
            if student_div[stu].subject[sub].grd == 'O':
                subjectsAnalysis[sub].grades.o_grade += 1
            elif student_div[stu].subject[sub].grd == 'A+':
                subjectsAnalysis[sub].grades.ap_grade += 1
            elif student_div[stu].subject[sub].grd == 'A':
                subjectsAnalysis[sub].grades.a_grade += 1
            elif student_div[stu].subject[sub].grd == 'B+':
                subjectsAnalysis[sub].grades.bp_grade += 1
            elif student_div[stu].subject[sub].grd == 'B':
                subjectsAnalysis[sub].grades.b_grade += 1
            elif student_div[stu].subject[sub].grd == 'C':
                subjectsAnalysis[sub].grades.c_grade += 1
            elif student_div[stu].subject[sub].grd == 'P':
                subjectsAnalysis[sub].grades.p_grade += 1
            elif student_div[stu].subject[sub].grd == 'F':
                subjectsAnalysis[sub].grades.f_grade += 1
            if student_div[stu].subject[sub].th == 'AB':
                subjectsAnalysis[sub].grades.absent += 1
            elif student_div[stu].subject[sub].pr == 'AB':
                subjectsAnalysis[sub].grades.absent += 1

            subjectsAnalysis[sub].sub_nm = student_div[stu].subject[sub].sub_nm

            sub += 1

        if student_div[stu].sgpa == '--':
            total_fails += 1
        else:
            total_pass += 1

        stu += 1

    f_class = 0
    hs_class = 0
    s_class = 0
    atkt = 0
    failure = 0
    def getSGPA(s):
        if s.sgpa == '--':
            return 0
        else:
            return float(s.sgpa)


    # simple calculations of distinction and all things
    stu = 0
    while stu < len(student_div):
        if getSGPA(student_div[stu])>= 7.75:
            distinction += 1
        elif 6.75<=getSGPA(student_div[stu])< 7.75:
            f_class += 1
        elif 6.25<=getSGPA(student_div[stu])< 6.75:
            hs_class += 1
        elif 5.5<=getSGPA(student_div[stu])< 6.25:
            s_class += 1
        elif getSGPA(student_div[stu]) == 0 and int(student_div[stu].credits)>=25:
            atkt += 1
        else:
            failure += 1
        stu += 1



    # writing to xlsx file, firstly open it and take it in a variable named "workbook"
    workbook = xlsxwriter.Workbook('./Grade_analysis_' + student_div[0].div + '.xlsx')
    worksheet = workbook.add_worksheet()

    worksheet.write(1, 2, "Subject")
    worksheet.write(1, 3, "O Grade")
    worksheet.write(1, 4, "A+ Grade")
    worksheet.write(1, 5, "A Grade")
    worksheet.write(1, 6, "B+ Grade")
    worksheet.write(1, 7, "B Grade")
    worksheet.write(1, 8, "C Grade")
    worksheet.write(1, 9, "P Grade")
    worksheet.write(1, 10, "Fail")
    worksheet.write(1, 11, "Absent")
    worksheet.write(1, 12, "Appeared")
    worksheet.write(1, 13, "Passing %")

    row = 2
    col = 2
    i = 0

    while i < len(subjectsAnalysis):
        worksheet.write(row, col, subjectsAnalysis[i].sub_nm)
        worksheet.write(row, col + 1, int(subjectsAnalysis[i].grades.o_grade))
        worksheet.write(row, col + 2, int(subjectsAnalysis[i].grades.ap_grade))
        worksheet.write(row, col + 3, int(subjectsAnalysis[i].grades.a_grade))
        worksheet.write(row, col + 4, int(subjectsAnalysis[i].grades.bp_grade))
        worksheet.write(row, col + 5, int(subjectsAnalysis[i].grades.b_grade))
        worksheet.write(row, col + 6, int(subjectsAnalysis[i].grades.c_grade))
        worksheet.write(row, col + 7, int(subjectsAnalysis[i].grades.p_grade))
        worksheet.write(row, col + 8, int(subjectsAnalysis[i].grades.f_grade))
        worksheet.write(row, col + 9, int(subjectsAnalysis[i].grades.absent))
        worksheet.write(row, col + 10,int(len(student_div)-subjectsAnalysis[i].grades.absent))
        worksheet.write(row, col + 11, 100*float('%.4f' % (float((len(student_div)-int(subjectsAnalysis[i].grades.absent)-int(subjectsAnalysis[i].grades.f_grade)))/int(len(student_div)-subjectsAnalysis[i].grades.absent))))
        row += 1
        i += 1

    workbook.close()




    print(('./Grade_analysis_' + student_div[0].div + '.xlsx'), " file generated")








    # Topper analysis
    def getKey(s):
        if s.sgpa == '--':
            return 0
        return float(s.sgpa)

    sorted_stu = sorted(student_div, key=getKey, reverse=True)

    workbook = xlsxwriter.Workbook('./Toppers_analysis_' + student_div[0].div + '.xlsx')
    worksheet = workbook.add_worksheet()

    i = 0

    row = 0
    col = 0

    while i < 40:
        if sorted_stu[i].sgpa == '--':
            sorted_stu[i].sgpa = 0
        worksheet.write(row, col, float(sorted_stu[i].sgpa))
        worksheet.write(row, col + 1, sorted_stu[i].seat_no)
        worksheet.write(row, col + 2, sorted_stu[i].name)
        i += 1
        row += 1

    workbook.close()





    print(('./Toppers_analysis_' + student_div[0].div + '.xlsx'), " file generated")



    workbook = xlsxwriter.Workbook("Overall_analysis_" + student_div[0].div + '.xlsx')
    worksheet = workbook.add_worksheet()

    row = 0
    col = 0

    analysis = (
        ['Total Failures', total_fails],
        ['Total No. of clear Pass', total_pass],
        ['% of Failure', float('%.2f' % (total_fails * 100 / (total_pass + total_fails)))],
        ['% of clear Passing', float('%.2f' % (total_pass * 100 / (total_pass + total_fails)))],
        ['Distinction', int(distinction)],
        ['First Class', int(f_class)],
        ['Higher second Class', int(hs_class)],
        ['Second class', int(s_class)],
        ['ATKT', int(atkt)],
        ['Failure (No ATKT)', int(failure)]
    )

    for key, value in (analysis):
        worksheet.write(row, col, key)
        worksheet.write(row, col + 1, value)
        row += 1

    workbook.close()


    print(("Overall_analysis_" + student_div[0].div + '.xlsx'), " file generated")

# analysis(studentsA)
# analysis(studentsB)
# analysis(studentsC)
analysis(students_without_div)

