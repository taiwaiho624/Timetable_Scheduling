import itertools
import csv
import sys
import xlwt
from xlwt import Workbook 

class Course:
    def __init__(self, code, name, schedule):
        self.code = code
        self.name = name
        self.schedule = schedule
    def __str__(self):
        return self.code + " " 


class Timetable:
    def __init__(self, courselist = []):
        self.courselist = courselist

    def __str__(self):
        output = ""
        for course in self.courselist:
            output = output + course.code + " "
        return output

    def add_course(self, course):
        return self.courselist.append(course)

    def check_free_day(self):
        daylist = []
        free_day_list = []
        for course in self.courselist:
            schedule = course.schedule
            for day in schedule:
                if (day['day'] not in daylist):
                    daylist.append(day['day'])
        for i in range(1,6):
            if ( i not in daylist):
                free_day_list.append(i)
        return free_day_list

def check_two_course_conflict(courseA, courseB):
    conflict = False
    scheduleA = courseA.schedule
    scheduleB = courseB.schedule
    for courseA_day in scheduleA:
        for courseB_day in scheduleB:
            if courseA_day['day'] == courseB_day['day']:
                courseA_starttime = courseA_day['time'][0]
                courseA_endtime = courseA_day['time'][1]
                courseB_starttime = courseB_day['time'][0]
                courseB_endtime = courseB_day['time'][1]
                if( (courseB_starttime >= courseA_starttime and courseB_starttime < courseA_endtime) or
                    (courseA_starttime >= courseB_starttime and courseA_starttime < courseB_endtime )
                  ):
                    conflict = True
                break
    return conflict

def display_result_list(result_list):
    for timetable in result_list:
        print(timetable)

def generate_all_combinations(courselist, num_course):
    result_list = []
    timetable_list = list(itertools.combinations(courselist, num_course))
    for timetable in timetable_list:
        out = list(itertools.chain(timetable)) 
        temp_timetable = Timetable(out)
        result_list.append(temp_timetable)
    return result_list
        
def generate_all_dependency(courselist):
    result_dict = {}
    for course_a in courselist:  
        dependency_list = []
        for course_b in courselist:
            if (course_a != course_b):
                if(check_two_course_conflict(course_a, course_b)):
                    dependency_list.append(course_b)
        result_dict[course_a.code] = dependency_list
    return result_dict

def remove_conflict(all_timetables, dependencies):
    remove_list = []
    result_list = []
    for timetable in all_timetables:
        remove = False
        for course_a in timetable.courselist:
            for course_b in timetable.courselist:
                if (course_a != course_b):
                    if (course_b in dependencies[course_a.code]):
                        remove = True
        if(remove):
            remove_list.append(timetable)
    for timetable in all_timetables:
        if timetable not in remove_list:
            result_list.append(timetable)
    return result_list


def chunks(lst, n):
    for i in range(0, len(lst), n):
        yield lst[i:i + n]


def write_to_excel(result_list=[]):
    wb = Workbook() 
    i = 1

    for timetable in result_list:
        sheet1 = wb.add_sheet('時間表' + str(i))   
        style = xlwt.easyxf('font: bold 1') 
        style_head = xlwt.easyxf('alignment: wrap True')

        sheet1.col(1).width = 256 * 30
        sheet1.col(2).width = 256 * 30
        sheet1.col(3).width = 256 * 30
        sheet1.col(4).width = 256 * 30
        sheet1.col(5).width = 256 * 30
        sheet1.col(6).width = 256 * 30
        sheet1.col(7).width = 256 * 50
        
        sheet1.row(0).height_mismatch = True
        sheet1.row(0).height = 256*3
        sheet1.row(1).height_mismatch = True
        sheet1.row(1).height = 256*3
        sheet1.row(2).height_mismatch = True
        sheet1.row(2).height = 256*3
        sheet1.row(3).height_mismatch = True
        sheet1.row(3).height = 256*3
        sheet1.row(4).height_mismatch = True
        sheet1.row(4).height = 256*3
        sheet1.row(5).height_mismatch = True
        sheet1.row(5).height = 256*3
        sheet1.row(6).height_mismatch = True
        sheet1.row(6).height = 256*3
        sheet1.row(7).height_mismatch = True
        sheet1.row(7).height = 256*3
        sheet1.row(8).height_mismatch = True
        sheet1.row(8).height = 256*3
        sheet1.row(9).height_mismatch = True
        sheet1.row(9).height = 256*3
        sheet1.row(10).height_mismatch = True
        sheet1.row(10).height = 256*3


        sheet1.write(1, 0, '830 - 920', style) 
        sheet1.write(2, 0, '930 - 1020', style) 
        sheet1.write(3, 0, '1030 - 1120', style) 
        sheet1.write(4, 0, '1130 - 1220', style) 
        sheet1.write(5, 0, '1230 - 1320', style) 
        sheet1.write(6, 0, '1330 - 1420', style) 
        sheet1.write(7, 0, '1430 - 1520', style) 
        sheet1.write(8, 0, '1530 - 1620', style) 
        sheet1.write(9, 0, '1630 - 1720', style) 
        sheet1.write(10, 0, '1730 - 1820', style) 
        sheet1.write(0, 1, 'Monday', style) 
        sheet1.write(0, 2, 'Tuesday', style) 
        sheet1.write(0, 3, 'Wednesday', style) 
        sheet1.write(0, 4, 'Thursday', style) 
        sheet1.write(0, 5, 'Friday', style)  
        sheet1.write(0, 6, 'Saturday', style) 
        sheet1.write(0, 7, 'Sunday', style) 
        i = i + 1
        
        for course in timetable.courselist:
            display_name = course.code + "\n" + course.name
            for one_day in course.schedule:
                day = one_day['day']
                startbox = int((one_day['time'][0] - 830)/100 + 1 )
                endbox = int((one_day['time'][1] - 920)/100 + 2)
                for j in range(startbox, endbox):
                    sheet1.write(j, day, display_name, style_head)
    if(result_list != []):
        wb.save('output.xls')

input_file = sys.argv[1]
num_course = int(sys.argv[2])
courselist = []

with open(input_file) as csvfile:
    rows = csv.reader(csvfile)
    for row in rows:
        code = row[0]
        name = row[1]
        schedule = []
        l = chunks(row[2:], 3)
        for s in l:
            day = int(s[0])
            starttime = int(s[1])
            endtime = int(s[2])
            schedule.append({
                'day' : day,
                'time' : (starttime, endtime)
            })
        new_course = Course(code,name,schedule)
        courselist.append(new_course)    

all_time_timetables = generate_all_combinations(courselist, num_course)
dependencies = generate_all_dependency(courselist)
result_list = remove_conflict(all_time_timetables, dependencies)
write_to_excel(result_list)