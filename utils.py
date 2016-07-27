from tkinter import filedialog
from openpyxl import *
from datetime import datetime, timedelta

import functools
import string

# define options for opening or saving a file
options = {}
options['defaultextension'] = '.xlsx'
options['filetypes'] = [('Excel files', '.xlsx .xls'), ('all files', '.*')]


def get_file_path(title):
    options['title'] = title
    path = filedialog.askopenfilename(**options)
    return path

def get_file_directory(title):
    options['title'] = title
    directory = filedialog.askdirectory()
    return directory

def verify_paths(path1, path2, path3):
    first_valid = verify_excel_file(path1)
    second_valid = verify_excel_file(path2)
    third_valid = verify_excel_file(path3)
    return first_valid and second_valid and third_valid

def verify_excel_file(path):
    if path == '':
        return False
    file_ending = path.split('.')[1]
    if file_ending == 'xlsx' or file_ending == 'xls':
        return True
    return False 

def verify_admit_data(path):
    try:
        work_book = load_workbook(path, read_only=True)
        w_sheet = work_book.active
        print(type(w_sheet['A2'].value) is int)

        print(type(w_sheet['B2'].value) is datetime)
        
        print(type(w_sheet['C2'].value) is datetime)
    except FileNotFoundError:
        return False


def process_data(admit_path, line_path, event_path, out_path):
    events = {}

    patients = read_patient_data(admit_path)
    read_line_data(line_path, patients)    
    read_event_data(event_path, patients, events)
    generate_output(out_path, patients, events)

def read_patient_data(path):
    work_book = load_workbook(path, read_only=True)
    w_sheet = work_book.active
    patients = {}
    index = 2
    while index <= w_sheet.max_row:
        p_id = w_sheet['A' + str(index)].value
        in_date = w_sheet['B' + str(index)].value
        out_date = w_sheet['C' + str(index)].value
        if not p_id in patients:
            p = Patient(p_id)
            patients[p_id] = p
        if check_full_day_admit(in_date, out_date):
            p.add_visit(Visit(patients[p_id], in_date, out_date))
        index += 1
    return patients

def read_line_data(path, patients):
    work_book = load_workbook(path, read_only=True)
    w_sheet = work_book.active
    index = 2
    while index <= w_sheet.max_row:
        p_id = w_sheet['A' + str(index)].value
        line_id = w_sheet['B' + str(index)].value
        line_type = w_sheet['C' + str(index)].value
        lumens = w_sheet['D' + str(index)].value
        in_date = w_sheet['E' + str(index)].value
        out_date = w_sheet['F' + str(index)].value 
        if out_date is None:
            out_date = w_sheet['G' + str(index)].value
        
        l = Line(line_id, line_type, lumens, in_date, out_date)
        patients[p_id].add_line(l)
        index += 1

def read_event_data(path, patients, events):
    work_book = load_workbook(path, read_only=True)
    w_sheet = work_book.active
    index = 2
    while index <= w_sheet.max_row:
        p_id = w_sheet['A' + str(index)].value
        event_date = w_sheet['B' + str(index)].value
        event_type = w_sheet['C' + str(index)].value
        if event_type not in events:
            events[event_type] = [Event(p_id, event_date, event_type)]
        else:
            events[event_type].append(Event(p_id, event_date, event_type))
        if event_type not in patients[p_id].events:
            patients[p_id].events[event_type] = [Event(p_id, event_date, event_type)]
        else:
            patients[p_id].events[event_type].append(Event(p_id, event_date, event_type))
        index += 1


def check_full_day_admit(in_time, out_time):
    diff = out_time - in_time
    if diff.days == 0:
        return False
    return True

def generate_output(path, patients, events):
    work_book = Workbook(data_only=True)
    w_sheet = work_book.active

    #Column Titles
    w_sheet.title = 'Output Individual Patient'
    w_sheet['A1'] = 'Patient ID'
    w_sheet['B1'] = 'Total Lines'
    w_sheet['C1'] = 'Sum of all Line Days' 
    w_sheet['D1'] = 'Mean Duration of Line (Days)'
    w_sheet['E1'] = 'Total Days with any Catheter'
    w_sheet['F1'] = 'Catheter Density (Sum of all Line Days/Total Days with any catheter)' 
    w_sheet['G1'] = 'Sum of all Lumen Days'
    w_sheet['H1'] = 'Lumen Density (Sum of all Lumen Days/Total Days with any cather)'

    event_indecies = {}

    for event_type in events:
        bottom = len(patients) + 2
        first_col = w_sheet.max_column + 1
        second_col = w_sheet.max_column + 2
        event_indecies[event_type] = first_col
        w_sheet.cell(row=1, column=first_col).value = event_type + 's'
        w_sheet.cell(row=1, column=second_col).value = event_type + ' Rate'
        total = len(events[event_type])

        rate_formula = '=' + string.ascii_uppercase[first_col - 1] + str(bottom) + "/E" + str(bottom)
        w_sheet.cell(row=bottom, column=first_col).value = total
        w_sheet.cell(row=bottom, column=second_col).value = rate_formula

    index = 2
    for p_id in patients:
        p = patients[p_id]

        w_sheet['A' + str(index)] = p_id
        w_sheet['B' + str(index)] = len(p.lines)
        w_sheet['C' + str(index)] = p.total_line_time.days
        w_sheet['D' + str(index)] = p.total_line_time.days/len(p.lines)
        w_sheet['E' + str(index)] = calculate_total_cath_daysI(p)
        w_sheet['F' + str(index)] = w_sheet['C' + str(index)].value / w_sheet['E' + str(index)].value
        w_sheet['G' + str(index)] = p.total_lumen_time.days
        w_sheet['H' + str(index)] = w_sheet['G' + str(index)].value / w_sheet['E' + str(index)].value
        for event_type in patients[p_id].events:
            first_col = event_indecies[event_type]
            second_col = event_indecies[event_type] + 1
            rate_formula = '=' + string.ascii_uppercase[first_col - 1] + str(index) + "/E" + str(index)
            w_sheet.cell(row=index, column=first_col).value = len(p.events[event_type])
            w_sheet.cell(row=index, column=second_col).value = rate_formula
        index += 1



    #Summation Data
    max_index = str(w_sheet.max_row)
    bottom = w_sheet.max_row - 1
    w_sheet['A' + max_index] = 'Population Total'
    w_sheet['B' + max_index] = '=SUM(B2:B' + str(bottom) + ')' 
    w_sheet['C' + max_index] = '=SUM(C2:C' + str(bottom) + ')'
    w_sheet['D' + max_index] = '=C' + str(bottom) + '/B' + str(bottom)
    w_sheet['E' + max_index] = '=SUM(E2:E' + str(bottom) + ')'
    w_sheet['F' + max_index] = '=C'+ str(max_index) + '/E' + str(max_index)
    w_sheet['G' + max_index] = '=SUM(G2:G' +  str(bottom) + ')'
    w_sheet['H' + max_index] = '=G'+ str(max_index) + '/E' + str(max_index)


    #adjust cell width for titles
    index = 1
    for col in w_sheet.columns:
        w_sheet.column_dimensions[cell.get_column_letter(index)].width = len(col[0].value)
        index += 1

    work_book.save(path + "/Output Individual Patient.xlsx")

def calculate_total_cath_daysI(p):
    #need to test on large data sample
    sorted(p.lines)
    total_days = p.lines[0].out_date - p.lines[0].in_date
    try:
        earliest = p.lines[0]
        current = p.lines[0]
    except IndexError:
        return 0
    for l in p.lines:
        if current.out_date < l.in_date:
            total_days += l.out_date - l.in_date
            earliest = l
            current = l
        else:
            total_days += l.out_date - current.out_date
            current = l
    return total_days.days


class Patient():
    
    def __init__(self, patient_id):
        self.visits = []
        self.lines = []
        self.events = {}
        self.patient_id = patient_id
        self.total_visit_time = None
        self.total_line_time = None
        self.total_lumen_time = None

    def add_visit(self, v):
        assert isinstance(v, Visit), "new visits must be of type Visit"
        self.visits.append(v)
        if self.total_visit_time is None:
            self.total_visit_time = v.total_time
        else:
            self.total_visit_time += v.total_time

    def add_line(self, l):
        assert isinstance(l, Line), "new lines must be of type Line"
        self.lines.append(l)
        if self.total_line_time is None:
            self.total_line_time = l.total_time
            self.total_lumen_time = l.lumen_days
        else:
            self.total_line_time += l.total_time
            self.total_lumen_time += l.lumen_days


class Visit():

    def __init__(self, patient, in_date, out_date):
        self.patient = patient
        self.check_in_date = in_date
        self.check_out_date = out_date
        self.total_time = in_date - out_date

@functools.total_ordering
class Line():

    def __init__(self, line_id, line_type, lumens, in_date, out_date):
        self.in_date = in_date
        self.line_type = line_type
        self.out_date = out_date
        self.line_id = line_id
        self.lumens = lumens
        self.total_time = out_date - in_date
        self.lumen_days = self.total_time * self.lumens


    def __lt__(self, other):
        if self.in_date ==  other.in_date:
            return self.out_date < other.out_date
        return self.in_date < self.out_date

    def __eq__(self, other):
        return self.in_date == other.in_date and self.out_date == other.out_date

class Event():

    def __init__(self, patient, date, event_type):
        self.patient = patient
        self.event_date = date
        self.event_type = event_type

