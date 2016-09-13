"""Classes and methods for Central Line Event Calculator analysis"""

from tkinter import filedialog, messagebox
from openpyxl import Workbook, load_workbook, cell
from datetime import datetime

import functools
import string
import os

# define options for opening or saving a file
OPTIONS = {}
OPTIONS['defaultextension'] = '.xlsx'
OPTIONS['filetypes'] = [('Excel files', '.xlsx .xls'), ('all files', '.*')]

def error_message(title, message):
    messagebox.showwarning(title, message)

def get_file_path(title):
    """Opens TkInter dialogue window and returns user specefied file path."""
    OPTIONS['title'] = title
    path = filedialog.askopenfilename(**OPTIONS)
    return path

def get_file_directory(title):
    """Opens TkInter dialogue window and returns user specefied file directory."""
    OPTIONS['title'] = title
    directory = filedialog.askdirectory()
    return directory

def verify_paths(path1, path2, path3, path4, output):
    """Verifies all file paths are Excel files with the correct data formats."""
    first_valid = verify_excel_file(path1)
    second_valid = verify_excel_file(path2)
    third_valid = verify_excel_file(path3)
    fourth_valid = verify_excel_file(path4)
    if not first_valid or not second_valid or not third_valid or not fourth_valid:
        error_message("Invalid Data Supplied", "All input data must be Excel files.")
        return False
    if not file_exits(path1) or not file_exits(path2) or not file_exits(path3) \
    or not file_exits(path4):
        error_message(
            "Invalid Data Supplied",
            "One or more of the specified input files does not exist.\n" + \
            "Check the file locations and names and try again.")
        return False
    return True
        
def file_exits(path):
    return os.path.isfile(path)

def verify_excel_file(path):
    """Ensures that a path is an Excel file."""
    if path == '':
        return False
    try:
        file_ending = path.split('.')[1]
    except IndexError:
        return False
    if file_ending == 'xlsx' or file_ending == 'xls':
        return True
    return False

def process_data(title, admit_path, line_path, clabsi_path,  clanc_path, out_path):
    """Read in each file and writes results to the out_path."""
    events = {}
    patients = read_patient_data(admit_path)
    read_line_data(line_path, patients)
    read_clabsi_data(clabsi_path, patients)
    read_clanc_data(clanc_path, patients)
    generate_patient_output(title, out_path, patients, events)
    generate_line_output(title, out_path, patients, events)
    return True

def read_patient_data(path):
    """Read in patient admit data. Returns a dictionary of Patient objects (Key: ID Number)."""
    work_book = load_workbook(path, read_only=True)
    w_sheet = work_book.active
    patients = {}
    index = 2
    while index <= w_sheet.max_row:
        p_id = w_sheet['A' + str(index)].value
        in_date = w_sheet['B' + str(index)].value
        out_date = w_sheet['C' + str(index)].value

        # Spreadsheet format check
        if not isinstance(p_id, int):
            raise BadFormatException("Patient ID Numbers in Column A of Patient Data must be numbers.")
        if not isinstance(in_date, datetime):
            raise BadFormatException("Patient Admission Dates in Column B of Patient Data must be dates.")
        if not isinstance(in_date, datetime):
            raise BadFormatException("Patient Disscharge Dates in Column C of Patient Data must be dates.")

        if not p_id in patients:
            p = Patient(p_id)
            patients[p_id] = p
        if check_full_day_admit(in_date, out_date):
            p.add_visit(Visit(patients[p_id], in_date, out_date))
        index += 1
    return patients

def read_line_data(path, patients):
    """Read in line data. Stores lines as Line objects associated with Patient IDs."""
    work_book = load_workbook(path, read_only=True)
    w_sheet = work_book.active
    index = 2
    while index <= w_sheet.max_row and w_sheet['A' + str(index)].value is not None:
        p_id = w_sheet['A' + str(index)].value
        line_id = w_sheet['B' + str(index)].value
        line_type = w_sheet['C' + str(index)].value
        lumens = w_sheet['D' + str(index)].value
        in_date = w_sheet['E' + str(index)].value
        out_date = w_sheet['F' + str(index)].value
        if out_date is None:
            out_date = w_sheet['G' + str(index)].value
        removal_reason = w_sheet['H' + str(index)].value

        # Spreadsheet format check
        if not isinstance(p_id, int):
            raise BadFormatException("Patient ID Numbers in Column A of Line Data must be numbers.")
        if not isinstance(line_id, int):
            raise BadFormatException("Line ID Numbers in Column B of Line Data must be numbers.")
        if not isinstance(line_type, str):
            raise BadFormatException("Line Types in Column C of Line Data must be text.")
        if not isinstance(lumens, int):
            raise BadFormatException("Lumen Count in Column D of Line Data must be numbers.")
        if not isinstance(in_date, datetime):
            raise BadFormatException("Patient Admission Dates in Column E of Patient Data must be dates.")
        if not isinstance(in_date, datetime):
            raise BadFormatException("Patient Disscharge Dates in Column F and G of Patient Data must be dates.")
        if not isinstance(removal_reason, str) and removal_reason is not None:
            raise BadFormatException("Reason For Removal in Column H of Line Data must be text.")
       
        l = Line(line_id, line_type, lumens, in_date, out_date, removal_reason)
        patients[p_id].add_line(l)
        index += 1

def read_clabsi_data(path, patients):
    work_book = load_workbook(path, read_only=True)
    w_sheet = work_book.active
    index = 2
    while index <= w_sheet.max_row:
        p_id = w_sheet['A' + str(index)].value
        clabsi_date = w_sheet['B' + str(index)].value

        if not isinstance(p_id, int):
            raise BadFormatException("Patient ID Numbers in Column A of CLABSI Data must be numbers.")
        if not isinstance(clabsi_date, datetime):
            raise BadFormatException("CLABSI Date in Column B of CLABSI Data must be a date.")

        p = patients[p_id]
        lines = []

        for l in p.lines:
            if l.in_date <= clabsi_date and l.out_date >= clabsi_date:
                lines.append(l)

        event = CLABSI(p, lines, clabsi_date)

        p.clabsis.append(event)
        index += 1


def read_clanc_data(path, patients):
    work_book = load_workbook(path, read_only=True)
    w_sheet = work_book.active
    index = 2
    while index <= w_sheet.max_row:
        p_id = w_sheet['A' + str(index)].value
        line_id = w_sheet['B' + str(index)].value
        clanc_date = w_sheet['C' + str(index)].value

        if not isinstance(p_id, int):
            raise BadFormatException("Patient ID Numbers in Column A of CLANC Data must be numbers.")
        if not isinstance(line_id, int):
            raise BadFormatException("Line ID Numbers in Column B of CLANC Data must be numbers.")
        if not isinstance(clanc_date, datetime):
            raise BadFormatException("CLABSI Date in Column C of CLANC Data must be a date.")
        
        p = patients[p_id]
        line = [l for l in p.lines if line_id == l.line_id][0]

        event = CLANC(p, line, clanc_date)

        p.clancs.append(event)
        line.clanc = event
        index += 1

def check_full_day_admit(in_time, out_time):
    """Returns True if Patient was admitted for at least 24 hours."""
    diff = out_time - in_time
    if diff.days == 0:
        return False
    return True

def generate_patient_output(title, path, patients, events):
    """Writes patient-only analysis to new Excel file."""
    work_book = Workbook()
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
    w_sheet['I1'] = 'CLABSIs'
    w_sheet['J1'] = "CLABSI Rate (x1000)"
    w_sheet['K1'] = 'CLANCs'
    w_sheet['L1'] = "CLANC Rate (x1000)"

    row = 2
    for p_id in patients:
        p = patients[p_id]

        w_sheet['A' + str(row)] = p_id
        w_sheet['B' + str(row)] = len(p.lines)
        w_sheet['C' + str(row)] = p.total_line_time.days
        w_sheet['D' + str(row)] = p.total_line_time.days/len(p.lines)
        w_sheet['E' + str(row)] = calculate_total_cath_days(p)
        w_sheet['F' + str(row)] = w_sheet['C' + str(row)].value / w_sheet['E' + str(row)].value
        w_sheet['G' + str(row)] = p.total_lumen_time.days
        w_sheet['H' + str(row)] = w_sheet['G' + str(row)].value / w_sheet['E' + str(row)].value
        w_sheet['I' + str(row)] = len(p.clabsis)
        w_sheet['J' + str(row)] = w_sheet['I' + str(row)].value / w_sheet['E' + str(row)].value * 1000
        w_sheet['K' + str(row)] = len(p.clancs)
        w_sheet['L' + str(row)] = w_sheet['K' + str(row)].value / w_sheet['E' + str(row)].value * 1000
        
        row += 1

    #Summation Data
    max_index = str(w_sheet.max_row)
    bottom = w_sheet.max_row - 1
    w_sheet['A' + max_index] = 'Population Total'
    w_sheet['B' + max_index] = '=SUM(B2:B' + str(bottom) + ')'
    w_sheet['C' + max_index] = '=SUM(C2:C' + str(bottom) + ')'
    w_sheet['D' + max_index] = '=C' + str(max_index) + '/B' + str(max_index)
    w_sheet['E' + max_index] = '=SUM(E2:E' + str(bottom) + ')'
    w_sheet['F' + max_index] = '=C'+ str(max_index) + '/E' + str(max_index)
    w_sheet['G' + max_index] = '=SUM(G2:G' +  str(bottom) + ')'
    w_sheet['H' + max_index] = '=G'+ str(max_index) + '/E' + str(max_index)
    w_sheet['I' + max_index] = '=SUM(I2:I' +  str(bottom) + ')'
    w_sheet['J' + max_index] = '=I'+ str(max_index) + '/E' + str(max_index)  + "* 1000"
    w_sheet['K' + max_index] = '=SUM(K2:K' +  str(bottom) + ')'
    w_sheet['L' + max_index] = '=K'+ str(max_index) + '/E' + str(max_index) + "* 1000"

    #adjust cell width for titles
    index = 1
    for col in w_sheet.columns:
        w_sheet.column_dimensions[cell.get_column_letter(index)].width = len(col[0].value)
        index += 1

    work_book.save(path + "/" + title + " - Output Individual Patient.xlsx")

def generate_line_output(title, path, patients, events):
    """Writes line-only analysis to new Excel file."""
    work_book = Workbook()
    w_sheet = work_book.active

     #Column Titles
    w_sheet.title = 'Output Individual Line'
    w_sheet['A1'] = 'Line ID'
    w_sheet['B1'] = 'Patient ID'
    w_sheet['C1'] = 'Number of Lumens'
    w_sheet['D1'] = 'Date of Insertion (or first evaluation)'
    w_sheet['E1'] = 'Date of Removal (or last evalulation)'
    w_sheet['F1'] = 'Line Days (any catheter)'
    w_sheet['G1'] = 'Lumen Days (Line Days x Number of Lumens)'

    w_sheet['H1'] = "Number CLABSIs"
    w_sheet['I1'] =  "Number of CLANCs"
    w_sheet['J1'] =  "Time from CLANC to line removal (Days)"
    w_sheet['K1'] =  "Reason For Line Removal"

    w_sheet['L1'] =  "ALL EVENTS"
    w_sheet['M1'] =  "ALL EVENT RATE (x1000)"

    row = 2
    for p_id in patients:
        p = patients[p_id]
        for l in p.lines:
            w_sheet['A' + str(row)] = l.line_id
            w_sheet['B' + str(row)] = p_id
            w_sheet['C' + str(row)] = l.lumens
            w_sheet['D' + str(row)] = l.in_date
            w_sheet['E' + str(row)] = l.out_date
            w_sheet['F' + str(row)] = l.total_time.days
            w_sheet['G' + str(row)] = l.lumen_days.days

            num_events = 0
            for e in p.clabsis:
                if l in e.lines:
                    num_events += 1/(len(e.lines))

            w_sheet['H' + str(row)] = num_events

            num_clancs = 0
            if l.clanc is not None:
                num_clancs = 1
            w_sheet['I' + str(row)] = num_clancs
            if l.clanc.date is not None:
                diff = l.out_date - l.clanc.date
                w_sheet['J' + str(row)] = (l.out_date - l.clanc.date).days
            else:
                 w_sheet['J' + str(row)] = "No CLANC Reported"
            w_sheet['K' + str(row)] = l.removal_reason

            total_events = num_clancs + num_events
            w_sheet['L' + str(row)] = total_events
            w_sheet['M' + str(row)] = (total_events / l.total_time.days) * 1000

            w_sheet['D' + str(row)].number_format = 'dd-mmm-yy'
            w_sheet['E' + str(row)].number_format = 'dd-mmm-yy'
            # w_sheet['J' + str(row)].number_format = 'dd'

            row += 1

    #adjust cell width for titles
    index = 1
    for col in w_sheet.columns:
        w_sheet.column_dimensions[cell.get_column_letter(index)].width = max(10, len(col[0].value))
        index += 1

    work_book.save(path + "/" + title + " - Output Individual Line.xlsx")

def calculate_total_cath_days(p):
    """Returns the total number of days a Patient has ANY catheter."""
    #need to test on large data sample
    sorted(p.lines)
    total_days = p.lines[0].out_date - p.lines[0].in_date
    try:
        current = p.lines[0]
    except IndexError:
        return 0
    for l in p.lines:
        if current.out_date < l.in_date:
            total_days += l.out_date - l.in_date
            current = l
        else:
            total_days += l.out_date - current.out_date
            current = l
    return total_days.days

class Patient():
    """Patient Class contains lists of Visits, Lines and a Dictionary of Events."""
    def __init__(self, patient_id):
        self.visits = []
        self.lines = []
        #self.events = {} # possibly becoming depricated. Review for later. Comment meant to trigger pylint.

        #list of clabsis and possible clanc events
        self.clabsis = []
        self.clancs = []

        self.patient_id = patient_id
        self.total_visit_time = None
        self.total_line_time = None
        self.total_lumen_time = None

    def add_visit(self, v):
        """Adds a Visit object to the list of Visits and adds the time to total_visit_time."""
        assert isinstance(v, Visit), "new visits must be of type Visit"
        self.visits.append(v)
        if self.total_visit_time is None:
            self.total_visit_time = v.total_time
        else:
            self.total_visit_time += v.total_time

    def add_line(self, l):
        """Adds a Line object to the list of Visits and adds the time to total_line_time."""
        assert isinstance(l, Line), "new lines must be of type Line"
        self.lines.append(l)
        if self.total_line_time is None:
            self.total_line_time = l.total_time
            self.total_lumen_time = l.lumen_days
        else:
            self.total_line_time += l.total_time
            self.total_lumen_time += l.lumen_days

class Visit():
    """Visit Class stores datetime info for a single Patient Visit."""
    def __init__(self, patient, in_date, out_date):
        self.patient = patient
        self.check_in_date = in_date
        self.check_out_date = out_date
        self.total_time = in_date - out_date

@functools.total_ordering
class Line():
    """Line Class stores data for a single Line in a Patient. Records a Dictionary of Events."""
    def __init__(self, line_id, line_type, lumens, in_date, out_date, removal_reason):
        self.line_type = line_type
        self.in_date = in_date
        self.out_date = out_date
        self.line_id = line_id
        self.lumens = lumens
        self.total_time = out_date - in_date
        self.lumen_days = self.total_time * self.lumens
        self.removal_reason = removal_reason


        #list of clabsis and variable for possible clanc event
        self.clabsis = []
        self.clanc = None

    def __lt__(self, other):
        if self.in_date == other.in_date:
            return self.out_date < other.out_date
        return self.in_date < self.out_date

    def __eq__(self, other):
        return self.in_date == other.in_date and self.out_date == other.out_date

class CLABSI():
    """Class for CLABSI event. used becuase required infectious information is more complicated"""
    def __init__(self, patient, lines, date):
        self.patient = patient
        self.lines = lines
        self.date = date

class CLANC(): 
    """Class for CLANC event. used becuase required non-infect information is more complicated"""
    def __init__(self, patient, line, date):
        self.patient = patient
        self.line = line
        self.date = date

class BadFormatException(Exception):
    def __init__(self, value):
        self.parameter = value
    def __str__(self):
        return self.parameter
