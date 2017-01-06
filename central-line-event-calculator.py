from tkinter import *
from tkinter import ttk, filedialog
from datetime import datetime

from utils import *
import os, sys

###Options###
start_range = datetime.min
end_range = datetime.max

def analyze(*args):
    admit = admit_entry.get()   
    line = line_entry.get()
    clabsi = clabsi_entry.get()
    clanc = clanc_entry.get()
    output = output_entry.get()

    title = project_title.get()
    if title == "":
        error_message("Empty Project Title", "A Project Title is Required.")
        return

    if not verify_paths(admit, line, clabsi, clanc, output):
        #ends the analysis if the paths are not valid
        return

    # verify_admit_data(admit_entry.get())
    try:
        process_data(title, admit, line, clabsi, clanc, output, start_range, end_range)
    except BadFormatException as e:
        error_message("Invalid SpreadSheet Format", str(e))
        return
    except Exception as e:
        #comment out for terminal tracebacks.
        #always uncomment for production releases.
        error_message("Execution Error", str(e))
        return

    os.startfile(output)

def admit_path(*args):
    admit_data_loc = get_file_path("Patient Admission Data")
    admit_entry.delete(0, 'end')
    admit_entry.insert(0, admit_data_loc)

def line_path(*args):
    line_data_loc = get_file_path("Line Data")
    line_entry.delete(0, 'end')
    line_entry.insert(0, line_data_loc)

def clabsi_path(*args):
    clabsi_data_loc = get_file_path("CLABSI Data")
    clabsi_entry.delete(0, 'end')
    clabsi_entry.insert(0, clabsi_data_loc)

def clanc_path(*args):
    clanc_data_loc = get_file_path("CLANC Data")
    clanc_entry.delete(0, 'end')
    clanc_entry.insert(0, clanc_data_loc)

def output_path(*args):
    output_loc = get_file_directory("Output Directory")
    output_entry.delete(0, 'end')
    output_entry.insert(0, output_loc)



root = Tk()
root.title("Central Line Event Calculator")

options_open = False


def display_options():
    global options_open
    if options_open:
        return
    options_open = True

    option = Tk()
    option.title("Options")

    def save_options(*args):
        tmp_start = datetime.min
        tmp_end = datetime.max 
        
        start_mo = ("00" + start_mo_entry.get())[-2:]
        start_day = ("00" + start_day_entry.get())[-2:]
        start_yr = ("0000" + start_yr_entry.get())[-4:]
        end_mo = ("00" + end_mo_entry.get())[-2:]
        end_day = ("00" + end_day_entry.get())[-2:]
        end_yr = ("0000" + end_yr_entry.get())[-4:]
        try:
            start = start_mo + start_day + start_yr
            tmp_start = datetime.strptime(start, '%m%d%Y')
        except ValueError as e:
            print(e)
        try:
            global end_range
            end = end_mo + end_day + end_yr
            tmp_end = datetime.strptime(end, '%m%d%Y')
        except ValueError as e:
            print(e)
        
        if tmp_start >= tmp_end:
            error_message("Invalid Date Range", "Start date must be before end date.")
        else:
            global start_range, end_range
            start_range, end_range = tmp_start, tmp_end
        close_options()
        return

    def close_options():
        global options_open
        option.destroy()
        options_open = False

    opt_frame = ttk.Frame(option, padding="3 3 12 12")
    opt_frame.grid(column=0, row=0, sticky=(N, W, E, S))
    opt_frame.columnconfigure(0, weight=1)
    opt_frame.rowconfigure(0, weight=1)

    start_date_mo = StringVar()
    start_date_day = StringVar()
    start_date_yr = StringVar()
    end_date_mo = StringVar()
    end_date_day = StringVar()
    end_date_yr = StringVar()

    ttk.Label(opt_frame, text="Start Date:").grid(column=1, row=2, sticky=(W, E))
    ttk.Label(opt_frame, text="End Date:").grid(column=1, row=3, sticky=(W, E))

    ttk.Label(opt_frame, text="MM").grid(column=2, row=1, sticky=(E))
    ttk.Label(opt_frame, text="DD").grid(column=3, row=1, sticky=(E))
    ttk.Label(opt_frame, text="YYYY").grid(column=4, row=1, sticky=(E))

    start_mo_entry = ttk.Entry(opt_frame, width=5, textvariable=start_date_mo)
    start_day_entry = ttk.Entry(opt_frame, width=5, textvariable=start_date_day)
    start_yr_entry = ttk.Entry(opt_frame, width=10, textvariable=start_date_yr)
    end_mo_entry = ttk.Entry(opt_frame, width=5, textvariable=end_date_mo)
    end_day_entry = ttk.Entry(opt_frame, width=5, textvariable=end_date_day)
    end_yr_entry = ttk.Entry(opt_frame, width=10, textvariable=end_date_yr)

    start_mo_entry.insert(0, start_range.month)
    start_day_entry.insert(0, start_range.day)
    start_yr_entry.insert(0, start_range.year)
    end_mo_entry.insert(0, end_range.month)
    end_day_entry.insert(0, end_range.day)
    end_yr_entry.insert(0, end_range.year)

    start_mo_entry.grid(column=2, row=2, sticky=(W, E))
    start_day_entry.grid(column=3, row=2, sticky=(W, E))
    start_yr_entry.grid(column=4, row=2, sticky=(W, E))
    end_mo_entry.grid(column=2, row=3, sticky=(W, E))
    end_day_entry.grid(column=3, row=3, sticky=(W, E))
    end_yr_entry.grid(column=4, row=3, sticky=(W, E))

    ttk.Button(opt_frame, text="Save", command=save_options).grid(column=4, row=4, sticky=E)
    ttk.Button(opt_frame, text="Cancel", command=close_options).grid(column=3, row=4, sticky=E)


    for child in opt_frame.winfo_children():
        child.grid_configure(padx=5, pady=5)

    option.protocol("WM_DELETE_WINDOW", close_options)
    option.bind('<Return>', save_options)
    root.bind('<Return>', save_options)

mainframe = ttk.Frame(root, padding="3 3 12 12")
mainframe.grid(column=0, row=0, sticky=(N, W, E, S))
mainframe.columnconfigure(0, weight=1)
mainframe.rowconfigure(0, weight=1)

project_title = StringVar()
admit_data_loc = StringVar()
line_data_loc = StringVar()
clabsi_data_loc = StringVar()
clanc_data_loc = StringVar()
output_loc = StringVar()

admit_entry = ttk.Entry(mainframe, width=70, textvariable=project_title)
admit_entry.grid(column=2, row=1, sticky=(W, E))

ttk.Label(mainframe, text="Select the files to be analyzed below:").grid(column=2, row=0, sticky=W)

admit_entry = ttk.Entry(mainframe, width=70, textvariable=admit_data_loc)
admit_entry.grid(column=2, row=2, sticky=(W, E))
admit_browse_button = Button(mainframe, text='Browse', command=admit_path)
admit_browse_button.grid(column=3, row=2, sticky=(W, E))

line_entry = ttk.Entry(mainframe, width=7, textvariable=line_data_loc)
line_entry.grid(column=2, row=3, sticky=(W, E))
line_browse_button = Button(mainframe, text='Browse', command=line_path)
line_browse_button.grid(column=3, row=3, sticky=(W, E))

clabsi_entry = ttk.Entry(mainframe, width=7, textvariable=clabsi_data_loc)
clabsi_entry.grid(column=2, row=4, sticky=(W, E))
clabsi_browse_button = Button(mainframe, text='Browse', command=clabsi_path)
clabsi_browse_button.grid(column=3, row=4, sticky=(W, E))

clanc_entry = ttk.Entry(mainframe, width=7, textvariable=clanc_data_loc)
clanc_entry.grid(column=2, row=5, sticky=(W, E))
clanc_browse_button = Button(mainframe, text='Browse', command=clanc_path)
clanc_browse_button.grid(column=3, row=5, sticky=(W, E))

output_entry = ttk.Entry(mainframe, width=7, textvariable=output_loc)
output_entry.grid(column=2, row=6, sticky=(W, E))
output_browse_button = Button(mainframe, text='Browse', command=output_path)
output_browse_button.grid(column=3, row=6, sticky=(W, E))


# Local testing paths for sample data.  DO NOT include in production
# admit_entry.insert(0, "D:/projects/med/sampledata/in/Admit and Discharge Input Data.xlsx")
# line_entry.insert(0, "D:/projects/med/sampledata/in/Sample Input Line Data.xlsx")
# clabsi_entry.insert(0, "D:/projects/med/sampledata/in/CLABSI Date Input.xlsx")
# clanc_entry.insert(0, "D:/projects/med/sampledata/in/CLANC Date Input.xlsx")
# output_entry.insert(0, "D:/projects/med/sampledata/out")


ttk.Label(mainframe, text="Project Title").grid(column=1, row=1, sticky=W)
ttk.Label(mainframe, text="Patient Admission Data").grid(column=1, row=2, sticky=W)
ttk.Label(mainframe, text="Line Data").grid(column=1, row=3, sticky=W)
ttk.Label(mainframe, text="CLABSI Data").grid(column=1, row=4, sticky=W)
ttk.Label(mainframe, text="CLANC Data").grid(column=1, row=5, sticky=W)


ttk.Label(mainframe, text="Output Destination").grid(column=1, row=6, sticky=W)

ttk.Button(mainframe, text="Continue", command=analyze).grid(column=3, row=7, sticky=E)
root.bind('<Return>', analyze)

ttk.Button(mainframe, text="Options", command=display_options).grid(column=2, row=7, sticky=E)


for child in mainframe.winfo_children():
    child.grid_configure(padx=5, pady=5)

def on_close():
    sys.exit()

root.protocol("WM_DELETE_WINDOW", on_close)

root.mainloop()

