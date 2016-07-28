from tkinter import *
from tkinter import ttk
from tkinter import filedialog

from utils import *
import os

def analyze(*args):
    if not verify_paths(admit_entry.get(), line_entry.get(), event_entry.get()):
        #Alert window to catch non Excel files
        alert = Toplevel()
        alert.title("Invalid Data Supplied")
        return
    if output_entry.get() == '':
        #validate output destination is supplied
        alert = Toplevel()
        alert.title("Invalid Output Destination")
        return
    verify_admit_data(admit_entry.get())
    process_data(admit_entry.get(), line_entry.get(), event_entry.get(), output_entry.get())
    os.startfile(output_entry.get())

def admit_path(*args):
    admit_data_loc = get_file_path("Patient Admission Data")
    admit_entry.insert(0, admit_data_loc)

def line_path(*args):
    line_data_loc = get_file_path("Line Data")
    line_entry.insert(0, line_data_loc)

def event_path(*args):
    event_data_loc = get_file_path("Event Data")
    event_entry.insert(0, event_data_loc)

def output_path(*args):
    output_loc = get_file_directory("Output Directory")
    output_entry.insert(0, output_loc)

root = Tk()
root.title("Central Line Event Calculator")

mainframe = ttk.Frame(root, padding="3 3 12 12")
mainframe.grid(column=0, row=0, sticky=(N, W, E, S))
mainframe.columnconfigure(0, weight=1)
mainframe.rowconfigure(0, weight=1)

admit_data_loc = StringVar()
line_data_loc = StringVar()
event_data_loc = StringVar()
output_loc = StringVar()

ttk.Label(mainframe, text="Select the files to be analyzed below:").grid(column=2, row=1, sticky=W)

admit_entry = ttk.Entry(mainframe, width=70, textvariable=admit_data_loc)
admit_entry.grid(column=2, row=2, sticky=(W, E))
admit_browse_button = Button(mainframe, text='Browse', command=admit_path)
admit_browse_button.grid(column=3, row=2, sticky=(W, E))

line_entry = ttk.Entry(mainframe, width=7, textvariable=line_data_loc)
line_entry.grid(column=2, row=3, sticky=(W, E))
line_browse_button = Button(mainframe, text='Browse', command=line_path)
line_browse_button.grid(column=3, row=3, sticky=(W, E))

event_entry = ttk.Entry(mainframe, width=7, textvariable=event_data_loc)
event_entry.grid(column=2, row=4, sticky=(W, E))
event_browse_button = Button(mainframe, text='Browse', command=event_path)
event_browse_button.grid(column=3, row=4, sticky=(W, E))

output_entry = ttk.Entry(mainframe, width=7, textvariable=output_loc)
output_entry.grid(column=2, row=5, sticky=(W, E))
output_browse_button = Button(mainframe, text='Browse', command=output_path)
output_browse_button.grid(column=3, row=5, sticky=(W, E))


#TODO: Remove default paths.
admit_entry.insert(0, "D:/projects/med/sampledata/in/Admit and Discharge Input Data.xlsx")
line_entry.insert(0, "D:/projects/med/sampledata/in/Sample Input Line Data.xlsx")
event_entry.insert(0, "D:/projects/med/sampledata/in/combineddata.xlsx")
output_entry.insert(0, "D:/projects/med/sampledata/out")


ttk.Label(mainframe, text="Patient Admission Data").grid(column=1, row=2, sticky=W)
ttk.Label(mainframe, text="Line Data").grid(column=1, row=3, sticky=W)
ttk.Label(mainframe, text="Event Data").grid(column=1, row=4, sticky=W)

ttk.Label(mainframe, text="Output Destination").grid(column=1, row=5, sticky=W)

ttk.Button(mainframe, text="Continue", command=analyze).grid(column=3, row=6, sticky=E)
root.bind('<Return>', analyze)

for child in mainframe.winfo_children():
    child.grid_configure(padx=5, pady=5)

root.mainloop()