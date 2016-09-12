from tkinter import *
from tkinter import ttk, filedialog

from utils import *
import os

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
        process_data(title, admit, line, clabsi, clanc, output)
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

for child in mainframe.winfo_children():
    child.grid_configure(padx=5, pady=5)

root.mainloop()