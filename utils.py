from tkinter import filedialog

def get_file_path():
    path = filedialog.askopenfilename()
    return path

def get_file_directory():
    directory = filedialog.askdirectory()
    return directory