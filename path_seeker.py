"""
Path Seeker - Search for occurence of text through files

Author: Stivan Kitchoukov
"""
#C:\Users\skitchoukov\Desktop\Python\venv\Scripts\python.exe "$(FULL_CURRENT_PATH)"
"""
IMPORTING LIBRARIES
"""
import os
# import sys
import pythoncom
# import glob
# from win32com.shell import shell, shellcon
from tkinter import filedialog
from tkinter import messagebox
from tkinter import *

# Creating window
master = Tk()
master.title("Path Seeker")
master.configure(background="grey")

# Window Height/Width & X,Y Coordinates
w = 750
h = 700
x = 50
y = 100

# Apply Dimensions
master.geometry("%dx%d+%d+%d" % (w, h, x, y))

# Prevent Resize
master.resizable(False, False)

'''
PATH TYPE: DROP DOWN
'''
# Label for path type
path_type_label = Label(
    master,
    bg="grey",
    anchor="w",
    text="Path Type:"
)
path_type_label.place(x=10, y=10)

# Setting default for dropdown menu
path_type_value = StringVar(master)
path_type_value.set("Directory")

# Creating dropdown with values
path_type = OptionMenu(
    master,
    path_type_value,
    "Directory",
    "File"
)
path_type.config(
    bg="grey",
    width="10"
)
path_type.place(x=100, y=10)

'''
SEARCH TYPE: DROP DOWN
'''
# Label for search type
search_type_label = Label(
    master,
    bg="grey",
    anchor="w",
    text="Search Type:"
)
search_type_label.place(x=240, y=10)

# Default for Search type drop down
search_type_value = StringVar(master)
search_type_value.set("First occurrence")

# Drop down with values
search_type = OptionMenu(
    master,
    search_type_value,
    "All occurrences",
    "First occurrence"
)
search_type.config(
    bg="grey",
    width="15"
)
search_type.place(x=340, y=10)

'''
DIRECTORY/FILE: BROWSER
'''
# Label for drop down
path_label = Label(
    master,
    bg="grey",
    anchor="w",
    text="Path:"
)
path_label.place(x=10, y=40)

# Path Entry field
path_text = StringVar(master)
path_entry = Entry(master, textvariable=path_text)
path_entry.config(
    state="disabled",
    highlightbackground="grey",
    width="50"
)
path_entry.place(x=100, y=40)


# Function to populate path based on dropdown
def browser_type():
    browser = path_type_value.get()

    if browser == "Directory":
        master.directory = filedialog.askdirectory()
        path_text.set(str(master.directory))
    elif browser == "File":
        master.filename = filedialog.askopenfilename()
        path_text.set(str(master.filename))
    return


# Browse Button
path_browse = Button(
    master,
    text="Browse",
    highlightbackground="grey",
    command=browser_type
)
path_browse.place(x=580, y=40)


def path_type_changed(*args):
    path_text.set("")


path_type_value.trace("w", path_type_changed)

'''
Search String: Search field and Button
'''
# Search label
search_label = Label(
    master,
    bg="grey",
    anchor="w",
    text="Find:"
)
search_label.place(x=10, y=70)

# Search text field
search_entry = Entry(master)
search_entry.config(
    highlightbackground="grey",
    width="50"
)
search_entry.place(x=100, y=70)

'''
Output and Errors
'''
# Result Label
result_label = Label(
    master,
    bg="grey",
    anchor="w",
    text="Result:"
)
result_label.place(x=10, y=100)

# Result Text field
result = Text(master)
result.config(
    highlightbackground="grey",
    width="90",
    height="25"
)
result.place(x=10, y=130)

# Error label
error_label = Label(
    master,
    bg="grey",
    anchor="w",
    text="Errors/Warnings:"
)
error_label.place(x=10, y=550)

# Error text field
error_text = Text(master)
error_text.config(
    highlightbackground="grey",
    width="90",
    height="5"
)
error_text.place(x=10, y=580)

'''
Search Button and Search functionality 
'''
# Function to search for occurrence in files
def search():
    result.delete("1.0", END)
    error_text.delete("1.0", END)

    if len(str(path_entry.get())) < 1 or len(str(search_entry.get())) < 1:
        messagebox.showinfo("Error", "Path and Search cannot be empty")
        return

    search_string = search_entry.get().lower()
    file_path = path_entry.get()
    search_string_type = search_type_value.get()

    if path_type_value.get() == "File":
        try:
            with open(file_path) as f:
                found = False
                for line in f:
                    if search_string in line.lower():
                        found = True
                        result.insert(END, line)
                        if search_string_type == "First occurrence":
                            break
                if not found:
                    result.insert(END, "Could not find " + search_string + " in file(s).")
        except:
            error_text.insert(END, "Error opening file: " + file_path)

    elif path_type_value.get() == "Directory":
        file_paths = [os.path.join(file_path, fn) for fn in next(os.walk(file_path))[2]]
        for i in file_paths:
            try:
                # START
                if i.endswith("lnk"):
                    from win32com.shell import shell, shellcon
                    def shortcut_target(filename):
                        link = pythoncom.CoCreateInstance(
                            shell.CLSID_ShellLink,
                            None,
                            pythoncom.CLSCTX_INPROC_SERVER,
                            shell.IID_IShellLink
                        )
                        link.QueryInterface(pythoncom.IID_IPersistFile).Load(filename)
                        #
                        # GetPath returns the name and a WIN32_FIND_DATA structure
                        # which we're ignoring. The parameter indicates whether
                        # shortname, UNC or the "raw path" are to be
                        # returned. Bizarrely, the docs indicate that the
                        # flags can be combined.
                        #
                        name, _ = link.GetPath(shell.SLGP_UNCPRIORITY)
                        return name

                    i = shortcut_target(i)
                    print(i)

                # END
                with open(i) as f:
                    found = False
                    result.insert(END, ("#########" + i + "##########" + "\n"))
                    for line in f:
                        if search_string in line.lower():
                            found = True
                            result.insert(END, line)
                            if search_string_type == "First occurrence":
                                break
                    if search_string_type == "First occurrence" and found == True:
                        return

                    if not found:
                        result.insert(END, ("Could not find " + search_string + "\n"))
            except:
                error_text.insert(END, ("Error opening file: " + i + "\n"))

    return


# Search Button
search_button = Button(
    master,
    text="Search",
    highlightbackground="grey",
    command=search
)
search_button.place(x=580, y=70)

# Running tkinter window
mainloop()
