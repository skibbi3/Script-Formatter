import tkinter
import backend
import docx
import tkinter.filedialog

window = tkinter.Tk()
window.title("Script Formatter")

def select_file():
    # Takes the input file and stores it
    global document_file
    document_file = docx.Document(tkinter.filedialog.askopenfilename(defaultextension=".docx", filetypes=(("Word Document", "*.docx"),("All Files", "*.*") )))

def run():
    # Accepts the location to save the document as user input
    save_file_name = tkinter.filedialog.asksaveasfilename(defaultextension=".docx", filetypes=(("Word Document", "*.docx"),("All Files", "*.*") ))

    # Passes the input file and topmatter to the backend to format and save the document
    success = backend.process_word(filename=document_file, 
        save_name = save_file_name,
        title = title_entry.get(),
        prog_code = code_entry.get(),
        ver = version_entry.get(),
        written_by = writer_entry.get(),
        edited_by = editor_entry.get()
    )

    # Exception handling to provide feedback of success or failure
    if success:
        tkinter.messagebox.showinfo("Success", "File saved successfully!")
    else:
        tkinter.messagebox.showerror("Error", "There appears to have been an error!")

# Declare input labels and entries
greeting_label = tkinter.Label(text = "Welcome to Skibinski's Script formatter.")
select_script_label = tkinter.Label(text = "Please select your script")
select_script_button = tkinter.Button(text = "Open", command=select_file)
file_name_label = tkinter.Label(text = "")

title_label = tkinter.Label(text = "Title of Program")
title_entry = tkinter.Entry()

code_label = tkinter.Label(text = "Program Code")
code_entry = tkinter.Entry()

version_label = tkinter.Label(text = "Version")
version_entry = tkinter.Entry()

writer_label = tkinter.Label(text = "Script Written By")
writer_entry = tkinter.Entry()

editor_label = tkinter.Label(text = "Script Edited By")
editor_entry = tkinter.Entry()

go_button = tkinter.Button(text = "Go", command=run)

# Packing the labels, input fields, and the button into the grid
greeting_label.grid(row=0, columnspan = 2)

select_script_label.grid(row = 1, column = 0, sticky = "W", pady = 2)
select_script_button.grid(row = 1, column = 1 , sticky = "W", pady = 2)

title_label.grid(row = 2, column = 0, sticky = "W", pady = 2)
title_entry.grid(row = 2, column = 1, sticky = "W", pady = 2)

code_label.grid(row = 3, column = 0, sticky = "W", pady = 2)
code_entry.grid(row = 3, column = 1, sticky = "W", pady = 2)

version_label.grid(row = 4, column = 0, sticky = "W", pady = 2)
version_entry.grid(row = 4, column = 1, sticky = "W", pady = 2)

writer_label.grid(row = 5, column = 0, sticky = "W", pady = 2)
writer_entry.grid(row = 5, column = 1, sticky = "W", pady = 2)

editor_label.grid(row = 6, column = 0, sticky = "W", pady = 2)
editor_entry.grid(row = 6, column = 1, sticky = "W", pady = 2)

go_button.grid(row = 7, column = 1, sticky = "W", pady = 2)

window.mainloop()