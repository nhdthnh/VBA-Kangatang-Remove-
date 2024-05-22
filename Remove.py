import win32com.client
import os
import tkinter as tk
from tkinter import filedialog
import psutil
from tkinter import scrolledtext
import threading
for proc in psutil.process_iter():
        if proc.name() == "EXCEL.EXE":
            proc.kill()

def update_text(message):
    text_box.insert(tk.END, message + '\n')
    text_box.see(tk.END)
def remove_module(filepath, module_name):
    excel = win32com.client.Dispatch("Excel.Application")
    try:
        excel.Visible = False  # Set to True if you want to see the Excel application window
        excel.ScreenUpdating = False  # Disable screen updating to improve performance

        workbook = excel.Workbooks.Open(os.path.abspath(filepath), False, True)  # Open the workbook in read-only mode
        macro_modules = workbook.VBProject.VBComponents

        module_found = False
        if macro_modules.Count > 0:
            for module in macro_modules:
                if module.Type == 1 and module.Name == module_name:  # Check if it's a module component with the specified name
                    macro_modules.Remove(module)  # Remove the module
                    module_found = True
                    break

        if module_found:
            selected_option = var.get()
            if (selected_option == "Another Folder"):
                namefile = filepath.split("\\")[-1]
                new_folder_path = Saventry.get()
                new_file_path = new_folder_path +"\\"+namefile.split(".")[0]+ "_modified.xls"  # Generate new file path
                if (os.path.exists(new_file_path)):
                    os.remove(new_file_path)
                workbook.SaveAs(new_file_path)  
                print(f"Module '{module_name}' removed and saved as '{new_file_path}'")
                update_text(f"Module '{module_name}' removed and saved as '{new_file_path}")
            else:
                new_file_path = os.path.splitext(filepath)[0] + "_modified.xls"  # Generate new file path
                if os.path.exists(new_file_path):
                    os.remove(new_file_path)
                workbook.SaveAs(new_file_path)
                update_text(f"Module '{module_name}' removed and saved as '{new_file_path}'")
        else:
            # print(f"Module '{module_name}' not found in the workbook.")
            update_text(f"Module '{module_name}' not found in the {filepath}.")
        workbook.Close(False)
        
    except Exception as e:
        print(f"Error occurred: {str(e)}")
    finally:
        excel.Quit()
    update_text("DONE")

def find_in_folder(folder_path, module_name):
    for filename in os.listdir(folder_path):
        if filename.endswith(".xls"):
            file_path = os.path.join(folder_path, filename)
            threading.Thread(target=remove_module, args=(file_path, module_name)).start()

def open_folder():
    directory = filedialog.askdirectory()
    entry1.delete(0, tk.END)
    entry1.insert(0, directory.replace('/', '\\'))

def open_save_folder():
    directory = filedialog.askdirectory()
    Saventry.delete(0, tk.END)
    Saventry.insert(0, directory.replace('/', '\\'))

def remove():
    folderPath= entry1.get()
    name = entry2.get()
    find_in_folder(folderPath,name)


    
window = tk.Tk()
window.title("Remove kangatang")
window.resizable(width=False,height=False)
folderLabel = tk.Label(text="Folder path",font = '15')
folderLabel.grid(row =0,column=0)

# Create the entry boxes
entry1 = tk.Entry(window,font= "15",width=50)
entry1.grid(row=0, column=1)

SaveLabel = tk.Label(text="Save folder",font = '15')
SaveLabel.grid(row =1,column=0)
Saventry = tk.Entry(window,font= "15",width=50)
Saventry.grid(row=1, column=1)

folderLabel = tk.Label(text="Macro name",font = '15')
folderLabel.grid(row =2,column=0)
entry2 = tk.Entry(window,font= "15")
entry2.insert(0, "Kangatang")  # Set the default value
entry2.grid(row=2, column=1)

Browse = tk.Button(window, text="Browse",font= "15",command=open_folder)
Browse.grid(row=0, column=2)
BrowseSave = tk.Button(window, text="Browse",font= "15",command=open_save_folder)
BrowseSave.grid(row=1, column=2)
var = tk.StringVar(value="Directly")
options = ["Directly", "Another Folder"]
for index, option in enumerate(options):
    radio_button = tk.Radiobutton(window, text=option, variable=var, value=option,font='13')
    radio_button.grid(row=index+6, column=0, sticky=tk.W)


Run = tk.Button(window, text="Remove",font= "15",command=remove)
Run.grid(row=9, column=0)
text_box = tk.Text(window)
text_box.grid(row=12, column=0, columnspan=2)
# update_text("abcdef")
# Start the Tkinter event loop
window.mainloop()
# filepath = r"C:\Users\thanh\OneDrive\Desktop\Covirus\Covirus"
# module_name = "Kangatang"
# find_in_folder(filepath, module_name)