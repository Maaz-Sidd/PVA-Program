import tkinter
from datetime import date
from tkinter import *
from tkinter import filedialog
import os
import openpyxl
from openpyxl import *
from openpyxl.styles import numbers

check = 0
def run_analysis():
    
    WO_number = WO_number_Entry.get()

    # Load the destination workbook
    PVA_WB = openpyxl.load_workbook('PVA_Template.xlsx')
    PVA_WS = PVA_WB.active

    PVA_WS['B4'].value = WO_number
    PVA_WS['H2'].value = date.today()

    global Estimates_WB
    global Estimates_file
    global Actuals_WB

    
    tasks = [130, 131, 200, 205, 225, 235, 245 , 255, 270, 280, 295, 315, 380, 400, 415, 430, 440, 450, 455, 465, 485, 500, 510, 515, 999]

    personnel = 0
    personnel_brdn = 0
    lines_hours = 0
    lines_hours_OT = 0
    des_hours = 0
    des_hours_OT = 0
    tools = 0


    if Estimates_WB is not None:
        Estimates_WS = Estimates_WB.active
        
        #for i in range (1, Estimates_WS.max_row + 1):
         #   Estimates_WS.cell(row= i, column= 3).number_format = "#,##"

        #Estimates_WB.save(Estimates_file)

        for i in range (17, 36):
            for k in range (2, 4):
                if k == 3:
                    continue
                PVA_WS.cell(row= i, column= k ).value = None

        for i in range (44, 55):
            for k in range (2, 3):
                PVA_WS.cell(row= i, column= k ).value = None

        for i in range (1, Estimates_WS.max_row + 1):

            sub = Estimates_WS.cell(row= i, column= 12).value
            des = Estimates_WS.cell(row= i, column= 13).value
            Wdes = Estimates_WS.cell(row= i, column= 17).value
            dollars = Estimates_WS.cell(row= i, column= 32).value
            brdn = Estimates_WS.cell(row= i, column= 34).value
            fringe = Estimates_WS.cell(row= i, column= 35).value
            hours = Estimates_WS.cell(row= i, column= 28).value
            hours_OT = Estimates_WS.cell(row= i, column= 29).value

            if sub == 1  and des == 'Direct Materials':
                mat_dir = dollars
                PVA_WS['B20'] = mat_dir
                print(mat_dir)
            elif sub == 1 and des == "Miscellaneous Cost":
                ex_cont = dollars
                PVA_WS['B28'] = ex_cont
                if brdn != None or fringe != None:
                    ex_cont_brdn = brdn
                    PVA_WS['B29'] = ex_cont_brdn
            elif sub == 1:
                mat_brdn = brdn
                mat = dollars
                PVA_WS['B18'] = mat
                if fringe != None:
                    mat_brdn += fringe
                PVA_WS['B19'] = mat_brdn
                print(mat)
                print(mat_brdn)
            if sub == 100 and Wdes == 'Travel Time':
                travel = dollars 
                if brdn != None:
                    travel += brdn
                if fringe != None:
                    travel += fringe
                PVA_WS['B24'] = travel
                print(travel)
            elif sub == 100:
                tools += dollars
                lines_hours += hours
                PVA_WS['B26'] = tools
                print(tools)
                if hours_OT != None:
                        lines_hours_OT += hours_OT
            if sub == 199:
                ICI_Credit = dollars
                PVA_WS['B31'] = ICI_Credit
                print(ICI_Credit)
            for k in range(len(tasks)):
                if sub == tasks[k]:
                    personnel += dollars
                    personnel_brdn += brdn + fringe
                    lines_hours += hours
                    PVA_WS['B22'] = personnel
                    PVA_WS['B23'] = personnel_brdn 
                    print(personnel)
                    print(personnel_brdn)
                    if hours_OT != None:
                        lines_hours_OT += hours_OT
            if sub == 170 or sub == 115:
                des_hours += hours
                des_brdn = brdn + fringe
                PVA_WS['B46'] = des_hours + des_brdn
                if hours_OT != None:
                        des_hours_OT += hours_OT
                        PVA_WS['B47'] = des_hours_OT
            if sub == 120:
                ins_hours = hours
                PVA_WS['B48'] = ins_hours
                if hours_OT != None:
                        ins_hours_OT = hours_OT
                        PVA_WS['B49'] = ins_hours_OT
            if sub == 800:
                met_hours = hours
                PVA_WS['B54'] = met_hours
                if hours_OT != None:
                    met_hours_OT = hours_OT
                    PVA_WS['B54'] = met_hours_OT

            PVA_WS['B50'] = lines_hours
            PVA_WS['B51'] = lines_hours_OT



    if Estimates_WB is not None:
        Estimates_WS = Estimates_WB.active

    download_Directory = filedialog.askdirectory()
    file_name = str(WO_number) + " - PVA.xlsx"
    file_path = os.path.join(download_Directory, file_name)
    PVA_WB.save(file_path)
    PVA_WB.close()
    Estimates_WB.close()
    #Actuals_WB.close()

    #openpyxl.writer.excel.save_workbook(PVA, WO_number)



def browseFiles_Estimates():
    global Estimates_WB

    global Estimates_file

    Estimates_file = filedialog.askopenfilename(initialdir = "/", title = "Select a File",filetypes = (("Excel Workbooks", "*.xlsx*"),("all files", "*.*")))

    # Change label contents
    label_file_explorer_estimates.configure(text="File Opened: "+ Estimates_file, fg = "blue")

    if Estimates_file:
        #check += 1
        Estimates_WB = openpyxl.load_workbook(Estimates_file, data_only=True)
    else:
        label_file_explorer_estimates.configure(text="Error loading file, please try again", fg= "red")

def browseFiles_Actuals():
    global Actuals_WB

    Actuals_file = filedialog.askopenfilename(initialdir = "/", title = "Select a File",filetypes = (("Excel Workbooks", "*.txt*"),("all files", "*.*")))

    # Change label contents
    label_file_explorer_actuals.configure(text="File Opened: "+ Actuals_file, fg = "blue")

    if Actuals_file:
        #check += 1
        Actuals_WB = openpyxl.load_workbook(Actuals_file)
    else:
        label_file_explorer_actuals.configure(text="Error loading file, please try again", fg= "red")

window = tkinter.Tk()
window.title("Project Variance Analysis")

frame = tkinter.Frame(window)
frame.pack() 

#reading estimate file
info_frame = tkinter.LabelFrame(frame, text="Work order Information")
info_frame.grid(row= 0, column= 0, padx = 20, pady= 10)

WO_number_label = tkinter.Label(info_frame, text= "Enter Work Order Number: ")
WO_number_label.grid(row= 0, column= 0)

WO_number_Entry = tkinter.Entry(info_frame)
WO_number_Entry.grid(row= 0, column= 1)

for widget in info_frame.winfo_children():
    widget.grid_configure(padx=10, pady=5)

file_frame = tkinter.LabelFrame(frame, text="File uploads")
file_frame.grid(row= 1, column= 0, sticky= "news", padx = 20, pady= 10)

Estimates_label = tkinter.Label(file_frame, text= "Upload Estimate file: ")
Estimates_label.grid(row= 0, column= 0)
Actuals_label = tkinter.Label(file_frame, text= "Upload Actuals file: ")
Actuals_label.grid(row= 1, column= 0)

Estimates_button = tkinter.Button(file_frame, text= "Upload File", command=browseFiles_Estimates)
Estimates_button.grid(row= 0, column= 1)
Actuals_button = tkinter.Button(file_frame, text= "Upload File", command=browseFiles_Actuals)
Actuals_button.grid(row= 1, column= 1)

label_file_explorer_estimates = Label(file_frame, text = "",fg = "blue")
label_file_explorer_estimates.grid(row=0, column=2)

label_file_explorer_actuals = Label(file_frame, text = "",fg = "blue")
label_file_explorer_actuals.grid(row=1, column=2)

for widget in file_frame.winfo_children():
    widget.grid_configure(padx=10, pady=5)

run_button = tkinter.Button(frame, text="Run analysis", command = run_analysis)
run_button.grid(row=3, column=0, sticky="news", padx=20, pady=10)

#if check == 2:
 #   run_button["state"] = tk.ENABLED

window.mainloop()
