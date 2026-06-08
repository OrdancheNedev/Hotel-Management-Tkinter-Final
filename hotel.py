import tkinter as tk
from tkinter import ttk
import openpyxl 
from tkcalendar import DateEntry

def load_data():
    path = "C:/Users/Lenovo/Desktop/Hotel-Management-Tkinter-Final/hotel_management_data.xlsx"
    try:
        workbook = openpyxl.load_workbook(path)
        sheet = workbook.active
        list_values = list(sheet.values)
        
    
        if not list_values:
            return

       
        for col_name in list_values[0]:
            treeview.heading(col_name, text=col_name)

     
        for item in treeview.get_children():
            treeview.delete(item)

        for value_tuple in list_values[1:]:
            treeview.insert('', tk.END, values=value_tuple)
    except FileNotFoundError:
        print(f"Error: The file at {path} was not found. Treeview left empty.")

def insert_row():
   
    new_data = (
        first_name_entry.get(), last_name_entry.get(), country.get(),
        age_spinbox.get(), adress.get(), id_type.get(), id_number.get(),
        room_number.get(), date_in.get(), date_out.get(), rate.get(),
        "Yes" if a.get() else "No"
    )
    
   
    treeview.insert('', tk.END, values=new_data)
    first_name_entry.delete(0, "end")
    first_name_entry.insert(0, "Name")
    last_name_entry.delete(0, "end")
    last_name_entry.insert(0, "Surname")
    country.delete(0, "end")
    country.insert(0, "Country")
    age_spinbox.delete(0, "end")
    age_spinbox.insert(0, "Age")
    adress.delete(0, "end")
    adress.insert(0, "Address")
    id_type.set("")
    id_type.current(0)
    id_number.delete(0, "end")
    id_number.insert(0, "ID Number")    
    room_number.delete(0, "end")
    room_number.insert(0, "Room Number")
    date_in.delete(0, "end")
    date_out.delete(0, "end")
    rate.delete(0, "end")
    rate.insert(0, "Rate")
    checkbutton.state(["!selected"]) 
    
    path = "C:/Users/Lenovo/Desktop/Hotel-Management-Tkinter-Final/hotel_management_data.xlsx"
    try:
        workbook = openpyxl.load_workbook(path)
        sheet = workbook.active
        sheet.append(new_data)
        workbook.save(path)
        print("Data successfully inserted and saved!")
    except Exception as e:
        print(f"Error saving to Excel: {e}")

def toggle_mode():
    if mode_switch.instate(["selected"]):
        style.theme_use("forest-light")
    else:
        style.theme_use("forest-dark")


root = tk.Tk()
root.title("Hotel Management System")
root.geometry("1600x900")


root.columnconfigure(0, weight=1)
root.rowconfigure(0, weight=1)

style = ttk.Style(root)
try:
    root.tk.call("source", "forest-light.tcl")
    root.tk.call("source", "forest-dark.tcl")
    style.theme_use("forest-dark")
except Exception:
    print("Forest themes not found. Using default fallback.")
    style.theme_use("clam")


frame = ttk.Frame(root, padding=20)
frame.grid(row=0, column=0, sticky="nsew")


frame.columnconfigure(0, weight=4) 
frame.columnconfigure(1, weight=1)
frame.rowconfigure(0, weight=1)

id_list = ["Passport", "ID Card", "Driver's License", "Other"]


widgets_frame = ttk.LabelFrame(frame, text="Check In/Out", padding=10)
widgets_frame.grid(row=0, column=1, padx=(20, 0), pady=10, sticky="nsew")
widgets_frame.columnconfigure(1, weight=1)


first_name_entry = ttk.Entry(widgets_frame)
first_name_entry.insert(0, "Name")
first_name_entry.bind("<FocusIn>", lambda e: first_name_entry.delete('0', 'end') if first_name_entry.get() == "Name" else None)
first_name_entry.grid(row=0, column=1, padx=5, pady=5, sticky="ew")

last_name_entry = ttk.Entry(widgets_frame)
last_name_entry.insert(0, "Surname")
last_name_entry.bind("<FocusIn>", lambda e: last_name_entry.delete('0', 'end') if last_name_entry.get() == "Surname" else None)
last_name_entry.grid(row=1, column=1, padx=5, pady=5, sticky="ew")

country = ttk.Entry(widgets_frame)
country.insert(0, "Country")
country.bind("<FocusIn>", lambda e: country.delete('0', 'end') if country.get() == "Country" else None)
country.grid(row=2, column=1, padx=5, pady=5, sticky="ew")

age_spinbox = ttk.Spinbox(widgets_frame, from_=0, to=100)
age_spinbox.set("Age")
age_spinbox.grid(row=3, column=1, padx=5, pady=5, sticky="ew")

adress = ttk.Entry(widgets_frame)
adress.insert(0, "Address")
adress.bind("<FocusIn>", lambda e: adress.delete('0', 'end') if adress.get() == "Address" else None)
adress.grid(row=4, column=1, padx=5, pady=5, sticky="ew")

id_type = ttk.Combobox(widgets_frame, values=id_list, state="readonly")
id_type.current(0)
id_type.grid(row=5, column=1, padx=5, pady=5, sticky="ew")

id_number = ttk.Entry(widgets_frame)
id_number.insert(0, "ID Number")
id_number.bind("<FocusIn>", lambda e: id_number.delete('0', 'end') if id_number.get() == "ID Number" else None)
id_number.grid(row=6, column=1, padx=5, pady=5, sticky="ew")

room_number = ttk.Spinbox(widgets_frame, from_=1, to=1000)
room_number.set("Room Number")
room_number.grid(row=7, column=1, padx=5, pady=5, sticky="ew")

date_in = DateEntry(widgets_frame, date_pattern='yyyy-mm-dd')
date_in.grid(row=8, column=1, padx=5, pady=5, sticky="ew")

date_out = DateEntry(widgets_frame, date_pattern='yyyy-mm-dd')
date_out.grid(row=9, column=1, padx=5, pady=5, sticky="ew")

rate = ttk.Spinbox(widgets_frame, from_=0, to=5000)
rate.set("Rate")
rate.grid(row=10, column=1, padx=5, pady=5, sticky="ew")

a = tk.BooleanVar()
checkbutton = ttk.Checkbutton(widgets_frame, text="Paid", variable=a)
checkbutton.grid(row=11, column=1, padx=5, pady=5, sticky="ew")

button = ttk.Button(widgets_frame, text="Insert", command=insert_row)
button.grid(row=12, column=1, padx=5, pady=10, sticky="ew")

separator = ttk.Separator(widgets_frame)
separator.grid(row=13, column=1, padx=10, pady=10, sticky="ew")

mode_switch = ttk.Checkbutton(widgets_frame, text="Mode", style="Switch", command=toggle_mode)
mode_switch.grid(row=14, column=1, padx=5, pady=10, sticky="ew")



treeFrame = ttk.Frame(frame)
treeFrame.grid(row=0, column=0, pady=10, sticky="nsew")
treeFrame.columnconfigure(0, weight=1)
treeFrame.rowconfigure(0, weight=1)

treeScroll = ttk.Scrollbar(treeFrame)
treeScroll.grid(row=0, column=1, sticky="ns")

cols = ("Name","Surname","Country", "Age", "Address", "ID Type", "ID Number", "Room Number", "Date In", "Date Out", "Rate", "Paid")
treeview = ttk.Treeview(treeFrame, show="headings", yscrollcommand=treeScroll.set, columns=cols)
treeview.grid(row=0, column=0, sticky="nsew")
treeScroll.config(command=treeview.yview)


for col in cols:
    treeview.column(col, width=100, anchor="center")

load_data()

root.mainloop()