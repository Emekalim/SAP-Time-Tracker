import tkinter as tk
from tkinter import ttk
from tkinter import font
from tkinter import messagebox
import time
import pandas as pd
from functions import *
from pathlib import Path
import module
from datetime import datetime
import subprocess
import win32com.client
from win32com.client import Dispatch, constants

class CatsTimeTracker:
    def __init__(self, root):
        self.root = root
        self.setup_menu()
        self.total_rows = 6
        self.start_times = {}  # Record the start time for each stopwatch
        self.elapsed_times = {i: 0 for i in range(1, self.total_rows)}  # Initialize elapsed time for each row
        self.total_day_time = 0
        self.documents_path = Path.home() / "Documents"/ "SAP Time Tracker"
        self.initial_label_current = "Activity XX\n Time Spent: "
        self.initial_label_total = "Total Time Elapsed Today: \n "
        self.setup_window()
        self.create_widgets()
        self.load_icons()
        self.create_activity_rows()    
        self.active_stopwatch = None

    def setup_menu(self):
        menubar = tk.Menu(self.root)

        #Create File Menu 
        filemenu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label ='File', menu = filemenu) 
        filemenu.add_command(label="Modify Chargelines", command=self.open_modify_chargelines)
        filemenu.add_command(label="Export Timeline")
        
        #Create help Menu 
        helpmenu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label ='Help', menu = helpmenu) 
        helpmenu.add_command(label="User Guide", command=self.open_help)
        helpmenu.add_command(label="Report a Bug", command=self.send_bug_mail)
        # Insert the menubar in the main window.
        self.root.config(menu=menubar)
        
    
    def setup_window(self):
        self.root.title("CATS Time Tracker")
        self.root.geometry("")  # set starting size of window
        default_font = font.nametofont("TkDefaultFont")
        default_font.configure(size=9)  # This is 20 pixels
        self.root.option_add("*Font", default_font)
        self.root.resizable(False, False)

    def create_widgets(self):
        self.label_frame = tk.Frame(self.root)
        self.label_frame.grid(row=0, column=0, rowspan=20, columnspan=2, sticky='news')
        self.label_frame.grid_rowconfigure(0, weight=1)
        self.label_frame.grid_rowconfigure(3, weight=1)

        self.label_current_time = tk.Label(self.label_frame, width=45, 
                                           text=self.initial_label_current, bg="navajowhite2", relief="raised")
        
        self.label_current_time.grid(row=0, column=0, rowspan=2, columnspan=2,
                                      ipadx=5, ipady=25, sticky='news')

        self.label_total_time = tk.Label(self.label_frame, width=45, 
                                         text=self.initial_label_total, bg="cornsilk2", relief="raised")
        
        self.label_total_time.grid(row=2, column=0, rowspan=2, columnspan=2, ipadx=5, ipady=25, sticky='news')

        self.label_combo = tk.Label(self.root, text="Activity Charge Lines")
        self.label_combo.grid(row=0, column=2, rowspan=1, columnspan=4,ipady=10, sticky='news')

    def open_help(self):
        document_path = resource_path(Path("Documents/User Guide.pdf"))
        print(document_path)
        os.startfile(document_path)

    def send_bug_mail(self):
        const = win32com.client.constants
        olMailItem = 0x0
        obj = win32com.client.Dispatch("Outlook.Application")
        newMail = obj.CreateItem(olMailItem)
        newMail.Subject = "Report a Bug"
        newMail.BodyFormat = 2 # olFormatHTML https://msdn.microsoft.com/en-us/library/office/aa219371(v=office.11).aspx
        newMail.HTMLBody = "<HTML><BODY>Please enter the details of your bug or error.</BODY></HTML>"
        newMail.To = "chukwuemekalim.nwaezeigwe@prattwhitney.com"
        newMail.display()

    def open_modify_chargelines(self):
        num_rows = len(self.chargelines) # Number of rows of entry boxes
        self.modify_chargeline_window = tk.Toplevel(self.root)
        self.modify_chargeline_window.title("Modify Chargelines")
        self.modify_chargeline_window.resizable(False, False)
        
        labels = ["Description", "LDN", "Rec. Order", "Network", "Operation", "Sub-O"]
        self.entries = []
        self.edit_buttons = []
        self.delete_button = []
        
        # Create labels
        for i, label in enumerate(labels):
            lbl = ttk.Label(self.modify_chargeline_window, text=label)
            lbl.grid(row=0, column=i, padx=5, pady=5)
        
        # Create rows of entry boxes and buttons
        if num_rows > 0:
            for row in range(num_rows):
                row_entries = []
                for col in range(len(labels)):
                    entry = ttk.Entry(self.modify_chargeline_window)
                    entry.grid(row=row + 1, column=col, padx=5, pady=5)
                    entry.insert(0, self.chargelines[row][col])
                    entry.config(state='disabled')
                    row_entries.append(entry)
                
                self.entries.append(row_entries)
                
                # Add buttons to Edit Chargeline 
                button_edit_chargeline = tk.Button(self.modify_chargeline_window, image=self.photo_edit, borderwidth=0, highlightthickness=1, 
                                                    command=lambda r=row: self.toggle_entries(r))
                button_edit_chargeline.grid(row=row + 1, column=len(labels), padx=5, pady=5, sticky='news')
                self.edit_buttons.append(button_edit_chargeline)

                # Add buttons to Delete Chargelines
                button_delete_chargeline = tk.Button(self.modify_chargeline_window, image=self.photo_delete, borderwidth=0, highlightthickness=1, 
                                                    command=lambda r=row: self.delete_chargeline(r))
                button_delete_chargeline.grid(row=row + 1, column=len(labels)+1, padx=(2, 10), pady=5, sticky='ns')
                self.delete_button.append(button_delete_chargeline)
        else:
            self.no_charge_lbl_frame = tk.Frame(self.modify_chargeline_window)
            self.no_charge_lbl_frame.grid(row=1, column=0, rowspan=20, columnspan=5, sticky='se')
            self.no_charge_lbl = ttk.Label(self.no_charge_lbl_frame, text="No Chargeline Exists, Click Add Chargeline")
            self.no_charge_lbl.grid(row=0, column=0, padx=5, pady=5)


        # Add "Add Chargeline" and "Save and Exit" buttons at the bottom
        self.mod_button_frame = tk.Frame(self.modify_chargeline_window)
        self.mod_button_frame.grid(row=50, column=0, rowspan=20, columnspan=5, sticky='se')
        button_add_chargeline = tk.Button(self.mod_button_frame, text ="Add Chargeline", 
                                   command=self.add_new_charge_line)
        
        button_add_chargeline.grid(row=0, column=1, sticky='nw', pady=(5,5), padx=2)
        
        button_mod_reset = tk.Button(self.mod_button_frame, text ="Reset", width=10, height= 1,
                                  command=self.reset_chargeline)
        
        button_mod_reset.grid(row=0, column=2, sticky='ne', pady=(5,5), padx=2)

        button_save_exit = tk.Button(self.mod_button_frame, text ="Save & Exit", 
                                          width=11, height=1, command=self.save_and_exit)
        
        button_save_exit.grid(row=0, column=3, sticky='nw', pady=(5,5), padx=1)

        # Make the second window modal
        self.modify_chargeline_window.transient(self.root)
        self.modify_chargeline_window.grab_set()
        self.root.wait_window(self.modify_chargeline_window)


    def toggle_entries(self, row):
        if self.edit_buttons[row].cget('image') == str(self.photo_edit):
            self.enable_entries(row)
        else:
            self.save_entries(row)

    def enable_entries(self, row):
        for entry in self.entries[row]:  # Adjust for 0-based index
            entry.config(state='normal')
        self.edit_buttons[row].config(image=self.photo_save)
    
    def save_entries(self, row):
        for entry in self.entries[row]:  
            entry.config(state='disabled')
        self.edit_buttons[row].config(image=self.photo_edit)

    def delete_chargeline(self,row):
        del self.chargelines[row]
        del self.entries[row]
        self.modify_chargeline_window.destroy()
        self.open_modify_chargelines()
    
    def reset_chargeline(self):
        file_path = self.documents_path/ "chargelines.xlsx"      
        self.chargelines = read_excel(file_path)
        self.modify_chargeline_window.destroy()
        self.open_modify_chargelines()

    def save_and_exit(self):
        # Save data from the entry boxes to self.data
        new_row = []
        self.chargelines =[]
        for row_entries in self.entries:
            new_row = []
            for entry in row_entries:
                new_row.append(entry.get())
            if is_array_completely_empty(new_row):
                pass
            else:
                self.chargelines.append(new_row)

        print(self.chargelines)

        columns =['Description', 'LDN', 'Rec. Order', 'Network', 'Operation', 'Sub-O',]
        df = pd.DataFrame(self.chargelines, columns=columns)
        print(df)
        file_path = f"{self.documents_path}/chargelines.xlsx"
        # print(file_path)
        df.to_excel(file_path, index=False)

        # Close the second window
        self.modify_chargeline_window.destroy()

    def add_new_charge_line(self):    
        next_row = len(self.entries) + 1
        if next_row == 1:
            self.no_charge_lbl.destroy()
        print(next_row)
        if next_row <= 15:
            row_entries = []
            for col in range(6):
                entry = ttk.Entry(self.modify_chargeline_window)
                entry.grid(row=next_row, column=col, padx=5, pady=5)
                entry.config(state='normal')
                row_entries.append(entry)
            
            self.entries.append(row_entries)
            
            # Add buttons to Edit Chargeline 
            button_edit_chargeline = tk.Button(self.modify_chargeline_window, image=self.photo_save, borderwidth=0, highlightthickness=1, 
                                                command=lambda r=next_row-1: self.toggle_entries(r))
            button_edit_chargeline.grid(row=next_row, column=6, padx=5, pady=5, sticky='news')
            self.edit_buttons.append(button_edit_chargeline)
            print(len(self.edit_buttons))

            # Add buttons to Delete Chargelines
            button_delete_chargeline = tk.Button(self.modify_chargeline_window, image=self.photo_delete, borderwidth=0, highlightthickness=1, 
                                                  command=lambda r=next_row-1: self.toggle_entries(r))
            button_delete_chargeline.grid(row=next_row, column=7, padx=(2, 10), pady=5, sticky='ns')
            self.delete_button.append(button_delete_chargeline)

        
    def load_icons(self):
        # Get the directory where the script is running
        self.icon_path_start = resource_path(Path("Icons/play.png"))
        self.photo_start = resize_image(self.icon_path_start)

        self.icon_path_stop = resource_path(Path("Icons/stop.png"))
        self.photo_stop = resize_image(self.icon_path_stop)

        self.icon_path_edit = resource_path(Path("Icons/edit.png"))
        self.photo_edit = resize_image(self.icon_path_edit)

        self.icon_path_delete = resource_path(Path("Icons/delete.png"))
        self.photo_delete = resize_image(self.icon_path_delete)

        self.icon_path_save = resource_path(Path("Icons/save.png"))
        self.photo_save = resize_image(self.icon_path_save)



    def create_activity_rows(self):
        file_path = self.documents_path/ "chargelines.xlsx"      
        self.chargelines = read_excel(file_path) 
        self.hour_entry_fields = {}
        self.start_buttons = {}
        # self.stop_buttons = {}
        self.combo_boxes = {}
        self.chargeline_key = {}

        chargline_activities = []

        for rows in self.chargelines:
            chargline_activities.append(rows[0])
            # print(chargline_activities)
        
        for i in range(1, self.total_rows):
            combo_activity = ttk.Combobox(self.root, values=chargline_activities, width=30)
            combo_activity.grid(row=i, column=2, rowspan=1, columnspan=3, padx=(5,5), pady=(5,0), sticky='n')
            combo_activity['values']= chargline_activities
            combo_activity['state']= 'readonly'

            self.combo_boxes[i] = combo_activity
            hour_entry = tk.Entry(self.root, width=7)
            hour_entry.grid(row=i, column=5, pady=(5,0), padx=(0,5), sticky='nw')
            hour_entry.configure(state='normal')
            self.hour_entry_fields[i] = hour_entry 

            button_start_stop = tk.Button(self.root, image=self.photo_start, borderwidth=0, highlightthickness=1, 
                                     command=lambda i=i: self.toggle_start_stop(i))
            
            button_start_stop.grid(row=i, column=6, sticky='news', padx=(0,5))
            self.start_buttons[i] = button_start_stop

            # button_stop = tk.Button(self.root, image=self.photo_stop, borderwidth=0, highlightthickness=0, 
            #                         state="disabled", command=lambda i=i: self.on_stop_button_click(i))
            
            # button_stop.grid(row=i, column=7, sticky='news', padx=(0,10))
            # self.stop_buttons[i] = button_stop
            # button_stop.configure(state="disabled")


        self.button_frame = tk.Frame(self.root)
        self.button_frame.grid(row=50, column=1, rowspan=20, columnspan=5, sticky='se')

        button_add_row = tk.Button(self.button_frame, text ="Add Row", 
                                   command=self.on_add_row_click, width=10, height=1)
        
        button_add_row.grid(row=0, column=2, sticky='nw', pady=(5,5), padx=2)
        
        button_reset = tk.Button(self.button_frame, text ="Reset", width=10, height= 1,
                                  command=self.on_reset_click)
        
        button_reset.grid(row=0, column=1, sticky='ne', pady=(5,5), padx=2)

        button_export_to_sap = tk.Button(self.button_frame, text ="Export to SAP", 
                                         command=self.final_time, width=11, height=1)
        
        button_export_to_sap.grid(row=0, column=4, sticky='nw', pady=(5,5), padx=1)

        # self.total_rows = 6

    def toggle_start_stop(self, row):
        if self.start_buttons[row].cget('image') == str(self.photo_start):
            self.on_start_button_click(row)
        else:
            print('Stop Button Clicked')
            self.on_stop_button_click(row)

    def on_start_button_click(self, row):
        print(f"Start button in row {row} clicked.")
        self.start_buttons[row].config(image=self.photo_stop)
        label_activity = self.combo_boxes[row].get()
        print(label_activity)
        self.initial_ref_time = time.time()
        try:
            self.elapsed_times[row] = float(self.hour_entry_fields[row].get())
        except ValueError:
            self.elapsed_times[row] = 0  # Default to 0 if parsing fails
            self.hour_entry_fields[row].insert(0, "0")
        
        _day_time = 0
        for times in self.hour_entry_fields:
            # print(self.hour_entry_fields[times].get())
            try:
                _day_time += float(self.hour_entry_fields[times].get())
            except ValueError:
                _day_time += 0
                
        self.total_day_time = _day_time
        print(f'Hours {self.total_day_time:.5f}')
        # print(self.elapsed_times[row])
        for rows in self.start_buttons:
            self.start_buttons[rows].configure(state="disabled")
            # self.stop_buttons[rows].configure(state="disabled")
            self.hour_entry_fields[rows].configure(state='disabled')

        self.start_buttons[row].configure(state="normal")
        # self.stop_buttons[row].configure(state="normal")
        self.start_times[row] = time.time() - self.elapsed_times[row]
        self.active_stopwatch = row
        self.update_stopwatch()

    def on_stop_button_click(self, row):
        self.start_buttons[row].config(image=self.photo_start)
        print(f"Stop button in row {row} clicked.")

    # Convert the total elapsed time in seconds to hours (including fractions)
        # self.elapsed_times[rows] = self.total_seconds
        total_hours = self.total_seconds / 3600.0

    # Update the entry field to show the total elapsed time in hours and fractional hours
        self.hour_entry_fields[row].configure(state='normal')
        self.active_stopwatch = None
        self.hour_entry_fields[row].delete(0, tk.END)
        self.hour_entry_fields[row].insert(0, f"{total_hours:.5f}")
        
        # self.stop_buttons[row].configure(state="disabled")
        # self.start_buttons[row].configure(state="normal")
        for rows in self.start_buttons:
            self.start_buttons[rows].configure(state="normal")
            self.hour_entry_fields[rows].configure(state='normal')
        self.active_stopwatch = None

    def update_stopwatch(self):
       
        # total_day_time = 0
        if self.active_stopwatch is not None:
            _day_time = self.total_day_time *3600
            # print(f'This is the total time in seconds {_day_time:.5f}')
            start_seconds = self.elapsed_times[self.active_stopwatch]* 3600
            # Calculate the current elapsed time since the stopwatch started
            current_time = time.time()
            elapsed_time = current_time - self.initial_ref_time
            # Calculate the total time in seconds, including the start offset
            self.total_seconds = start_seconds + elapsed_time
            _day_time = _day_time + elapsed_time
            print(f'This is the total time in seconds {_day_time:.5f}')
            # Convert total_seconds back to hours, minutes, and seconds for display
            # print(start_seconds)
            # print(self.total_seconds)
            activity_label = self.combo_boxes[self.active_stopwatch].get()
            hours_curent, minutes_current, seconds_current = convert_to_time_format(self.total_seconds)
            hours_total, minutes_total, seconds_total = convert_to_time_format(_day_time)
            self.label_current_time.configure(text=f"Activity: {activity_label} \n Time Spent:  {hours_curent:02}:{minutes_current:02}:{seconds_current:02}")
            self.label_total_time.configure(text=f"Total Time Elapsed Today: \n {hours_total:02}:{minutes_total:02}:{seconds_total:02}")
            self.hour_entry_fields[self.active_stopwatch].configure(state='normal')
            self.hour_entry_fields[self.active_stopwatch].delete(0, tk.END)
            self.hour_entry_fields[self.active_stopwatch].insert(0, f"{self.total_seconds/3600:.2f}")
            self.hour_entry_fields[self.active_stopwatch].configure(state='disabled')
            self.root.after(1000, self.update_stopwatch)  # Schedule the next update

    def final_time(self):
        place_holder = self.show_custom_messagebox(title="Warning", message='This will upload your time to SAP', box_type='okcancel')
        if place_holder:
            export_time = []
            chargelines_export = []
            row_count = len(self.hour_entry_fields)
            export_count = 0
            for rows in self.hour_entry_fields:
                if self.hour_entry_fields[rows].get() != "":
                    print(self.combo_boxes[rows].current())
                    chargline_hour = self.hour_entry_fields[rows].get()
                    index = self.combo_boxes[rows].current() # Gives positional index of selection
                    if index != -1:
                        chargelines_export = self.chargelines[index].copy()
                        chargelines_export.append(chargline_hour)
                        export_time.append(chargelines_export)
                    else: 
                        self.show_custom_messagebox("Error", f"Activity Chargline Not Valid. Excluded From Export\n Ref: Row {rows}", "info")
                else:
                    export_count = export_count + 1 # Error Handler to determine if all rows are empty

            # print(f'<<{export_count}>>  {len(self.hour_entry_fields)}')

            if export_count != row_count:
                columns =['Description', 'LDN', 'Rec. Order', 'Network', 'Operation', 'Sub-O', 'Time']
                df = pd.DataFrame(export_time, columns=columns)
                SAP = module.SapApi()
                SAP.cat2_input_time(df)
                self.show_custom_messagebox('Operation Complete', 'Succesfully Uploaded Time', 'info')
                place_holder = self.show_custom_messagebox('Save Time Log', 'Do You Want To Save Time Log?', 'yesno')
                if place_holder:
                    df['Time'] = df['Time'].astype(float)
                    df['Time'] = df['Time'].round(1)
                    self.save_time_log_to_excel(df, open=False)
                # print(df)
            else:
                self.show_custom_messagebox('Error', 'No Time to Export', 'info')

    def on_add_row_click(self):
        # Determine the next available row
        next_row = max(self.combo_boxes.keys(), default=0) + 1
        if next_row <= 15:
            chargline_activities = [row[0] for row in self.chargelines]

            # Create and place a new widgets
            combo_activity = ttk.Combobox(self.root, values=chargline_activities, width=30, state='readonly')
            combo_activity.grid(row=next_row, column=2, rowspan=1, columnspan=3, padx=(5, 5), pady=(5, 0), sticky='n')
            self.combo_boxes[next_row] = combo_activity

            hour_entry = tk.Entry(self.root, width=7)
            hour_entry.grid(row=next_row, column=5, pady=(5, 0), padx=(0, 5), sticky='nw')
            self.hour_entry_fields[next_row] = hour_entry

            button_start_stop = tk.Button(self.root, image=self.photo_start, borderwidth=0, highlightthickness=1,
                                    command=lambda i=next_row: self.toggle_start_stop(i))
            button_start_stop.grid(row=next_row, column=6, sticky='news', padx=(0,5))
            self.start_buttons[next_row] = button_start_stop

            # button_stop = tk.Button(self.root, image=self.photo_stop, borderwidth=0, highlightthickness=0, state="disabled",
            #                         command=lambda i=next_row: self.on_stop_button_click(i))
            # button_stop.grid(row=next_row, column=7, sticky='news', padx=(0,10))
            # self.stop_buttons[next_row] = button_stop
            # button_stop.configure(state="disabled")
        else:
            print('Max number of rows reached')
            self.show_custom_messagebox(title='ERROR', message="Max Number of Rows Reached", box_type='err')

        if self.active_stopwatch:
            self.start_buttons[next_row].configure(state="disabled")
            # cha[next_row].configure(state="disabled")
            self.hour_entry_fields[next_row].configure(state='disabled')

    def on_reset_click(self):
        self.active_stopwatch = None
        place_holder = self.show_custom_messagebox(title='Warning', message='This will reset window to default. Continue?', box_type='yesno')
        if place_holder:
            # Destroy all dynamically added widgets
            for widget_dict in (self.combo_boxes, self.hour_entry_fields, self.start_buttons): #,self.stop_buttons):
                for widget in widget_dict.values():
                    widget.destroy()
            # Clear the dictionaries to remove references to the destroyed widgets and reinitialize the widgets
            self.combo_boxes.clear()
            self.hour_entry_fields.clear()
            self.start_buttons.clear()
            # self.stop_buttons.clear()
            self.create_activity_rows()
            self.label_current_time.configure(text=self.initial_label_current)
            self.label_total_time.configure(text=self.initial_label_total)

    def show_custom_messagebox(self, title="Default Title", message="Your custom message here.", box_type="info"):
        """
        Displays a Tkinter messagebox with a custom title, message, and type.
        
        Parameters:
        - title (str): The title of the messagebox.
        - message (str): The message to be displayed in the messagebox.
        - box_type (str): The type of messagebox to display ('info', 'warning', 'error').
        """
        # Map the box_type to the appropriate messagebox function
        box_type_map = {
            'info': tk.messagebox.showinfo,
            'warning': tk.messagebox.showwarning,
            'error': tk.messagebox.showerror,
            'yesno': tk.messagebox.askyesno,
            "okcancel": tk.messagebox.askokcancel
        }
        
        # Fetch the appropriate messagebox function based on box_type
        messagebox_function = box_type_map.get(box_type, tk.messagebox.showinfo)

        # For 'yesno', directly return the result of the messagebox function call
        if box_type in {'yesno', 'okcancel'}:
            return messagebox_function(title, message)
        else:
            messagebox_function(title, message)
            return None  # For informational, warning, or error message boxes, return None or a default value

    def save_time_log_to_excel(self, df, open=False):
        today = datetime.now()
        date_str = today.strftime('%m_%d_%y')
        file_path = f"{self.documents_path}/Time Records/time_charge_{date_str}.xlsx"
        print(file_path)
        df.to_excel(file_path, index=False)
        if open:
            os.startfile(file_path)