import tkinter as tk
from tkinter import ttk
from tkinter import font
import time
from gui import CatsTimeTracker
import functions

if __name__ == "__main__":
    functions.create_sap_time_tracker_folder()
    root = tk.Tk()
    app = CatsTimeTracker(root)
    root.mainloop()

