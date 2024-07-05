import pandas as pd
import numpy as np
from pathlib import Path
import platform

import tkinter as tk
import win32com.client
from win32com.client import Dispatch, constants

const=win32com.client.constants
olMailItem = 0x0
obj = win32com.client.Dispatch("Outlook.Application")
newMail = obj.CreateItem(olMailItem)
newMail.Subject = "Report a Bug"
newMail.BodyFormat = 2 # olFormatHTML https://msdn.microsoft.com/en-us/library/office/aa219371(v=office.11).aspx
newMail.HTMLBody = "<HTML><BODY>Please enter the details of your bug or error.</BODY></HTML>"
newMail.To = "chukwuemekalim.nwaezeigwe@prattwhitney.com"
newMail.display()