import win32com.client
import subprocess
import time
import pandas as pd
import time
import re
# import os
import numpy as np

from functions_SAP import *

class SapApi:
    def __init__(self):
        # Class Attributes
        self.bln_dlt = None 
        self.bln_sn = None
        self.Connection = None
        self.li_info = None
        self.qn_info = None
        self.qn_number = None
        self.SapGuiAuto = None
        self.session = None
        self.sorted_sns_li = None
        self.user = ""

        self.connect_sap()


    def connect_sap(self):
            try:
                self.SapGuiAuto = win32com.client.GetObject('SAPGUI')
                # print (SapGuiAuto)
                if isinstance(self.SapGuiAuto, win32com.client.CDispatch):
                        print("SAP IS OPEN")
            except: 
                print("SAP IS NOT OPEN")
                subprocess.Popen("C:\\Program Files (x86)\\SAP\\FrontEnd\\SAPgui\\saplgpad.exe")
                time.sleep(4)
                self.SapGuiAuto = win32com.client.GetObject('SAPGUI')
                pass
            App = self.SapGuiAuto.GetScriptingEngine

            try:
                self.Connection = App.Children(0)
                
                #print("Number of connections open" , Connection.Children.Count)
                if isinstance(self.Connection, win32com.client.CDispatch):
                          print("LOGGED IN")
            except Exception as error:
                    print("SAP IS NOT LOGGED IN")
                    #print("An exception occurred:", error)
                    self.Connection = App.openconnection("P10 (SSO) - PW ECC Production System", True)
                    pass
            self.session = self.Connection.Children(0)
            self.user = self.session.Info.User

    def get_transaction_name(self):
        """Retrieves the current transaction name."""
        return self.session.Info.Transaction
       
    def open_transaction(self, transaction):
        """Opens a given transaction."""
        self.session.StartTransaction(transaction)

    def new_session(self):
        """Creates a new session."""
        self.session.createSession()
        time.sleep(4)

    def get_data_for_serial_number(self, sn):
        """Fetch data for a single serial number."""
        try:
            self.session.findById("wnd[0]/usr/txtP_SERIAL").Text = sn
            self.session.findById("wnd[0]/tbar[1]/btn[8]").press()
            
            if self.session.ActiveWindow.Name == "wnd[1]":
                self.session.findById("wnd[1]/tbar[0]/btn[12]").press()
            
            self.session.findById("wnd[0]/tbar[1]/btn[13]").press()
            
            row_count = self.session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").VisibleRowCount
            self.select_layout("ZSNQ_Layout_Macro")  # Set layout only if needed and not set
            
            rows_data = []
            for i in range(row_count):
                rows_data.append([
                    self.session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").getCellValue(i, col_name)
                    for col_name in ["QMNUM", "ZZMQI", "FETXT", "QMTXT", "OTKTXTCD", "STRMN"]
                ])
                
            # Navigate back
            self.session.findById("wnd[0]/tbar[0]/btn[3]").press()
            self.session.findById("wnd[0]/tbar[0]/btn[3]").press()
            
            # print(rows_data)

            return rows_data
        except Exception as e:
            # Handle or log the error appropriately
            return []

    def zsnq_search (self, input_sheet_path,output_sheet_path):
        """Performs a search based on the provided serial number."""
        df_input= pandas_read_file(input_sheet_path).dropna(axis='columns', how='all')

        # print(len(df_input.iloc[:, 0]))
        
        # Prepare for bulk data collection
        all_data = []
        total_sns = len(df_input)

        for index, sn in enumerate(df_input.iloc[:, 0], 1):  # Assuming serial numbers are in the first column
            sn_data = self.get_data_for_serial_number(sn)
            print(f"Processing SN {index} of {total_sns}... ({total_sns - index} SNs left)")
            for row in sn_data:
                all_data.append(row)

        df_output = pd.DataFrame(all_data, columns=["QN", "MQI", "Defect Text", "Detailed Defect Text", "Disposition", "Required Start Date"])
        
        # print(df_output)

        df_output.to_excel(output_sheet_path, index=False)

    def select_layout(self, layout_name):
        """Set a specific layout in SAP GUI."""
        try:
            self.session.findById("wnd[0]/tbar[0]/btn[3]")
            self.session.findById("wnd[0]/tbar[1]/btn[23]").press()
            self.session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").setCurrentCell(0, "TEXT")
            self.session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").contextMenu()
            self.session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").selectContextMenuItem("&FILTER")
            self.session.findById("wnd[2]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-LOW").Text = layout_name
            self.session.findById("wnd[2]/tbar[0]/btn[0]").press()
            self.session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").clickCurrentCell()
        except Exception as e:
            if str(e).find('-2147024809') != -1:  # This is a simplified check for the specific error
                print("Layout Doesn't exist. New Layout to be Created")
                self.create_new_layout(layout_name)
            else:
                print("An unexpected error occurred:", str(e))

    def create_new_layout(self, layout_name):
        """Create a new layout if the specified one doesn't exist."""
        table_col = ["Required Start", "Description", "Object Part Code Text"]
        for value in table_col:
            self.session.findById("wnd[0]/tbar[1]/btn[32]").press()
            self.session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER1_LAYO/shellcont/shell").pressToolbarButton("&FIND")
            self.session.findById("wnd[2]/usr/chkGS_SEARCH-EXACT_WORD").Selected = True
            self.session.findById("wnd[2]/usr/txtGS_SEARCH-VALUE").Text = value
            self.session.findById("wnd[2]/tbar[0]/btn[0]").press()
            self.session.findById("wnd[2]/tbar[0]/btn[12]").press()
            self.session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/btnAPP_WL_SING").press()
        
        self.session.findById("wnd[1]/tbar[0]/btn[5]").press()
        self.session.findById("wnd[2]/usr/tabsG50_TABSTRIP/tabpTAB_D0501/ssubD0505_SUBSCREEN:SAPLSLVC_DIALOG:0501/txtLTDX-VARIANT").Text = layout_name
        self.session.findById("wnd[2]/usr/tabsG50_TABSTRIP/tabpTAB_D0501/ssubD0505_SUBSCREEN:SAPLSLVC_DIALOG:0501/txtLTDXT-TEXT").Text = "ZSNQ_Layout_Macro"
        self.session.findById("wnd[2]/tbar[0]/btn[0]").press()
        self.session.findById("wnd[1]/tbar[0]/btn[0]").press()
    
    def _get_total_rows_cat2(self):
        """Get the total number of relevant rows in an SAP table."""
        total_rows = 0  # Initialize the total rows counter

        # Get the count of all rows in the SAP table
        rows_count = self.session.findById("wnd[0]/usr/subCATS002:SAPLCATS:2200/tblSAPLCATSTC_CATSD").Rows.Count
        # print(rows_count)
        # Loop through each row in the table, starting from row 2 to rows_count - 1 (inclusive)
        for i in range(2, rows_count):
            # Check if specific fields in the row are not empty
            lstar_text = self.session.findById(f"wnd[0]/usr/subCATS002:SAPLCATS:2200/tblSAPLCATSTC_CATSD/ctxtCATSD-LSTAR[1,{i}]").Text
            awart_text = self.session.findById(f"wnd[0]/usr/subCATS002:SAPLCATS:2200/tblSAPLCATSTC_CATSD/ctxtCATSD-AWART[6,{i}]").Text
            
            if lstar_text != "" or awart_text != "":
                total_rows += 1  # Increment the counter if the condition is met

        # Optionally, print the total row count for debugging
        # print(total_rows)

        return total_rows
    
    def _fetch_sap_row_data(self, row_index):
        """Fetch data from specific fields in an SAP table row.
        
        Args:
            row_index (int): The index of the row from which to fetch data, adjusted for 1-based indexing in SAP.

        Returns:
            tuple: A tuple containing the text values of the Rec_Check, Network_Check, Op_Check, and SubO_Check fields.
        """
        # Adjusting row_index for 1-based indexing used by SAP GUI Scripting since data starts from Col 2
        row_index_adj = row_index + 1
        
        # Fetch values from SAP GUI
        rec_check = self.session.findById(f"wnd[0]/usr/subCATS002:SAPLCATS:2200/tblSAPLCATSTC_CATSD/ctxtCATSD-RAUFNR[2,{row_index_adj}]").Text
        network_check = self.session.findById(f"wnd[0]/usr/subCATS002:SAPLCATS:2200/tblSAPLCATSTC_CATSD/ctxtCATSD-RNPLNR[3,{row_index_adj}]").Text
        op_check = self.session.findById(f"wnd[0]/usr/subCATS002:SAPLCATS:2200/tblSAPLCATSTC_CATSD/ctxtCATSD-VORNR[4,{row_index_adj}]").Text
        subo_check = self.session.findById(f"wnd[0]/usr/subCATS002:SAPLCATS:2200/tblSAPLCATSTC_CATSD/ctxtCATSD-UVORN[5,{row_index_adj}]").Text
        
        return rec_check, network_check, op_check, subo_check

    def _update_sap_row(self, row_index, time_value, day_index, column_index):
        """Update a specific field in an SAP table row with a new time value.
        
        Args:
            row_index (int): The 0-based index of the row to update.
            time_value (float or str): The new time value to set in the field.
            table_day (int): The day offset used to identify the correct day column.
            column_index (int): The index of the column where the time value should be updated.
        """
        # Adjust row_index for 1-based indexing used by SAP GUI Scripting
        row_index_adj = row_index + 1
        
        # Construct the SAP GUI Scripting ID for the specific field
        field_id = f"wnd[0]/usr/subCATS002:SAPLCATS:2200/tblSAPLCATSTC_CATSD/txtCATSD-DAY{day_index}[{column_index},{row_index_adj}]"
        
        # Update the field with the new time value
        self.session.findById(field_id).Text = str(time_value)

    def cat2_input_time(self, data_frame, date):
        """Inputs time into SAP based on input data frame."""
        # Load Excel workbook and select the CATS sheet
        df_input = data_frame
        # df_input = pandas_read_file(input_sheet_path,dtype=str)
        # df_input = df_input.replace(np.nan, "")

        # Initialize necessary variables
        # addresses = []
        day_of_week = pd.Timestamp(date).dayofweek + 1  # Adjusted to match VBA Weekday function, assuming Monday as first day
        
        # Find cells with data in the specified range
        # for index, value in enumerate(df_input.loc[:, 'Time'],0):  # Assuming data starts from G11 to G200
        #     value = value.strip()
        #     if value != "":
        #         addresses.append(df_input.index[index])

        self.open_transaction("CAT2")
        employee_number = ''.join(re.findall(r'\d+', self.user))
        self.session.findById("wnd[0]/usr/ctxtCATSFIELDS-PERNR").Text = employee_number
        self.session.findById("wnd[0]/tbar[1]/btn[5]").press()
        self._update_time_records(df_input, day_of_week)
        
    
    def _check_sap_access(self):
        try:
            self.session.findById("wnd[0]").Iconify()
            self.open_transaction("ZHPR")
            self.session.findById("wnd[0]/usr/ctxtS_BNAME-LOW").text = self.user
            self.session.findById("wnd[0]/tbar[1]/btn[8]").press()
            array =["Id", "Text"]
            JSONTree = self.session.GetObjectTree("wnd[0]", array)
            if 'ZS_SCR' in JSONTree:
                scripting_access = True
            else:
                scripting_access = False
        except:
            scripting_access = False
        
        self.session.findById("wnd[0]/tbar[0]/btn[3]").press()
        self.session.findById("wnd[0]/tbar[0]/btn[3]").press()

        return scripting_access




    def _update_time_records(self, df_input, day_index):
        """Update or insert time records in SAP based on preprocessed input."""
        total_rows = self._get_total_rows_cat2()  # Placeholder for method to fetch total rows in SAP table
        print('Initial Toal Rows ', total_rows)

        for idx,rows in df_input.iterrows():
            # Extract values for each relevant column
            time_value = round(float(rows['Time']), 1)
            record_number = rows['Rec. Order']
            network = rows['Network']
            operation = rows['Operation']
            sub_operation = rows['Sub-O']

            found = False
            i = 1  
            while not found and i <= total_rows:
                # Fetch values from SAP to check against
                rec_check, network_check, op_check, sub_op_check = self._fetch_sap_row_data(i)
                
                if network_check == network and op_check == operation and sub_op_check == sub_operation and rec_check == record_number:
                    self._update_sap_row(i, time_value, day_index, day_index + 8)
                    found = True
                    i += 1
                elif i == total_rows:
                    print("record not found")
                    total_rows += 1
                    print("Row count ", total_rows)
                    self._insert_new_record(total_rows, record_number, network, operation, sub_operation, time_value, day_index)
                    i += 1
                    found = True                     
                else:
                    i += 1
    
    def _insert_new_record(self, row_index, record_number, network, operation, sub_operation, time_value, day_index):
            """Insert a new record into SAP at the specified row index."""
            field_id_record =f"wnd[0]/usr/subCATS002:SAPLCATS:2200/tblSAPLCATSTC_CATSD/ctxtCATSD-RAUFNR[2,{row_index + 1}]" #To account for SAP based indexing
            field_id_network =f"wnd[0]/usr/subCATS002:SAPLCATS:2200/tblSAPLCATSTC_CATSD/ctxtCATSD-RNPLNR[3,{row_index + 1}]"
            field_id_operation =f"wnd[0]/usr/subCATS002:SAPLCATS:2200/tblSAPLCATSTC_CATSD/ctxtCATSD-VORNR[4,{row_index + 1}]"
            field_id_suboperation = f"wnd[0]/usr/subCATS002:SAPLCATS:2200/tblSAPLCATSTC_CATSD/ctxtCATSD-UVORN[5,{row_index + 1}]"

            self.session.findById(field_id_record).text = record_number
            self.session.findById(field_id_network).text = network
            self.session.findById(field_id_operation).text = operation
            self.session.findById(field_id_suboperation).text = sub_operation

            sap_row = row_index #To account for SAP based indexing
            self._update_sap_row(sap_row, time_value, day_index, day_index + 8)
    
    def fetch_data(self,quality_notification , bln_sn=False, bln_dlt=False):
        self.bln_sn = bln_sn
        self.bln_dlt = bln_dlt
        #Initialize Empty Data Frame
        self.qn_number = quality_notification
        self.qn_info = pd.DataFrame()

        self.session.findById(r"wnd[0]/usr/ctxtRIWO00-QMNUM").text = quality_notification 
        self.session.findById(r"wnd[0]").sendVKey(0)

        # Notification status and description
        notification_status = self.session.findById(r"wnd[0]/usr/subSCREEN_1:SAPLIQS0:1070/txtRIWO00-STTXT").text
        description = self.session.findById(r"wnd[0]/usr/subSCREEN_1:SAPLIQS0:1070/txtVIQMEL-QMTXT").text
        
        # Reference Object Tab
        self.session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\\TAB01").select()

        notification_type = self.session.findById(r"wnd[0]/usr/subSCREEN_1:SAPLIQS0:1070/subNOTIF_TYPE:SAPLIQS0:1071/cmbRIWO00-QMARTE").text.strip()

        # print(notification_type)

        # Configuration dictionary for each notification type
        config_notification_type = {
            "Vendor Error Manual": {
                "material_path_suffix": "3010",
                "vendor_path_suffix": "0500",
                "vendor_name_suffix":"txtRQM02",
                "vendor_name_id": "NAME_LIEF",
                "vendor_number_id": "LIFNUM"
            },
            "Internal Qual Notif.": {
                "material_path_suffix": "3020",
                "vendor_path_suffix": "0600",
                "vendor_name_suffix":"ctxtRQM02",
                "vendor_name_id": "NAME_AUTOR",
                "vendor_number_id": "BUNAME"
            }
        }

        # Determine which configuration to use based on notification_type
        cfg = config_notification_type.get(notification_type)
        # print(f"wnd[0]/usr/tabsTAB_GROUP_10/tabp10\\TAB03/ssubSUB_GROUP_10:SAPLIQS0:7235/subCUSTOM_SCREEN:SAPLIQS0:7212/subSUBSCREEN_1:SAPLIQS0:7730/subANSPRECH:SAPLIPAR:{cfg['vendor_path_suffix']}/txtRQM02-{cfg['vendor_name_id']}")
        
        if cfg:  # Check if the notification_type is in the config
            print(f'{notification_type} loop')
            
            # Fetch material details
            material_number = self.session.findById(f"wnd[0]/usr/tabsTAB_GROUP_10/tabp10\\TAB01/ssubSUB_GROUP_10:SAPLIQS0:7235/subCUSTOM_SCREEN:SAPLIQS0:7212/subSUBSCREEN_1:SAPLIQS0:7322/subOBJEKT:SAPMQM00:{cfg['material_path_suffix']}/ctxtRQM00-MATNR").text
            material_name = self.session.findById(f"wnd[0]/usr/tabsTAB_GROUP_10/tabp10\\TAB01/ssubSUB_GROUP_10:SAPLIQS0:7235/subCUSTOM_SCREEN:SAPLIQS0:7212/subSUBSCREEN_1:SAPLIQS0:7322/subOBJEKT:SAPMQM00:{cfg['material_path_suffix']}/txtRIWO00-MATKTX").text
            operation_number = self.session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\\TAB01/ssubSUB_GROUP_10:SAPLIQS0:7235/subCUSTOM_SCREEN:SAPLIQS0:7212/subSUBSCREEN_5:SAPLIQS0:7900/subUSER0001:SAPLXQQM:9101/txtVIQMEL-ZZOPRNR").text
            military_commercial = self.session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\\TAB01/ssubSUB_GROUP_10:SAPLIQS0:7235/subCUSTOM_SCREEN:SAPLIQS0:7212/subSUBSCREEN_5:SAPLIQS0:7900/subUSER0001:SAPLXQQM:9101/cmbVIQMEL-ZZMIL").text.strip()

            # Contact Persons Tab
            self.session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\\TAB03").Select() 
            
            # Fetch vendor details
            vendor_name = self.session.findById(f"wnd[0]/usr/tabsTAB_GROUP_10/tabp10\\TAB03/ssubSUB_GROUP_10:SAPLIQS0:7235/subCUSTOM_SCREEN:SAPLIQS0:7212/subSUBSCREEN_1:SAPLIQS0:7730/subANSPRECH:SAPLIPAR:{cfg['vendor_path_suffix']}/{cfg['vendor_name_suffix']}-{cfg['vendor_name_id']}").text
            
            vendor_number = self.session.findById(f"wnd[0]/usr/tabsTAB_GROUP_10/tabp10\\TAB03/ssubSUB_GROUP_10:SAPLIQS0:7235/subCUSTOM_SCREEN:SAPLIQS0:7212/subSUBSCREEN_1:SAPLIQS0:7730/subANSPRECH:SAPLIPAR:{cfg['vendor_path_suffix']}/ctxtVIQMEL-{cfg['vendor_number_id']}").text
        else:
            print("Invalid notification type")

       #Dictionary with data
        data_row = {"Material": material_number, "Material Name": material_name, "Vendor Name":vendor_name, "Vendor Number":vendor_number,\
                    "Operation Number": operation_number, "Military or Commercial": military_commercial,\
                    "Notification Type": notification_type, "Description": description}
            
        #Defects Tab
        self.session.findById(r"wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB10").Select() 

        li_qty = self.session.findById(r"wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB10/ssubSUB_GROUP_10:SAPLIQS0:7210/tabsTAB_GROUP_20/tabp20\TAB01/ssubSUB_GROUP_20:SAPLIQS0:7110/txtRIWO00-TOPOS").text
        data_row.update({"LI Quantity": li_qty})

        # Append the row to the DataFrame
        self.qn_info = pd.concat([self.qn_info,pd.DataFrame([data_row])],ignore_index=True)
        # print(self.qn_info)  

        self.li_info = pd.DataFrame(columns=['Line Item','Disposition','Defect Type','Short Text','Long Text','Defect Class',\
                                             'MQI', 'QN', 'Drawing Location','Engineering Change Number','Quantity'])
        data_list = []
        self.total_serial_numbers = []
        for i in range(0,int(li_qty)):
            data_list.append(self._read_defect(i))
            
        self.li_info = pd.concat([self.li_info,pd.DataFrame(data_list)],ignore_index=True)

        # self.defect_texts = self.li_info.loc['Long Text'] (work on this)

        #Sorts the serial number and exports it to an excel file 
        # self.serial_number_sorter()
        # print(self.li_info)  
    
    def _read_defect(self, row, bln_cause=False, bln_caa=False):
        """Reads a single line in the defect tab, with optional arguments for cause, CAA, and DLT."""

        self.session.findById(r"wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB10/ssubSUB_GROUP_10:SAPLIQS0:7210/tabsTAB_GROUP_20/tabp20\TAB01/ssubSUB_GROUP_20:SAPLIQS0:7110/tblSAPLIQS0POSITION_VIEWER").verticalScrollbar.Position = row

        data = {
            "Line Item": self.session.findById(r"wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB10/ssubSUB_GROUP_10:SAPLIQS0:7210/tabsTAB_GROUP_20/tabp20\TAB01/ssubSUB_GROUP_20:SAPLIQS0:7110/tblSAPLIQS0POSITION_VIEWER/txtVIQMFE-POSNR[0,0]").text,
            "Disposition": self.session.findById(r"wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB10/ssubSUB_GROUP_10:SAPLIQS0:7210/tabsTAB_GROUP_20/tabp20\TAB01/ssubSUB_GROUP_20:SAPLIQS0:7110/tblSAPLIQS0POSITION_VIEWER/ctxtRIWO00-TXTCDOT[3,0]").text,
            "Defect Type": self.session.findById(r"wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB10/ssubSUB_GROUP_10:SAPLIQS0:7210/tabsTAB_GROUP_20/tabp20\TAB01/ssubSUB_GROUP_20:SAPLIQS0:7110/tblSAPLIQS0POSITION_VIEWER/ctxtRIWO00-TXTCDFE[6,0]").text,
            "Short Text": self.session.findById(r"wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB10/ssubSUB_GROUP_10:SAPLIQS0:7210/tabsTAB_GROUP_20/tabp20\TAB01/ssubSUB_GROUP_20:SAPLIQS0:7110/tblSAPLIQS0POSITION_VIEWER/txtVIQMFE-FETXT[7,0]").text,
            "Long Text": "",
            "Defect Class": self.session.findById(r"wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB10/ssubSUB_GROUP_10:SAPLIQS0:7210/tabsTAB_GROUP_20/tabp20\TAB01/ssubSUB_GROUP_20:SAPLIQS0:7110/tblSAPLIQS0POSITION_VIEWER/cmbVIQMFE-FEQKLAS[11,0]").text
        }

        if self.bln_dlt:
            long_text_button = self.session.findById(r"wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB10/ssubSUB_GROUP_10:SAPLIQS0:7210/tabsTAB_GROUP_20/tabp20\TAB01/ssubSUB_GROUP_20:SAPLIQS0:7110/tblSAPLIQS0POSITION_VIEWER/btnQMICON-LTFEHLER[8,0]")
            data["Long Text"] = self._copy_long_text(long_text_button)

        

        self.session.findById(r"wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB10/ssubSUB_GROUP_10:SAPLIQS0:7210/tabsTAB_GROUP_20/tabp20\TAB01/ssubSUB_GROUP_20:SAPLIQS0:7110/tblSAPLIQS0POSITION_VIEWER").getAbsoluteRow(row).selected = True
        self.session.findById(r"wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB10/ssubSUB_GROUP_10:SAPLIQS0:7210/tabsTAB_GROUP_20/tabp20\TAB01/ssubSUB_GROUP_20:SAPLIQS0:7110/btnDETAIL").press()
        
        qty = int(float(self.session.findById("wnd[1]/usr/txtVIQMFE-FMGFRD").text) + float(self.session.findById("wnd[1]/usr/txtVIQMFE-FMGEIG").text))
        
        
        data.update({
            "MQI": self.session.findById("wnd[1]/usr/subUSER0002:SAPLXQQM:9210/txtVIQMFE-ZZMQI").text,
            "QN": self.session.findById("wnd[0]/usr/subSCREEN_1:SAPLIQS0:1070/ctxtVIQMEL-QMNUM").text,
            "Drawing Location": self.session.findById("wnd[1]/usr/subUSER0002:SAPLXQQM:9210/txtVIQMFE-ZZDRWLOC").text,
            "Engineering Change Number": self.session.findById("wnd[1]/usr/subUSER0002:SAPLXQQM:9210/txtVIQMFE-ZZECNR").text,
            "Quantity": str(qty)
            # "Serial Numbers": serial_numbers
        })

        if self.bln_sn:
            serial_numbers = self._retrieve_serial_numbers()
            for sn in serial_numbers:
                self.total_serial_numbers.append(sn)

            data.update({"Serial Numbers": serial_numbers})

        self.session.findById("wnd[1]/tbar[0]/btn[12]").press()
            
        return data
    
    def _copy_long_text(self, gui_button):
        """
        Fetches and returns the long text from an SAP GUI text window.

        Args:
            gui_button: The SAP GUI button object that opens the long text window.

        Returns:
            A string containing the long text, or "NO LONG TEXT" if not applicable.
        """
        long_text = ""
        
        if gui_button.iconName == "B_TXDP":  # Check if the correct icon is present
            gui_button.setFocus()
            gui_button.press()

            # Navigate to Download Text as RTF file
            self.session.findById("wnd[0]/mbar/menu[0]/menu[4]").select()
            self.session.findById("wnd[1]/tbar[0]/btn[0]").press()
            path = r'C:\Temp\temp_dlt.rtf'
            self.session.findById(r"wnd[2]/usr/ctxtITCTK-TDFILENAME").text = path
            self.session.findById("wnd[2]/tbar[0]/btn[0]").press()
            
            # If it opens a dialog asking to save in temp folder
            if self.session.ActiveWindow.Name == "wnd[3]":
                self.session.findById("wnd[3]/tbar[0]/btn[1]").press()

            time.sleep(2)
            
            # Reads Long Text File and deletes the file once closed.
            long_text = read_rich_text_file(path)

            self.session.findById("wnd[0]/tbar[0]/btn[3]").press()
        else:
            long_text = "NO LONG TEXT"

        return long_text

    def _retrieve_serial_numbers(self):
        serial_numbers = []
        self.session.findById("wnd[1]/usr/subUSER0002:SAPLXQQM:9210/btnSERIAL_NUMBERS").press()
        sn_quantity = int(self.session.findById("wnd[2]/usr/txtGV_QTY").text)
        if sn_quantity == 1:
            self.session.findById("wnd[2]/tbar[0]/btn[12]").press()
            self.session.findById("wnd[1]/usr/subUSER0002:SAPLXQQM:9210/btnSERIAL_HSFX").press()
            serial_numbers.append(self.session.findById("wnd[2]/usr/tblSAPLZ20SME701TC1000/txtI_LISTTAB-FIELD[0,0]").text.strip())
        else:
            for serial_number in range(0, sn_quantity):
                self.session.findById("wnd[2]/usr/tblSAPLZ_QME409_SERNR_POPUPTC_SERNR").verticalScrollbar.position = serial_number
                serial_numbers.append(self.session.findById("wnd[2]/usr/tblSAPLZ_QME409_SERNR_POPUPTC_SERNR/txtGK_SERNR-VALUE[0,0]").text.strip())
        self.session.findById("wnd[2]/tbar[0]/btn[12]").press()
        # print(serial_numbers)
        return serial_numbers

    def serial_number_sorter(self, open = True):
        df = pd.DataFrame(self.total_serial_numbers, columns=['Serial Numbers']).drop_duplicates()

        index_all_sns = []

        for _, li_info_row in self.li_info.iterrows():
            # Assuming 'Serial Numbers' is a list of serial numbers in each row
            index_li_sns = []
            serial_numbers = li_info_row['Serial Numbers']
            for sn in serial_numbers:
                index_li_sns.extend(df.index[df['Serial Numbers'] == sn])
                # print(store)
            index_all_sns.append(index_li_sns)

        # print(store_big)


        for i,sns in enumerate(index_all_sns,0):
            line_item = f"LI {self.li_info.iloc[i]['Line Item']}"
            df[line_item] = ' '
            df.loc[sns, line_item] = "X"        
        print('Checking')
        self.sorted_sns_li = df
        # Get the path to the desktop directory

        desktop_path = Path.home() / 'Desktop'

        detail_sn_path = f'{desktop_path}\\{self.qn_number}_Detail_Serial.xlsx'
        self.sorted_sns_li.to_excel(detail_sn_path, index=False)
        if open:   
          os.startfile(detail_sn_path)  

# test = SapApi()
# test._check_sap_access()
# print(test.get_transaction_name())
# test.open_transaction("CAT2")
# test.new_session()
# # print(test.get_transaction_name())
# SAP = SapApi()
# SAP.cat2_input_time(r"C:\Users\M337199\Desktop\Excel\inputtime.csv")
# SAP.open_transaction("QM03")
# SAP.fetch_data("5002847508", bln_sn=True)
# SAP.serial_number_sorter()
# print(SAP.sorted_sns_li)
# # text = SAP.defect_texts
# process_defects(text.replace('\n', ' '))

