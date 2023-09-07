import os
import time
import win32com.client as win32
import comtypes.client as cc

from dataclasses import dataclass




@dataclass
class RunExcelMarcos:
    '''
    Macro runner for Excel

    Parameters
    ----------
    macros (list | str): Name of macro(s) to run
    file_path (str): Path to Excel file
    refresh_workbook (bool): Refresh workbook calculations after running macro(s)
    save_workbook (bool): Save workbook after running macro(s)
    disable_alerts (bool): Disable Excel pop-up alerts
    sleep_time (int): Time to sleep between macros
    setting_level (int): Excel security setting level
    '''

    macros: list | str = None
    file_path: str = None
    refresh_workbook: bool = True
    save_workbook: bool = False
    sleep_time: int = None

    def run_function(self) -> bool:
        
        '''
        Run the RunExcelMarco which runs macro(s) in Excel

        Returns
        -------
        bool: True if macro(s) ran successfully
        '''

        try:
            # Launch Excel and Disable Alerts
            excel_app = cc.CreateObject("Excel.Application")
            excel_app.DisplayAlerts = False            
            # excel_app.Application.AutomationSecurity = 1 # Security Setting Level
            
            workbook = excel_app.Workbooks.Open(os.path.abspath(self.file_path))

            # Run macro(s) and Sleep between macros
            if self.macros is not None:
                for macro in list([self.macros]): 
                    excel_app.Run(macro)
                    
                    # Refresh the workbook and Sleep between macros
                    if self.refresh_workbook: 
                        workbook.RefreshAll()
                    
                    if isinstance(self.sleep_time, int): 
                        time.sleep(self.sleep_time) 

                    print(f'{macro} ran successfully and slept for {self.sleep_time} seconds.')
            
            # Refresh and Save the workbook
            if self.refresh_workbook: 
                workbook.RefreshAll()
            
            if self.save_workbook: 
                workbook.Save()
            
            return True

        except Exception as e: 
            print(f'Error: {e}')
            return False
        
        finally: 
            excel_app.Quit()
