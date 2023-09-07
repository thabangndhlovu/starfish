# Import Modules
import time
import win32com.client as win32

from dataclasses import dataclass



@dataclass
class RunOutlookMacros:
    """
    Runs Macro(s) in Outlook
    
    Parameters
    ----------
    macro_names (str): Name of macro(s) to run
    sleep_time (int): Time to sleep between macros in seconds
    """
    
    macro_names: list[str] | str
    sleep_time: int = None


    def run_function(self) -> bool:
        '''
        Run the RunOutlookMacros which runs macro(s) in Outlook

        Returns
        -------
        bool: True if macro(s) ran successfully
        '''

        try:
            # Initialise Outlook and Run macro(s)
            outlook_app = win32.Dispatch("Outlook.Application")
            
            for macro in list([self.macro_names]):
                outlook_app.Run(macro)
                
                if isinstance(self.sleep_time, int): 
                    time.sleep(self.sleep_time)

            print(f'Macro(s) {self.macro_names} ran successfully!')
            return True

        except Exception as e:
            print(f'{e} occured for {self.macro_names}.')
            return False
        
        finally:
            outlook_app.Quit() # Close Outlook