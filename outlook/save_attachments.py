# Import Modules
import os
import win32com.client as win32
from datetime import datetime, timedelta

from dataclasses import dataclass



PATH = 'C:\\Users\\Thabang Ndhlovu\\Desktop\\Work\\Projects\\startfish\\'
FILE_NAMES = open(f'{PATH}funds.txt').read().split('\n')
TODAY = datetime.today().strftime('%Y-%m-%d')
PREVIOUS_DAY = (datetime.today() - timedelta(days=1)).strftime('%Y-%m-%d')


def make_dir(*args) -> str:
    '''Create a directory if it does not exist'''
    
    location = os.path.join(*args)

    # Create folder if it does not exist
    if not os.path.exists(location):
        os.makedirs(location)
    return location

def messages_filter(inbox, **kwargs):
    '''
    Return filtered messages from Outlook inbox
    
    Parameters
    ----------
    inbox (win32com.client.CDispatch): Outlook inbox object
    kwargs (dict): 
        folder_name - Name of folder to filter messages by
        sender_email - Email address of sender to filter messages by
        from_date - From Date to filter messages by default is previous day
        to_date - To Date to filter messages by default is today
        desending_sort - Sort messages in descending order

    Returns
    -------
    messages (win32com.client.CDispatch): Outlook messages object
    '''
    folder_name = kwargs.get('folder_name')
    sender_email = kwargs.get('sender_email')
    from_date = kwargs.get('from_date')
    to_date = kwargs.get('to_date')
    desending_sort = kwargs.get('desending_sort')

    # filter messages by folder, sender, etc
    if folder_name is not None:
        inbox = inbox.Folders(folder_name)    
    messages = inbox.Items
    
    messages.Sort("[ReceivedTime]", desending_sort)

    if sender_email is not None:
        messages = messages.Restrict(f"[SenderEmailAddress] = {sender_email}")
    
    if from_date is not None:
        messages = messages.Restrict(f"[ReceivedTime] >= '{from_date}'")
    else:
        messages = messages.Restrict(f"[ReceivedTime] >= '{PREVIOUS_DAY}'")

    if to_date is not None:
        messages = messages.Restrict(f"[ReceivedTime] <= '{to_date}'")
    else:
        messages = messages.Restrict(f"[ReceivedTime] <= '{TODAY}'")
    
    return messages

def required_file(filename, file_names) -> bool:
    '''
    Check if filename is required
    
    Parameters
    ----------
    filename (str): Name of file to check
    file_names (list): List of file names to check against

    Returns
    -------
    bool: True if filename is required else False

    '''
    try:
        io_code = len(file_names[0]) # IO code refers to Maitland sheet Code
        sheet_code = [num for num in filename.split('_') if num.isdigit() and len(num) <= io_code]
        return True if sheet_code[0] in file_names else False
    except: return False 

def save_attachments(messages, file_names, *directory) -> list[str]:
    ''' 
    Save attachments from Outlook messages to specified folder.
    
    Parameters
    ----------
    messages (win32com.client.CDispatch): Outlook messages to save attachments from 
    directory (dict): Contains the path to the folder where the files will be saved

    Returns
    -------
    saved_files (list[str]): List of file paths to saved files
    '''    
    
    saved_files = []
    
    # Loop through messages and attachments then save files
    for message in messages:
        path = make_dir(
            *directory, message.CreationTime.strftime('%Y'), message.CreationTime.strftime('%B %Y')
            )            
        
        for attachment in message.Attachments:
            if required_file(str(attachment.FileName), file_names):
                filename_path = os.path.join(path, str(attachment.FileName))
                attachment.SaveAsFile(filename_path)
                saved_files.append(filename_path)

    print(f'{len(saved_files)}: Files saved to {path}')
    return saved_files


@dataclass
class SaveEmailAttachments:
    '''Save attachments from Outlook messages to specified folder
    
    Parameters
    ----------
    folder_name (str): Name of folder to filter messages by
    sender_email (str): Email address of sender to filter messages by
    from_date (str): From Date to filter messages by default is previous day
    to_date (str): To Date to filter messages by default is today
    desending_sort (bool): Sort messages in descending order
    file_names (list): List of file names to save
    path_extension (str): Name of folder to save files to
    '''
    
    folder_name: str = None
    sender_email: str = None
    from_date: str = None
    to_date: str = None
    desending_sort: bool = True
    path_extension: str = 'Reports'
    file_names = FILE_NAMES


    def run_function(self, return_saved_files: bool = True) -> list[str] | None:
        '''
        Run the SaveAttachments which saves attachments from Outlook messages to specified folder

        Parameters
        ----------
        return_saved_files (bool): Return saved files if True

        Returns
        -------
        saved_files (list[str]): List of file paths to saved files
        '''
        try:
            outlook_app = win32.Dispatch("Outlook.Application")
            namespace = outlook_app.GetNamespace("MAPI")
            inbox = namespace.GetDefaultFolder(6)
            full_path = os.path.join(PATH, self.path_extension)
            
            _filter_dict = {
                'folder_name': self.folder_name,
                'sender_email': self.sender_email, 
                'from_date': self.from_date,
                'to_date': self.to_date,
                'desending_sort': self.desending_sort
            }
            messages = messages_filter(inbox, **_filter_dict)
            saved_files = save_attachments(messages, self.file_names, full_path)
            if return_saved_files: return saved_files
        
        except Exception as e:
            print(f'Error: {e}')
        
        finally:
            outlook_app.Quit()
            