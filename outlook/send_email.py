from datetime import datetime
import win32com.client as win32
from dataclasses import dataclass


PATH = 'C:\\Users\\Thabang Ndhlovu\\Desktop\\Work\\Projects\\startfish\\'
DISTRIBUTION = open(f'{PATH}email_distribution.txt').read().split('\n')
TODAY = datetime.today().strftime("%d %B %Y")

@dataclass
class SendEmails:
        '''
        Send emails using Outlook
    
        Parameters
        ----------S
        recipients (str): Recipients of the email default is the email distribution list
        subject (str): Subject of the email
        body (str): Body of the email
        attachment (str): Path to attachment
        '''
    
        recipients: str = None
        subject: str = ''
        body: str = ''
        attachment: str = None

        def __post_init__(self) -> None:
            self.recipients = '; '.join(DISTRIBUTION) if self.recipients is None else self.recipients
            

        def run_function(self) -> bool:
            
            '''
            Run the SendEmails which sends emails using Outlook
    
            Returns
            -------
            bool: True if email sent successfully
            '''
            try:
                # Initialize Outlook
                outlook_app = win32.Dispatch("Outlook.Application")
                mail = outlook_app.CreateItem(0)  # 0 represents an email item
                
                # Send email
                mail.To = self.recipients
                mail.Subject = f'{self.subject} {TODAY}'
                mail.Body = self.body
                if self.attachment is not None: 
                    mail.Attachments.Add(self.attachment)
                mail.Send()

                print(f'Email sent to {self.recipients} successfully!')
                return True
            
            except Exception as e: 
                print(f'{e} occured for {self.recipients}.')
                return False

            finally:
                outlook_app.Quit() # Close Outlook
