import datetime, xlrd, sys, os
import win32com.client as win32

data_list = []
# External files. They must have the same name as given in the code.
file = 'Resources/Clients.xlsx'
file_email = 'Resources/email_body.txt'

# File information. Extract from the file first row and the number of rows.
xl_workbook = xlrd.open_workbook(file)
xl_sheet = xl_workbook.sheet_by_name('Sheet1')
xl_rows = xl_sheet.nrows

# Attachment for email
path = os.getcwd()
attachment = '\\Resources\VELFAC Direct - Reporting Damage Policy.pdf'
full_path = '{}{}'.format(path, attachment)

# Function for clearing the data from unnecessary informations.
start = 'Do you want to start creating emails? There will be ' + str(xl_rows) + ' emails.'

# Email data. Preparing the email body from .txt file
topic = "Questions regarding the delivery of your VELFAC ORDER ({})"
body = open(file_email, 'r', encoding='UTF-8')
contents = body.read()

# Dates needed for the email content.
today = datetime.date.today()
next_monday = today + datetime.timedelta(days=-today.weekday(), weeks=2)
prior_friday = next_monday - datetime.timedelta(days=3)


class DataPreparation:
    def __init__(self, max_range, storage):
        self.max_range = max_range
        self.storage = storage

    def datamixer(self):
        for i in range(1, self.max_range):
            xl_row = xl_sheet.row_values(i)
            del xl_row[1]
            del xl_row[6:15]
            self.storage.append(xl_row)


# Emailer class with creation method.
class Emailer:
    def __init__(self, recipient, subject, body,):
        self.recipient = recipient
        self.subject = subject
        self.body = body
        self.attachment = full_path

    # Creates an email within outlook. Three parameters mus be given.
    # Subject, recipient and body.

    def create(self):
        print(full_path)
        print(path)
        outlook = win32.Dispatch('outlook.application')
        email = outlook.CreateItem(0)
        email.To = self.recipient
        email.Subject = self.subject
        email.htmlBody = self.body
        email.Attachments.Add(Source=self.attachment)
        email.Display(True)


class EmailBuilder:
    def __init__(self, range):
        self.range = range

    def create(self):
        for i in range(1, xl_rows):
            address = data_list[0][5]
            topic_formated = topic.format(int(data_list[0][0]))
            # Email creation with data extracted from the file. Emailer class was used here.
            email = Emailer(address, topic_formated, contents.format(int(data_list[0][0]),
                                                                     next_monday,
                                                                     data_list[0][1],
                                                                     data_list[0][2],
                                                                     data_list[0][4],
                                                                     data_list[0][3],
                                                                     prior_friday))
            email.create()

