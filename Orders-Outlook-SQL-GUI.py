import sys
import csv
import pyodbc
import os.path
import re
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QTableWidget, QTableWidgetItem, QCheckBox, QPushButton
import subprocess
import tempfile
import os
import shutil
from pathlib import Path
import datetime
import math
import win32com.client  #pip install pywin32
import time

start_time = time.time()
# Define the connection parameters
server_name = 'Your-SQLserver-name'
database_name = 'SQL-Database-name'
username = 'sql-login'
password = 'sql-password'

# Create a connection to your SQL Server database
connection = pyodbc.connect(f'DRIVER=SQL Server;'
                      f'SERVER={server_name};'
                      f'DATABASE={database_name};'
                      f'UID={username};'
                      f'PWD={password}')

cursor = connection.cursor()

# Create output folder
output_dir = Path.cwd() / "Output"
print("Cleared folder Path")
shutil.rmtree(output_dir)
output_dir.mkdir(parents=True, exist_ok=True)

# Connect to outlook
outlook = win32com.client.Dispatch("Outlook.Application")
namespace = outlook.GetNamespace("MAPI")
# Connect to folder

Sample = namespace.Folders['Sample Prints']
Sample_inbox = Sample.Folders['Inbox']
subfolder_names = ["First Type Samples", "Second Type Samples", "Third Type Sampless", "Fourth Type Samples"]
for subfolder_name in subfolder_names:
    subfolder = Sample_inbox.Folders(subfolder_name)
    messages = subfolder.Items
    # Get messages
    for message in messages:
        subject = message.Subject
        body = message.body
        attachments = message.Attachments

        # https://docs.microsoft.com/en-us/office/vba/api/outlook.oldefaultfolders
        # DeletedItems=3, Outbox=4, SentMail=5, Inbox=6, Drafts=16, FolderJunk=23

        # Create separate folder for each message, exclude special characters and timestampe
        unique_identifier = datetime.datetime.now().strftime("%Y%m%d%H%M%S%f")  # Using timestamp
        target_folder = output_dir / re.sub('[^0-9a-zA-Z]+', '', subject) / unique_identifier
        target_folder.mkdir(parents=True, exist_ok=True)

        # Write body to text file
        Path(target_folder / "EMAIL_BODY.txt").write_text(str(subject)+str(body))

        # Save attachments and exclude special
        for attachment in attachments:
            filename = re.sub('[^0-9a-zA-Z\.]+', '', attachment.FileName)
            attachment.SaveAsFile(target_folder / filename)
Folders_paths = []
output_dir = "Output"
for root, dirs, files in os.walk(output_dir):
    for dir_name in dirs:
        folder_path = os.path.join(root, dir_name)
        Folders_paths.append(folder_path)
filtered_paths = [path for path in Folders_paths if path.count('\\') > 1]
print("Saved new folder tree")
for every_emailbody_txt in filtered_paths:

    Item_position = 0

    with open(every_emailbody_txt + "\EMAIL_BODY.txt", 'r') as in_file:
        stripped = [line.strip() for line in in_file.readlines()]
        lines = [line.split(",") for line in stripped if line]
        subject_of_email = lines[0][0]

        if "First Type Samples" in subject_of_email:
            Part_Reference_line = 0
            If_Note_In = False
            # Define name of outlook subfolder
            outlook_sub_folder = "First Type Samples"
            # Define Order ID
            Order_ID_line = subject_of_email
            Order_ID_Search = re.search("\(\d*\)", Order_ID_line)
            Order_ID_With_Brackets = Order_ID_Search.group(0)
            Order_ID = int(Order_ID_With_Brackets[1:-1])
            for line in lines:
                if "Note" in line[0]:
                    If_Note_In = True
                if line[0] == 'Product\tQuantity\tPrice' or line[0] == 'Product\t Quantity\t Price':
                    break
                Part_Reference_line += 1

            # Define how many lines of order
            lines_in_email = len(lines)
            If_Item_in_same_line_with_price = False
            if "£" in lines[Part_Reference_line + 1][0]:
                If_Item_in_same_line_with_price = True
            else:
                If_Item_in_same_line_with_price = False

            if If_Item_in_same_line_with_price == True:
                Order_Lines = (lines_in_email - (Part_Reference_line + 4))
            elif If_Item_in_same_line_with_price == False:
                Order_Lines = int((lines_in_email - (Part_Reference_line + 4)) / 2)


            # Define Customer Name
            Customer = lines[2][0].replace("'","")

            # Define Addressline1
            Addressline1 = lines[3][0].replace("'","")
            # Define City
            City = lines[Part_Reference_line - 3][0].replace("'","")


            # Define Postcode
            Postcode = lines[Part_Reference_line - 2][0]
            if If_Note_In == True:
                City = lines[Part_Reference_line - 4][0]
                Postcode = lines[Part_Reference_line - 3][0]

            # Define Item name and Quantity
            Item_position = Part_Reference_line + 1


            for every_line in range(Order_Lines):
                if If_Item_in_same_line_with_price == True:
                    Item = lines[Item_position][0].split(" ")
                    Quantity = Item[-2]
                    Item = Item[:-2]
                    Item = ' '.join(Item)

                else:
                    Item = lines[Item_position][0]
                    Item_position += 1
                    Quantity = lines[Item_position][0][0]
                    Item_position += 1
                try:
                    cursor.execute(
                        # It is missing $Item
                        f"INSERT INTO  [Production].[Samp].[SampleOrders]      "
                        f"     (Order_ID, Line, Item, Quantity, Customer, Complete, Addressline1, City, PostCode, Name) "
                        f" VALUES    (  {Order_ID} , {every_line + 1}, '{Item}', {Quantity}, '{Customer}', 0, '{Addressline1}', '{City}', '{Postcode}', '{outlook_sub_folder}' )"

                    )
                    connection.commit()
                except pyodbc.IntegrityError as e:
                    # Handle the IntegrityError (duplicate key error) here
                    print(f"Error: {e}")
        elif "Second Type Samples" in subject_of_email:
            outlook_sub_folder = "Second Type Samples"

            # Define Order ID from subject
            Order_ID_line = lines[1][0]
            Order_ID_line_table = Order_ID_line.split(" ")
            Order_ID = Order_ID_line_table[2]

            lines_in_email = len(lines)



            #Finding a line where order starts
            Part_Reference_line = 0
            for whatever_value in lines:
                if whatever_value[0] == "Part Reference\tDescription\tQuantity":
                    Item_position = Part_Reference_line + 1
                    break

                Part_Reference_line += 1

            # Define how many lines of order
            Order_Lines = lines_in_email - Item_position

            # Define Customer Name
            Customer = lines[2][0]

            # Define Addressline1
            Addressline1 = lines[4][0]
            # Define City

            City = lines[Part_Reference_line - 2][0]

            # Define Postcode
            Postcode = lines[Part_Reference_line - 1][0]

            #Loop as many times as there is orders
            for every_line in range(Order_Lines):

                Part_Reference = lines[Item_position][0]
                Part_Reference = Part_Reference.replace('\t', ' ')
                Item = Part_Reference.split(" ")
                Part_Reference = Item[0]
                Quantity = Item[-1]
                Item = Item[1:-1]
                Item = ' '.join(Item)

                try:
                    cursor.execute(
                        # It is missing $Item
                        f"INSERT INTO  [Production].[Samp].[SampleOrders]      "
                        f"     (Order_ID, Line, Item, Quantity, Customer, Complete, Addressline1, City, PostCode, Name, Part_Reference) "
                        f" VALUES    (  {Order_ID} , {every_line + 1}, '{Item}', {Quantity}, '{Customer}', 0, '{Addressline1}', '{City}', '{Postcode}', '{outlook_sub_folder}', '{Part_Reference}' )"

                    )
                    connection.commit()
                except pyodbc.IntegrityError as e:
                    # Handle the IntegrityError (duplicate key error) here
                    print(f"Error: {e}")
        elif "Third Type Sampless" in subject_of_email:
            outlook_sub_folder = "Third Type Sampless"
            # Finding a line where order starts
            Part_Reference_line = 0
            for whatever_value in lines:
                if "https://www."  in whatever_value[0] or "https://protect" in whatever_value[0]:
                    Date_line = Part_Reference_line
                if whatever_value[0] == "Product\t Quantity\t Price" or whatever_value[0] == 'Product\tQuantity\tPrice':
                    Item_position = Part_Reference_line + 1
                    break
                Part_Reference_line += 1

            # Define Order ID
            Order_ID_line = lines[0][0]
            Order_ID_Search = re.search("\(\d*\)", Order_ID_line)
            Order_ID_With_Brackets = Order_ID_Search.group(0)
            Order_ID = int(Order_ID_With_Brackets[1:-1])
            # Define how many lines of order
            lines_in_email = len(lines)
            Order_Lines = int((lines_in_email - (Item_position + 3)) / 2)

            # Define Customer Name
            Customer = lines[2][0]

            # Define Addressline1
            Addressline1 = lines[3][0]

            # Use regular expressions to remove special characters
            Addressline1 = re.sub(r'[^a-zA-Z0-9\s]', '', Addressline1)

            # Define City
            City = ' '.join(lines[Date_line - 2])
            City = re.sub(r'[^a-zA-Z0-9\s]', '', City)

            # Define Postcode
            Postcode = lines[Date_line - 1][0]

            for every_line in range(Order_Lines):
                Item_line = lines[Item_position][0]
                Item_line = Item_line.split('(')
                Item = Item_line[0]
                Part_Reference = Item_line[1][1:-1]
                Item_position += 1
                Quantity = lines[Item_position][0][0]
                Item_position += 1

                try:
                    cursor.execute(
                        # It is missing $Item
                        f"INSERT INTO  [Production].[Samp].[SampleOrders]      "
                        f"     (Order_ID, Line, Item, Quantity, Customer, Complete, Addressline1, City, PostCode, Name, Part_Reference) "
                        f" VALUES    (  {Order_ID} , {every_line + 1}, '{Item}', {Quantity}, '{Customer}', 0, '{Addressline1}', '{City}', '{Postcode}', '{outlook_sub_folder}', '{Part_Reference}' )"

                    )
                    connection.commit()
                except pyodbc.IntegrityError as e:
                    # Handle the IntegrityError (duplicate key error) here
                    print(f"Error: {e}")

        elif "Fourth Type Sampless" in subject_of_email:
            outlook_sub_folder = "Fourth Type Sampless"
            Part_Reference_line = 0
            for whatever_value in lines:
                if whatever_value[0] == "Qty \tSample \tTotal":
                    #Item_position is going to point to first line of a item but it will be Quantity
                    Item_position = Part_Reference_line + 1

                    break
                Part_Reference_line += 1

            # Define Order ID
            for any_number in range(lines_in_email):
                if "Order" in lines[Part_Reference_line - 3 - any_number][0]:
                    Order_ID_line = lines[Part_Reference_line - 3 - any_number][0]
                    Order_ID_line_table = Order_ID_line.split(" ")
                    Order_ID = Order_ID_line_table[2][1:]
                    break

            # Define how many lines of order
            lines_in_email = len(lines)
            Order_Lines = int(((lines_in_email - 9) - Item_position) / 2)

            # Define Customer Name
            Customer = lines[1][0].replace("'","")

            # Define Addressline1
            Addressline1 = lines[3][0].replace("'","")

            # Define City
            City = lines[Part_Reference_line - 6][0].replace("'","")

            # Define Postcode
            Postcode = lines[Part_Reference_line - 5][0]


            # Loop as many times as there is orders
            for every_line in range(Order_Lines):

                Quantity = lines[Item_position][0][0]
                Item = lines[Item_position][0][3:]
                Item_position += 1
                Part_Reference = lines[Item_position][0][:-6]
                Item_position += 1

                try:
                    cursor.execute(
                        # It is missing $Item
                        f"INSERT INTO  [Production].[Samp].[SampleOrders]      "
                        f"     (Order_ID, Line, Item, Quantity, Customer, Complete, Addressline1, City, PostCode, Name, Part_Reference) "
                        f" VALUES    (  {Order_ID} , {every_line + 1}, '{Item}', {Quantity}, '{Customer}', 0, '{Addressline1}', '{City}', '{Postcode}', '{outlook_sub_folder}', '{Part_Reference}' )"

                    )
                    connection.commit()
                except pyodbc.IntegrityError as e:
                    # Handle the IntegrityError (duplicate key error) here
                    print(f"Error: {e}")
    print("Data Saved to SQL and locally to folder tree")

    '''











                '''
    '''
        Order_number = re.search(r'\((.*?)\)', lines[3][0]).group(1)
        with open('log.csv','w') as out_file:
            writer = csv.writer(out_file)
            writer.writerow(('title','intro'))

    '''


def fetch_data():
    cursor.execute("SELECT [SampI], [Order_ID], [Line], [Item], [Quantity], [Customer], [Complete], [Addressline1], [City], [PostCode], [companyName], [Name] ,[Part_Reference] FROM [Production].[Samp].[SampleOrders] WHERE [Complete] = 0 ")
    return cursor.fetchall()

class DatabaseViewer(QWidget):
    def __init__(self):
        super().__init__()
        self.show_completed = False  # Initialize to show non-completed items
        self.init_ui()

    @staticmethod
    def fetch_completed_data():
        cursor.execute(
            "SELECT [SampI], [Order_ID], [Line], [Item], [Quantity], [Customer], [Complete], [Addressline1], [City], [PostCode], [companyName], [Name] ,[Part_Reference]  FROM [Production].[Samp].[SampleOrders] WHERE [Complete] = 1")
        return cursor.fetchall()

    def on_print_clicked(self):
        print_button = self.sender()
        row = self.table_widget.indexAt(print_button.pos()).row()
        Order_ID = self.table_widget.item(row, 1).text()  # Assuming Order_ID is in column 1
        Folder_name = self.table_widget.item(row, 11).text()  # Assuming Folder_name is in column 11

        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")

        # Access the desired mailbox folder (e.g., Sample Prints)
        sample_prints = namespace.Folders['Sample Prints']
        sample_inbox = sample_prints.Folders['Inbox']
        #Check if email in Inbox
        in_the_inbox = False
        # Iterate through the subfolders within the Sample Prints Inbox
        for folder in sample_inbox.Folders:
            # Check if the current folder's name matches the desired folder name
            if folder.name == Folder_name:
                # Access the email messages within the current folder
                messages = folder.Items
                # Iterate through the email messages
                for message in messages:
                    # Check if the order number is present in the email body
                    if Order_ID in message.body:
                        # Check if from Style Studio
                        '''
                        # Orders stopped coming with attached order to it, email is now an order itself
                        
                        if Folder_name == "Style Studio Samples":
                            # Can be changed in browser setting to print every PDF opened in the browser

                            # Print attachment instead of the actual email
                            for attachment in message.Attachments:
                                # Create a temporary directory to save the attachment
                                temp_dir = tempfile.mkdtemp()
                                attachment_path = os.path.join(temp_dir, attachment.FileName)

                                # Save the attachment to the temporary directory
                                attachment.SaveAsFile(attachment_path)

                                print(attachment_path)
                                print(temp_dir)

                                # Use subprocess to open and print the saved attachment with Microsoft Edge
                                subprocess.run(["C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe",
                                                attachment_path, "--print"])

                                print(f"Attachment '{attachment.FileName}' printed.")

                                # Clean up the temporary directory and its contents
                                shutil.rmtree(temp_dir)

                    else:'''
                        # Print the current email message
                        message.PrintOut()
                        print(f"Email with order '{Order_ID}' printed.")
                        in_the_inbox = True
                        break  # Exit the loop after finding the desired folder
        if in_the_inbox == False:
            sample_inbox = sample_prints.Folders['Archive']
            for folder in sample_inbox.Folders:
                # Iterate through the subfolders within the Sample Prints Archive
                if folder.name == Folder_name:
                    # Access the email messages within the current folder
                    messages = folder.Items
                    # Iterate through the email messages
                    for message in messages:
                        # Check if the order number is present in the email body
                        if Order_ID in message.body:
                            # Print the current email message
                            message.PrintOut()
                            print(f"Email with order '{Order_ID}' printed.")
                    break  # Exit the loop after finding the desired folder

    # Modify refresh_table_data
    def refresh_table_data(self):
        if self.show_completed:
            data = self.fetch_completed_data()  # Use fetch_completed_data() as a static method
        else:
            data = fetch_data()
        self.table_widget.setRowCount(0)  # Clear the existing rows

        for row_num, row_data in enumerate(data):
            self.table_widget.insertRow(row_num)
            for col_num, col_data in enumerate(row_data):

                if col_num == 6:
                    item = QTableWidgetItem()
                    item.setFlags(item.flags() | 32)  # Make the "Check" column editable
                    self.table_widget.setItem(row_num, col_num, item)
                    check = QCheckBox()
                    check.stateChanged.connect(self.on_check_changed)
                    self.table_widget.setCellWidget(row_num, 6, check)
                elif col_num == 10:
                    button = QPushButton("Print")
                    button.setText("Print")
                    button.clicked.connect(self.on_print_clicked)
                    self.table_widget.setCellWidget(row_num,col_num, button)


                else:
                    item = QTableWidgetItem(str(col_data))
                    item.setFlags(item.flags() & ~32)  # Make other columns non-editable
                    self.table_widget.setItem(row_num, col_num, item)


    def init_ui(self):
        self.setWindowTitle('BEST IT PROGRAM')
        self.setGeometry(100, 100, 1300, 400)

        layout = QVBoxLayout()

        self.table_widget = QTableWidget()
        self.table_widget.setColumnCount(13)
        self.table_widget.setHorizontalHeaderLabels(
            ["SampI", "Order_ID", "Line", "Item", "Quantity", "Customer", "Complete", "Addressline1", "City", "PostCode", "Reprint","Sub-Folder", "Part_Reference"])

        data = fetch_data()

        for row_num, row_data in enumerate(data):
            # Populate the table as before
            self.table_widget.insertRow(row_num)
            for col_num, col_data in enumerate(row_data):
                if col_num == 6:
                    item = QTableWidgetItem()
                    item.setFlags(item.flags() | 32)  # Make the "Check" column editable
                    self.table_widget.setItem(row_num, col_num, item)
                    check = QCheckBox()
                    check.stateChanged.connect(self.on_check_changed)
                    self.table_widget.setCellWidget(row_num, 6, check)
                elif col_num == 10:
                    button = QPushButton("Print")
                    button.setText("Print")
                    button.clicked.connect(self.on_print_clicked)
                    self.table_widget.setCellWidget(row_num,col_num, button)
                else:
                    item = QTableWidgetItem(str(col_data))
                    item.setFlags(item.flags() & ~32)  # Make other columns non-editable
                    self.table_widget.setItem(row_num, col_num, item)


        layout.addWidget(self.table_widget)

        # Create the "Toggle Show Completed" button and connect it
        self.toggle_button = QPushButton("Toggle Show Completed")
        self.toggle_button.clicked.connect(self.toggle_show_completed)
        layout.addWidget(self.toggle_button)

        self.setLayout(layout)

    def on_check_changed(self, state):
        checkbox = self.sender()
        row = self.table_widget.indexAt(checkbox.pos()).row()
        SampI = self.table_widget.item(row, 0).text()  # Assuming SampI is in column 1

        if state == 2:  # 2 corresponds to a checked checkbox
            if self.show_completed == False:
                # Set the value of the "Complete" column to 1 in the database here
                cursor.execute("UPDATE [Production].[Samp].[SampleOrders] SET [Complete] = 1 WHERE [SampI] = ?",
                               (SampI,))
                connection.commit()
                self.table_widget.removeRow(row)
            elif self.show_completed == True:
                # Set the value of the "Complete" column to 1 in the database here
                cursor.execute("UPDATE [Production].[Samp].[SampleOrders] SET [Complete] = 0 WHERE [SampI] = ?",
                               (SampI,))
                connection.commit()
                self.table_widget.removeRow(row)

    def toggle_show_completed(self):
        if self.show_completed:
            self.show_completed = False
            self.toggle_button.setText("Show Completed")
        else:
            self.show_completed = True
            self.toggle_button.setText("Show Non-Completed")

        self.refresh_table_data()


if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = DatabaseViewer()
    window.show()
    sys.exit(app.exec_())

    # Close the database connection
cursor.close()
connection.close()

end_time = time.time()
elapsed_time = end_time - start_time
print(f"Script execution time: {elapsed_time} seconds")
