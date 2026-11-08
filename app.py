# Insert import os
import openpyxl
import pandas as pd
import smtplib
import mimetypes
from email.message import EmailMessage
from analyticskit import *

pd.options.display.float_format = '{:,.2f}'.format # Display option to include 2 decimal places on all floats
pd.set_option('display.max_columns', None) # Display option to show all columns
pd.set_option('display.max_rows', None) # Display option to show all rowsyour python code here


directory = select_file("Please enter the directory")
while True:
    try:
        directory = directory.split("/")
        
        del directory[-1]
        directory = "/".join(directory)
        
        directory = directory + "/"
        
        filenames = os.listdir(directory)
        list = ['ProductData.xlsx', 'MasterInvoice.pdf', 'Order1.xlsx', 'Order2.xlsx', 'Order3.xlsx', 'Order4.xlsx']
        
        
        if "ProductData.xlsx" in filenames and "MasterInvoice.pdf" in filenames and 'Order1.xlsx' in filenames and 'Order2.xlsx' in filenames and 'Order3.xlsx' in filenames and 'Order4.xlsx' in filenames:
            break
        else:
            show_message("Error","Please choose the correct file path")
            
    except:
        show_message("Error", "Choose the correct file path")


sheet = pd.read_excel(directory + "ProductData.xlsx", sheet_name="Descriptions")
sheet.head()

product_descriptions = {}
for r, row in sheet.iterrows():
    product_descriptions[row["Unit number"]]=row["Description"]

sheet = pd.read_excel(directory + "ProductData.xlsx", sheet_name="Prices")
product_prices = {}
for r, row in sheet.iterrows():
    product_prices[row["Unit number"]] = row["Price per unit"]

def get_order_data(invoice_num, filename, directory):
    df_invoice = get_pdf(directory + "MasterInvoice.pdf")
    order_data = {}
    for r, row in df_invoice.iterrows():
        order_data[row['name']] = row["value"]
    path = directory + filename
    
    order = openpyxl.load_workbook(path)
    orderform = order['OrderForm']

    order_data["InvoiceNum"] = str(invoice_num)
    order_data['Date'] = str(orderform["C9"].value)
    
   
    order_data['Customer Address1'] = str(orderform['C16'].value)
    order_data['CustomerName'] = str(orderform['C14'].value)
    order_data['CustomerCompany'] = str(orderform["C15"].value)

    city = str(orderform["C17"])
    state = str(orderform['18'])
    zipcode = str(orderform['19'])
    order_data['CustomerAddress2'] = str(city + state + zipcode)

    order_data['Description0'] = str(product_descriptions[orderform['B24'].value])
    order_data['Price0'] = str(product_prices[orderform['B24'].value])
    order_data['Quantity0'] = str(orderform['C24'].value)
    amount0 = product_prices[orderform['B24'].value] * orderform['C24'].value
    # amount0 = float(order_data['Price0']) * float(order_data['Quantity0'])
    order_data['Amount0'] = str(round(amount0, 2))

    if orderform['B25'].value in product_descriptions.keys():
        order_data['Description1'] = str(product_descriptions[orderform['B25'].value])
        order_data['Price1'] = str(product_prices[orderform['B25'].value])
        order_data['Quantity1'] = str(orderform['C25'].value)
        amount1 = float(order_data['Price1']) * float(order_data['Quantity1'])
        order_data['Amount1'] = str(amount1)

    
    
    if orderform['B26'].value in product_descriptions.keys():
        order_data['Description2'] = str(product_descriptions[orderform['B26'].value])
        order_data['Price2'] = str(product_prices[orderform['B26'].value])
        order_data['Quantity2'] = str(orderform['C26'].value)
        amount2 = float(order_data['Price2']) * float(order_data['Quantity2'])
        order_data['Amount2'] = str(amount2)

    
    if orderform['B27'].value in product_descriptions.keys():
        order_data['Description3'] = str(product_descriptions[orderform['B27'].value])
        order_data['Price3'] = str(product_prices[orderform['B27'].value])
        order_data['Quantity3'] = str(orderform['C27'].value)
        amount3 = float(order_data['Price3']) * float(order_data['Quantity3'])
        order_data['Amount3'] = str(amount3)
    
    if orderform['B28'].value in product_descriptions.keys():
        order_data['Description4'] = str(product_descriptions[orderform['B28'].value])
        order_data['Price4'] = str(product_prices[orderform['B28'].value])
        order_data['Quantity4'] = str(orderform['C28'].value)
        amount4 = float(order_data['Price4']) * float(order_data['Quantity4'])
        order_data['Amount4'] = str(amount4)

    
    
    
    order_data["Total"] = str(str(order_data['Amount0']) + str(order_data['Amount1']) + str(order_data['Amount2']) + str(order_data['Amount3']) + str(order_data['Amount4']) )

    name = order_data['CustomerName']
    email = 'cash.schmutz@gmail.com'
    date = order_data['Date']


    return order_data, name, email, date
def create_invoice(invoice_num, order_data, directory):
    df_pdf = get_pdf(directory + 'MasterInvoice.pdf')
    for index, row in df_pdf.iterrows():
        update_data = {}
        order_data
        update_data["InvoiceNum"] = order_data['InvoiceNum']
        update_data['Date'] = order_data['Date']
            
           
        update_data['Customer Address1'] = order_data['CustomerAddress1']
        update_data['CustomerName'] = order_data['CustomerName']
        update_data['CustomerCompany'] = order_data['CustomerCompany']
        
        # city = str(orderform["C17"])
        # state = str(orderform['18'])
        # zipcode = str(orderform['19'])
        update_data['CustomerAddress2'] = order_data['CustomerAddress2']
        
        update_data['Description0'] = order_data['Description0']
        update_data['Price0'] = order_data['Price0']
        update_data['Quantity0'] = order_data['Quantity0']
        # amount = float(update_data['Price0']) * float(update_data['Quantity0'])
        update_data['Amount0'] = order_data['Amount0']
        
        update_data['Description1'] = order_data['Description1']
        update_data['Price1'] = order_data['Price1']
        update_data['Quantity1'] = order_data['Quantity1']
        # amount = float(update_data['Price1']) * float(update_data['Quantity1'])
        update_data['Amount1'] = order_data['Amount1']
        
    create_pdf(directory + "MasterInvoice.pdf", f'Invoice {invoice_num}.pdf', update_data)


def send_email(invoice_num, name, email, date, directory):
    # Enter your gmail user name and password

    gmail_user = "cash.schmutz@mgmail.com"
    gmail_password_path = select_file('Please select the gmail password file')

    gmail_password = open(gmail_password_path).read()



    # List of recipients

    recipients = email

    # Email subect

    subject = f'Your Order from PythonTech Co. - Invoice {invoice_num}'

    # Body of the email

    body = f"""Dear {name},
        
    Attached is your invoice dated {date}.

    Thank you for your business,
        
    PythonTech Co.
    10 Python Drive
    Pleasant Grove, UT 84062
    Phone: (801)555-2372
    Email: sales@pythontecho.com"""

    # Initialize the email message

    message = EmailMessage()
    message['From'] = gmail_user
    message['To'] = recipients
    message['Subject'] = subject
    message.set_content(body)

    # Add attachment

    path_to_file1 = directory + 

    messsage = add_attachment(path_to_file1, message)
            
    # Log in to the Gmail server

    mail_server = smtplib.SMTP_SSL('smtp.gmail.com')
    mail_server.login(gmail_user, gmail_password)

    # Send the email

    mail_server.send_message(message)
    mail_server.quit()


filenames = os.listdir(directory)
invoice_num = 8515
for name in filenames:
    
    if "Order" in name and "xlsx" in name:
        
        order_data, name, email, date = get_order_data(invoice_num, name, directory)
        
        create_invoice(invoice_num, order_data, directory)

        send_email(invoice_num, name, email, date, directory)

        invoice_num = invoice_num + 1

