while True:
    directory = single_input("Please enter the directory where the files are located.")
    if directory[-1] != '\\':
        directory = directory + "\\"
#print(directory)
        files_in_directory = os.listdir(directory)
#print(files_in_directory)
        if 'ProductData.xlsx' in files_in_directory and 'MasterInvoice.pdf' in files_in_directory and 'Order1.xlsx' in files_in_directory and 'Order2.xlsx' in files_in_directory and 'Order3.xlsx' in files_in_directory and 'Order4.xlsx' in files_in_directory:
            break
        else:
            show_message("Error!", "Please re-enter the directory.")# Insert your python code here


df = pd.read_excel(f'{directory}ProductData.xlsx',sheet_name="Descriptions")
df1 = pd.read_excel(f'{directory}ProductData.xlsx',sheet_name="Prices")

product_descriptions = {}
for index,row in df.iterrows():
    unit_number = row['Unit number']
    product_descriptions[unit_number] = row.Description


product_prices = {}
for index,row in df1.iterrows():
    unit_number = row['Unit number']
    product_prices[unit_number] = row['Price per unit']

print(product_descriptions)
print(product_prices)


def get_order_data(invoice_num, filename, directory):

    template = openpyxl.load_workbook(directory + filename, data_only=True)
    selected_sheet = template['OrderForm']

    name = selected_sheet['C14'].value
    email = selected_sheet['C13'].value
    address = selected_sheet['C16'].value
    company = selected_sheet['C15'].value
    city = selected_sheet['C17'].value
    state = selected_sheet['C18'].value
    zipcode = selected_sheet['C19'].value
    date = selected_sheet['C9'].value
    date = f'{date.month}/{date.day}/{date.year}'
    unit_number1 = selected_sheet['B24'].value
    unit_number2 = selected_sheet['B25'].value
    unit_number3 = selected_sheet['B26'].value
    Quantity1 = selected_sheet['C24'].value
    Quantity2 = selected_sheet['C25'].value
    Quantity3 = selected_sheet['C26'].value

    order_data = {}
    order_data['InvoiceNum'] = str(invoice_num)
    order_data['Date'] = date
    order_data['CustomerAddress1'] = str(address)
    order_data['CustomerName'] = str(name)
    order_data['CustomerCompany'] = str(company)
    order_data['CustomerAddress2'] = f'{city}, {state} {zipcode}'

    total = 0
    if pd.isnull(unit_number1) == False and pd.isnull(Quantity1) == False:
       price1 = float(product_prices[unit_number1])
       amount1 = price1 * Quantity1
       order_data['Description0'] = str(product_descriptions[unit_number1])
       order_data['Price0'] = f'${price1:,.2f}'
       order_data['Quantity0'] = str(Quantity1)
       order_data['Amount0'] = f'${amount1:,.2f}'
       total = total + amount1

    if pd.isnull(unit_number2) == False and pd.isnull(Quantity2) == False:
           price2 = float(product_prices[unit_number2])
           amount2 = price2 * Quantity2
           order_data['Description1'] = str(product_descriptions[unit_number2])
           order_data['Price1'] = f'${price2:,.2f}'
           order_data['Quantity1'] = str(Quantity2)
           order_data['Amount1'] = f'${amount2:,.2f}'
           total = total + amount2

    if pd.isnull(unit_number3) == False and pd.isnull(Quantity3) == False:
               price3 = float(product_prices[unit_number3])
               amount3 = price3 * Quantity3
               order_data['Description2'] = str(product_descriptions[unit_number3])
               order_data['Price2'] = f'${price3:,.2f}'
               order_data['Quantity2'] = str(Quantity3)
               order_data['Amount2'] = f'${amount3:,.2f}'
               total = total + amount3
    order_data['Total'] = f'${total:,.2f}'
    
    return order_data, name, email, date

invoice_num = 8515
filename = 'Order1.xlsx'
order_data, name, email, date = get_order_data(invoice_num, filename, directory)
print(name)
print(email)
print(date)
print(order_data)


def create_invoice(invoice_num, order_data, directory):
    create_pdf(directory+'MasterInvoice.pdf', directory+f'Invoice {invoice_num}.pdf', order_data)

create_invoice(8515, order_data, directory)


def send_email(invoice_num, name, email, date, directory):
    gmail_user = "peng.hailey@gmail.com"
    gmail_password_path = select_file('Please select the gmail password file')
    gmail_password = open(gmail_password_path).read()
    recipients = email
    subject = f"Your Order from PythonTech Co. - Invoice {invoice_num}"

    body = f"""Dear {name},
    
Attached is your invoice dated {date}.

Thank you for your business.
        
PythonTech Co.
10 Python Drive
Pleasant Grove, UT 84062
Phone: (801) 555-2372
Email: sales@pythontechco.com"""


    message = EmailMessage()  
    message['From'] = gmail_user   
    message['To'] = recipients
    message['Subject'] = subject
    message.set_content(body)

    path_to_file = directory+f'Invoice {invoice_num}.pdf'

    message = add_attachment(path_to_file, message)
            
    mail_server = smtplib.SMTP_SSL('smtp.gmail.com')
    mail_server.login(gmail_user, gmail_password)



    mail_server.send_message(message)
    mail_server.quit()

send_email(8515, name, email, date, directory)


filenames = os.listdir(directory)
#filenames
invoice_num = 8515
for selected_filename in filenames:
    if 'Order' in selected_filename and 'xlsx' in selected_filename:
        order_data, name, email, date = get_order_data (invoice_num, selected_filename, directory)
        create_invoice (invoice_num, order_data, directory)
        send_email(invoice_num, name, email, date, directory)
        invoice_num = invoice_num + 1
        