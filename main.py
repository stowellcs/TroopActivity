import pandas as pd
import re
import datetime
import logging
import numpy as np


# Creating a function which will remove extra leading
# and tailing whitespace from the data.
# pass dataframe as a parameter here
def whitespace_remover(dataframe):
    # iterating over the columns
    for i in dataframe.columns:

        # checking datatype of each column
        if dataframe[i].dtype == 'object':

            # applying strip function on column
            dataframe[i] = dataframe[i].map(str.strip)
        else:

            # if condn. is False then it will do nothing.
            pass


# Creating a function to parse the name field
#
#
def nameparser(namefield):
    splitted = namefield.str.split()
    global FNAME_CUST
    df['FNAME_CUST'] = splitted.str[0]
    global LNAME_CUST
    df['LNAME_CUST'] = splitted.str[-1]
    global MNAME_CUST
    middle = splitted.str[1]
    df['MNAME_CUST'] = middle.mask(middle == df['LNAME_CUST'], '')

# Define function to split name into first, middle, and last names
def split_name(s):
    name_parts = s.split()

    if len(name_parts) == 2:
        first_name, last_name = name_parts
        middle_name = None
    elif len(name_parts) == 3:
        first_name, middle_name, last_name = name_parts
    else:
        first_name = name_parts[0]
        middle_name = ' '.join(name_parts[1:-1])
        last_name = name_parts[-1]

    return pd.Series({'First Name': first_name, 'Middle Name': middle_name, 'Last Name': last_name})

# Define function to extract Scout Name
def extract_scout_name(s):
    if 'Scout Name' in s:
        return s.split(',')[0].split(': ')[1]
    else:
        return None

# Define function to extract Scout Rank
def extract_scout_rank(s):
    return s.split(',')[1].strip()

# Define function to extract Scout Patrol
def extract_scout_patrol(s):
    if 'Scout Name' in s:
        return s.split(',')[2].strip()
    else:
        return 'Rocking Chair'

# Define function to extract Emergency Contact Phone Number
# def extract_emergency_contact_phone_number(s):
#     if 'Emergency Contact Phone Number' in s:
#         return s.split(',')[3].split(': ')[1]
#     else:
#         return None

def extract_attendee_name(s):
    scout_name = re.search(r'Scout Name:\s*(\w+\s+\w+)', s)
    scouter_name = re.search(r'Scouter Name:\s*(\w+\s+\w+)', s)
    if scout_name:
        return scout_name.group(1)
    elif scouter_name:
        return scouter_name.group(1)
    else:
        return None

def extract_emergency_contact_and_phone_number(s):
    phone_number_string = re.search(r'Emergency Contact Phone Number:\s*(\d{10})', s)
    contact_name_string = re.search(r'Emergency Contact:\s*([\w\s]+),?', s)
    if phone_number_string and contact_name_string:
        phone_number_start_index = s.find(phone_number_string.group(1))
        contact_name_end_index = s.find(contact_name_string.group(1)) + len(contact_name_string.group(1))
        contact_name = s[contact_name_end_index:].strip()
        phone_number = format_phone_number(s[phone_number_start_index:phone_number_start_index+10])
        return contact_name, phone_number
    else:
        return None, None

def extract_emergency_contact(s):
    contact_name_string = re.search(r'Emergency Contact:\s*([\w\s]+),?', s)
    contact_name_end_index = s.find(contact_name_string.group(1)) + len(contact_name_string.group(1))
    contact_name = s[contact_name_end_index:].strip()
    return contact_name


def extract_emergency_contact_phone_number(s):
    phone_number_string = re.search(r'Emergency Contact Phone Number:\s*(\d{10})', s)
    phone_number_start_index = s.find(phone_number_string.group(1))
    phone_number = format_phone_number(s[phone_number_start_index:phone_number_start_index+10])
    return phone_number


# Define function to extract Emergency Contact
# def extract_emergency_contact(s):
#     contact_name_string = re.search(r'Emergency Contact:\s*([\w\s]+),?', s)
#     contact_string = re.search(r'Emergency Contact:\s*', s)
#     if 'Emergency Contact' in s:
#         return s.split(',')[4].split(': ')[1]
#     else:
#         return None

# Define function to extract Driving Permission
def extract_driving_permission(s):
    if 'Yes - driving' in s:
        return 'Yes-Driving'
    elif 'Yes - carpool' in s:
        return 'Yes-Carpool'
    else:
        return 'No'

# Define function to extract Scouter Name
def extract_scouter_name(s):
    if 'Scouter Name' in s:
        return s.split(',')[0].split(': ')[1]
    else:
        return None

# Define function to extract Cell Phone Number
def extract_cell_phone_number(s):
    if 'Cell phone number' in s:
        cell = s.split(',')[1].split(': ')[1]
        return cell
    else:
        return None

# Creating a function to format the phone field
#
#
def phoneparser(phonefield):
    prefix = phonefield.str[0:3]
    exchange = phonefield.str[3:6]
    basenum = phonefield.str[6:10]
    phonenum = prefix + "-" + exchange + "-" + basenum

    logging.info(phonefield)
    logging.info(prefix)
    logging.info(exchange)

    return phonenum

def format_phone_number(s):
    phone_number = re.findall(r'\d+', s)
    if len(phone_number) == 10:
        formatted_number = "({}) {}-{}".format(phone_number[0], phone_number[1], phone_number[2]+phone_number[3]+phone_number[4])
        return formatted_number
    else:
        return None


# #####################################################################################
# USER MODIFY: Input files defined
# #####################################################################################
myDirectory = 'C:/Users/chris/Downloads/'
myInputFile = f'{myDirectory}items-2023-04-20-2023-05-31.csv'

# #####################################################################################
# USER MODIFY: Registration Label to filter on
# #####################################################################################
filterChoice1 = ['Scout Registration - Ropes Course Campout - May 2023']
filterChoice2 = ['Scouter Registration - Ropes Course Campout - May 2023']
eventName = 'May 2023 Ropes Course and Recruiting'
# filterChoice1 = ['Scout Registration - Camp Alexander (initial deposit)']
# filterChoice2 = ['Scouter Registration - Camp Alexander (initial deposit)']
# eventName = '2023 Camp A'

# #####################################################################################
# File Names - Various Steps
# #####################################################################################
mySortedFile = f'{myInputFile}_sorted.csv'
filteredFile1 = f'{myDirectory}{filterChoice1}.csv'
filteredFile2 = f'{myDirectory}{filterChoice2}.csv'
current_date_time = format(datetime.datetime.now().strftime("%Y-%m-%d-%H.%M"))

# #####################################################################################
# Logging --**** more research needed...explore config file *****
# #####################################################################################
# logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(name)s', filename=f'{myDirectory}App_Log.txt')
logging.basicConfig(level=logging.INFO, filename=f'{myDirectory}App.Log', format='%(asctime)s - %(message)s')
logging.error("Starting process")

# #####################################################################################
# Read Input File  // Clean the column names // Parse Name
# #####################################################################################
df = pd.read_csv(myInputFile, header=0)
df = df.dropna(axis=1, how='all')
df.columns = [re.sub("[ ,-]", "_", re.sub("[.,`]", "", c)) for c in df.columns]
# nameparser(df['Customer_Name'])

# Split customer names into separate columns for sorting
df[['FirstName_Cust', 'MiddleName_Cust', 'LastName_Cust']] = df['Customer_Name'].apply(split_name)


# #####################################################################################
# Sort by sales item and last name of customer// Create Sorted File
# #####################################################################################
df.sort_values(by=['Item', 'LastName_Cust'], inplace=True, key=lambda col: col.str.lower())

df['Attendee'] = df['Modifiers_Applied'].apply(lambda x: extract_attendee_name(x))
df['Emer_Contact'], df['Emer_Contact_Ph'] = zip(*df['Modifiers_Applied'].apply(lambda x: extract_emergency_contact_and_phone_number(x)))
# df['Emer_Contact'] = df['Modifiers_Applied'].apply(lambda x: extract_emergency_contact(x))
# df['Emer_Contact_Ph'] = (df['Modifiers_Applied'].apply(lambda x: extract_emergency_contact_phone_number(x)))
print(df['Attendee'])
print(df['Emer_Contact'])
print(df['Emer_Contact_Ph'])


df.to_csv(mySortedFile)

df = df[df['Item'].isin(filterChoice1)]


# Add new columns to dataframe using the functions
# df['Attendee'] = df['Modifiers_Applied'].apply(extract_scout_name)
df['Rank'] = df['Modifiers_Applied'].apply(extract_scout_rank)
df['Patrol'] = df['Modifiers_Applied'].apply(extract_scout_patrol)
# df['Emer_Contact_Ph'] = df['Modifiers_Applied'].apply(lambda x: extract_emergency_contact_phone_number(x))
# df['Emer_Contact'] = df['Modifiers_Applied'].apply(extract_emergency_contact)
df['Driving_Status'] = df['Modifiers_Applied'].apply(extract_driving_permission)
# df['Scouter Name'] = df['Modifiers_Applied'].apply(extract_scouter_name)
df['Cell Phone Number'] = df['Modifiers_Applied'].apply(extract_cell_phone_number)
# Extract and format emergency contact name and phone number


# #####################################################################################
# Now using sorted dataframe/file for the rest of the process
# #####################################################################################

# #####################################################################################
# Parse the Modifiers_Applied column // Remove whitespace // Write Scout file
# #####################################################################################
# df = df[df['Item'].isin(filterChoice1)]
# df = df[(df.Item == filterChoice1)]
# df[['Scout_Label', 'Attendee', 'Rank', 'Patrol', 'Drive_Status', 'Emer_Label',
#     'Emer_Contact', 'Emer_Label2', 'Emer_Contact_Ph']] = df.Modifiers_Applied.str.split(",|:", expand=True, )
# df2 = df.Modifiers_Applied.str.split(",|:", expand=True, )
# print(df2)

df.to_csv(filteredFile2)

# whitespace_remover(df)
# df['Emer_Contact_Ph'] = phoneparser(df['Emer_Contact_Ph'])
# Extract and format emergency contact phone numbers
# df['Emer_Contact_Ph'] = df['Modifiers_Applied'].apply(lambda x: format_phone_number(x.split('Emergency Contact Phone Number: ')[-1]))

df.to_csv(filteredFile1, index=False, columns=['Date', 'Time', 'Item',
                                               'Qty', 'Net_Sales',
                                               'Customer_Name', 'Attendee',
                                               'Rank', 'Patrol',
                                               'Driving_Status', 'Emer_Contact',
                                               'Emer_Contact_Ph'])

# #####################################################################################
# Read sorted file for Data Frame 2
# #####################################################################################
df2 = pd.read_csv(mySortedFile, header=0)

# #####################################################################################
# Parse Modifiers_Applied column // Remove whitespace //Write Scouter file
# #####################################################################################
df2 = df2[df2['Item'].isin(filterChoice2)]
# df2 = df2[(df2.Item == filterChoice2)]
# df2[['Scouter_Label', 'Attendee', 'Cell_Label', 'Cell_Num', 'Driving_Status',
#      'Emer_Contact', 'Emer_Contact_Ph']] = df2.Modifiers_Applied.str.split(",|:", expand=True, )
# whitespace_remover(df2)
# new >>
df2['Driving_Status'] = df2['Modifiers_Applied'].apply(extract_driving_permission)
# df2['Attendee'] = df2['Modifiers_Applied'].apply(extract_scouter_name)
df2['Cell_Num'] = df2['Modifiers_Applied'].apply(extract_cell_phone_number)
df2['Rank'] = df2['Modifiers_Applied'].apply(extract_scout_rank)
df2['Patrol'] = df2['Modifiers_Applied'].apply(extract_scout_patrol)
# df['Emer_Contact_Ph'] = df['Modifiers_Applied'].apply(lambda x: extract_emergency_contact_phone_number(x))
# df2['Emer_Contact'] = df2['Modifiers_Applied'].apply(extract_emergency_contact)

# df2['Emer_Contact'], df2['Emer_Contact_Ph'] = zip(*df2['Modifiers_Applied'].apply(lambda x: extract_emergency_contact_phone_number(x)))
print(df2['Emer_Contact_Ph'])



df2['Cell_Num'] = phoneparser(df2['Cell_Num'])
# df2['Emer_Contact_Ph'] = phoneparser(df2['Emer_Contact_Ph'])


# df2 = df2.assign(Patrol='Rocking Chair')
df2.to_csv(filteredFile2, index=False, columns=['Date', 'Time', 'Item',
                                                'Qty', 'Net_Sales',
                                                'Customer_Name', 'Attendee',
                                                'Cell_Num', 'Driving_Status', 'Patrol',
                                                'Emer_Contact',
                                                'Emer_Contact_Ph'])

# #####################################################################################
# Read in the two separate files and concatenate into one output file
# #####################################################################################
csv1 = pd.read_csv(filteredFile1)
logging.info(csv1.head())
csv2 = pd.read_csv(filteredFile2)
logging.info(csv2.head())

registrationFull = f'{myDirectory}{eventName}_Registered_as_of'
concate_data = pd.concat([csv1, csv2])
concate_data.head()
# concate_data.to_csv(f'{registrationFull}_{current_date_time}.csv', index=False,
#                     columns=['Date', 'Time', 'Item', 'Qty', 'Net_Sales', 'Customer_Name', 'Attendee',
#                              'Cell_Num', 'Patrol', 'Rank', 'Driving_Status', 'Emer_Contact', 'Emer_Contact_Ph'],
#                     )

#
# Summary Pivot
#
print(concate_data)
sumtable = pd.pivot_table(concate_data, index=["Item", "Patrol"], values="Qty", aggfunc=[np.sum], fill_value=0,
                                    margins=True)

# >>>NEW>>>
writer = pd.ExcelWriter(f'{registrationFull}_{current_date_time}.xlsx', engine='xlsxwriter')
concate_data.to_excel(writer, sheet_name='RegistrationData', index=False)
sumtable.to_excel(writer, sheet_name='SummaryTables', index=True)
workbook = writer.book
worksheet = writer.sheets['RegistrationData']
worksheet2 = writer.sheets['SummaryTables']
format1 = workbook.add_format({'num_format': '#,##0.00'})
format2 = workbook.add_format({'num_format': '0%'})

worksheet.set_column('A:B', 12)
worksheet.set_column('C:C', 30)
worksheet.set_column('D:E', 8)
worksheet.set_column('F:M', 18)
worksheet2.set_column('A:A', 35)
worksheet2.set_column('B:B', 15)
worksheet2.set_column('C:C', 11)

writer.close()

# xl = pd.read_csv(f'{registrationFull}_{current_date_time}.csv')
# Loading Excel file
# wb = xlrd.open_workbook("C:/Users/chris/Downloads/Registrations.xls")
# Storing the first sheet into a variable
# sheet = wb.sheet_by_index(0)
# Printing various cell values
# print("Value of 0-0 cell: ",sheet.cell_value(0, 0))
# print("Value of 2-4 cell: ",sheet.cell_value(2, 4))
# # Get max no of rows and columns
# print("Number of Rows: ", sheet.nrows)
# print("Number of Columns: ",sheet.ncols)
# # Get all column names
# print("ALL COLUMN NAMES ARE: ")
# for i in range(sheet.ncols):
#     print(sheet.cell_value(0,i))
# # Get first 10 rows for 5 columns
# for i in range(11):
#     for j in range(5):
#         print(sheet.cell_value(i,j), end="\t\t\t")
#     print()


# #################################################################
# Read in the two separate files and merge into one output file
# #################################################################

# registrationMerge = f'{myDirectory}{eventName}Merged_as_of'
# merged_data = csv2.merge(csv1,on=["Customer_Name"])
# merged_data.head()
# merged_data.to_csv(f'{registrationMerge}_{current_date_time}.csv' , index=False)')

# See PyCharm help at https://www.jetbrains.com/help/pycharm/
