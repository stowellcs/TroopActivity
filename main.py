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


# #####################################################################################
# USER MODIFY: Input files defined
# #####################################################################################
myDirectory = 'C:/Users/chris/Downloads/'
myInputFile = f'{myDirectory}items-2022-11-11-2023-02-11b.csv'

# #####################################################################################
# USER MODIFY: Registration Label to filter on
# #####################################################################################
filterChoice1 = 'Scout Registration - NASA'
filterChoice2 = 'Scouter Registration - NASA'
eventName = 'March 2023 NASA Campout'
# filterChoice1 = 'Scout Registration - Mountain Bike'
# filterChoice2 = 'Scouter Registration - Mountain Bike'
# eventName = 'Feb 2023 Mountain Bike'

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
nameparser(df['Customer_Name'])

# #####################################################################################
# Sort by sales item and last name of customer// Create Sorted File
# #####################################################################################
df.sort_values(by=['Item', 'LNAME_CUST'], inplace=True, key=lambda col: col.str.lower())
# df.to_csv(mySortedFile)

# #####################################################################################
# Now using sorted dataframe/file for the rest of the process
# #####################################################################################

# #####################################################################################
# Parse the Modifiers_Applied column // Remove whitespace // Write Scout file
# #####################################################################################
df = df[(df.Item == filterChoice1)]
df[['Scout_Label', 'Attendee', 'Rank', 'Patrol', 'Drive_Status', 'Emer_Label',
    'Emer_Contact', 'Emer_Label2', 'Emer_Contact_Ph']] = df.Modifiers_Applied.str.split(",|:", expand=True, )
df.to_csv(filteredFile2)

whitespace_remover(df)
df['Emer_Contact_Ph'] = phoneparser(df['Emer_Contact_Ph'])
df.to_csv(filteredFile1, index=False, columns=['Date', 'Time', 'Item',
                                               'Qty', 'Net_Sales',
                                               'Customer_Name', 'Attendee',
                                               'Rank', 'Patrol',
                                               'Drive_Status', 'Emer_Contact',
                                               'Emer_Contact_Ph'])

# #####################################################################################
# Read sorted file for Data Frame 2
# #####################################################################################
df2 = pd.read_csv(mySortedFile, header=0)

# #####################################################################################
# Parse Modifiers_Applied column // Remove whitespace //Write Scouter file
# #####################################################################################
df2 = df2[(df2.Item == filterChoice2)]
df2[['Scouter_Label', 'Attendee', 'Cell_Label', 'Cell_Num', 'Drive_Status', 'Emer_Label',
     'Emer_Contact', 'Emer_Label2', 'Emer_Contact_Ph']] = df2.Modifiers_Applied.str.split(",|:", expand=True, )
whitespace_remover(df2)
# new >>
df2['Cell_Num'] = phoneparser(df2['Cell_Num'])
df2['Emer_Contact_Ph'] = phoneparser(df2['Emer_Contact_Ph'])

df2 = df2.assign(Patrol='Rocking Chair')
df2.to_csv(filteredFile2, index=False, columns=['Date', 'Time', 'Item',
                                                'Qty', 'Net_Sales',
                                                'Customer_Name', 'Attendee',
                                                'Cell_Num', 'Drive_Status', 'Patrol',
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
#                              'Cell_Num', 'Patrol', 'Rank', 'Drive_Status', 'Emer_Contact', 'Emer_Contact_Ph'],
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
