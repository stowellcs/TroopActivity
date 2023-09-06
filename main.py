import pandas as pd
import re
import datetime
import logging


# if you want a different format for the phone numbers, change it here
def fmt_phone_number(s):
    formatted_number = '({}) {}-{}'.format(s[0:3], s[3:6], s[6:10])
    return formatted_number


# get the cell and send it to the format function
def extract_cell_phone_number(s):
    phone_number_string = re.search(r'Cell phone number:\s([\d-]{10})', s)
    if phone_number_string:
        phone_number_start_index = s.find(phone_number_string.group(1))
        phone_number = fmt_phone_number(s[phone_number_start_index:phone_number_start_index + 10])
        return phone_number
    else:
        return None


# the idea of this function is to take the item name, strip out the "Scout|Scouter Registration" part for better reports
def get_event(s):
    scout_event_string = re.search(r'Scout Registration -\s*(\w+)', s)
    scouter_event_string = re.search(r'Scouter Registration -\s*(\w+)', s)
    if scout_event_string:
        event_index = s.find(scout_event_string.group(1))
        event = s[event_index:event_index + (len(s) - 21)]
        return event
    elif scouter_event_string:
        event_index = s.find(scouter_event_string.group(1))
        event = s[event_index:event_index + (len(s) - 23)]
        return event
    return s


# handle names, whether they are first and last or include middle
def split_name(s):
    name_parts = s.split()
    if len(name_parts) == 0:
        first_name = None
        middle_name = None
        last_name = None
    elif len(name_parts) == 2:
        first_name, last_name = name_parts
        middle_name = None
    elif len(name_parts) == 3:
        first_name, middle_name, last_name = name_parts
    else:
        first_name = name_parts[0]
        middle_name = ' '.join(name_parts[1:-1])
        last_name = name_parts[-1]
    return pd.Series({'First Name': first_name, 'Middle Name': middle_name, 'Last Name': last_name})


def parse_name(name):
    if name:
        first, *middle, last = name.split()
        return pd.Series({'First Name': first, 'Middle Name': ' '.join(middle), 'Last Name': last})


# find the attendee's name in the file and return it without the label
def extract_attendee_name(s):
    scout_name = re.search(r'Scout Name:\s*(\w+\s+\w+)', s)
    scouter_name = re.search(r'Scouter Name:\s*(\w+\s+\w+)', s)
    if scout_name:
        return scout_name.group(1)
    elif scouter_name:
        return scouter_name.group(1)
    else:
        return extract_abbrev_name(s)


def extract_abbrev_name(s):
    scout_name = re.search(r'Scout Name:\s*(\w+)', s)
    scouter_name = re.search(r'Scouter Name:\s*(\w+)', s)
    if scout_name:
        return scout_name.group(1)
    else:
        return scouter_name.group(1)


# pull out the emergency contact without the label
def extract_emergency_contact(s):
    contact_name_string = re.search(r'Emergency Contact:\s*([\w\s]+)?', s)
    if contact_name_string:
        return contact_name_string.group(1)
    else:
        return "Unknown"


# pull out the e-contact's phone number and send it to the phone formatter, without the label
def extract_emergency_contact_phone_number(s):
    if s:
        phone_number_string = re.search(r'Emergency Contact Phone Number:\s*([\d-]{10})', s)
        if phone_number_string:
            phone_number_start_index = s.find(phone_number_string.group(1))
            phone_number = fmt_phone_number(s[phone_number_start_index:phone_number_start_index + 10])
            return phone_number
        else: "NA"
    else:
        return "Unknown"


# Define function to extract Driving Permission
def extract_driving_permission(s):
    if 'Yes - driving' in s:
        return 'Yes-Driving'
    elif 'Yes - carpool' in s:
        return 'Yes-Carpool'
    else:
        return 'No'


# Define function to extract Scout Rank. Because position in Square file changes, use list as lookup values
def extract_scout_rank(s):
    if 'Scout Name' in s:
        ranks = [' Prospective scout / new to troop', ' Scout', ' Tenderfoot', ' Second Class', ' First Class', ' Star',
                 ' Life', ' Eagle']
        parts = s.split(',')
        for g in ranks:
            if g in parts:
                return g
        else:
            return ''


# Define function to extract Scout Patrol. Because position in Square file changes, use list as lookup values
def extract_scout_patrol(s):
    if 'Scout Name' in s:
        patrols = [' Cobras', ' Pedros', ' Green Frogs', ' Vikings']
        parts = s.split(',')
        for g in patrols:
            if g in parts:
                return g
    else:
        return 'Rocking Chair'


# if the filter variable is on, create two extra files - one for filter choice 1 and one for choice 2
def filterfiles():
    dfscout = df[(df.Item == filterChoice1)]
    dfscout.to_csv(filteredFile1, index=False, columns=column_list)
    dfscouter = df[(df.Item == filterChoice2)]
    dfscouter.to_csv(filteredFile2, index=False, columns=column_list)


# I put the summary table definition here
def summarizetbl():
    dfsumgrp = df.groupby(['Item', 'Driving Status', 'Patrol']).count()
    dfsum2grp = df.groupby(['Item', 'Patrol']).count()
    return dfsumgrp, dfsum2grp


# #####################################################################################
# USER MODIFY: Input files defined
# #####################################################################################
myDirectory = 'C:/Users/chris/Downloads/'
fileName = 'items-2023-05-31-2023-09-29'

myInputFile = f'{myDirectory}{fileName}.csv'

# If you want to filter the file, enter the name of the event(s) here.  Item name (without Scout/Scouter designation)
filterChoice1 = 'End of Year Pool Party 2023'
filterChoice2 = 'Aquatics Campout - September 2023'

# eventName is essentially the prefix on your file name
eventName = 'Aquatics Sep 2023'

# output variables to set whether you want CSV and or Excel.  Excel is formatted.
outputTypeCSV = False
outputTypeExcel = True
applyfilter = False

# the filteredFileOut variable determines if you want a separate scout and scouter file for research.
# This is used for testing changes
filteredFileOut = False

# #####################################################################################
# File Names - Various Steps
# #####################################################################################
mySortedFile = f'{myInputFile}_sorted.csv'

filteredFile1 = f'{myDirectory}{filterChoice1}.csv'
filteredFile2 = f'{myDirectory}{filterChoice2}.csv'

current_date_time = format(datetime.datetime.now().strftime("%Y-%m-%d-%H.%M"))

# List of columns to output on report
column_list = ['Date', 'Time', 'Item',
               'Qty', 'Price_Point_Name', 'Net_Sales',
               'Customer_Name', 'Attendee Name',
               'AttendLastName', 'AttendFirstName',
               'Scout Rank', 'Patrol',
               'Driving Status', 'Emergency Contact',
               'Emergency Contact Phone Number']
column_list2 = ['Qty']
patrol_categories = ['Green Frogs', 'Cobras', 'Pedros', 'Vikings', 'Rocking Chair', 'Unknown']

# #####################################################################################
# Unit testing here
# #####################################################################################

# tcase1 = scout
# tcase2 = scouter
# tcase3 = scout alternative fmt
# tcase4 = 'Scott Stowell'
# tcase5 = alternate layout with econtact first
# tcase6 = ditto
# tcase7 = 'Scouter Registration - Wilderness Survival - April 2023'
# tcase8 = 'Scout Registration - NASA'
# tcase9 = 'Emergency Contact: Sara Hale, Scouter Name: John Hale, Yes - driving, Cell phone number: 2144053412, Emergency Contact Phone Number: 21468-8468'

# #####################################################################################
# Logging
# #####################################################################################
logging.basicConfig(level=logging.INFO, filename=f'{myDirectory}App.log', filemode='w',
                    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
logging.info('Admin logged in')
logging.debug('debug sample')

# #####################################################################################
# Read Input File  // Clean the column names // Parse Name
# #####################################################################################

df = pd.read_csv(myInputFile, header=0)
df = df.dropna(axis=1, how='all')
df.columns = [re.sub("[ ,-]", "_", re.sub("[.,`]", "", c)) for c in df.columns]

df['Customer_Name'] = df['Customer_Name'].fillna('Unknown')
df[['CustFirstName', 'CustMiddleName', 'CustLastName']] = df['Customer_Name'].apply(split_name)

df['Item'] = df['Item'].replace('\n', " ", regex=True)
df['Item'] = df['Item'].apply(get_event)

# df.to_csv(f'{myDirectory}test.csv')

# #####################################################################################
# Sort by sales item and last name of customer// Create Sorted File
# #####################################################################################

# df.sort_values(by=['Item', 'Price_Point_Name', 'CustLastName'], inplace=True,
#                key=lambda col: col.str.lower())
#
# Add new columns to dataframe using the functions
try:
    df['Emergency Contact Phone Number'] = df['Modifiers_Applied'].apply(extract_emergency_contact_phone_number)
except:
    df['Emergency Contact Phone Number'] = 'NA'

try:
    df['Emergency Contact'] = df['Modifiers_Applied'].apply(extract_emergency_contact)
except:
    df['Emergency Contact'] = 'NA'

df['Attendee Name'] = df['Modifiers_Applied'].apply(extract_attendee_name)

df[['AttendFirstName', 'AttendMiddleName', 'AttendLastName']] = df['Modifiers_Applied'].apply(split_name)

df['Driving Status'] = df['Modifiers_Applied'].apply(extract_driving_permission)
df['Cell Phone Number'] = df['Modifiers_Applied'].apply(extract_cell_phone_number)
df['Patrol'] = df['Modifiers_Applied'].apply(extract_scout_patrol).fillna("Unknown")
df['Patrol'] = df['Patrol'].str.strip()
df['Scout Rank'] = df['Modifiers_Applied'].apply(extract_scout_rank)
df['Scout Rank'] = df['Scout Rank'].str.strip()

# #####################################################################################
# Sort by sales item and last name of attendee // Create sorted file
# #####################################################################################

df.sort_values(by=['Item', 'Price_Point_Name', 'AttendLastName', 'AttendFirstName'], inplace=True,
               key=lambda col: col.str.lower())

if applyfilter:
    filtersused = [filterChoice1, filterChoice2]
    df = df[df['Item'].isin(filtersused)]

# df.to_csv(f'{mySortedFile}.csv')


# #####################################################################################
# Now using sorted dataframe/file for the rest of the process
# #####################################################################################

if filteredFileOut:
    filterfiles()

registrationFull = f'{myDirectory}{eventName}_Registered_as_of_{current_date_time}'

if outputTypeCSV:
    df.to_csv(f'{registrationFull}.csv', index=False, columns=column_list)
    logging.info('CSV output is on in ' + myDirectory)
else:
    logging.info('CSV output is off')

rowCount = len(df.index)
suminfo = [rowCount]
column_names = ['Total']
df2 = pd.DataFrame(suminfo, columns=column_names)
df2.index = ['Total Attendees']
print(df2)

if outputTypeExcel:
    writer = pd.ExcelWriter(f'{registrationFull}.xlsx', engine='xlsxwriter')
    df.to_excel(writer, sheet_name='SignUps', index=False, startcol=0, startrow=0, columns=column_list)
    df['Patrol'] = pd.Categorical(df['Patrol'], categories=patrol_categories)
    df.sort_values(by='Patrol')
    [dfSum, dfSum2] = summarizetbl()
    df2.to_excel(writer, sheet_name='Summary', index=True, startcol=0, startrow=0)
    dfSum2.to_excel(writer, sheet_name='Summary', index=True, startcol=0, startrow=3, columns=column_list2)
    dfSum.to_excel(writer, sheet_name='Summary', index=True, startcol=5, startrow=3, columns=column_list2)
    logging.info('Excel output is on in ' + myDirectory)
    worksheet = writer.sheets['SignUps']
    worksheet2 = writer.sheets['Summary']
    worksheet.set_column('A:B', 10)
    worksheet.set_column('C:C', 45)
    worksheet.set_column('D:D', 6)
    worksheet.set_column('E:E', 20)
    worksheet.set_column('F:F', 10)
    worksheet.set_column('G:I', 25)
    worksheet.set_column('J:K', 23)
    worksheet.set_column('L:L', 25)
    worksheet.set_column('M:O', 28)
    worksheet2.set_column('A:A', 35)
    worksheet2.set_column('B:B', 20)
    worksheet2.set_column('C:C', 15)
    worksheet2.set_column('F:F', 30)
    worksheet2.set_column('G:H', 20)
    worksheet2.set_column('I:I', 15)
    writer.close()
else:
    print('Excel output is off')

# unused section that creates custom formats by type

# workbook = writer.book

# format1 = workbook.add_format({'num_format': '#,##0.00'})

# format2 = workbook.add_format({'num_format': '0%'})
# add_format({'num_format': '0%'})
