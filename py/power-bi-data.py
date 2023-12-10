import pandas as pd
import seaborn as sns
import geopandas as gpd

from dateutil import parser

# Import dataset
path = r'C:\Users\niels\OneDrive\7. Portfolio Data Analytics\Challenges\Maven Analytics - Power Outage\Data\DOE_Electric_Disturbance_Events.xlsx'
wb = pd.read_excel(path, sheet_name=None)

# Prepare dataframes
data = pd.DataFrame()
data_removed = pd.DataFrame()

# Lists for cleaning data

# List of values for filtering "Date Event Began" column
values_to_exclude = ['january', 'february', 'march', 'april', 'may', 'june', 'july', 'august', 'september', 'october',
                     'november', 'december', 'none', 'table b.2.', 'date', 'date/time', 'date event began', 'ongoing', 'na', 'note', 'continued', 'estimated', 'source', 'http', 'information']

# Values to format
values_to_replace = ['ongoing', 'unknown',
                     'unknown ', 'unkonwn', 'none', 'nan', 'NaT']

# Values to 0
strings_to_zero = ['unknown']


'''Functions'''

# Functions for cleaning and convert date, time and datetime values

# Parsing time strings


def parse_time_string(time_string, show_errors=False):
    if pd.notna(time_string):
        original_value = time_string

        # Remove incorrect string values before parsing
        try:
            time_string = time_string.lower().replace(',', ' ')
            time_string = time_string.lower().replace(': ', ':')
            time_string = time_string.lower().replace('noon', 'p.m.')
            time_string = time_string.lower().replace('unknown', '')
            time_string = time_string.lower().replace('12:00 midnight', '00:00')
            time_string = time_string.lower().replace('midnight', '00:00')
            time_string = time_string.lower().replace('evening', '5 p.m.')
            time_string = time_string.lower().replace('ongoing', ' ')
            time_string = time_string.lower().replace('approximately', '')
            time_string = time_string.lower().replace('approximately ', '')

            # Parse time with parser
            parsed_time = parser.parse(time_string)
            formatted_time = parsed_time.strftime('%H:%M:%S')

            return formatted_time
        except ValueError as e:
            try:
                # If parser gives an error, try pd.to_datetime() or print error
                numeric_time = pd.to_numeric(
                    original_value, errors='coerce') * 24 * 3600
                formatted_time = pd.to_datetime(
                    numeric_time, unit='s', errors='coerce').time()

                return formatted_time

            except ValueError as e:
                # If both approaches fail, print the error for checking and keep original value
                if show_errors and time_string != 'NaN' and time_string != 'nan' and time_string != 'nat':
                    print(f"Error parsing time string '{time_string}': {e}")
                return original_value

    else:
        return pd.NaT

# Parsing datetime strings


def parse_datetime_string(datetime_string, show_errors=False):
    if pd.notna(datetime_string):
        original_value = datetime_string

        # Remove incorrect string values before parsing
        try:
            datetime_string = datetime_string.lower().replace(',', ' ')
            datetime_string = datetime_string.lower().replace(': ', ':')
            datetime_string = datetime_string.lower().replace('noon', 'p.m.')
            datetime_string = datetime_string.lower().replace('midnight', 'a.m.')
            datetime_string = datetime_string.lower().replace('unknown', '')
            datetime_string = datetime_string.lower().replace('ongoing', '')
            datetime_string = datetime_string.lower().replace('12:00 midnight', '00:00')
            datetime_string = datetime_string.lower().replace('approximately', '')
            datetime_string = datetime_string.lower().replace('approximately ', '')
            datetime_string = datetime_string.lower().replace('(trans. only)', '')

            # Parse time with parser
            parsed_datetime = parser.parse(datetime_string)
            formatted_datetime = parsed_datetime.strftime('%Y-%m-%d %H:%M:%S')

            return pd.to_datetime(formatted_datetime)

        # Print error
        except ValueError as e:
            if show_errors and datetime_string != 'nan' and datetime_string != 'NaT' and datetime_string != 'nat':
                print(
                    f"Error parsing datetime string '{datetime_string}': {e}")
            return original_value

    else:
        return pd.NaT

# Parsing date strings


def parse_date_string(date_string, show_errors=False):
    if pd.notna(date_string):
        original_value = date_string

        # Remove incorrect string values before parsing
        try:
            date_string = date_string.lower().replace(',', ' ')
            date_string = date_string.lower().replace(': ', ':')
            date_string = date_string.lower().replace('//', '/')
            date_string = date_string.lower().replace('unknown', '')
            date_string = date_string.lower().replace('44641', '31-03-2022')

            # Parse time with parser
            parsed_date = parser.parse(date_string)
            formatted_date = parsed_date.strftime('%Y-%m-%d')

            return pd.to_datetime(formatted_date)

        # Print error
        except ValueError as e:
            if show_errors and date_string != 'NaT' and date_string != 'nat' and date_string != 'nan':
                print(f"Error parsing date string '{date_string}': {e}")

            return original_value
    else:
        return pd.NaT

# Functions for parsing Alert Criteria and Event Type column


def report_type_id(row):
    if pd.isna(row) or row == '':
        return None

    id_1 = ['physical attack that causes major interruptions or impacts']
    id_2 = ['cyber event that causes interruptions',
            'reportable cyber security incident']
    id_3 = ['complete operational failure']
    id_4 = ['islanding']
    id_5 = ['uncontrolled loss of 300 megawatts or more of firm system loads']
    id_6 = ['firm load shedding of 100 megawatts',
            'load shedding of 100 megawatts or more', 'load shed 100']
    id_7 = ['voltage reductions of 3 percent', 'voltage reduction']
    id_8 = ['public appeal to reduce']
    id_9 = ['physical attack that could potentially impact electric power system',
            'actual physical attack']
    id_10 = ['cyber event that could potentially impact electric power system',
             'cyber security incident that was an attempt to compromise', 'suspected cyber attack']
    id_11 = ['loss of electric service to more than 50,000 customers']
    id_12 = ['fuel supply emergencies that could impact electric power system']
    id_13 = ['damage or destruction of a facility within its reliability coordinator']
    id_14 = ['damage or destruction of its facility that results from actual or suspected intentional human action',
             'suspected physical attack']
    id_15 = [
        'physical threat to its facility excluding weather or natural disaster related threats']
    id_16 = [
        'physical threat to its bulk electric system control center, excluding weather']
    id_17 = ['bulk electric system emergency resulting in voltage deviation',
             'voltage deviation equal to or greater than 10%']
    id_18 = [
        'uncontrolled loss of 200 megawatts or more of firm system loads for 15 minutes or more']
    id_19 = [
        'total generation loss, within one minute of: greater than or equal to 2,000 megawatts']
    id_20 = ['affecting a nuclear generating station']
    id_21 = ['unexpected transmission loss within its area, contrary to design, of three or more bulk electric system facilities']
    id_22 = ['unplanned evacuation']
    id_23 = [
        'complete loss of interpersonal communication and alternative interpersonal communication']
    id_24 = ['loss of monitoring or control']

    if any(keyword in row.lower() for keyword in id_1):
        return 1

    elif any(keyword in row.lower() for keyword in id_2):
        return 2

    elif any(keyword in row.lower() for keyword in id_3):
        return 3

    elif any(keyword in row.lower() for keyword in id_4):
        return 4

    elif any(keyword in row.lower() for keyword in id_5):
        return 5

    elif any(keyword in row.lower() for keyword in id_6):
        return 6

    elif any(keyword in row.lower() for keyword in id_7):
        return 7

    elif any(keyword in row.lower() for keyword in id_8):
        return 8

    elif any(keyword in row.lower() for keyword in id_9):
        return 9

    elif any(keyword in row.lower() for keyword in id_10):
        return 10

    elif any(keyword in row.lower() for keyword in id_11):
        return 11

    elif any(keyword in row.lower() for keyword in id_12):
        return 12

    elif any(keyword in row.lower() for keyword in id_13):
        return 13

    elif any(keyword in row.lower() for keyword in id_14):
        return 14

    elif any(keyword in row.lower() for keyword in id_15):
        return 15

    elif any(keyword in row.lower() for keyword in id_16):
        return 16

    elif any(keyword in row.lower() for keyword in id_17):
        return 17

    elif any(keyword in row.lower() for keyword in id_18):
        return 18

    elif any(keyword in row.lower() for keyword in id_19):
        return 19

    elif any(keyword in row.lower() for keyword in id_20):
        return 20

    elif any(keyword in row.lower() for keyword in id_21):
        return 21

    elif any(keyword in row.lower() for keyword in id_22):
        return 22

    elif any(keyword in row.lower() for keyword in id_23):
        return 23

    elif any(keyword in row.lower() for keyword in id_24):
        return 24

    else:
        return 24


def emergency_cause_id(row):
    if pd.isna(row) or row == '':
        return

    id_1 = ["unknown"]
    id_2 = ['physical attack', 'sabotage', 'actual physical event',
            'suspected physical attack', 'suspected sabotage', 'suspected telecommunications attack']
    id_3 = ['threat of physical', 'potential physical attack']
    id_4 = ['vandalsim', 'vandalism']
    id_5 = ['theft']
    id_6 = ['suspicious activity']
    id_7 = []
    id_8 = ['cyber']
    id_9 = ["fuel supply"]
    id_10 = ['generator', 'generation inadequacy']
    id_11 = ['transmission equipment', 'transmission  equipment', 'transmission system', 'transmission level',
             'equipment trip', 'equipment failure', 'equipment malfunction', 'equipment faulted', 'transformer failure']
    id_12 = ["switch", "failure at high voltage substation", 'substation']
    id_13 = ['weather', 'natural disaster', 'storm', 'lightning', 'wind', 'tornado', 'hurricane', 'heat wave', 'heatwave'
             'earthquake', 'earthquake', 'wildfire', 'brush fire', 'tropical', 'ice', 'flood', 'rain', 'wild fire', 'high winds', 'high temperatures', 'wild land fire']
    id_14 = ["operator"]
    id_15 = ['other']

    if any(keyword in row.lower() for keyword in id_1):
        return 1

    elif any(keyword in row.lower() for keyword in id_3):
        return 3

    elif any(keyword in row.lower() for keyword in id_5):
        return 5

    elif any(keyword in row.lower() for keyword in id_4):
        return 4

    elif any(keyword in row.lower() for keyword in id_6):
        return 6

    elif any(keyword in row.lower() for keyword in id_2):
        return 2

    elif any(keyword in row.lower() for keyword in id_7):
        return 7

    elif any(keyword in row.lower() for keyword in id_8):
        return 8

    elif any(keyword in row.lower() for keyword in id_13):
        return 13

    elif any(keyword in row.lower() for keyword in id_9):
        return 9

    elif any(keyword in row.lower() for keyword in id_10):
        return 10

    elif any(keyword in row.lower() for keyword in id_11):
        return 11

    elif any(keyword in row.lower() for keyword in id_12):
        return 12

    elif any(keyword in row.lower() for keyword in id_14):
        return 14

    elif any(keyword in row.lower() for keyword in id_15):
        return 15

    else:
        return 15


def emergency_impact_id(row):
    if pd.isna(row) or row == '':
        return 17

    id_1 = ['none']
    id_2 = ['unplanned evacuation from its bulk electric system control center']
    id_3 = ['Complete loss of Interpersonal Communication and Alternative Interpersonal Communication capability',
            'complete loss of monitoring or control capability']
    id_4 = ['damage', 'destruction']
    id_5 = ['electrical system separation',
            'electric system separation', 'islanding', 'electrical separation']
    id_6 = ['complete operational failure',
            'complete operational failure or shut down of the transmission and/or distribution electrical system', 'complete electric system failure']
    id_7 = ['three or more BES elements']
    id_8 = ['major distribution system']
    id_9 = ['uncontrolled loss of 200 mw']
    id_10 = ['loss of electric service to more than 50,000 customers']
    id_11 = ['voltage reductions of 3 percent']
    id_12 = ['bulk electric system emergency resulting in voltage deviation',
             'voltage deviation equal to or greater than 10%']
    id_13 = ['inadequate electric resources to serve load']
    id_14 = ['capacity loss of 1,400 mw']
    id_15 = ['capacity loss of 2,000 mw']
    id_16 = ['nuclear generating']
    id_17 = ['other']

    if any(keyword in row.lower() for keyword in id_1):
        return 1

    elif any(keyword in row.lower() for keyword in id_2):
        return 2

    elif any(keyword in row.lower() for keyword in id_3):
        return 3

    elif any(keyword in row.lower() for keyword in id_4):
        return 4

    elif any(keyword in row.lower() for keyword in id_5):
        return 5

    elif any(keyword in row.lower() for keyword in id_6):
        return 6

    elif any(keyword in row.lower() for keyword in id_7):
        return 7

    elif any(keyword in row.lower() for keyword in id_8):
        return 8

    elif any(keyword in row.lower() for keyword in id_9):
        return 9

    elif any(keyword in row.lower() for keyword in id_10):
        return 10

    elif any(keyword in row.lower() for keyword in id_11):
        return 11

    elif any(keyword in row.lower() for keyword in id_12):
        return 12

    elif any(keyword in row.lower() for keyword in id_13):
        return 13

    elif any(keyword in row.lower() for keyword in id_14):
        return 14

    elif any(keyword in row.lower() for keyword in id_15):
        return 15

    elif any(keyword in row.lower() for keyword in id_16):
        return 16

    elif any(keyword in row.lower() for keyword in id_17):
        return 17

    else:
        return 17


def emergency_action_id(row):
    if pd.isna(row) or row == '':
        return 9
    id_1 = ['none']
    id_2 = ['load shedding of 100 megawatt',
            'load shed 100+', 'load shed of 100+', ]
    id_3 = ['public appeal to reduce']
    id_4 = ['warning', 'alert', 'contingency plan', 'implementation of stage 2 electrical emergency plan',
            'declaration of  transmission emergency', 'declared stage 1 electric emergency']
    id_5 = ['voltage reduction']
    id_6 = ['shed interruptible load',
            'interruptible load shed', '/interruptible load shed']
    id_7 = ['repaired', 'restored']
    id_8 = ['mitigation implemented',
            'initiated interruption of air conditioner']
    id_9 = ['other']

    if any(keyword in row.lower() for keyword in id_1):
        return 1

    elif any(keyword in row.lower() for keyword in id_2):
        return 2

    elif any(keyword in row.lower() for keyword in id_3):
        return 3

    elif any(keyword in row.lower() for keyword in id_4):
        return 4

    elif any(keyword in row.lower() for keyword in id_5):
        return 5

    elif any(keyword in row.lower() for keyword in id_6):
        return 6

    elif any(keyword in row.lower() for keyword in id_7):
        return 7

    elif any(keyword in row.lower() for keyword in id_8):
        return 8

    elif any(keyword in row.lower() for keyword in id_9):
        return 9

    else:
        return 9


def update_impact_id(row):
    if row['Impact ID'] == 17 and row['Report Type ID'] == 22:
        return 2
    elif row['Impact ID'] == 17 and (row['Report Type ID'] == 23 or row['Report Type ID'] == 24):
        return 3
    elif row['Impact ID'] == 17 and (row['Report Type ID'] == 13 or row['Report Type ID'] == 14):
        return 4
    elif row['Impact ID'] == 17 and row['Report Type ID'] == 4:
        return 5
    elif row['Impact ID'] == 17 and row['Report Type ID'] == 18:
        return 9
    elif row['Impact ID'] == 17 and row['Report Type ID'] == 11:
        return 10
    elif row['Impact ID'] == 17 and row['Report Type ID'] == 7:
        return 11
    elif row['Impact ID'] == 17 and row['Report Type ID'] == 17:
        return 12
    elif row['Impact ID'] == 17 and row['Report Type ID'] == 20:
        return 16
    else:
        return row['Impact ID']


def update_action_id(row):
    if row['Action ID'] == 9 and row['Report Type ID'] == 6:
        return 2
    elif row['Action ID'] == 9 and row['Report Type ID'] == 8:
        return 3
    elif row['Action ID'] == 9 and row['Report Type ID'] == 7:
        return 5
    else:
        return row['Action ID']


def replace_values(value):
    lower_value = str(value).lower().strip()
    if lower_value in [val.lower().strip() for val in values_to_replace]:
        return pd.NaT
    return value


def replace_string_to_zero(value):
    lower_value = str(value).lower()
    if lower_value in [val.lower() for val in strings_to_zero]:
        return 0
    return value


'''Cleaning sheets and concatenate data'''

# Iterate through individual sheets in the Excel workbook
for sheet_name, sheet_data in wb.items():

    # Find the index of the first row containing strings
    first_string_row = sheet_data[sheet_data.map(
        lambda x: isinstance(x, str)).all(axis=1)].index[0]

    # Use the first string row as column names
    sheet_data.columns = sheet_data.iloc[first_string_row]

    # Remove rows before the first column name row
    sheet_data = sheet_data.iloc[first_string_row + 1:]

    # Rename column names
    for old_column, new_column in {
        'Date': 'Date Event Began',
        'Time': 'Time Event Began',
        'Restoration': 'Restoration Time',
        'Type of Disturbance': 'Event Type',
        'Loss (megawatts)': 'Demand Loss (MW)',
        'Number of Customers Affected 1[1]': 'Number of Customers Affected 1',
        'Area': 'Area Affected',
        ' NERC Region': 'NERC Region'
    }.items():
        if old_column in sheet_data.columns:
            sheet_data = sheet_data.rename(columns={old_column: new_column})

   # Store the original data before filtering
    original_data = sheet_data.copy()

    # Filter out all empty rows
    sheet_data = sheet_data.dropna(how='all')

    # Filter out rows where date event began is NA
    sheet_data = sheet_data[~pd.isna(sheet_data['Date Event Began'])]

   # Filter out rows with values to exclude using from date event began column
    sheet_data = sheet_data[sheet_data['Date Event Began'].apply(
        lambda x: all(val.lower() not in str(x).lower() for val in values_to_exclude))]

    # Calculate what rows are filtered out and store it in data_removed dataframe
    removed_rows = original_data.merge(
        sheet_data, how='outer', indicator=True).loc[lambda x: x['_merge'] == 'left_only'].drop('_merge', axis=1)

    data_removed = pd.concat([data_removed, removed_rows], ignore_index=True)

    # Handle the sheets where the restoration date and time are in single column and split the column
    if 'Restoration Time' in sheet_data.columns:

        # Clean rows with string and no date and time values to pd.NaT
        sheet_data['Restoration Time'] = sheet_data['Restoration Time'].apply(
            replace_values)

        # Parse the datetime column to datetime dtype and check for errors
        sheet_data['Restoration Time'] = sheet_data['Restoration Time'].astype(
            str).apply(parse_datetime_string, show_errors=True)

        # Split columns if all errors are solved
        sheet_data[['Date of Restoration', 'Time of Restoration']] = sheet_data['Restoration Time'].apply(
            lambda x: pd.Series(str(x).split(
                ' ', 1) if pd.notna(x) else [None, None])
        )

    # Parse Date and time columns

    # Clean rows with string and no date and time values to pd.NaT
    sheet_data['Date of Restoration'] = sheet_data['Date of Restoration'].apply(
        replace_values)
    sheet_data['Time of Restoration'] = sheet_data['Time of Restoration'].apply(
        replace_values)

    # Parse date and time columns
    sheet_data["Date Event Began"] = sheet_data["Date Event Began"].astype(
        str).apply(parse_date_string, show_errors=True)

    sheet_data["Time Event Began"] = sheet_data["Time Event Began"].astype(
        str).apply(parse_time_string, show_errors=True)

    sheet_data["Date of Restoration"] = sheet_data["Date of Restoration"].astype(
        str).apply(parse_date_string, show_errors=True)

    sheet_data['Time of Restoration'] = sheet_data['Time of Restoration'].astype(
        str).apply(parse_time_string, show_errors=True)

    sheet_data["Date of Restoration"] = pd.to_datetime(
        sheet_data["Date of Restoration"])

    # Concatenate the modified sheet data to the data dataframe
    data = pd.concat([data, sheet_data], ignore_index=True)


# Filter out unnecesary columns
data = data[['Date Event Began', 'Time Event Began', 'Date of Restoration',	'Time of Restoration',	'Area Affected',
             'NERC Region',	'Alert Criteria', 'Event Type', 'Demand Loss (MW)', 'Number of Customers Affected']]

# Sorting dataset and resetting index
data.sort_values(by='Date Event Began', ascending=True, inplace=True)
data.reset_index(drop=True, inplace=True)

# Storing data from >= 2015
data_from_2015 = data[data['Date Event Began'] >= '01-01-2015'].copy()

# Cleaning NERC Region column

data_from_2015["NERC Region"] = data_from_2015["NERC Region"].str.strip()

nerc_values_replace = {
    "RF": "RFC",
    "RE": "TRE",
    "FRCC": "SERC",
    "SPP RE": "SPP/TRE",
    "RF/SERC": "RFC/SERC",
    "SERC/RF":  "RFC/SERC",
    "SERC / RF": "RFC/SERC",
    "SERC/MRO": "MRO/SERC",
    "MRO/RF": "MRO/RFC",
    "MRO / RF": "MRO/RFC",
    "WECC/MRO": "MRO/WECC",
    "RF/MRO": "MRO/RFC",
    "SPP, SERC, TRE": "SPP/SERC/TRE",
    "WECC/SERC": "SERC/WECC"
}

data_from_2015["NERC Region"] = data_from_2015["NERC Region"].replace(
    nerc_values_replace)

# Add HI and PR if NERC Region is NA
data_from_2015.loc[data_from_2015["NERC Region"].isna() & data_from_2015["Area Affected"].str.lower(
).str.contains("puerto rico", case=False, na=False), "NERC Region"] = "PR"
data_from_2015.loc[data_from_2015["NERC Region"].isna() & data_from_2015["Area Affected"].str.lower(
).str.contains("hawaii", case=False, na=False), "NERC Region"] = "HI"


# Parsing Alert Criteria and Event Type columns into seperate ID columns

# Report Type ID
data_from_2015["Report Type ID"] = data_from_2015["Alert Criteria"].apply(
    report_type_id)

# Emergency Cause ID
data_from_2015["Cause ID"] = data_from_2015["Event Type"].apply(
    emergency_cause_id)

# Emergency Impact ID
data_from_2015["Impact ID"] = data_from_2015["Event Type"].apply(
    emergency_impact_id)
data_from_2015['Impact ID'] = data_from_2015.apply(update_impact_id, axis=1)

# Emergency Action ID
data_from_2015["Action ID"] = data_from_2015["Event Type"].apply(
    emergency_action_id)
data_from_2015['Action ID'] = data_from_2015.apply(update_action_id, axis=1)

# Replacing Unknown and NA values in Demand Loss (MW) and Number of Customers Affected
data_from_2015['Demand Loss (MW)'] = data_from_2015['Demand Loss (MW)'].replace(
    'Unknown', 0)
data_from_2015['Number of Customers Affected'] = data_from_2015['Number of Customers Affected'].replace(
    'Unknown', 0)
data_from_2015['Number of Customers Affected'] = data_from_2015['Number of Customers Affected'].replace(
    pd.NA, 0)

# Adding column Outage Time in minutes
data_from_2015.loc[2525, 'Date of Restoration'] = pd.to_datetime('2019-08-18')
data_from_2015['Outage Time'] = (data_from_2015['Date of Restoration'] + pd.to_timedelta(data_from_2015['Time of Restoration'])
                                 ) - (data_from_2015['Date Event Began'] + pd.to_timedelta(data_from_2015['Time Event Began']))
data_from_2015['Outage Time'] = data_from_2015['Outage Time'].dt.total_seconds() / 60

# Removing outliers and duplicates
data_from_2015[data_from_2015['Outage Time'] < 0]

# Filter duplicate rows
indices_to_filter = [1691, 2037, 2062, 2063, 2134, 2147, 2409, 2575, 2590, 2602, 2673, 3051, 3069,
                     3070, 3096, 3227, 3299, 3312, 3377, 3377, 3378, 3473, 3475, 3507, 3518, 3601, 3624, 3640, 3858, 3859]

data_from_2015 = data_from_2015.drop(indices_to_filter)
data_from_2015.loc[2039, 'NERC Region'] = "SERC/TRE"

# Sort Data and add Event ID Column
data_from_2015.sort_values(by=['Date Event Began', 'Time Event Began'], inplace=True)

data_from_2015.reset_index(inplace=True, drop=True)
data_from_2015.index += 1 
data_from_2015.reset_index(drop=False,inplace=True)
data_from_2015.rename(columns={'index': 'Event ID'}, inplace=True)        




data_from_2015
