import pandas as pd
import math
from openpyxl import load_workbook
from excel_file_address import excel_file_path
# Read excel file

df = pd.read_excel(excel_file_path, sheet_name='summary')

# List of months, years, groups exist in excel file

print('Prepare the list of moth, year, groups...', end = " ")
months = []
for i in df.index:
    months.append(df['Month'][i])
months = list(set(months))  # remove duplicate

years = []
for i in df.index:
    years.append(df['Year'][i])
years = list(set(years))  # remove duplicate

time = {}
for i in df.index:
    month = df['Month'][i]
    year = df['Year'][i]
    try:
        time[year].append(month)
    except KeyError:
        time[year] = []
        time[year].append(month)

for year in years:
    time[year] = list(set(time[year]))

# List of group exist in excel file
groups = []
for i in df.index:
    groups.append(df['Group'][i])
groups = list(set(groups))  # remove duplicate
print('Done')

# Create dictionary with format {'Group name': {employee id : employee name}}
print('Create dictionary of Employee with group,id,name...', end=' ')
employee = {}
for i in df.index:
    name = df['Name'][i]
    group_name = df['Group'][i]
    id = df['Enum'][i]
    try:
        employee[group_name]
    except KeyError:
        employee[group_name] = {}
    employee[group_name][id] = name

# List of columns
output_group_list = []
output_name_list = []
output_id_list = []

# Create 3 list group -> id -> name from the dictionary upon
for group in groups:
    for emp_id in employee[group]:
        output_group_list.append(group)
        output_id_list.append(emp_id)
        output_name_list.append(employee[group][emp_id])
print('Done')

print('Create dictionary of Employee Prime Effort with group, id, month, prime effort...', end=' ')
prime_effort = {}
for i in df.index:
    month = df['Month'][i]
    group_name = df['Group'][i]
    id = df['Enum'][i]
    try:
        prime_effort[group_name]
    except KeyError:
        prime_effort[group_name] = {}
    try:
        prime_effort[group_name][id]
    except KeyError:
        prime_effort[group_name][id] = {}
    prime_effort[group_name][id][month] = 0  # set the effort of all employee initially = 0


# Parsing effort to the prime_effort dictionary format {'id': {month : effort}}
for i in df.index:
    group_name = df['Group'][i]
    effort = df['Project Timesheet Effort'][i]
    id = df['Enum'][i]
    month = df['Month'][i]
    try:
        if id in output_id_list and month in months:
            prime_effort[group_name][id][month] += effort
    except KeyError:
            print('Error at line: ', i)
print('Done')

print('Create dictionary of Employee non Prime Effort with group, id, month, prime effort...', end=' ')
non_prime_effort = {}
for i in df.index:
    month = df['Month'][i]
    group_name = df['Group'][i]
    id = df['Enum'][i]
    try:
        non_prime_effort[group_name]
    except KeyError:
        non_prime_effort[group_name] = {}
    try:
        non_prime_effort[group_name][id]
    except KeyError:
        non_prime_effort[group_name][id] = {}
    non_prime_effort[group_name][id][month] = 0


# Parsing effort to the prime_effort dictionary format {'id': {month : effort}}
for i in df.index:
    group_name = df['Group'][i]
    non_effort = df['Non Project Timesheet Effort'][i]
    id = df['Enum'][i]
    month = df['Month'][i]
    try:
        if id in output_id_list and month in months:
            if math.isnan(non_effort):
                non_effort = 0
            non_prime_effort[group_name][id][month] += non_effort
    except KeyError:
            print('Error at line: ', i)
print('Done')

print('Prepare the final data for writing in excel...', end=" ")
output_prime_effort = {}
for group in groups:
    for id in prime_effort[group].keys():
        for month in months:
            effort = prime_effort[group][id][month]
            try:
                output_prime_effort[month].append(effort)
            except KeyError:
                output_prime_effort[month] = []
                output_prime_effort[month].append(effort)

output_non_prime_effort = {}
for group in groups:
    for id in non_prime_effort[group].keys():
        for month in months:
            non_effort = non_prime_effort[group][id][month]
            try:
                output_non_prime_effort[month].append(non_effort)
            except KeyError:
                output_non_prime_effort[month] = []
                output_non_prime_effort[month].append(non_effort)

print('Done')

print("Write into excel file...", end=" ")
for month in months:
    df = pd.DataFrame({'Group': output_group_list,
                       'Id': output_id_list,
                       'Resource': output_name_list,
                       'Project Effort': output_prime_effort[month],
                       'Non Project Effort': output_non_prime_effort[month]
                        })

    # Create a Pandas Excel writer using XlsxWriter as the engine.
    book = load_workbook(excel_file_path)
    writer = pd.ExcelWriter(excel_file_path, engine='openpyxl')
    writer.book = book
    writer.sheets = dict((ws.title, ws) for ws in book.worksheets)
    # Convert the dataframe to an XlsxWriter Excel object.
    for year in years:
        if month in time[year]:
            sheet_name = str(month) + '-' + str(year)
            df.to_excel(writer, sheet_name, index=False)
    #
    # # Close the Pandas Excel writer and output the Excel file.
    writer.save()

print("Done")
