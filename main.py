import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import glob
from datetime import datetime, time, timedelta

# function to combine files, a new file named new.csv would be created.
folder_path = "./partner_files/*.xlsx"
files = []
files.extend(glob.glob(folder_path))
combined_dataframes = []

if len(files) > 0:
    for file in files:
        print(file)
        data = pd.read_excel(file)
        combined_dataframes.append(data)

combined_data = pd.concat(combined_dataframes, ignore_index=True)
combined_data.to_csv("new.csv", index=False)

mif = pd.read_excel('mif.xlsx')

# functions.
def time_difference(call_time, arrival_time):
    day_diff = abs((arrival_time.date() - call_time.date()).days)

    if day_diff == 0:
        hours_diff = (arrival_time - call_time).total_seconds() / 3600
        if call_time.hour < 9:
            hours_before_9_am = 9 - call_time.hour
            hours_diff -= hours_before_9_am
        if arrival_time.hour >= 18:
            hours_after_18 = arrival_time.hour - 18
            hours_diff -= hours_after_18
    else:
        non_working_hours = 15 * day_diff
        hours_diff = (arrival_time - call_time).total_seconds() / 3600
        hours_diff = hours_diff - non_working_hours
        if call_time.hour < 9:
            call_to_9_am = (min(call_time.replace(hour=9), arrival_time) - call_time).total_seconds() / 3600
            hours_diff = hours_diff - call_to_9_am
        if arrival_time.hour >= 18:
            after_6_pm = (arrival_time - max(arrival_time.replace(hour=18), call_time)).total_seconds() / 3600
            hours_diff -= after_6_pm
        if call_time.hour >= 18:
            hours_after_6_pm = call_time.hour - 18
            hours_diff += hours_after_6_pm

    
    return hours_diff

def format_hours_decimal(hours_decimal):
    # Split the hours_decimal into whole hours and remaining minutes
    hours, remainder = divmod(hours_decimal, 1)
    minutes = remainder * 60

    # Format the result as HH:MM
    formatted_time = "{:02d}:{:02d}".format(int(hours), int(minutes))
    return formatted_time

# Making the partner dashboard..
all_partners = np.unique(combined_data['Partner Name'])
partner_info_list = []

for partner in all_partners:
    partner_info = {}
    rts = [] # for response times.
    service_times = [] # for service times.
    down_times = [] # for down times = completion time - call time.
    FTF = 0

    # partner information.
    partner_data = data[data['Partner Name'] == partner]

    # partner calls.
    total_calls = len(partner_data)

    # partner response time.
    for i in range(total_calls):
        call_date = partner_data['Call Date'].iloc[i]
        arrival_date = partner_data['Arrival Date'].iloc[i]
        rt = time_difference(call_date, arrival_date)
        rts.append(rt)
    
    rts = pd.Series(rts)

    # average service time.
    for i in range(total_calls):
        arrival_date = partner_data['Arrival Date'].iloc[i]
        completion_date = partner_data[' Completion Date'].iloc[i]
        service_time = (completion_date - arrival_date).total_seconds() / 3600
        service_times.append(service_time)

    service_times = pd.Series(service_times)

    # average down time.
    for i in range(total_calls):
        call_date = partner_data['Call Date'].iloc[i]
        completion_date = partner_data[' Completion Date'].iloc[i]
        downtime = time_difference(call_date, completion_date)
        down_times.append(downtime)
        remote_resolution_calls = len(partner_data[partner_data['Call Resolution Type'] == 'Remote Resolution'])
        # FTF = arrival date and completion date are the same.
        if abs(arrival_date.date() == completion_date.date()):
            FTF = FTF + 1

    down_times = pd.Series(down_times)

    # average productivty
    on_board_engineers = np.unique(partner_data['Engineer Name'])
    engineer_calls = []
    for engineer in on_board_engineers:
        total_calls_by_engineer = len(partner_data[partner_data['Engineer Name'] == engineer])
        engineer_calls.append(total_calls_by_engineer)
    
    engineer_calls = pd.Series(engineer_calls)

    # Service Coverage Ratio
    total_mif = mif['MIF'][mif['Partner Name'] == partner].item()

    # C%M Ratio
    total_CM_calls = len(partner_data['Contract Status'] == 'C&M')
    total_cm_mif = mif['C&M'][mif['Partner Name'] == partner].item()

    # CM, PM Ratio
    total_maintainance_calls = len(partner_data[partner_data['Call Type'] == 'CM'])
    total_pm_calls = len(partner_data[partner_data['Call Type'] == 'PM'])
    total_cc_calls = len(partner_data[partner_data['Call Type'] == 'CC'])

    partner_info["Partner Name"] = partner
    partner_info["Total Calls"] = total_calls
    partner_info["Average Response Time (Hours)"] = format_hours_decimal(rts.mean())
    partner_info['Average Service Time (Hours)'] = format_hours_decimal(service_times.mean())
    partner_info['Average Downtime (Hours)'] = format_hours_decimal(down_times.mean())
    partner_info['Remote Resolution Ratio'] = remote_resolution_calls / total_calls
    partner_info['Per Day Productivity'] = (total_calls / len(np.unique(partner_data['Engineer Name']))) / 25
    partner_info['Service Coverage Ratio %'] = (total_calls / total_mif) * 100
    partner_info['On Board Engineer'] = len(np.unique(partner_data['Engineer Name']))
    partner_info['Calls Per Day (Mean) %'] = (engineer_calls.mean() / 25)
    partner_info['Calls Per Day (Median) %'] = (engineer_calls.median() / 25)
    partner_info['C&M Connect Ratio %'] = (total_CM_calls / total_cm_mif) * 100
    partner_info['CM Ratio'] = total_maintainance_calls / total_calls
    partner_info['PM Ratio'] = total_pm_calls / total_calls
    partner_info['FTS Percentage'] = (FTF / total_maintainance_calls) * 100
    partner_info['CC Ratio'] = total_cc_calls / total_calls
    partner_info_list.append(partner_info)

partner_dashboard = pd.DataFrame(partner_info_list)
partner_dashboard.to_csv('partner_dashboard.csv', index=False)
