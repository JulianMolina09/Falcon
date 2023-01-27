import requests
import datetime
import pytz
import pandas as pd

# Create a session to handle the requests
session = requests.Session()
aps_key = "4438~2vKWjGpXWcWyM8n380l6hJTsHov081P0cE2ITbinoTDGXPuvEcbE39XraAZFMFIn"
nvcc_key = "13096~Wq8gUpaLcgpXs7DNySIWuHz4FlHDdlNKmgE4jTrwCDDGxcajoHXmBw91U9SQBgqT"

aps_url = "https://apsva.instructure.com/api/v1"
nvcc_url = "https://learn.vccs.edu/api/v1"

aps_ids = ['85633', '85762', '85157']
nvcc_ids = ['479410']

def format_date(date_str):
    date_time_obj = datetime.datetime.strptime(date_str, '%Y-%m-%dT%H:%M:%SZ') - datetime.timedelta(hours=5)
    est_timezone = pytz.timezone('EST')
    date_time_obj_est = est_timezone.localize(date_time_obj)
    return date_time_obj_est.strftime('%Y-%m-%d %I:%M:%S %p')

# Get the current date
# Get the current date
now = datetime.datetime.now()

# Get the start and end of the current week
week_start = now - datetime.timedelta(days=now.weekday())
week_end = week_start + datetime.timedelta(days=6)

# Create a dataframe to store the assignments
# Try to read the existing excel sheet
try:
    df = pd.read_excel("assignments.xlsx")
except:
    # If the sheet doesn't exist, create a new dataframe with the status column
    df = pd.DataFrame(columns=['Name', 'Due Date', 'Course', 'Status'])

# Iterate through each course id
for course in aps_ids + nvcc_ids:
    if course in aps_ids:
        api_url = aps_url
        key = aps_key
    else:
        api_url = nvcc_url
        key = nvcc_key
    # make API call to get the assignment
    response = session.get(f"{api_url}/courses/{course}/assignments?per_page=100", headers={"Authorization": f"Bearer {key}"})
    if response.status_code == 200:
        # parse the response
        assignments = response.json()
        for assignment in assignments:
            if assignment["due_at"]:
                date_time_str = str(assignment['due_at'])
                due_date = datetime.datetime.strptime(date_time_str, '%Y-%m-%dT%H:%M:%SZ')
                # check if the due date is within the current week
                if week_start <= due_date <= week_end:
                    formatted_date = format_date(date_time_str)
                    status = ""
                    if due_date < now:
                        status = "late"
                    else:
                        status = "on time"
                    # Add the assignment to the dataframe
                    if (course) == '479410':
                        (course) = "English"
                    elif (course) == '85762':
                        (course) = "History"
                    elif (course) == '85633':
                        (course) = "Math"
                    elif (course) == '85157':
                        (course) = "Physics"
                    df = df.append({'Name': assignment['name'], 'Due Date': formatted_date, 'Course': course, 'Status': status}, ignore_index=True)
df = df.drop_duplicates()
df = df.style.set_properties(**{'text-align': 'left', 'white-space': 'pre-wrap'})

# Export the dataframe to an Excel sheet
df.to_excel("assignments.xlsx", index=False)
