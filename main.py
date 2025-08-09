from selenium import webdriver
from selenium.common import TimeoutException
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
import os
from dotenv import load_dotenv
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from datetime import date, datetime
import pytz

# Load environment variables from .env file
load_dotenv()

# Define scopes (read/write access)
SCOPES = ['https://www.googleapis.com/auth/calendar.events']


def authenticate():
    flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
    creds = flow.run_local_server(port=0)  # Opens browser for auth
    return creds


def get_credentials(option):
    if option == 1:
        return os.getenv('JESUS_EMAIL'), os.getenv('JESUS_PASSWORD')
    elif option == 2:
        return os.getenv('ENRIQUE_EMAIL'), os.getenv('ENRIQUE_PASSWORD')
    elif option == 3:
        # Return both sets of credentials
        return (
            (os.getenv('JESUS_EMAIL'), os.getenv('JESUS_PASSWORD')),
            (os.getenv('ENRIQUE_EMAIL'), os.getenv('ENRIQUE_PASSWORD'))
        )


def acc_log_out(driver):
    try:
        acc_btn = WebDriverWait(driver, 2).until(
            EC.element_to_be_clickable((By.ID, "ChangeUsernameId1"))
        )
        acc_btn.click()

        logout_link = WebDriverWait(driver, 2).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "#logoutForm a"))
        )
        logout_link.click()

    except TimeoutException:
        print("Failed to Log Out!")


def get_table(email, password) -> (list[str], list[list[str]]):
    # Launch browser
    driver = webdriver.Chrome()
    driver.get("https://acumen.dcisoftware.com/")

    # Fill credentials and click submit
    driver.find_element(By.ID, "Email").send_keys(email)
    driver.find_element(By.ID, "Password").send_keys(password)
    driver.find_element(By.ID, "btnSubmit").click()

    # Checks for confirm continue login popup
    try:
        WebDriverWait(driver, 1).until(
            lambda d: d.find_element(By.ID, "confirmContinueLogin").is_displayed()
        )
        driver.find_element(By.ID, "btnContnueLogin").click()

    except TimeoutException:
        print("Login Ok?")

    try:
        # Wait for home page to load to find Entries Button
        entries_btn = WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable((By.ID, "leftmenuLinkEmployerPunches"))
        )
        entries_btn.click()
        print("Login is Ok.")
    except TimeoutException:
        print("We have failed to load into the home page. Did we sign in correctly?")
        return  # Terminate Script

    try:
        # Wait for home page to load to find Entries Button
        loadMoreBtn = WebDriverWait(driver, 1).until(
            EC.element_to_be_clickable((By.ID, "btnLoadmore"))
        )
        loadMoreBtn.click()
    except TimeoutException:
        print("No more new entries loaded.")

    # Extract table's header from thead
    table_thead = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, "#tblPunches thead"))
    )
    table_ths = table_thead.find_elements(By.CSS_SELECTOR, "tr th")
    header = [header.text for header in table_ths]

    # Extract table's data from tbody
    punches_table = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, "#tblPunches tbody"))
    )
    rows = punches_table.find_elements(By.TAG_NAME, "tr")
    data = []
    for row in rows:
        cells = row.find_elements(By.TAG_NAME, "td")
        data.append([cell.text for cell in cells])

    # Sign out and quit driver
    acc_log_out(driver)
    driver.quit()

    return header, data


def to_dataframe(table_header, table_data) -> pd.DataFrame:
    df = pd.DataFrame(data=table_data, columns=table_header)
    df = df.set_index('Id')  # Set Id column as the index

    # Get rid of any currently Opened or Rejected Services
    df = df.drop(df[(df['Status'] == 'Open') | (df['Status'] == 'Rejected')].index, axis='index')

    # Service is of type int
    df['Service Code'] = df['Service Code'].astype(int)

    # Create Hours and Minutes columns
    df['Hours'] = df['Amount'].apply(lambda x: x.split(':')[1] if x else None)
    df['Minutes'] = df['Amount'].apply(lambda x: x.split(':')[2] if x else None)

    # Combine 'Hours' and 'Minutes' into an 'HH:MM:00' format
    df['Amount'] = df['Hours'].astype(str) + ':' + df['Minutes'].astype(str) + ':00'

    # Convert to timedelta
    df['Amount'] = pd.to_timedelta(df['Amount'])

    # Converting to datetime
    # Define datetime formats explicitly
    date_format = '%b %d, %Y'  # e.g. "Jul 28, 2025"
    time_format = '%I:%M %p'  # e.g. "03:38 PM"

    # Convert to datetime with explicit formats
    df['Start Time'] = pd.to_datetime(
        df['Service Date'] + ' ' + df['Start Time'],
        format=f'{date_format} {time_format}'
    )
    df['End Time'] = pd.to_datetime(
        df['Service Date'] + ' ' + df['End Time'],
        format=f'{date_format} {time_format}'
    )

    # Convert Service Date to pure date (no time component)
    df['Service Date'] = pd.to_datetime(df['Service Date'],
                                        format=date_format).dt.normalize()  # Normalize since we don't care about time

    # Drops unwanted columns
    keep_columns = ['Service Date', 'Start Time', 'End Time', 'Amount', 'Service Code', 'Status']
    df = df[keep_columns]

    return df


def format_hhmmss(td) -> str:
    total_seconds = int(td.total_seconds())
    hours = total_seconds // 3600
    minutes = (total_seconds % 3600) // 60
    seconds = total_seconds % 60
    return f"{hours:02d}:{minutes:02d}:{seconds:02d}"


def get_month_data(df, target_date):
    """
    Get data for the month of the specified date (defaults to current month if not specified)

    Args:
        df: DataFrame containing the data
        target_date: Date object or string representing the target month (defaults to today)

    Returns:
        Filtered DataFrame for the target month
    """
    month_start = pd.to_datetime(target_date.replace(day=1)).normalize()
    next_month_start = (month_start + pd.offsets.MonthBegin(1)).normalize()

    return df[(df['Service Date'] >= month_start) &
              (df['Service Date'] < next_month_start)]


def get_biweekly_data(df, target_date):
    """
    Get data split into two biweekly periods for the specified month
    (defaults to current month if not specified)

    Args:
        df: DataFrame containing the data
        target_date: Date object or string representing the target month (defaults to today)

    Returns:
        Tuple of two DataFrames (first biweekly, second biweekly)
    """
    month_start = pd.to_datetime(target_date.replace(day=1)).normalize()
    next_month_start = (month_start + pd.offsets.MonthBegin(1)).normalize()

    # First biweekly period (1st-15th)
    first_end = month_start.replace(day=15)

    # Second biweekly period (16th-end of month)
    second_start = month_start.replace(day=16)
    second_end = next_month_start - pd.Timedelta(days=1)

    # Filter for the entire month first
    month_df = df[(df['Service Date'] >= month_start) &
                  (df['Service Date'] < next_month_start)]

    first_biweekly = month_df[month_df['Service Date'] <= first_end]
    second_biweekly = month_df[month_df['Service Date'] >= second_start]

    return first_biweekly, second_biweekly


def calculate_hours(df, target_date):
    df = get_month_data(df, target_date)
    first_biweekly_df, second_biweekly_df = get_biweekly_data(df, target_date)

    first_biweekly_hours = {
        '331': first_biweekly_df[first_biweekly_df['Service Code'] == 331]['Amount'].sum(),
        '320': first_biweekly_df[first_biweekly_df['Service Code'] == 320]['Amount'].sum(),
        '310': first_biweekly_df[first_biweekly_df['Service Code'] == 310]['Amount'].sum(),
        'Total Hours': first_biweekly_df['Amount'].sum()
    }

    second_biweekly_hours = {
        '331': second_biweekly_df[second_biweekly_df['Service Code'] == 331]['Amount'].sum(),
        '320': second_biweekly_df[second_biweekly_df['Service Code'] == 320]['Amount'].sum(),
        '310': second_biweekly_df[second_biweekly_df['Service Code'] == 310]['Amount'].sum(),
        'Total Hours': second_biweekly_df['Amount'].sum()
    }

    month_hours = {
        '331': df[df['Service Code'] == 331]['Amount'].sum(),
        '320': df[df['Service Code'] == 320]['Amount'].sum(),
        '310': df[df['Service Code'] == 310]['Amount'].sum(),
        'Total Hours': df['Amount'].sum()
    }

    # Convert to HH:MM:SS strings
    first_biweekly_formatted = {
        key: format_hhmmss(td)  # Takes the last part (HH:MM:SS)
        for key, td in first_biweekly_hours.items()
    }
    second_biweekly_formatted = {
        key: format_hhmmss(td)  # Takes the last part (HH:MM:SS)
        for key, td in second_biweekly_hours.items()
    }
    month_formatted = {
        key: format_hhmmss(td)  # Takes the last part (HH:MM:SS)
        for key, td in month_hours.items()
    }

    return first_biweekly_formatted, second_biweekly_formatted, month_formatted


def print_hours(j_df, e_df, target_date):
    j_first_biweekly_formatted, j_second_biweekly_formatted, j_month_formatted = calculate_hours(j_df, target_date)
    e_first_biweekly_formatted, e_second_biweekly_formatted, e_month_formatted = calculate_hours(e_df, target_date)
    print(f"Jesus' First Biweekly Hours:  {j_first_biweekly_formatted}")
    print(f"Jesus' Second Biweekly Hours: {j_second_biweekly_formatted}")
    print(f"Jesus' Month Hours:           {j_month_formatted}", end='\n\n')
    print(f"Enrique's First Biweekly Hours:  {e_first_biweekly_formatted}")
    print(f"Enrique's Second Biweekly Hours: {e_second_biweekly_formatted}")
    print(f"Enrique's Month Hours:           {e_month_formatted}")


def event_exists(service, start_time: datetime, end_time: datetime):
    # Localizes timezone to calendar's timezone
    la_tz = pytz.timezone('America/Los_Angeles')
    time_min = la_tz.localize(start_time)
    time_max = la_tz.localize(end_time)

    # Searches for events within time range interval -> (timeMin, timeMax)
    events_result = service.events().list(
        calendarId=os.getenv('CALENDAR_ID'),
        timeMin=time_min.isoformat(),  # Must be RFC3339 formatted
        timeMax=time_max.isoformat(),  # Must be RFC3339 formatted
        singleEvents=True,
        orderBy='startTime'
    ).execute()

    # False if there are no entries within time range
    return len(events_result.get('items', [])) > 0


def add_calendar_event(service, summary, start_time, end_time, color_id):
    event = {
        'summary': summary,
        'start': {
            'dateTime': start_time.isoformat(),
            'timeZone': 'America/Los_Angeles'
        },
        'end': {
            'dateTime': end_time.isoformat(),
            'timeZone': 'America/Los_Angeles'
        },
        'colorId': color_id
    }

    event = service.events().insert(
        calendarId=os.getenv('CALENDAR_ID'),
        body=event
    ).execute()


def process_punch_data(df, target_date, employee, service):
    EMPLOYEE_COLORS = {
        "Jesus": 2,  # Green
        "Enrique": 9  # Blue
    }

    df = get_month_data(df, target_date)

    # Skip if event exists
    for _, row in df.iterrows():
        if event_exists(service, row['Start Time'], row['End Time']):
            continue

        add_calendar_event(
            service=service,
            summary=f"{employee} ({row['Service Code']}) {format_hhmmss(row['Amount'])[:-3]}hrs",
            start_time=row['Start Time'],
            end_time=row['End Time'],
            color_id=EMPLOYEE_COLORS.get(employee, 1)
        )


def main():
    input_msg = '''Who do you want to get entries from?
1. Jesus Donate
2. Enrique Donate
3. Both
'''
    option = int(input(input_msg))

    # Ask for month selection
    month_msg = '''Select month option:
1. Current month
2. Specific month (YYYY-MM)
'''
    month_option = int(input(month_msg))

    if month_option == 1:  # Current month
        target_date = date.today()
    elif month_option == 2:  # Specific month
        month_input = input("Enter month (YYYY-MM): ")
        target_date = pd.to_datetime(month_input).date()
    else:
        print("Invalid month option")
        return

    # Google Authentication
    creds = authenticate()
    service = build('calendar', 'v3', credentials=creds)

    if option in {1, 2}:
        email, password = get_credentials(option)
        table_header, table_data = get_table(email=email, password=password)
        df = to_dataframe(table_header, table_data)
        first_biweekly_formatted, second_biweekly_formatted, month_formatted = calculate_hours(df, target_date)
        print(f'\n\n ** Date: {target_date.year}/{target_date.month} **')
        print(f'First Biweekly Hours:  {first_biweekly_formatted}')
        print(f'Second Biweekly Hours: {second_biweekly_formatted}')
        print(f'Month Hours:           {month_formatted}')

        # Adds entries to Acumen Punch-In Times Google Calendar
        if option == 1:
            process_punch_data(df, target_date, 'Jesus', service)
        else:
            process_punch_data(df, target_date, 'Enrique', service)

    elif option == 3:
        # Handle both cases
        jesus_creds, enrique_creds = get_credentials(option)
        # Process Jesus
        j_table_header, j_table_data = get_table(*jesus_creds)
        j_df = to_dataframe(j_table_header, j_table_data)
        # Process Enrique
        e_table_header, e_table_data = get_table(*enrique_creds)
        e_df = to_dataframe(e_table_header, e_table_data)

        print_hours(j_df, e_df, target_date)

        # Adds entries to Acumen Punch-In Times Google Calendar
        process_punch_data(j_df, target_date, 'Jesus', service)
        process_punch_data(e_df, target_date, 'Enrique', service)

    else:
        print("Invalid option")


if __name__ == '__main__':
    main()
