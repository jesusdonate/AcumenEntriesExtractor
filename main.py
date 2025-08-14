from selenium import webdriver
from selenium.common.exceptions import TimeoutException, WebDriverException
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
import pandas as pd
import os
import sys
from dotenv import load_dotenv
from google.oauth2 import service_account
from googleapiclient.discovery import build
from datetime import date, datetime
import pytz
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

# Load environment variables from .env file
load_dotenv()

SCOPES = ['https://www.googleapis.com/auth/calendar.events']
SERVICE_ACCOUNT_PATH = 'acumen-entries-extractor-d2640e073bc3.json'


def authenticate():
    creds = service_account.Credentials.from_service_account_file(
        SERVICE_ACCOUNT_PATH,
        scopes=['https://www.googleapis.com/auth/calendar']
    )
    return creds


def get_credentials():
    return (
        (os.getenv('JESUS_USERNAME'), os.getenv('JESUS_PASSWORD')),
        (os.getenv('ENRIQUE_USERNAME'), os.getenv('ENRIQUE_PASSWORD'))
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


def get_chrome_options():
    """Configure Chrome options for both Windows and Linux"""
    chrome_options = Options()

    # Platform-specific configurations
    if sys.platform == 'linux':
        chrome_options.add_argument("--headless")
        chrome_options.add_argument("--no-sandbox")
        chrome_options.add_argument("--disable-dev-shm-usage")
        chrome_options.add_argument("--window-size=1920,1080")

    # Common configurations
    chrome_options.add_argument("--log-level=3")
    chrome_options.add_experimental_option('excludeSwitches', ['enable-logging'])
    chrome_options.add_argument("--disable-background-networking")
    chrome_options.add_argument("--disable-default-apps")
    chrome_options.add_argument("--disable-extensions")

    return chrome_options


def initialize_driver():
    """Initialize WebDriver with platform-specific settings"""
    chrome_options = get_chrome_options()

    try:
        if sys.platform == 'linux':
            return webdriver.Chrome(
                service=Service("/usr/local/bin/chromedriver"),
                options=chrome_options
            )
        else:  # Windows/Mac
            return webdriver.Chrome(options=chrome_options)
    except WebDriverException as e:
        print(f"Failed to initialize WebDriver: {str(e)}")
        raise


def handle_login(driver, email, password):
    """Handle the login process with error checking"""
    try:
        driver.get("https://acumen.dcisoftware.com/")

        # Fill credentials
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "Email"))
        ).send_keys(email)

        driver.find_element(By.ID, "Password").send_keys(password)
        driver.find_element(By.ID, "btnSubmit").click()

        # Handle potential confirmation popup
        try:
            WebDriverWait(driver, 2).until(
                lambda d: d.find_element(By.ID, "confirmContinueLogin").is_displayed()
            )
            driver.find_element(By.ID, "btnContnueLogin").click()
        except TimeoutException:
            pass

        return True
    except Exception as e:
        print(f"Login failed: {str(e)}")
        return False


def get_table_data(driver):
    """Extract table headers and data"""
    try:
        # Wait for and click entries button
        WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, "leftmenuLinkEmployerPunches"))
        ).click()

        # Extract headers
        table_thead = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "#tblPunches thead"))
        )
        header = [header.text for header in table_thead.find_elements(By.CSS_SELECTOR, "tr th")]

        # Extract data
        punches_table = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "#tblPunches tbody"))
        )
        rows = punches_table.find_elements(By.TAG_NAME, "tr")
        data = [[cell.text for cell in row.find_elements(By.TAG_NAME, "td")] for row in rows]

        return header, data
    except TimeoutException:
        print("Failed to locate table elements")
        return None, None


def get_table(email: str, password: str) -> tuple[list[str], list[list[str]]]:
    """Main function to retrieve table data"""
    driver = None
    try:
        driver = initialize_driver()

        if not handle_login(driver, email, password):
            return [], []

        header, data = get_table_data(driver)
        if header is None or data is None:
            return [], []

        return header, data

    except Exception as e:
        print(f"Unexpected error: {str(e)}")
        return [], []
    finally:
        if driver:
            try:
                # Try to sign out if possible
                acc_log_out(driver)
            except:
                pass
            driver.quit()


def to_dataframe(table_header, table_data, emp_name) -> pd.DataFrame:
    df = pd.DataFrame(data=table_data, columns=table_header)
    df = df.set_index('Id')  # Set Id column as the index

    # Adds employee name for all rows
    df['Employee Name'] = emp_name

    # Get rid of any currently Opened or Rejected Services
    df = df.drop(df[(df['Status'] == 'Open') | (df['Status'] == 'Rejected') | (df['Status'] == 'Unvalidated')].index,
                 axis='index')

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

    # Convert Service Date to pure date
    df['Service Date'] = pd.to_datetime(df['Service Date'],
                                        format=date_format).dt.normalize()  # Normalize since we don't care about time

    # Drops unwanted columns
    keep_columns = ['Service Date', 'Start Time', 'End Time', 'Amount', 'Service Code', 'Status', 'Employee Name']
    df = df[keep_columns]

    return df


def save_new_entries(df: pd.DataFrame) -> pd.DataFrame:
    file = 'entries.csv'

    # If file doesn't exist, save the new DataFrame directly
    if not os.path.exists(file):
        df.to_csv(file)
        return df

    # Read existing data with proper type conversion
    old_df = pd.read_csv(
        file,
        dtype={'Id': 'int64', 'Amount': 'str'},
        parse_dates=['Service Date', 'Start Time', 'End Time']  # Convert to datetime
    )
    old_df['Amount'] = pd.to_timedelta(old_df['Amount'])
    old_df = old_df.set_index('Id')

    # Ensure new dataframe's index is integer type
    df.index = df.index.astype('int64')

    # Combine and remove Id duplicates
    new_df = pd.concat([old_df, df])
    new_df = new_df[~new_df.index.duplicated(keep='first')]

    # Sort and save
    new_df.sort_index(ascending=False, inplace=True)
    new_df.to_csv(file)
    return new_df


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


def calculate_hours(df, target_date) -> (dict, dict, dict):
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


def process_punch_data(df: pd.DataFrame, target_date, service):
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
            summary=f"{row['Employee Name']} ({row['Service Code']}) {format_hhmmss(row['Amount'])[:-3]}hrs",
            start_time=row['Start Time'],
            end_time=row['End Time'],
            color_id=EMPLOYEE_COLORS.get(row['Employee Name'], 1)
        )


def send_email(
        sender_email: str,
        sender_password: str,
        recipient_email: str,
        subject: str,
        message: str,
        smtp_server: str = "smtp.gmail.com",
        smtp_port: int = 587
):
    """Send a regular email using SMTP."""

    html_message = message.replace('\n', '<br>')  # Newlines

    # Create email with both plain and HTML versions
    msg = MIMEMultipart('alternative')
    msg['From'] = sender_email
    msg['To'] = recipient_email
    msg['Subject'] = subject

    # Plain-text fallback
    msg.attach(MIMEText(message, 'plain'))

    # HTML version
    msg.attach(MIMEText(f"""
        <html>
          <body>
            <p>{html_message}</p>
          </body>
        </html>
        """, 'html'))

    # Send email
    try:
        with smtplib.SMTP(smtp_server, smtp_port) as server:
            server.starttls()
            server.login(sender_email, sender_password)
            server.sendmail(sender_email, recipient_email, msg.as_string())
    except Exception as e:
        print(f"Failed to send email: {str(e)}")


def email_employee(df: pd.DataFrame, employee_name: str, target_date: datetime.date):
    recipient_email = os.getenv(f'{employee_name.strip().upper()}_EMAIL')

    if recipient_email is None:
        print('Failed to email employee. Employee\'s not in .env file.')
        return

    first_biweekly_formatted, second_biweekly_formatted, month_formatted = calculate_hours(df, target_date)
    subject = f"{target_date.strftime('%m/%d/%Y')} - Acumen horas de {employee_name}"
    message = f"""<strong>Primera Quincena:</strong>
331 -> {first_biweekly_formatted['331'][:-3]}hrs
320 -> {first_biweekly_formatted['320'][:-3]}hrs
310 -> {first_biweekly_formatted['310'][:-3]}hrs

<strong>Segunda Qunicena:</strong>
331 -> {second_biweekly_formatted['331'][:-3]}hrs
320 -> {second_biweekly_formatted['320'][:-3]}hrs
310 -> {second_biweekly_formatted['310'][:-3]}hrs

<strong>Total Horas del mes:</strong>
331 -> {month_formatted['331'][:-3]}hrs
320 -> {month_formatted['320'][:-3]}hrs
310 -> {month_formatted['310'][:-3]}hrs
"""

    send_email(
        sender_email=os.getenv('JESUS_EMAIL'),
        sender_password=os.getenv('JESUS_GMAIL_APP_PASSWORD'),
        recipient_email=recipient_email,
        subject=subject,
        message=message
    )


def main():
    target_date = date.today()

    # If Mac or Windows, ask user for month option
    if sys.platform != 'linux':
        month_msg = '''Select month option:
1. Current month
2. Specific month (YYYY-MM)
'''
        month_option = int(input(month_msg))

        if month_option == 1:  # Current month
            target_date = date.today()
        elif month_option == 2:  # Specific month
            while True:  # Keep asking until valid input
                month_input = input("Enter month (YYYY-MM): ")
                try:
                    target_date = pd.to_datetime(month_input).date()
                    break
                except ValueError:
                    print("Invalid format. Please enter in YYYY-MM format (e.g., 2025-06).")
        else:
            print("Invalid month option. Choose 1 (current) or 2 (specific).")
            return

    # Service Account Authentication
    creds = authenticate()
    service = build('calendar', 'v3', credentials=creds)

    # Handle both cases
    jesus_creds, enrique_creds = get_credentials()
    # Process Jesus
    j_table_header, j_table_data = get_table(*jesus_creds)
    j_df = to_dataframe(j_table_header, j_table_data, 'Jesus')

    # Process Enrique
    e_table_header, e_table_data = get_table(*enrique_creds)
    e_df = to_dataframe(e_table_header, e_table_data, 'Enrique')

    save_new_entries(j_df)
    df = save_new_entries(e_df)

    j_df = df[df['Employee Name'] == 'Jesus']
    e_df = df[df['Employee Name'] == 'Enrique']
    print_hours(j_df, e_df, target_date)

    if sys.platform == 'linux':  # Send emails when using EC2 instance only
        email_employee(j_df, 'Jesus', target_date)
        email_employee(e_df, 'Enrique', target_date)

    # Adds entries to Acumen Punch-In Times Google Calendar
    process_punch_data(df, target_date, service)


if __name__ == '__main__':
    main()