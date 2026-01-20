# AcumenEntriesExtractor

# Acumen Work Hours Automation

## Overview
A Python automation solution that extracts work hour entries from Acumen DCI software, processes time data, 
and provides automated reporting for employee hours management. The system eliminates manual tracking and 
calculation of biweekly hours, replacing error-prone spreadsheets with reliable automation.

## Business Problem
Before the creation of this script, employees needed to keep track and manage their work hours by manually 
recording work entries using spreadsheets or paper-based systems, which lacked any automated calculation 
of biweekly totals. This manual process required time-consuming reconciliation of hours across different 
periods and was susceptible to human error in calculations, potentially leading to inaccurate payroll 
reporting. Furthermore, there was no centralized notification system to provide employees with timely 
hour summaries.


## Solutions

### **Automated Data Extraction**
- Daily retrieval of work entries from Acumen DCI
- Automatic handling of login sessions and data scraping

### **Time Processing**
- Biweekly hour calculations (1st-15th and 16th-end of month)
- Categorization by service codes (310, 320, 331)

### **Data Synchronization**
- MongoDB database integration for persistent storage
- Conflict resolution between new and existing entries
- Automatic cleanup of rejected/duplicate records

### **Google Calendar Integration**
- Automatic creation of calendar events for work shifts

### **Automated Notifications**
- Email summaries sent to employees provide biweekly and monthly hour summaries
- Customizable email templates with hour breakdowns


## Business Impact

### **Time Savings**
- **2+ hours monthly** saved per employee on manual tracking
- **100% reduction** in calculation errors
- **Instant access** to hour summaries vs. manual compilation

### **Process Improvements**
- Eliminated spreadsheet dependency
- Automated compliance with biweekly reporting requirements
- Centralized data management

### **Employee Benefits**
- Automated email notifications with clear hour breakdowns
- Calendar integration for shift visualization
- No manual data entry required
- Consistent, reliable reporting



## Installation

### **Packages/Tools Used**
- **Python 3.11** with Selenium for web automation
- **MongoDB Atlas** for cloud data storage
- **Google Calendar API** for event management
- **GitHub Actions** for scheduled execution
- **Chrome Headless** for browser automation
- **SMTP** for email notifications

### **Steps to Run Script**

1. Clone the repository:
```bash
git clone https://github.com/jesusdonate/AcumenEntriesExtractor.git
cd acumen-entries-extractor
```

2. Environment Configuration
Create a .env file in the project root with the following variables:

```
# Acumen DCI Login Credentials
JESUS_USERNAME=my_acumen_username
JESUS_PASSWORD=my_acumen_password
ENRIQUE_USERNAME=fathers_acumen_username
ENRIQUE_PASSWORD=fathers_acumen_password

# Employee Email Addresses (for hour reports)
JESUS_EMAIL=jesus@example.com
ENRIQUE_EMAIL=enrique@example.com

# Gmail Configuration for Email Notifications
JESUS_GMAIL_APP_PASSWORD=16_digit_gmail_app_password

# Google Calendar Integration
CALENDAR_ID=google_calendar_id_here

#MongoDB URI connection string to collection
MONGODB_URI="mongodb+srv://example"

```
**Important Note:** This is a personal automation script built specifically for family use. 
It was designed with the explicit purpose of automating work hour tracking for **only two individuals** 
(myself and my father). As such, the code contains hardcoded elements and assumptions that reflect this limited, 
personal scope.


3. Google Cloud API Setup
Required Google Cloud Files:
   * Service Account JSON Key:

      * Download from Google Cloud Console -> IAM & Admin -> Service Accounts

      * Place in project root directory

      * Change SERVICE_ACCOUNT_PATH (main.py) to the path of your Service Account JSON Key


**Google Cloud Configuration Steps:**
1. Enable Google Calendar API in Google Cloud Console

2. Create a Service Account with "Calendar Editor" permissions

3. Download the JSON credentials file

4. Share your Google Calendar with the service account email address

5. Locate your Calendar ID in Google Calendar settings


4. Gmail App Password Setup
   1. Enable 2-Factor Authentication on your Gmail account

   2. Generate an App Password:

      * Go to Google Account -> Security -> 2-Step Verification -> App passwords

      * Select "Mail" and "Other (Custom name)"

      * Name it "Acumen Extractor" or whatever

      * Use the generated 16-character password in your .env file


## Automated Daily Execution via GitHub Actions

The `.github/workflows/acumen.yml` file sets an automated workflow that executes the script daily. 
This workflow automatically runs at 11:56 PM Los Angeles time each day. Note that it requires an annual 
adjustment for daylight saving time. When triggered, the workflow sets up a complete execution environment 
that includes Python 3.11, the Chrome browser, and all required Python dependencies before running the main 
script.

To use this automated system, users must configure several GitHub secrets in their repository 
settings under "Actions secrets and variables". These secrets include Acumen login credentials for users, 
MongoDB connection strings, Google Calendar configuration details, email addresses for notifications, 
and the complete service account JSON key for Google API access. Once these secrets are properly configured, 
users can manually trigger the workflow for testing by navigating to the Actions tab, selecting the "Acumen Daily 
Automation" workflow, and clicking the "Run workflow" button to verify everything functions 
correctly before relying on the automated schedule.

**Personal Note:** Before switching to GitHub Actions, I had used an AWS EC2 instance that would run my program 
every 24 hours at midnight, which incurred monthly costs of approximately **$12**. 
While functional, this approach represented unnecessary ongoing expenses for a task that only 
required a few minutes of daily execution. When researching for free alternatives, I 
discovered GitHub Actions, a powerful CI/CD platform with generous free tier allowances. 
Learning that I could run scheduled workflows at no cost, I quickly switched to GitHub 
Actions.

**Migration Results:**
- **Monthly Cost**: $12 -> **$0** (100% reduction)
- **Resource Efficiency**: 24/7 server -> ~2 minutes/day execution
- **Reliability**: Comparable uptime to EC2
- **Learning Outcome**: Gained valuable experience with GitHub Actions workflows



### GitHub Secrets Setup

Before using GitHub Actions, configure these secrets in your repository:

1. **Go to your repository -> Settings -> Secrets and variables -> Actions**
2. **Click "New repository secret"** and add each of these:

| Secret Name                | Value to Paste                                              |
|----------------------------|-------------------------------------------------------------|
| `JESUS_USERNAME`           | Employee 1 Acumen username                                  | 
| `JESUS_PASSWORD`           | Employee 1 Acumen password                                  | 
| `ENRIQUE_USERNAME`         | Employee 2 Acumen username                                  |
| `ENRIQUE_PASSWORD`         | Employee 2 Acumen password                                  |
| `MONGODB_URI`              | MongoDB connection string                                   |
| `CALENDAR_ID`              | Google Calendar ID                                          | 
| `JESUS_EMAIL`              | Email address                                               |
| `JESUS_GMAIL_APP_PASSWORD` | 16-character Gmail app password                             |
| `ENRIQUE_EMAIL`            | Employee 2 email address                                    |
| `SERVICE_ACCOUNT_JSON`     | **Entire content** of your Google Service Account JSON file |

### Important: Service Account JSON Setup

For the `SERVICE_ACCOUNT_JSON` secret:
1. Open your service account `.json` file in a text editor
2. **Copy the entire JSON content** (including all `{}` brackets)
3. Paste it as the secret value


