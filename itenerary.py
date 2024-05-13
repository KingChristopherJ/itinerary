from datetime import datetime, timedelta
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from dateutil.parser import parse
import os
from dotenv import load_dotenv
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import time
import re

# Setup Chrome options
chrome_options = Options()
chrome_options.add_argument("--disable-gpu")
chrome_options.add_argument("--ignore-certificate-errors")

# Initialize WebDriver with options
try:
    driver = webdriver.Chrome(options=chrome_options)
    print("WebDriver initialized successfully.")
except Exception as e:
    print(f"Failed to initialize WebDriver: {e}")
    exit(1)

# Load environment variables from .env file
load_dotenv()

# Email credentials
sender_email = os.getenv('EMAIL')
password = os.getenv('EMAIL_PASSWORD')

def remove_ordinal_suffix(date_str):
    return re.sub(r'(\d+)(st|nd|rd|th)', r'\1', date_str)

def format_appointment(appointment):
    formatted_str = (
        f"<li><strong>Student Name:</strong> {appointment['student_name']}<br>"
        f"<strong>Meeting Date:</strong> {appointment['meeting_date']}<br>"
        f"<strong>Coach Name:</strong> {appointment['coach_name']}<br>"
        f"<strong>Student Year:</strong> {appointment['student_year']}<br>"
        f"<strong>Majors/Minors:</strong> {appointment['majors_minors']}<br>"
        f"<strong>Coaching Notes:</strong> {appointment['coaching_notes']}</li>"
    )
    return formatted_str

def generate_handshake_url():
    today = datetime.now()
    day_after_tomorrow = today + timedelta(days=2)
    the_day_after_that = day_after_tomorrow + timedelta(days=1)

    start_str = day_after_tomorrow.strftime('%Y-%m-%dT05:00:00.000Z')
    end_str = the_day_after_that.strftime('%Y-%m-%dT04:59:59.999Z')

    url = f"https://vanderbilt.joinhandshake.com/edu/appointments?page=1&per_page=25&sort_direction=desc&sort_column=default&staff_members%5B%5D=41763197&include_past_appointments=true&status%5B%5D=approved&start_date_start={start_str}&start_date_end={end_str}"
    return url

driver.get(generate_handshake_url())
time.sleep(35)  # Adjust based on your actual page load times

appointments_data = []
appointment_links = set(link.get_attribute('href') for link in driver.find_elements(By.TAG_NAME, 'a') if 'https://vanderbilt.joinhandshake.com/edu/appointments/' in link.get_attribute('href') and link.get_attribute('href').split('/')[-1].isdigit())

coach_user_id = '41763197'

for appointment_link in appointment_links:
    driver.get(appointment_link)
    time.sleep(3)

    all_links = driver.find_elements(By.TAG_NAME, 'a')
    user_links = [link.get_attribute('href') for link in all_links if link.get_attribute('href') and 'https://vanderbilt.joinhandshake.com/edu/users/' in link.get_attribute('href') and 'edit' not in link.get_attribute('href') and 'null' not in link.get_attribute('href')]
    user_ids = set(link.split('/')[-1] for link in user_links if link.split('/')[-1].isdigit() and link.split('/')[-1] != coach_user_id)

    for user_id in user_ids:
        student_appointments_url = f"https://vanderbilt.joinhandshake.com/edu/appointments?students[]={user_id}&include_past_appointments=true"
        driver.get(student_appointments_url)
        time.sleep(3)

        try:
            xpath_expression = "//div[contains(@class, 'style__text___2ilXR') and contains(@class, 'style__small___1Nyai') and contains(@class, 'style__tight___RF4uH')]"
            student_name = driver.find_element(By.XPATH, xpath_expression).text
        except Exception:
            student_name = "Unknown Student"

        student_appointment_links = set(link.get_attribute('href') for link in driver.find_elements(By.TAG_NAME, 'a') if link.get_attribute('href') and 'https://vanderbilt.joinhandshake.com/edu/appointments/' in link.get_attribute('href') and link.get_attribute('href').split('/')[-1].isdigit())
        for appt_link in student_appointment_links:
            driver.get(appt_link)
            time.sleep(3)

            appointment_dict = {
                'student_name': student_name,
                'appointment_link': appt_link,
                'coaching_notes': "No notes added.",
                'coach_name': "Unknown Coach",
                'student_year': "Unknown Year",
                'majors_minors': "Unknown Majors/Minors"
            }

            try:
                meeting_date = driver.find_element(By.XPATH, "//h4[text()='When']/following-sibling::p").text
                appointment_date = parse(remove_ordinal_suffix(meeting_date.split(' at ')[0]), fuzzy=True)
                appointment_dict['meeting_date'] = meeting_date
                appointment_dict['appointment_status'] = "Future" if appointment_date > datetime.now() else "Completed"
            except Exception:
                appointment_dict['meeting_date'] = "Failed to extract"

            try:
                notes_xpath = "//p[@class='respect-newlines margin-bottom'][@data-bind='html: safe_content_html']"
                coaching_notes = driver.find_element(By.XPATH, notes_xpath).text
                appointment_dict['coaching_notes'] = coaching_notes if coaching_notes.strip() else "No notes added."
            except Exception:
                pass

            try:
                coach_xpath = "//h4[text()='Staff Member']/following-sibling::p//a"
                coach_name = driver.find_element(By.XPATH, coach_xpath).text
                appointment_dict['coach_name'] = coach_name
            except Exception:
                pass

            try:
                details_xpath = "//h4[text()='Student Details']/following-sibling::p[@class='text']"
                student_details = driver.find_element(By.XPATH, details_xpath).text
                details_list = student_details.split("\n")
                appointment_dict['student_year'] = details_list[0].strip()
                appointment_dict['majors_minors'] = "; ".join(details_list[1:])
            except Exception:
                pass

            appointments_data.append(appointment_dict)

today_date = datetime.now().date()
formatted_date = today_date.strftime('%Y-%m-%d')

print(f"Today's appointments for {formatted_date}:\n")
html_content = "<html><head></head><body><h1>Today's appointments for " + formatted_date + ":</h1><ul>"
todays_appointments = [appt for appt in appointments_data if datetime.strptime(remove_ordinal_suffix(appt['meeting_date'].split(' at ')[0]), "%A, %B %d %Y").date() == today_date]

for appointment in todays_appointments:
    html_content += format_appointment(appointment)

html_content += "</ul><h1>History of coaching appointments:</h1><ul>"

for appointment in appointments_data:
    appointment_date = datetime.strptime(remove_ordinal_suffix(appointment['meeting_date'].split(' at ')[0]), "%A, %B %d %Y")
    if appointment_date.date() < today_date:
        html_content += format_appointment(appointment)

html_content += "</ul></body></html>"

msg = MIMEMultipart()
msg['From'] = sender_email
msg['To'] = 'christopher.j.king@vanderbilt.edu'
msg['Subject'] = "Today's appointments"
msg.attach(MIMEText(html_content, 'html'))

server = smtplib.SMTP('smtp.office365.com', 587)
server.starttls()
server.login(sender_email, password)
server.sendmail(sender_email, 'christopher.j.king@vanderbilt.edu', msg.as_string())
server.quit()

driver.quit()

print("Script completed and email sent successfully.")
