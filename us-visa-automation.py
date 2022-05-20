from time import sleep
from datetime import datetime
from selenium.webdriver import Chrome
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options as ChromeOptions
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import Select
from email.mime.text import MIMEText

import smtplib
import constants
import sys
import logging


def setup_logging():
    now = datetime.now()
    formatter = logging.Formatter('%(asctime)s | %(levelname)s | %(message)s', '%m-%d-%Y %H:%M:%S')
    logger = logging.getLogger("US VISA SCHEDULE CHECK")
    logger.setLevel(logging.INFO)

    file_handler = logging.FileHandler(".\\%s.log" % now.strftime("%Y%m%d_%H"))
    file_handler.setLevel(logging.INFO)
    file_handler.setFormatter(formatter)

    logger.addHandler(file_handler)
    return logger


def do_send_error_email(error: Exception, logger):
    sent_from = constants.GMAIL_USERNAME
    sent_to = ["<list_of_emails_goes_here>"]
    sent_subject = "Schedule Check Script Error"
    sent_body = ("<your_name_goes_here>,\n\n"
                "An error has occurred in the schedule check script: \n\n"
                "%s \n\n"
                ":(\n") % error
    message = MIMEText(sent_body)
    message['Subject'] = sent_subject
    send_email(sent_from, sent_to, message, logger)


def do_send_email(logger):
    sent_from = constants.GMAIL_USERNAME
    sent_to = ["<list_of_emails_goes_here>"]
    sent_subject = "There is a new spot available for scheduling"
    sent_body = ("<your_name_goes_here>,\n\n"
                "There is a new opening for scheduling your visa appointment. Check it out as soon as possible!\n\n"
                ":)\n")
    message = MIMEText(sent_body)
    message['Subject'] = sent_subject
    send_email(sent_from, sent_to, message, logger)



def do_send_process_ran_email(logger):
    sent_from = constants.GMAIL_USERNAME
    sent_to = ["<list_of_emails_goes_here>"]
    sent_subject = "Script just ran..."
    sent_body = ("<your_name_goes_here>,\n\n"
                "The schedule chek script just finished running, but unfortunately there were no available spots at the moment.\n\n"
                ":(\n")
    message = MIMEText(sent_body)
    message['Subject'] = sent_subject
    send_email(sent_from, sent_to, message, logger)


def send_email(sent_from: str, sent_to, message: MIMEText, logger):
    logger.info("Sending email FROM: {email_from}, TO: {email_to}".format(email_from = sent_from, email_to = ", ".join(sent_to)))
    mail_server = smtplib.SMTP(constants.GMAIL_SERVER, constants.GMAIL_PORT)
    mail_server.ehlo()
    mail_server.starttls()
    mail_server.login(constants.GMAIL_USERNAME, constants.GMAIL_PASSWORD)
    mail_server.send_message(message, sent_from, sent_to)
    mail_server.close()
    logger.info("Email successfully sent!")


def start_browser() -> Chrome:
    options = ChromeOptions()
    options.add_argument("--headless")

    browser = Chrome(service=Service(), options=options)
    return browser


def do_login(browser: Chrome):
    # Gets the US Visa login page
    browser.get("<url_goes_here>")
    sleep(1) # Wait for the page to be loaded

    # Closes the pop-up
    popup_button = browser.find_element(By.CSS_SELECTOR, ".ui-dialog-buttonset > .ui-button.ui-corner-all.ui-widget")
    popup_button.click()
    sleep(0.5) # Wait for the popup to be closed

    # Fill the username
    login_form_username_input = browser.find_element(By.ID, "user_email")
    login_form_username_input.send_keys(constants.US_VISA_USERNAME)

    # Fill the password
    login_form_password_input = browser.find_element(By.ID, "user_password")
    login_form_password_input.send_keys(constants.US_VISA_PASSWORD)

    # Clicks on the policy checkbox
    login_form_policy_checkbox = browser.find_element(By.ID, "policy_confirmed")
    if (not login_form_policy_checkbox.is_selected()):
        login_form_policy_checkbox.send_keys(Keys.SPACE)

    # Clicks login button
    login_form_sign_in_button = browser.find_element(By.NAME, "commit")
    login_form_sign_in_button.click()
    sleep(2) # Wait for the new page to be loaded


def do_continue_to_schedule(browser: Chrome):
    # Click the continue button
    summary_continue_button = browser.find_element(By.CSS_SELECTOR, ".dropdown .button.primary.small")
    summary_continue_button.click()
    sleep(2) # Wait for the new page to load

    # Open the accordion
    my_page_schedule_accordion = browser.find_element(By.CSS_SELECTOR, ".accordion .accordion-item:nth-child(1)")
    my_page_schedule_accordion.click()
    sleep(0.5) # Wait for the accordion to expand

    # Click the schedule button
    my_page_schedule_button = browser.find_element(By.CSS_SELECTOR, ".accordion .accordion-item.is-active .button")
    my_page_schedule_button.click()
    sleep(2) # Wait for the next page to load


def do_check_schedule(browser: Chrome, logger):
    # Select Halifax location
    appointment_location_select = Select(browser.find_element(By.ID, "appointments_consulate_appointment_facility_id"))
    appointment_location_select.select_by_visible_text("Halifax")
    sleep(2) # Wait for the text to appear

    appointment_location_error = browser.find_element(By.ID, "consulate_date_time_not_available")
    if (appointment_location_error.is_displayed()):
        logger.info("Appointment not available :(")
        do_send_process_ran_email(logger)
    else:
        logger.info("Appointment AVAILABLE. Sending emails...")
        do_send_email(logger)


def check_visa_schedule():
    logger = setup_logging()
    try:
        logger.info("Starting the US Visa Schedule Check...")
        browser = start_browser()
        logger.info("Logging in...")
        do_login(browser)
        logger.info("Accessing the schedule page...")
        do_continue_to_schedule(browser)
        logger.info("Checking if there is schedule available...")
        do_check_schedule(browser, logger)
        browser.quit()
        logger.info("Finishing the US Visa Schedule Check...")
    except Exception as exception:
        logger.error(exception)
        do_send_error_email(exception, logger)


if __name__ == "__main__":
    check_visa_schedule()