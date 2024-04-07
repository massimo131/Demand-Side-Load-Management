from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
from datetime import date, datetime
import urllib
import urllib.parse
import urllib.request
import requests
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

# Infinite loop to continuously monitor and control appliances
while True:
    # Fetching current date
    current_date = date.today()
    current_day = current_date.strftime("%d")
    current_day_int = int(current_day)
    current_month = current_date.strftime("%m")
    current_month_int = int(current_month)
    current_year = current_date.strftime("%y")
    current_year_int = int(current_year)
    previous_day = current_day_int - 1

    # Handling edge cases for calculating previous day and month
    if (previous_day == 0 and current_month_int in [2, 4, 6, 8, 9, 11]) or (previous_day == 0 and current_month_int == 1):
        previous_day_final = 31
        previous_month_final = current_month_int - 1
    elif (previous_day == 0 and current_month_int in [5, 7, 10, 12]):
        previous_day_final = 30
        previous_month_final = current_month_int - 1
    elif (previous_day == 0 and current_month_int == 3):
        previous_day_final = 28
        previous_month_final = current_month_int - 1
    else:
        previous_day_final = previous_day
        previous_month_final = current_month_int

    # Fetching current time
    now = datetime.now()
    current_hour = now.hour
    current_hour_int = int(current_hour)
    previous_hour = current_hour_int - 1
    current_time_slot = (13 + (12 * previous_hour))

    chrome_driver_path = "C:\\chromedriver.exe"

    # Fetching yesterday's peak load value
    driver = webdriver.Chrome(chrome_driver_path)
    driver.get("https://www.delhisldc.org/Redirect.aspx?Loc=0805")
    month_dropdown = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, f"/html/body/form/div[3]/table/tbody/tr[4]/td/select[1]/option[{previous_month_final}]")))
    month_dropdown.click()
    day_dropdown = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, f"/html/body/form/div[3]/table/tbody/tr[4]/td/select[2]/option[{previous_day_final}]")))
    day_dropdown.click()
    peak_load_element = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "/html/body/form/div[3]/table/tbody/tr[4]/td/table[2]/tbody/tr/td[2]/font/font/font/b/font/font/font/font/b/font/table/tbody/tr[2]/td[2]")))
    print("Yesterday's Peak Load Value:", peak_load_element.text)
    yesterday_peak_load = float(peak_load_element.text)
    driver.quit()

    # Fetching yesterday's value at the current time slot
    driver = webdriver.Chrome(chrome_driver_path)
    driver.get(f"https://www.delhisldc.org/Loaddata.aspx?mode={previous_day_final:02d}/{previous_month_final:02d}/20{current_year_int}")
    value_element = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, f"/html/body/form/div[2]/table/tbody/tr[4]/td/div/table/tbody/tr[{current_time_slot}]/td[2]")))
    yesterday_value = float(value_element.text)
    print("Yesterday's Value:", yesterday_value)
    driver.quit()

    # Fetching present load value
    driver = webdriver.Chrome(chrome_driver_path)
    if current_hour_int == 0:
        driver.get(f"https://www.delhisldc.org/Loaddata.aspx?mode={previous_day_final:02d}/{previous_month_final:02d}/20{current_year_int}")
        present_value_element = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "/html/body/form/div[2]/table/tbody/tr[4]/td/div/table/tbody/tr[289]/td[2]")))
        present_value = float(present_value_element.text)
        print("Present Value:", present_value)
        driver.quit()
    else:
        driver.get(f"https://www.delhisldc.org/Loaddata.aspx?mode={current_day}/{current_month}/20{current_year_int}")
        present_value_element = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, f"/html/body/form/div[2]/table/tbody/tr[4]/td/div/table/tbody/tr[{current_time_slot}]/td[2]")))
        present_value = float(present_value_element.text)
        print("Present Value:", present_value)
        driver.quit()

    # Have used 2 cases for the below task as the free API only allows 8 fields per channel.
    last_status = 0
    last_status1 = 0

    # Function to check status of appliances
    def check_appliance_status(x):
        appliances = ["Light 1", "Light 2", "Fan", "Iron", "Microwave", "Dish Washer", "Water Heater", "Washing Machine", "Clothes Dryer"]
        if x <= 8:
            global last_status
            data_url = "https://api.thingspeak.com/channels/1338183/fields/{0}/last.json?api_key=34NL1A31263L46L6".format(x)
            json_data = requests.get(data_url).json()
            json_status = json_data['field{}'.format(x)]
            last_status = int(json_status)
            # If the "Keep ON" button for any particular appliance on the Mobile App is toggled "ON", then keep it on. Otherwise,turn it off.
            if last_status == 1:
                print("Keeping On " + appliances[x - 1])
            elif last_status == 0:
                try:
                    urllib.request.urlopen('http://192.168.0.115/Relay{0}OFF'.format(x))
                except:
                    pass
                print("Turning Off " + appliances[x - 1])
        if x == 9:
            global last_status1
            data1_url = "https://api.thingspeak.com/channels/1391356/fields/1/last.json?api_key=Z0B1K93A6DXOC9G4"
            json_data1 = requests.get(data1_url).json()
            json_status1 = json_data1['field1']
            last_status1 = int(json_status1)
            # If the "Keep ON" button for any particular appliance on the Mobile App is toggled "ON", then keep it on. Otherwise,turn it off.
            if last_status1 == 1:
                print("Keeping On " + appliances[x - 1])
            elif last_status1 == 0:
                try:
                    urllib.request.urlopen('http://192.168.0.115/Relay{0}OFF'.format(x))
                except:
                    pass
                print("Turning Off " + appliances[x - 1])
        return x

    def send_notification(y):
        global subject
        global body
        # Notification for Non Peak Hours
        if y == 0:
            subject = "Update: Peak Hours are now Inactive"
            body = "Dear Customer, The peak hours are now inactive. You can use your appliances."
        # Notification for Peak Hours
        elif y == 1:
            subject = "Update: Peak Hours are now Active"
            body = "Dear Customer, The peak hours are now active. Please restrict the use of your appliances."
        # from_address == your configured email address
        from_address = "demo@gmail.com"
        # to_address == customer's email address
        to_address = "demo@gmail.com"
        msg = MIMEMultipart()
        msg['From'] = from_address
        msg['To'] = to_address
        msg['Subject'] = subject
        body_text = body
        msg.attach(MIMEText(body_text, 'plain'))
        smtp_server = smtplib.SMTP('smtp.gmail.com', 587)
        smtp_server.starttls()
        # "password" == your email password
        smtp_server.login(from_address, "password")
        text = msg.as_string()
        smtp_server.sendmail(from_address, to_address, text)
        smtp_server.quit()
        return msg

    if yesterday_peak_load < 3000:
        yesterday_peak_load += 500
        threshold = 150
        difference = yesterday_peak_load - present_value
        if difference >= threshold:
            print("Non-Peak Hours")
            urllib.request.urlopen("https://api.thingspeak.com/update?api_key=MN890J7216MMBDKQ&field1=1")
            send_notification(1)
        else:
            print("Peak Hours")
            urllib.request.urlopen("https://api.thingspeak.com/update?api_key=MN890J7216MMBDKQ&field1=0")
            send_notification(0)
        print("Waiting for 30sec...")
        time.sleep(30)
        for i in range(9):
            i = i + 1
            appliance_status = check_appliance_status(i)
            print(appliance_status)
    else:
        threshold = 400
        difference = yesterday_peak_load - present_value
        if difference >= threshold:
            if present_value > yesterday_value:
                difference2 = present_value - yesterday_value
                if difference2 >= 100:
                    print("Non-Peak Hours")
                    urllib.request.urlopen("https://api.thingspeak.com/update?api_key=MN890J7216MMBDKQ&field1=1")
                    send_notification(1)
                else:
                    print("Peak Hours")
                    urllib.request.urlopen("https://api.thingspeak.com/update?api_key=MN890J7216MMBDKQ&field1=0")
                    send_notification(0)
                print("Waiting for 30sec...")
                time.sleep(30)
                for i in range(9):
                    i = i + 1
                    appliance_status = check_appliance_status(i)
                    print(appliance_status)
            else:
                difference2 = yesterday_value - present_value
                if difference2 >= 50:
                    print("Non-Peak Hours")
                    urllib.request.urlopen("https://api.thingspeak.com/update?api_key=MN890J7216MMBDKQ&field1=1")
                    send_notification(1)
                else:
                    print("Peak Hours")
                    urllib.request.urlopen("https://api.thingspeak.com/update?api_key=MN890J7216MMBDKQ&field1=0")
                    send_notification(0)
                print("Waiting for 30sec...")
                time.sleep(30)
                for i in range(9):
                    i = i + 1
                    appliance_status = check_appliance_status(i)
                    print(appliance_status)
        else:
            print("Peak Hours")
            urllib.request.urlopen("https://api.thingspeak.com/update?api_key=MN890J7216MMBDKQ&field1=0")
            send_notification(0)
            print("Waiting for 30sec...")
            time.sleep(30)
            for i in range(9):
                i = i + 1
                appliance_status = check_appliance_status(i)
                print(appliance_status)
    # Waiting for 1 hour before fetching the next set of values.
    time.sleep(3600)