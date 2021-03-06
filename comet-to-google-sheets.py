from pprint import pprint
from googleapiclient import discovery
import os
import pickle
from pymodbus.client.sync import ModbusTcpClient
import time
from datetime import datetime, timedelta

from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
__location__ = os.path.realpath(
    os.path.join(os.getcwd(), os.path.dirname(__file__)))


def google_api_init():
    SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

    credentials = None
    creds = None
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists(os.path.join(__location__, 'token.pickle')):
        with open(os.path.join(__location__, 'token.pickle'), 'rb') as token:
            creds = pickle.load(token)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                os.path.join(__location__, 'credentials.json'), SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open(os.path.join(__location__, 'token.pickle'), 'wb') as token:
            pickle.dump(creds, token)
    credentials = creds
    service = discovery.build('sheets', 'v4', credentials=credentials)
    return service


def google_api_insert_row(google_api_instance, myarray):
    # The ID of the spreadsheet to update.
    spreadsheet_id = '1MgSxyI_AJyMk7-ZpGUoRCUcpONm7CMwWDo2diG-71IU'

    # The A1 notation of a range to search for a logical table of data.
    # Values will be appended after the last row of the table.
    range_ = 'A1:C1'

    # How the input data should be interpreted.
    value_input_option = 'USER_ENTERED'

    # How the input data should be inserted.
    insert_data_option = 'INSERT_ROWS'

    value_range_body = {
        "range": 'A1:C1',
        "values": [
            myarray
        ]
    }

    request = google_api_instance.spreadsheets().values().append(
        spreadsheetId=spreadsheet_id,
        range=range_,
        valueInputOption=value_input_option,
        insertDataOption=insert_data_option,
        body=value_range_body)
    response = request.execute()
    return response


def comet_init(ip_address):

    while True:
        try:
            comet_client = ModbusTcpClient(ip_address)
        except Exception as e:
            time.sleep(1)
            print("Failed to init Comet Modbus TCP connection. Trying again.")
            print(e)
            continue
        break
    return comet_client


def comet_read_microamp_int(comet_client):
    # result = client.read_holding_registers(0x9c22,2,unit=1) # works returns 2.3 as 23
    # result = client.read_input_registers(39977, 2 , unit=1) # returns float
    # 2.3
    while True:
        try:
            result = comet_client.read_input_registers(
                40002, 1, unit=1)  # returns [uA] int16
        except Exception as e:
            time.sleep(1)
            print("Failed get Comet reading. Trying again.")
            print(e)
            continue

        # check if read has error to break out of while loop
        if not result.isError():
            break
        else:
            print("modbus result is error. Trying again.")
            continue

    return result.registers[0]


def scale_420_to_sensor_range(sensor_min, sensor_max, microAmpReading):
    milliampFloat = microAmpReading / 1000

    # perform scale equation
    iMin = int(4)
    iMax = int(20)

    oMin = int(sensor_min)
    oMax = int(sensor_max)

    ispan = int(iMax - iMin)
    ospan = int(oMax - oMin)

    milliampFloat_scaled = (milliampFloat - iMin) / ispan
    output = oMin + (milliampFloat_scaled * ospan)
    return output


def log_sensor_reading(timestamp, microamp, feet):
    print("[" + timestamp
          + "] Current reading is: " + str(microamp) + " [microamps]"
          + " Water height is: " + str(round(feet, 3)) + " [ft]")


def get_water_reading(comet_client):
    # get comet reading with timestamp
    microampInt = comet_read_microamp_int(comet_client)
    current_time = datetime.now().isoformat()

    # scale output to water depth sensor with range 0 to 30 feet.
    feet_output = scale_420_to_sensor_range(0, 30, microampInt)

    # print output to terminal
    log_sensor_reading(current_time, microampInt, feet_output)

    water_array = [current_time, microampInt, feet_output]
    return water_array


def publish_data_to_google_sheets(google_api_instance, data_array):
    print("Sending data to Google sheet, response follows: ")
    response = google_api_insert_row(google_api_instance, data_array)
    pprint(response)


def loop(comet_client, google_api_instance):

    while True:

        # get time. add 1 hour. remove minutes, seconds, and microseconds.
        dt = datetime.now() + timedelta(hours=1)
        dt = dt.replace(minute=0)
        dt = dt.replace(second=0)
        dt = dt.replace(microsecond=0)

        # wait until next hour on the dot.
        print("Waiting for time: " + dt.isoformat() + " to publish data.")
        while datetime.now() < dt:
            time.sleep(1)

        # get sensor reading, and publish to google sheet
        water_array = get_water_reading(comet_client)
        publish_data_to_google_sheets(google_api_instance, water_array)


def main():
    while True:
        try:
            google_api_instance = google_api_init()
        except Exception as e:
            time.sleep(1)
            print("Failed to init Google API. Trying again.")
            print(e)
            continue
        break

    ip_address = '192.168.88.46'
    comet_client = comet_init(ip_address)
    get_water_reading(comet_client)

    loop(comet_client, google_api_instance)
    print("Loop exited due to unknown exception error.")
    input("Press any key to quit.") 

if __name__ == '__main__':
    main()
