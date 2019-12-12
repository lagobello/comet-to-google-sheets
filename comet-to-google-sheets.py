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

def google_api_init():
    SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

    credentials = None
    creds = None
        # The file token.pickle stores the user's access and refresh tokens, and is
        # created automatically when the authorization flow completes for the first
        # time.
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
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

    request = google_api_instance.spreadsheets().values().append(spreadsheetId=spreadsheet_id, range=range_, valueInputOption=value_input_option, insertDataOption=insert_data_option, body=value_range_body)
    response = request.execute()
    return response

def comet_init(ip_address):
    
    comet_client = ModbusTcpClient(ip_address)
    return comet_client

def comet_read_microamp_int(comet_client):
    # result = client.read_holding_registers(0x9c22,2,unit=1) # works returns 2.3 as 23
    # result = client.read_input_registers(39977, 2 , unit=1) # returns float 2.3
    result = comet_client.read_input_registers(40002, 1 , unit=1) # returns [uA] int16
    return result.registers[0]

def scale_420_to_sensor_range(sensor_min, sensor_max, microAmpReading):
    milliampFloat = microAmpReading/1000
    
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
    print("[" + timestamp\
        + "] Current reading is: " + str(microamp) + " [microamps]" \
        + " Water height is: " + str(round(feet,3)) + " [ft]")


def get_water_reading(comet_client):
    # get comet reading with timestamp
    microampInt = comet_read_microamp_int(comet_client)
    current_time = datetime.now().isoformat()
    
    # scale output to water depth sensor with range 0 to 30 feet.
    feet_output = scale_420_to_sensor_range(0, 30, microampInt)

    # print output to terminal
    log_sensor_reading(current_time, microampInt, feet_output)

    water_array=[current_time, microampInt, feet_output]
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
    
    google_api_instance = google_api_init()
    
    ip_address = '192.168.88.240'
    comet_client = comet_init(ip_address)
    get_water_reading(comet_client)

    loop(comet_client, google_api_instance)
    print("Loop exited due to unknown exception error.")



if __name__ == '__main__':
    main()
    input("Press any key to quit.")