import requests
from requests.auth import HTTPBasicAuth
import time
import datetime
import json, os
import pyodbc
import re
import sys
import subprocess
import socket
from tkinter import messagebox
import tkinter as tk
import multiprocessing
from multiprocessing import Manager
from icecream import ic
import urllib3

# Add the parent directory ('NCC Automations') to the Python path
# This allows us to import the 'PythonTools' package from there.
parent_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.append(parent_dir)
from PythonTools import CREDS, EMAILS, AE_HARDWARE_MAP, PausableTimer 


def check_sql_server_status(service_name):
    """
    Checks if the SQL Server service is installed and running.
    Returns:
        'running': if the service is running.
        'stopped': if the service is installed but not running.
        'not_installed': if the service is not installed.
        'error': for other errors.
    """
    try:
        # Use CREATE_NO_WINDOW flag to prevent console window from appearing
        startupinfo = subprocess.STARTUPINFO()
        startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
        
        output = subprocess.check_output(
            ['sc', 'query', service_name], 
            text=True, 
            stderr=subprocess.STDOUT,
            startupinfo=startupinfo
        )
        if "RUNNING" in output:
            return 'running'
        elif "STOPPED" in output:
            return 'stopped'
        else:
            return 'error'
    except subprocess.CalledProcessError as e:
        if "The specified service does not exist as an installed service." in e.output:
            return 'not_installed'
        else:
            if "STOPPED" in e.output: # Sometimes sc query can return error on stopped service
                return 'stopped'
            return 'error'
    except FileNotFoundError:
        # This would happen if 'sc.exe' is not in the system's PATH.
        return 'error'

def check_and_handle_sql_server():
    """Checks for SQL Server and prompts user if needed."""
    # Create a temporary root for message boxes and hide it.
    temp_root = tk.Tk()
    temp_root.withdraw()

    sql_service_names = ['MSSQL$SQLEXPRESS001', 'MSSQL$SQLEXPRESS']
    status = None
    for service_name in sql_service_names:
        status = check_sql_server_status(service_name)
        if status == 'running':
            break  # If one is running, stop checking

    if status == 'running':
        temp_root.destroy()
        return # SQL Server is running, proceed with app
    
    messagebox.showerror("SQL Server Not Found", 
                         f"SQL Server Express (SQLEXPRESS) is not installed or the service 'MSSQL$SQLEXPRESS' cannot be found.\n\n{status}\n\nPlease install SQL Server Express Edition 2022 and then run the DB Table Creation Tool and try again.", parent=temp_root)
    
    temp_root.destroy()
    sys.exit(1)


urllib3.disable_warnings()
#Attmepted
#os.environ['REQUESTS_CA_BUNDLE'] = certifi.where()
#os.environ['SSL_CERT_FILE'] = certifi.where()
dataPullTime = 1




email = EMAILS['NCC Desk']
password = CREDS['AlsoEnergy']
base_url = "https://api.alsoenergy.com"
token_endpoint = "/Auth/token"

sites_endpoint = "/Sites"


def check_fail_loop(auth_file):
        auth_file.seek(0)
        txt = auth_file.read().strip()
        print(txt)  # Read once and store the content
        count = int(txt) - 1
        return count
def reset_count(auth_file):
        auth_file.seek(0)
        auth_file.write("1")
        auth_file.truncate()
def counting_fails(auth_file):
        auth_file.seek(0)
        content = auth_file.read().strip()  # Read once and store the content
        current_value = int(content)
        new_value = current_value + 1

        # Write incremented value back to file
        auth_file.seek(0)
        auth_file.write(str(new_value))
        auth_file.truncate()
        return new_value

def token_management():
    if 'access_token' in globals():
        pass
    else:
        print("Requesting Access Token")
        get_access_token()

# Make the POST request to obtain an access token
def get_access_token():
    headers = {
        'Content-Type': 'application/x-www-form-urlencoded',
    }
    data = {
        'grant_type': 'password',
        'username': email,
        'password': password,
    }
    access_token_url = base_url + token_endpoint
    globals()['initial_response'] = requests.post(access_token_url, headers=headers, data=data, verify=False) #I know that I shouldn't set verify to false for security reasons. But the Code worked flawlessly for over a year being set to True by Default and this was the only thing out of 10 things I tried that worked. 🤷‍♂️ Adivce is welcomed to joseph.lang@narenco.com. If you message me, please include 'Found on GitHub, looking to help', so that I know it's not spam.
    print("Starting response", initial_response.status_code)
    if initial_response.status_code == 200:
        globals()['access_token'] = initial_response.json().get('access_token')
    else:
        new_value = counting_fails(auth_file)
        print(f"Failed Attempts: {new_value}")


# Function to make API request with authentication
def make_api_request(get_hardware_url, access_token):
    headers = {
        'Content-Type': 'application/x-www-form-urlencoded',
    }
    headers['Authorization'] = f"Bearer {access_token}"
    return requests.get(get_hardware_url, headers=headers, verify=False) #I know that I shouldn't set verify to false for security reasons. But the Code worked flawlessly for over a year being set to True by Default and this was the only thing out of 10 things I tried that worked. 🤷‍♂️ Adivce is welcomed to joseph.lang@narenco.com. If you message me, please include 'Found on GitHub, looking to help', so that I know it's not spam.

def get_data_for_site(site, site_data, api_data, AE_HARDWARE_MAP, start, base_url, access_token):
    troubleshooting_file = r"C:\Users\omops\Documents\Automations\Troubleshooting.txt"
    
    current = time.perf_counter()
    #print("Start Processing:", site, round(current-start, 2))
    for category, hardware_data in site_data.items():
        for hardware_id, hdname in hardware_data.items():
            get_hardware_url = f"{base_url}/Hardware/{hardware_id}"
            hardware_response = make_api_request(get_hardware_url, access_token)
            #print(f"Hardware Response {site} | {hdname}: {hardware_response.status_code}")

            if hardware_response.status_code == 200:
                register_values = {}
                hardware_data_response = hardware_response.json()
                #with open(troubleshooting_file, "a") as tbfile: 
                #    json.dump(hardware_data_response, tbfile, indent=2)

                hdtimestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")


                # AE Register Names
                breaker_register_names = ['Status', 'Breaker Open/Closed', 'Status Closed']
                inverterKW_register_names = ['Grid power', 'Line kW', 'AC Real Power', '3phase Power', 'Active Power', 'Total Active Power', 'Pac', 'AC Power']
                inverterDC_register_names = ['DC Power Total', 'DC Voltage1', 'DC Input Voltage', 'DC voltage (average)', 'DC Voltage Average', 'Bus Voltage', 'Total DC Power', 'DC Voltage', 'Input Voltage', 'Vpv']
                amps_a_register_names = ['Phase current, A', 'AC Current A', 'Current A', 'Amps A', 'AC Phase A Current']
                amps_b_register_names = ['Phase current, B', 'AC Current B', 'Current B', 'Amps B', 'AC Phase B Current']
                amps_c_register_names = ['Phase current, C', 'AC Current C', 'Current C', 'Amps C', 'AC Phase C Current']

                volts_a_register_names = ['Volts A-N', 'Volts A', 'AC Voltage A', 'Voltage AN', 'AC Voltage A (Line-Neutral)', 'Voltage, A-N', 'AC Phase A Voltage', 'AC Voltage AN']
                volts_b_register_names = ['Volts B-N', 'Volts B', 'AC Voltage B', 'Voltage BN', 'AC Voltage B (Line-Neutral)', 'Voltage, B-N', 'AC Phase B Voltage', 'AC Voltage BN']
                volts_c_register_names = ['Volts C-N', 'Volts C', 'AC Voltage C', 'Voltage CN', 'AC Voltage C (Line-Neutral)', 'Voltage, C-N', 'AC Phase C Voltage', 'AC Voltage CN']
                meterkw_register_names = ['Active Power', 'Real power', 'Real Power', 'Total power']    #Real Power is probably not used but lowercase power is.            
                
                weather_station_register_names = ['POA Irradiance', 'Plane of Array Irradiance',  'GHI Irradiance', 'Sun (GHI)', 'Sun (POA Temp comp)', 'GHI', 'POA', 'POA irradiance', 'Sun (POA)']

                # Iterate over register groups for the current hardware
                for register_group in hardware_data_response.get('registerGroups', []):
                    for register in register_group.get('registers', []):
                        # Check if the register name is in the list of register names for the category
                        if register['name'] in breaker_register_names:
                            register_values['Status'] = register['value']
                        elif register['name'] in weather_station_register_names:
                            register_values['POA'] = register['value']
                        elif register['name'] in inverterKW_register_names:
                            register_values['KW'] = register['value']
                        elif register['name'] in inverterDC_register_names:
                            register_values['DC V'] = register['value']
                        elif register['name'] in volts_a_register_names:
                            register_values['Volts A'] = register['value']
                        elif register['name'] in volts_b_register_names:
                            register_values['Volts B'] = register['value']
                        elif register['name'] in volts_c_register_names:
                            register_values['Volts C'] = register['value']
                        elif register['name'] in amps_a_register_names:
                            register_values['Amps A'] = register['value']
                        elif register['name'] in amps_b_register_names:
                            register_values['Amps B'] = register['value']
                        elif register['name'] in amps_c_register_names:
                            register_values['Amps C'] = register['value']
                        elif register['name'] in meterkw_register_names:
                            register_values['KW'] = register['value']
                        



                register_values['pytimestamp'] = hdtimestamp
                try:
                    # Save to Dictionary
                    dvtimestamp = hardware_data_response.get('lastUpload')
                    datetime_obj = datetime.datetime.strptime(dvtimestamp, "%Y-%m-%dT%H:%M:%S%z")
                    aetimestamp = datetime_obj.strftime("%Y-%m-%d %H:%M:%S")
                    register_values['aetimestamp'] = aetimestamp
                except Exception as e:
                    print("Error parsing timestamp:", e)
                    print("Timestamp value:", dvtimestamp)

                api_data[hardware_id] = register_values

                #Added to reduce strain on AE API server
                time.sleep(.01)

                #with open(troubleshooting_file, 'a') as tfile:
                #    json.dump(register_values, tfile, indent=2)

            elif hardware_response.status_code == 401: # Unauthorized or Forbidden
                print(f"Failed to retrieve hardware data for {hardware_id}. Status Code: {hardware_response.status_code}")
                sys.exit() #Also Energy has been locking our account so I hope this will prevent that from happening as often. 
            elif hardware_response.status_code > 401:
                print(f"Failed to retrieve hardware data for {hardware_id} at {site} in {category}. Status code: {hardware_response.status_code}")
                time.sleep(60)
 
    end = time.perf_counter()
    print(f"Pulled Data: {site:<15}| Time Taken: {round(end-current, 2)}")

if __name__ == '__main__': #This is absolutely necessary due to running the async pool.
    def my_main():
        global AE_HARDWARE_MAP, sites_endpoint, dataPullTime, today_date, start, api_data, access_token
        def error_callback(e):
            """A callback function to handle and print errors from worker processes."""
            print(f"ERROR in worker process: {e}")
   
        
        # Get today's date
        today_date = datetime.date.today()

        # Your existing code for obtaining the access token and initializing dictionaries
        start = time.perf_counter()

        api_data = multiprocessing.Manager().dict()
        
        #Calls get access token if access token is not already defined. Hopoing this avoid multiple authentications and therefor AE locking out our account. 
        token_management()


        if access_token:
            pool = multiprocessing.Pool()

            for site, site_data in AE_HARDWARE_MAP.items():
                pool.apply_async(get_data_for_site, args=(site, site_data, api_data, AE_HARDWARE_MAP, start, base_url, access_token), error_callback=error_callback)
                time.sleep(1) #Again added to reduce strain on AE API Server
            pool.close()
            pool.join()

            api_data_dict = dict(api_data)
            # This was temporary just so I could visualize the output
            json_loop_exit_file = r"G:\Shared drives\O&M\NCC Automations\Notification System\api_data_visualized.json"
            # Convert and write JSON object to file
            #with open(json_loop_exit_file, "w+") as outfile: 
            #    json.dump(api_data_dict, outfile, indent=2)
            
            #print(api_data_dict)
            hostname = socket.gethostname()
            if hostname == "NAR-OMOPSXPS":
                server = "SQLEXPRESS01"
            else:
                server = "SQLEXPRESS"
            # Create a connection to the Access database
            connection_string = (
                r'DRIVER={ODBC Driver 18 for SQL Server};'
                fr'SERVER=localhost\{server};'
                r'DATABASE=NARENCO_O&M_AE;'
                r'Trusted_Connection=yes;'
                r'Encrypt=no;'
            )
            dbconnection = pyodbc.connect(connection_string)
            cursor = dbconnection.cursor()

            data_start = time.perf_counter()
            #Variables Defined for the right now Data
            if api_data_dict:
                for site, site_data in AE_HARDWARE_MAP.items():
                    inv_num = 1   
                    for device_category, devices in site_data.items():
                        for hardwareid, device_name in devices.items():
                            if device_category == "breakers":
                                if hardwareid == "390118":
                                    violet_Exception = " 2"
                                elif hardwareid == "390117":
                                    violet_Exception = " 1"
                                else:
                                    violet_Exception = ""
                                try:
                                    hdtimestamp_Relay = api_data_dict[f'{hardwareid}']['pytimestamp']
                                    aetimestamp_Relay = '1999-01-01 00:00:00' if api_data_dict[f'{hardwareid}']['aetimestamp'] == '0001-01-01 00:00:00' else api_data_dict[f'{hardwareid}']['aetimestamp']
                                    relay_stat = api_data_dict[f'{hardwareid}']['Status']
                                    valid_values = ['1', '240', 'closed']
                                    openClose = True if relay_stat.lower().strip() in valid_values else False
                                    cursor.execute(f"""INSERT INTO [{site} Breaker Data{violet_Exception}]
                                                (Timestamp, [Last Upload], Status, HardwareId)
                                                    VALUES (?,?,?,?)""", hdtimestamp_Relay, aetimestamp_Relay, openClose, hardwareid)
                                    dbconnection.commit()
                                except KeyError:
                                    print(f"No Comms with {site} Relay") 
                            if device_category == "meters":
                                try:
                                    voltsA = float(re.search(r'\d+', api_data_dict[f'{hardwareid}']['Volts A']).group()) if re.search(r'\d+', api_data_dict[f'{hardwareid}']['Volts A']) else 0
                                    voltsB = float(re.search(r'\d+', api_data_dict[f'{hardwareid}']['Volts B']).group()) if re.search(r'\d+', api_data_dict[f'{hardwareid}']['Volts B']) else 0
                                    voltsC = float(re.search(r'\d+', api_data_dict[f'{hardwareid}']['Volts C']).group()) if re.search(r'\d+', api_data_dict[f'{hardwareid}']['Volts C']) else 0
                                    ampsA = float(re.search(r'\d+', api_data_dict[f'{hardwareid}']['Amps A']).group()) if re.search(r'\d+', api_data_dict[f'{hardwareid}']['Amps A']) else 0
                                    ampsB = float(re.search(r'\d+', api_data_dict[f'{hardwareid}']['Amps B']).group()) if re.search(r'\d+', api_data_dict[f'{hardwareid}']['Amps B']) else 0
                                    ampsC = float(re.search(r'\d+', api_data_dict[f'{hardwareid}']['Amps C']).group()) if re.search(r'\d+', api_data_dict[f'{hardwareid}']['Amps C']) else 0
                                    if api_data_dict[f'{hardwareid}']['KW'].startswith('-'):
                                        meterkw = 0
                                    else:
                                        meterkw = float(re.search(r'\d+', api_data_dict[f'{hardwareid}']['KW']).group()) if re.search(r'\d+', api_data_dict[f'{hardwareid}']['KW']) else 0

                                    if "kW" in api_data_dict[f'{hardwareid}']['KW']:
                                        meterkw = meterkw*1000
                                    elif "MW" in api_data_dict[f'{hardwareid}']['KW']:
                                        meterkw = meterkw*1000000
                                    

                                    hdtimestamp_Meter = api_data_dict[f'{hardwareid}']['pytimestamp']
                                    aetimestamp_Meter = '1999-01-01 00:00:00' if api_data_dict[f'{hardwareid}']['aetimestamp'] == '0001-01-01 00:00:00' else api_data_dict[f'{hardwareid}']['aetimestamp']
                                    cursor.execute(f""" INSERT INTO [{site} Meter Data] (Timestamp, [Last Upload], [Volts A], [Volts B], [Volts C], [Amps A], [Amps B], [Amps C], Watts, HardwareId) VALUES (?,?,?,?,?,?,?,?,?,?)"""
                                                , hdtimestamp_Meter, aetimestamp_Meter, voltsA, voltsB, voltsC, ampsA, ampsB, ampsC, meterkw, hardwareid)
                                    dbconnection.commit()
                                except KeyError as e:
                                    print(f"No Comms with {site} Meter", e)
                                except AttributeError:
                                    print(f"{site} Meter Volt or Amp value is Null")
                            if device_category == "weather_stations":
                                #POA
                                try:
                                    hdtimestamp_POA = api_data_dict[f'{hardwareid}']['pytimestamp']
                                    aetimestamp_POA = '1999-01-01 00:00:00' if api_data_dict[f'{hardwareid}']['aetimestamp'] == '0001-01-01 00:00:00' else api_data_dict[f'{hardwareid}']['aetimestamp']
                                    poa = float(re.search(r'\d+', api_data_dict[f'{hardwareid}']['POA']).group())
                                    cursor.execute(f"""INSERT INTO [{site} POA Data] (Timestamp, [Last Upload], [W/M²], HardwareId) VALUES (?,?,?,?)""", hdtimestamp_POA, aetimestamp_POA, poa, hardwareid)
                                    dbconnection.commit()
                                except KeyError:
                                    print(f"No Comms with {site} POA")
                                except AttributeError:
                                    print(f"{site} POA 0 W/M²")
                                    poa = 0
                                    hdtimestamp_POA = api_data_dict[f'{hardwareid}']['pytimestamp']
                                    aetimestamp_POA = '1999-01-01 00:00:00' if api_data_dict[f'{hardwareid}']['aetimestamp'] == '0001-01-01 00:00:00' else api_data_dict[f'{hardwareid}']['aetimestamp']
                                    cursor.execute(f"""INSERT INTO [{site} POA Data] (Timestamp, [Last Upload], [W/M²], HardwareId) VALUES (?,?,?,?)""", hdtimestamp_POA, aetimestamp_POA, poa, hardwareid)
                                    dbconnection.commit()
                            if device_category == "inverters":
                                # Duplin String/Central Inv Work Around
                                str_invs = ["94056", "94057", "94058", "94059", "94060", "94061", "94062", "94063", "94064", "94065", "94066", "94067", "94068", "94069", "94070", "94071", "94072", "94073"]
                                cent_invs = ["94053", "94055", "94054"]
                                if hardwareid in str_invs:
                                    duplin_exception = " String"
                                elif hardwareid in cent_invs:
                                    duplin_exception = " Central"
                                else:
                                    duplin_exception = ""

                                #Inverters
                                try:
                                    if api_data_dict[f'{hardwareid}']['KW'].startswith('-'):
                                        invkW = 0
                                    else:
                                        invkW = float(re.search(r'\d+', api_data_dict[f'{hardwareid}']['KW']).group())
                                    invcomms = True
                                
                                except KeyError:
                                    print(f"No Comms with {site}{duplin_exception} Inv {inv_num} kW")
                                    invcomms = False
                                except AttributeError:
                                    print(f"{site}{duplin_exception} Inv {inv_num} kW not reporting to AE")
                                    invkW = 0
                                    invcomms = True

                                if invcomms:
                                    if "kW" in api_data_dict[f'{hardwareid}']['KW']:
                                        invkW = invkW*1000
                                    elif "MW" in api_data_dict[f'{hardwareid}']['KW']:
                                        invkW = invkW*1000000
                                try:
                                    hdtimestamp_inv = api_data_dict[f'{hardwareid}']['pytimestamp']
                                    
                                except:
                                    print("I Never thought this would happen, INV timestamp error")

                                try:
                                    aetimestamp_inv = '1999-01-01 00:00:00' if api_data_dict[f'{hardwareid}']['aetimestamp'] == '0001-01-01 00:00:00' else api_data_dict[f'{hardwareid}']['aetimestamp']
                                except:
                                    print("AE INV Timestamp Error")

                                try:
                                    invDCV = float(re.search(r'\d+', api_data_dict[f'{hardwareid}']['DC V']).group())
                                except KeyError:
                                    print(f"No Comms with {site}{duplin_exception} Inv {inv_num}")
                                except AttributeError:
                                    print(f"{site}{duplin_exception} Inv {inv_num} DC V not reporting to AE")
                                    invDCV = 0
                                if duplin_exception == " String":  #Adjusts the inv_num for namng convention when Central and String inverters in use.   
                                    cursor.execute(f""" INSERT INTO [{site}{duplin_exception} INV {inv_num - 3} Data] (Timestamp, [Last Upload], Watts, [dc V], HardwareID) VALUES (?,?,?,?,?)""",
                                                (hdtimestamp_inv, aetimestamp_inv, invkW, invDCV, hardwareid))
                                    dbconnection.commit()
                                else:
                                    cursor.execute(f""" INSERT INTO [{site}{duplin_exception} INV {inv_num} Data] (Timestamp, [Last Upload], Watts, [dc V], HardwareID) VALUES (?,?,?,?,?)""",
                                                (hdtimestamp_inv, aetimestamp_inv, invkW, invDCV, hardwareid))
                                    dbconnection.commit()
                                inv_num += 1


                finish = time.perf_counter()
                print("Data Injection Time:", round(finish - data_start, 5))
                # Close the connection after all data is inserted
                dbconnection.close()

                end = time.perf_counter()
                dataPullTime = round((end - start)/60, 3)
                print("Total Time:", round((end - start)/60, 3), "Minutes")
                reset_count(auth_file)


    loop_exit_file = r"G:\Shared drives\O&M\NCC Automations\Notification System\APISiteStat\Exiting Loop due to Failed Authentications.txt"
    auth_file = open(loop_exit_file, "r+")
    
    #Checks for Instacne of SQL Server and if Running. If not installed locally, it will close program.
    check_and_handle_sql_server()


    my_main()
    wait_time = datetime.datetime.now().replace(hour=8, minute=30, second=0, microsecond=0)
    auth_file.close()
    while True:
        auth_file = open(loop_exit_file, "r+")
        current_time = datetime.datetime.now()
        count = check_fail_loop(auth_file)
        if count > 20:
            #Failed too many times
            reset_count(auth_file)
            os._exit(0)
            break
        #winsound.Beep(250, 500)
        if count > 4 and count < 10:
            #Failed lots, so wait longer
            print("Waiting 120 Seconds then pulling data again")
            time.sleep(120)
        elif count > 10:
            #Failed lots, so wait longer
            print("Waiting 180 Seconds then pulling data again")
            time.sleep(180)     
        else:
            wait = (2 - dataPullTime) * 60
            if wait <= 10:
                wait = 10  
            print(f"Waiting {round(wait, 2)} Seconds then pulling data again")
            time.sleep(wait)

        print("Looping")
        my_main()
        auth_file.close()
