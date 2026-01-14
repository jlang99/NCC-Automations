import requests
import json
import os, sys, ctypes
import urllib.parse
import pyodbc
import tkinter as tk
import time as ty
from tkinter import messagebox, ttk # Added for showing selected values


# Add a input to clipboard feature

parent_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.append(parent_dir)
from PythonTools import API_KEYS


myappid = 'NCC.Routes.Estimator'
ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)

def dbcnxn():
    #Connect to DB
    db = r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=G:\Shared drives\O&M\NCC\NCC 039.accdb;'
    connect_db = pyodbc.connect(db)
    c = connect_db.cursor()
    return c, connect_db


def get_site_coordinates(site_name):
    """Fetches latitude and longitude for a given site or technician name from the database."""
    try:
        c, connect_db = dbcnxn()
        # First, check the TechnicianList table for technicians
        c.execute("SELECT Latitude, Longitude FROM TechnicianList WHERE Technician = ?", site_name)
        result = c.fetchone()
        if result:
            connect_db.close()
            return f"{result[0]},{result[1]}"
        
        # If Not Found, check the LocationList table for sites
        c.execute("SELECT Latitude, Longitude FROM LocationList WHERE Location = ?", site_name)
        result = c.fetchone()
        if result:
            connect_db.close()
            return f"{result[0]},{result[1]}"



        # If not found in either table
        messagebox.showwarning("Location Not Found", f"Coordinates for '{site_name}' not found in the database.")
        connect_db.close()
        return None
    except pyodbc.Error as e:
        messagebox.showerror("Database Error", f"Failed to fetch coordinates for '{site_name}': {e}")
        return None
def list_of_sites():
    """Fetches an ordered list of technicians and site locations from the database."""
    try:
        c, connect_db = dbcnxn()
        # Fetch technicians, ordered
        c.execute("SELECT Technician FROM TechnicianList ORDER BY Technician")
        tech_results = c.fetchall()
        technicians = [row[0] for row in tech_results]

        # Fetch sites, ordered
        c.execute("SELECT Location FROM LocationList ORDER BY Location")
        site_results = c.fetchall()
        sites = [row[0] for row in site_results]
        connect_db.close()
        # Combine lists, with technicians at the top
        return technicians + sites
    except pyodbc.Error as e:
        messagebox.showerror("Database Error", f"Failed to fetch site list: {e}")
        return []


root = tk.Tk()
root.title("Routing Tool")
root.iconbitmap("G:\\Shared drives\\O&M\\NCC Automations\\Icons\\driving1.ico")

# Checkbutton for computeBestOrder
compute_best_order_var = tk.BooleanVar()
# This button is to demonstrate how to get the values for the API call
compute_best_order_checkbox = ttk.Checkbutton(root, text="Compute Best Order", variable=compute_best_order_var)
compute_best_order_checkbox.pack(padx=5)

# OptionMenu for routeType
route_type_var = tk.StringVar(root)
route_type_var.set("fastest") # default value
route_type_options = ["fastest", "shortest"]
route_type_menu = ttk.Combobox(root, textvariable=route_type_var, state="readonly")
route_type_menu['values'] = route_type_options
route_type_menu.pack(padx=5)

# A frame to hold the dynamic option menus
options_frame = tk.Frame(root, padx=10, pady=5)
options_frame.pack()

info_label = tk.Label(options_frame, text="Select Locations in order")
info_label.pack()



# This list will hold the StringVar for each dropdown
option_vars = []
site_list = list_of_sites()

def on_keyrelease(event):
    """Filters the combobox dropdown list based on user input."""
    combobox = event.widget
    typed_text = combobox.get().lower()

    if not typed_text:
        filtered_list = site_list
    else:
        filtered_list = [site for site in site_list if typed_text in site.lower()]
    
    combobox['values'] = filtered_list

def on_enter_press(event):
    """Opens the combobox dropdown when Enter is pressed."""
    event.widget.event_generate('<Down>')

def add_location_dropdown():
    """Creates and adds a new location dropdown to the window."""
    if not site_list:
        # Don't add a dropdown if the site list is empty
        if not option_vars: # Show error only once
             messagebox.showwarning("No Sites", "Could not load any sites from the database.")
        return

    new_var = tk.StringVar()
    option_vars.append(new_var)

    # The *site_list unpacks the list of sites into individual arguments for the OptionMenu
    new_menu = ttk.Combobox(options_frame, textvariable=new_var)
    new_menu['values'] = site_list
    new_menu.pack(pady=2)
    new_menu.bind('<KeyRelease>', on_keyrelease)
    new_menu.bind('<Return>', on_enter_press)

def get_selected_values():
    """Collects and shows the currently selected values from all dropdowns."""
    selected_sites = [var.get() for var in option_vars]
    return selected_sites

def show_results_window(optimized_locations, result_string, result_distance):
    """Displays the route results in a new window with options to copy details."""
    result_window = tk.Toplevel(root)
    result_window.title("Route Estimate")
    try:
        result_window.iconbitmap("G:\\Shared drives\\O&M\\NCC Automations\\Icons\\driving1.ico")
    except tk.TclError:
        print("Icon not found, skipping.")
    result_window.wm_attributes("-topmost", True)

    main_frame = tk.Frame(result_window, padx=10, pady=10)
    main_frame.pack()

    locations_label = tk.Label(main_frame, text= f"{' -> '.join(optimized_locations)}\nDefaults to departing Now\nWhich can causing optimal route Variance", justify=tk.CENTER)
    locations_label.pack(pady=(0, 10))

    time_frame = tk.Frame(main_frame)
    time_frame.pack(fill='x', pady=2)
    tk.Label(time_frame, text=result_string).pack(side=tk.LEFT, padx=(0, 10))
    

    def copy_time():
        result_window.clipboard_clear()
        result_window.clipboard_append(result_string)
        copy_time_button.config(text="Copied!")
        result_window.after(1500, lambda: copy_time_button.config(text="Copy Time"))

    copy_time_button = ttk.Button(time_frame, text="Copy Time", command=copy_time)
    copy_time_button.pack(side=tk.RIGHT)

    distance_frame = tk.Frame(main_frame)
    distance_frame.pack(fill='x', pady=2)
    tk.Label(distance_frame, text=result_distance).pack(side=tk.LEFT, padx=(0, 10))

    def copy_distance():
        result_window.clipboard_clear()
        result_window.clipboard_append(result_distance)
        copy_dist_button.config(text="Copied!")
        result_window.after(1500, lambda: copy_dist_button.config(text="Copy Distance"))

    copy_dist_button = ttk.Button(distance_frame, text="Copy Distance", command=copy_distance)
    copy_dist_button.pack(side=tk.RIGHT)

    close_button = ttk.Button(main_frame, text="Close", command=result_window.destroy)
    close_button.pack(pady=(10, 0))

def reset_locations():
    """Clears all location dropdowns and resets to two."""
    for widget in options_frame.winfo_children():
        widget.destroy()
    
    option_vars.clear()

    # Re-add the info label that was destroyed
    info_label = tk.Label(options_frame, text="Select Locations in order")
    info_label.pack()

    add_location_dropdown()
    add_location_dropdown()

def route_estimation():
    """Constructs and sends a request to the TomTom Routing API."""    
    selected_locations = get_selected_values()
    if len(selected_locations) < 2:
        messagebox.showwarning("Insufficient Locations", "Please select at least two locations for routing.")
        return
    
    coordinates = []
    for site in selected_locations:
        coords = get_site_coordinates(site)
        if coords:
            coordinates.append(coords)
        else:
            # If any site's coordinates are not found, stop the process
            return
    
    if len(coordinates) < 2:
        messagebox.showwarning("Error", "Could not retrieve coordinates for all selected sites.")
        return

    locations_param = ":".join(coordinates)
    
    version_number = 1
    content_type = "json"
    
    base_url = f"https://api.tomtom.com/routing/{version_number}/calculateRoute/{locations_param}/{content_type}"
    
    params = {
        "key": API_KEYS["tomtom"],
        "routeType": route_type_var.get(), # Use the selected route type
        "traffic": "false",
        "avoid": "tollRoads",
        "travelMode": "car"  # Defaulting to car travel mode
    }
    # Only include computeBestOrder if there are more than two locations
    best_tf = False
    if len(coordinates) > 2:
        best_tf = compute_best_order_var.get()
        params["computeBestOrder"] = str(best_tf).lower()
    
    try:
        print(f"Trying to connect {base_url} with params {params}")
        response = requests.get(base_url, params=params)
        response.raise_for_status()  # Raise an exception for HTTP errors
        route_data = response.json()

        """
        Helped me define outputs from the API


        # Create a filename from the selected locations, replacing invalid characters
        safe_filename = "".join(c if c.isalnum() or c in "-_" else "_" for c in "-".join(selected_locations) + f"_{"T" if best_tf else "F"}")
        filename = f"{safe_filename}.json"
        
        # Get the directory of the script
        script_dir = os.path.dirname(os.path.abspath(__file__))
        output_path = os.path.join(script_dir, filename)
        
        with open(output_path, 'w') as f:
            json.dump(route_data, f, indent=4)
        """
        
        summary = route_data['routes'][0]['summary']

        # Get travel time in seconds
        travel_time_seconds = summary['travelTimeInSeconds']

        # Convert travel time to hours and minutes
        hours = travel_time_seconds // 3600
        minutes = (travel_time_seconds % 3600) // 60

        # Get length in meters
        length_meters = summary['lengthInMeters']

        # Convert meters to miles (1 meter = 0.000621371 miles)
        length_miles = length_meters * 0.000621371

        # Prepare the string for display and clipboard
        result_string = f"Travel Time: {hours} hours {minutes} minutes"
        result_distance = f"Distance: {length_miles:.2f} miles"

        # Determine the order of locations for display
        ordered_locations_for_display = selected_locations
        if best_tf and 'optimizedWaypoints' in route_data:
            optimized_waypoints = route_data['optimizedWaypoints']
            
            # 1. Grab the middle items (the strings)
            movable_points = selected_locations[1:-1]
            
            # 2. Create a NEW empty list of the correct size to hold the reordered strings
            # This prevents you from mixing dictionaries and strings
            reordered_middle = [None] * len(movable_points)
            
            # 3. Fill the new list using the mapping
            for waypoint in optimized_waypoints:
                old_idx = waypoint["optimizedIndex"]
                new_idx = waypoint["providedIndex"]
                reordered_middle[new_idx] = movable_points[old_idx]

            ordered_locations_for_display = [selected_locations[0]] + reordered_middle + [selected_locations[-1]]
        
        show_results_window(ordered_locations_for_display, result_string, result_distance)
            
        print(f"Travel Time: {hours} hours and {minutes} minutes")
        print(f"Distance: {length_miles:.2f} miles")
    except requests.exceptions.RequestException as e:
        messagebox.showerror("API Error", f"Failed to get route from TomTom API: {e}")
        print(e)




# --- Buttons ---
button_frame = tk.Frame(root, pady=10)
button_frame.pack()

add_button = ttk.Button(button_frame, text="Add Location", command=add_location_dropdown)
add_button.pack(side=tk.LEFT, padx=5)

reset_button = ttk.Button(button_frame, text="Reset", command=reset_locations)
reset_button.pack(side=tk.LEFT, padx=5)

get_values_button = ttk.Button(button_frame, text="Estimate Route", command=route_estimation)
get_values_button.pack(side=tk.LEFT, padx=5)

# Add the initial two dropdowns
add_location_dropdown()
add_location_dropdown()

root.mainloop()
