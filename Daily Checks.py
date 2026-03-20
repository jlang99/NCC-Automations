#This is the GUI for the Daily Data Analysis Tasks the NCC Desk performs
import os, sys, ctypes
import tkinter as tk
from tkinter import filedialog, messagebox

project_root = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.path.append(project_root)
from PythonTools import get_google_credentials
import TrackerDataUtils
import PerformanceDataUtils



myappid = 'NCC.Daily.Checks'
ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)


def run_tracker_check():
    """Handles the file selection and processing for the Tracker Check."""
    disable_buttons()
    root.update()  # Ensure GUI updates to show disabled state

    downloads_folder = os.path.join(os.path.expanduser("~"), "Downloads")
    loss_file_path = filedialog.askopenfilename(
        title="Select Tracker Loss Data Excel File",
        filetypes=[("Excel files", "*.xlsx")],
        initialdir=downloads_folder
    )
    angles_file_path = filedialog.askopenfilename(
        title="Select Tracker Data Excel File",
        filetypes=[("Excel files", "*.xlsx")],
        initialdir=downloads_folder
    )

    if angles_file_path and loss_file_path:
        try:
            creds = get_google_credentials()
            TrackerDataUtils.process_AE_Tracker_Loss_file(loss_file_path, creds)
            TrackerDataUtils.process_AE_Tracker_file(angles_file_path, creds)
            messagebox.showinfo("Success", "Tracker check processing complete!")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred during tracker check: {e}")
        finally:
            enable_buttons()
    else:
        enable_buttons()  # Re-enable if the user cancels the dialog


def run_inv_performance_check():
    """Handles the file selection and processing for the INV Performance Check."""
    disable_buttons()
    root.update()  # Ensure GUI updates to show disabled state

    downloads_folder = os.path.join(os.path.expanduser("~"), "Downloads")
    file_path = filedialog.askopenfilename(
        title="Select INV Performance Data Excel File",
        filetypes=[("Excel files", "*.xlsx")],
        initialdir=downloads_folder
    )

    if file_path:
        try:
            creds = get_google_credentials()
            PerformanceDataUtils.process_INV_performance_xlsx(file_path, creds)
            messagebox.showinfo("Success", "INV Performance check processing complete!")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred during INV performance check: {e}")
        finally:
            enable_buttons()
    else:
        enable_buttons()  # Re-enable if the user cancels the dialog


def run_tracker_reports():
    """Handles the file selection and processing for the Tracker Reports."""
    disable_buttons()
    root.update()  # Ensure GUI updates to show disabled state

    downloads_folder = os.path.join(os.path.expanduser("~"), "Downloads")
    file_path = filedialog.askopenfilename(
        title="Select Tracker WO's Excel File",
        filetypes=[("Excel files", "*.xlsx")],
        initialdir=downloads_folder
    )

    if file_path:
        try:
            creds = get_google_credentials()
            TrackerDataUtils.process_TR_report_wos(file_path, creds)
            messagebox.showinfo("Success", "Tracker reports processing complete!")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred during tracker reports processing: {e}")
        finally:
            enable_buttons()
    else:
        enable_buttons()  # Re-enable if the user cancels the dialog


def run_cb_check():
    """Handles the file selection and processing for the CB Check."""
    disable_buttons()
    root.update()  # Ensure GUI updates to show disabled state

    downloads_folder = os.path.join(os.path.expanduser("~"), "Downloads")
    file_path = filedialog.askopenfilename(
        title="Select Combiner Box Data Excel File",
        filetypes=[("Excel files", "*.xlsx")],
        initialdir=downloads_folder
    )

    if file_path:
        try:
            creds = get_google_credentials()
            PerformanceDataUtils.process_cb_file(file_path, creds)
            messagebox.showinfo("Success", "Combiner Box check processing complete!")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred during combiner box check: {e}")
        finally:
            enable_buttons()
    else:
        enable_buttons()  # Re-enable if the user cancels the dialog


def run_weekly_performance_updates():
    """Handles the file selection and processing for the Weekly Performance Updates."""
    disable_buttons()
    root.update()  # Ensure GUI updates to show disabled state

    downloads_folder = os.path.join(os.path.expanduser("~"), "Downloads")
    file_path = filedialog.askopenfilename(
        title="Select Weekly Performance Updates Excel File",
        filetypes=[("Excel files", "*.xlsx")],
        initialdir=downloads_folder
    )

    if file_path:
        try:
            creds = get_google_credentials()
            PerformanceDataUtils.process_WO_issue_tracking_file(file_path, creds)
            messagebox.showinfo("Success", "Weekly Performance Updates processing complete!")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred during Weekly Performance Updates: {e}")
        finally:
            enable_buttons()
    else:
        enable_buttons()  # Re-enable if the user cancels the dialog


def disable_buttons():
    for child in main_frame.winfo_children():
        if isinstance(child, tk.Button):
            child.config(state=tk.DISABLED)

def enable_buttons():
    for child in main_frame.winfo_children():
        if isinstance(child, tk.Button):
            child.config(state=tk.NORMAL)






# Create the main window
root = tk.Tk()
root.title("Daily Checks GUI")
try:
    root.iconbitmap(r"G:\Shared drives\O&M\NCC Automations\Icons\tracker_3KU_icon.ico")
except Exception as e:
    print(f"Error loading icon: {e}")


# Create a frame to hold the buttons
main_frame = tk.Frame(root, padx=10, pady=10)
main_frame.pack(expand=True, fill=tk.BOTH)

# Define button properties (text and color)
button_properties = [
    ("CB Check", "pink", run_cb_check),
    ("Tracker Check", "orange", run_tracker_check),
    ("INV Performance Check", "gold", run_inv_performance_check),
    ("Weekly Performance Updates", "green", run_weekly_performance_updates),
    ("Tracker Reports", "lightblue", run_tracker_reports),
    ("Send Reports Tool", "violet", lambda: os.startfile(r"G:\Shared drives\O&M\NCC Automations\Daily Automations\Daily Checks\Technician Data Delivery.pyw"))
]

# Create and pack the buttons
for text, color, function_call in button_properties:
    button = tk.Button(main_frame, text=text, font=("", 14), bg=color, height=2, width=40, command=function_call)
    button.pack(pady=5, fill=tk.X)

# Start the main loop
root.mainloop()