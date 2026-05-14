#This is the GUI for the Daily Data Analysis Tasks the NCC Desk performs
import os, sys, ctypes
import tkinter as tk
from tkinter import filedialog, messagebox
import concurrent.futures
import threading
import time

project_root = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.path.append(project_root)
from PythonTools import get_google_credentials
import TrackerDataUtils
import PerformanceDataUtils



myappid = 'NCC.Daily.Checks'
ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)

_creds = None
_executor = concurrent.futures.ThreadPoolExecutor(max_workers=3)

def _get_creds():
    """Returns cached Google credentials, initializing them on first call."""
    global _creds
    if _creds is None:
        _creds = get_google_credentials()
    return _creds


def run_all_checks():
    """Opens file dialogs for each check (cancel any to skip it) then runs all selected checks in parallel."""
    disable_buttons()
    root.update()

    # Initialise credentials on the main thread before any background tasks start.
    creds = _get_creds()
    downloads_folder = os.path.join(os.path.expanduser("~"), "Downloads")
    tasks = []

    cb_file = filedialog.askopenfilename(
        title="[1/6] CB Check — select file (Cancel to skip)",
        filetypes=[("Excel files", "*.xlsx")], initialdir=downloads_folder
    )
    if cb_file:
        tasks.append(('CB Check', lambda f=cb_file, c=creds: PerformanceDataUtils.process_cb_file(f, c)))

    inv_file = filedialog.askopenfilename(
        title="[2/6] INV Performance — select file (Cancel to skip)",
        filetypes=[("Excel files", "*.xlsx")], initialdir=downloads_folder
    )
    if inv_file:
        tasks.append(('INV Performance', lambda f=inv_file, c=creds: PerformanceDataUtils.process_INV_performance_xlsx(f, c)))

    loss_file = filedialog.askopenfilename(
        title="[3/6] Tracker Check — select Tracker Loss file (Cancel to skip)",
        filetypes=[("Excel files", "*.xlsx")], initialdir=downloads_folder
    )
    if loss_file:
        angles_file = filedialog.askopenfilename(
            title="[4/6] Tracker Check — select Tracker Angles file",
            filetypes=[("Excel files", "*.xlsx")], initialdir=downloads_folder
        )
        if angles_file:
            def tracker_task(l=loss_file, a=angles_file, c=creds):
                TrackerDataUtils.process_AE_Tracker_Loss_file(l, c)
                TrackerDataUtils.process_AE_Tracker_file(a, c)
            tasks.append(('Tracker Check', tracker_task))

    wo_file = filedialog.askopenfilename(
        title="[5/6] Weekly Performance Updates — select file (Cancel to skip)",
        filetypes=[("Excel files", "*.xlsx")], initialdir=downloads_folder
    )
    if wo_file:
        tasks.append(('Weekly Performance Updates', lambda f=wo_file, c=creds: PerformanceDataUtils.process_WO_issue_tracking_file(f, c)))

    tr_file = filedialog.askopenfilename(
        title="[6/6] Tracker Reports — select Tracker WOs file (Cancel to skip)",
        filetypes=[("Excel files", "*.xlsx")], initialdir=downloads_folder
    )
    if tr_file:
        tasks.append(('Tracker Reports', lambda f=tr_file, c=creds: TrackerDataUtils.process_TR_report_wos(f, c)))

    if not tasks:
        enable_buttons()
        return

    overall_start = time.perf_counter()

    def run_task(name, fn):
        t0 = time.perf_counter()
        try:
            fn()
            return (name, None, time.perf_counter() - t0)
        except Exception as e:
            return (name, str(e), time.perf_counter() - t0)

    # Tracker Reports must wait for Tracker Check to finish (SQL dependency).
    # Everything else — including Tracker Check itself — runs in parallel.
    _reports_entry = next(((n, f) for n, f in tasks if n == 'Tracker Reports'), None)
    _check_present = any(n == 'Tracker Check' for n, _ in tasks)

    def monitor():
        results = {}
        timings = {}

        if _reports_entry and _check_present:
            # Start all tasks except Tracker Reports immediately.
            # Launch Tracker Reports the moment Tracker Check completes.
            pending = {_executor.submit(run_task, n, f): n
                       for n, f in tasks if n != 'Tracker Reports'}
            reports_launched = False

            while pending:
                done, _ = concurrent.futures.wait(pending, return_when=concurrent.futures.FIRST_COMPLETED)
                for fut in done:
                    task_name = pending.pop(fut)
                    _, err, elapsed = fut.result()
                    results[task_name] = err
                    timings[task_name] = elapsed
                    if task_name == 'Tracker Check' and not reports_launched:
                        rn, rf = _reports_entry
                        pending[_executor.submit(run_task, rn, rf)] = rn
                        reports_launched = True
        else:
            # No Tracker Check / Reports dependency — run everything in parallel.
            for future in concurrent.futures.as_completed(
                [_executor.submit(run_task, n, f) for n, f in tasks]
            ):
                name, err, elapsed = future.result()
                results[name] = err
                timings[name] = elapsed

        total = time.perf_counter() - overall_start
        root.after(0, lambda r=results, t=timings, tot=total: on_all_done(r, t, tot))

    def on_all_done(results, timings, total):
        enable_buttons()
        print("\n" + "=" * 45)
        print("  DAILY CHECKS — PROCESS TIME LOG")
        print("=" * 45)
        for name, fn in tasks:
            elapsed = timings.get(name, 0)
            status = "ERROR" if results.get(name) else "OK"
            print(f"  {name:<30} {elapsed:>6.1f}s  [{status}]")
        print("-" * 45)
        print(f"  {'Total (wall clock)':<30} {total:>6.1f}s")
        print("=" * 45 + "\n")
        errors = {k: v for k, v in results.items() if v is not None}
        if errors:
            msg = "Some checks failed:\n" + "\n".join(f"• {k}: {v}" for k, v in errors.items())
            messagebox.showerror("Checks Completed with Errors", msg)
        else:
            messagebox.showinfo("Success", f"All {len(tasks)} checks completed successfully!")

    threading.Thread(target=monitor, daemon=True).start()


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
            TrackerDataUtils.process_AE_Tracker_Loss_file(loss_file_path, _get_creds())
            TrackerDataUtils.process_AE_Tracker_file(angles_file_path, _get_creds())
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
            PerformanceDataUtils.process_INV_performance_xlsx(file_path, _get_creds())
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
            TrackerDataUtils.process_TR_report_wos(file_path, _get_creds())
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
            PerformanceDataUtils.process_cb_file(file_path, _get_creds())
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
            PerformanceDataUtils.process_WO_issue_tracking_file(file_path, _get_creds())
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
    ("Run All Daily Checks", "lightyellow", run_all_checks),
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