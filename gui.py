import subprocess
import sys
import tkinter as tk
from tkinter import ttk, filedialog
from ttkthemes import ThemedTk
import main
import threading

# Approved users
approved_users = ["jana omaiche", "igor ferreira", "kirstie mclaren", "lana matteucci", "jai kite", "carolyn boddington", "john denton"]

def upload_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if file_path:
        file_label.config(text="File selected", foreground="green")
        if hasattr(main, 'set_file_path'):
            main.set_file_path(file_path)
        run_button.config(state=tk.NORMAL)

def update_analyte_options():
    if program_type.get() == "Single Analyte":
        analyte_menu['values'] = ["D-Dimer", "CRP", "INR", "Troponin", "NT-ProBNP", "Haemoglobin"]
    elif program_type.get() == "Multi Analyte":
        analyte_menu['values'] = ["Epoc", "i-STAT", "Lipids", "WCC"]

def check_user_approval(user_id):
    return user_id.lower() in approved_users

def run_program():
    user_id_val = user_id_entry.get().strip().lower()
    if not check_user_approval(user_id_val):
        print("Not Approved User")
        return

    def blink():
        if status_label.cget("text") == "Generating Reports":
            status_label.config(text="")
        else:
            status_label.config(text="Generating Reports")
        if main.is_running_flag.is_set():  # Stop blinking when the task is done
            status_label.config(text="Reports Generated Successfully", foreground="green")
            return
        status_label.after(500, blink)

    def run():
        try:
            program_type_val = program_type.get()
            analyte_val = analyte.get()
            cycle_method_val = cycle_method.get()
            cycle_val = cycle.get()
            
            # Run the main program in a thread
            main.run_program_in_thread(program_type_val, analyte_val, cycle_method_val, cycle_val, user_id_val)
            
            # Wait until the task is done
            main.is_running_flag.wait()
            
            # Ensure blinking stops and final status is displayed
            blink()
        except Exception as e:
            print(f"Error running program: {e}")
            status_label.config(text="Error Generating Reports", foreground="red")

    # Start blinking and run the program
    status_label.config(foreground="black")  # Reset color to default
    threading.Thread(target=run).start()
    blink()

def return_to_main_screen():
    window.destroy()
    subprocess.Popen([sys.executable, 'preqap.py'])

# Create the main window with the aqua theme
window = ThemedTk(theme="aqua")
window.title("iCCnet QAP Program")

window.geometry("600x300")

# Configure column weights for expansion
for i in range(3):
    window.columnconfigure(i, weight=1)

# Program Type
program_type_label = ttk.Label(window, text="Program type:")
program_type_label.grid(row=0, column=0, sticky="w", pady=2, padx=5)

program_type = tk.StringVar()
single_analyte_radio = ttk.Radiobutton(window, text="Single Analyte", variable=program_type, value="Single Analyte", command=update_analyte_options)
single_analyte_radio.grid(row=0, column=1, sticky="w", padx=5)
multi_analyte_radio = ttk.Radiobutton(window, text="Multi Analyte", variable=program_type, value="Multi Analyte", command=update_analyte_options)
multi_analyte_radio.grid(row=0, column=2, sticky="w", pady=2, padx=5)

# Analyte
analyte_label = ttk.Label(window, text="Analyte:")
analyte_label.grid(row=1, column=0, sticky="w", pady=2, padx=5)

analyte = tk.StringVar()
analyte_menu = ttk.Combobox(window, textvariable=analyte)
analyte_menu.grid(row=1, column=1, columnspan=2, sticky="ew", pady=2, padx=5)

# Cycle Method
cycle_method_label = ttk.Label(window, text="Cycle method:")
cycle_method_label.grid(row=2, column=0, sticky="w", pady=2, padx=5)

cycle_method = tk.StringVar()
single_cycle_radio = ttk.Radiobutton(window, text="Single Cycle", variable=cycle_method, value="Single Cycle")
single_cycle_radio.grid(row=2, column=1, sticky="w", padx=5)
multi_cycle_radio = ttk.Radiobutton(window, text="Multi Cycle", variable=cycle_method, value="Multi Cycle", state=tk.DISABLED)
multi_cycle_radio.grid(row=2, column=2, sticky="w", pady=1, padx=5)

# Select Cycle
cycle_label = ttk.Label(window, text="Select Cycle:")
cycle_label.grid(row=3, column=0, sticky="w", pady=1, padx=5)

cycle = tk.StringVar()
cycle_menu = ttk.Combobox(window, textvariable=cycle)
cycle_menu['values'] = ["EQA2401", "EQA2402", "EQA2403", "EQA2404", "EQA2405", "EQA2406", "EQA2407", "EQA2408", "EQA2409", "EQA2410", "EQA2411", "EQA2412"]
cycle_menu.grid(row=3, column=1, columnspan=4, sticky="ew", pady=10, padx=5)

# File Upload
file_upload_label = ttk.Label(window, text="Select XLSX file:")
file_upload_label.grid(row=4, column=0, sticky="w", pady=2, padx=5)

upload_button = ttk.Button(window, text="Upload", command=upload_file)
upload_button.grid(row=4, column=1, sticky="w", padx=5)

file_label = ttk.Label(window, text="")
file_label.grid(row=4, column=2, sticky="ew", pady=10, padx=5)

# User ID
user_id_label = ttk.Label(window, text="User ID:")
user_id_label.grid(row=5, column=0, sticky="w", pady=2, padx=5)

user_id_entry = ttk.Entry(window)
user_id_entry.grid(row=5, column=1, columnspan=2, sticky="ew", pady=2, padx=5)

# Run Button
run_button = ttk.Button(window, text="Generate reports", state=tk.DISABLED, command=run_program)
run_button.grid(row=6, column=1, columnspan=1, pady=10, padx=5)

# Status Label
status_label = ttk.Label(window, text="")
status_label.grid(row=7, column=1, columnspan=1, pady=5)

# Return Button
return_button = ttk.Button(window, text="\u2190 Return to Main Screen", command=return_to_main_screen)
return_button.grid(row=8, column=1, columnspan=1, pady=5, padx=5)

window.mainloop()