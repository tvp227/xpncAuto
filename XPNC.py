#XPNC auto - By Tom Porter.

import tkinter as tk
import subprocess

command = 'pip install -r"%USERPROFILE%/Documents/XPNCv2/Prereq/requirements.txt"'
try:
    output = subprocess.check_output(command, stderr=subprocess.STDOUT, shell=True, encoding='utf-8')
    print(output)
except subprocess.CalledProcessError as e:
    print(f"Error executing pip command: {e.output}")

#Defining the buttons and directory
def Scan_Receipts():
    file_path = "ReceiptScan.py"
    subprocess.run(["python", file_path])

def Convert_into_PDF():
    file_path = "PDFconvertion.py"
    subprocess.run(["python", file_path])

def Manual_data_entry():
    file_path = "DataEntry.py"
    subprocess.run(["python", file_path])

def Open_Excel():
    file_path = "expenses.xlsx"
    subprocess.run(["start", file_path], shell=True)

def quit_application():
    window.quit()

# Create the main window
window = tk.Tk()
window.title("XPNCauto by Tom Porter")
window.geometry("600x300")
window.resizable(False, False)
background_image = tk.PhotoImage(file="Prereq\Background.png")
background_label = tk.Label(window, image=background_image)
background_label.place(x=0, y=0, relwidth=1, relheight=1)
button_font = ("Impact", 18, "bold")

# Create the buttons
button1 = tk.Button(window, text="Scan receipts", command=Scan_Receipts)
button1.pack(pady=15)

button2 = tk.Button(window, text="Create PDF", command=Convert_into_PDF)
button2.pack(pady=15)

button3 = tk.Button(window, text="Manual data entry", command=Manual_data_entry)
button3.pack(pady=15)

button4 = tk.Button(window, text="Open expence form", command=Open_Excel)
button4.pack(pady=15)

quit_button = tk.Button(window, text="Quit", command=quit_application)
quit_button.pack(pady=10)

# Start the main GUI loop
window.mainloop()
