import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import pywhatkit as kit
import time

PHONE_NUMBER_COLUMNS = ['Phone Number', 'Phone', 'Mob. No.' 'Mobile', 'Contact Number']

def send_whatsapp_message(phone_numbers, message):
    for number in phone_numbers:
        try :
            print(number)
        # Send message at the current time + 1 minute
            kit.sendwhatmsg_instantly(number, message)
            time.sleep(10)  # Wait to ensure the message is sent before sending the next one
        except Exception as e :
            messagebox.showerror("Error", f"Failed to send message to {number}. Error: {e}")


def browse_file():
    filename = filedialog.askopenfilename(
        filetypes=[("Excel files", "*.xlsx *.xls")],  # Allow multiple extensions in a single filter
        defaultextension=".xlsx"  # Ensure the default extension is set to .xlsx
    )
    print(f"Selected file: {filename}")  # Debugging print to see the selected file
    if filename:
        file_entry.delete(0, tk.END)
        file_entry.insert(0, filename)

def find_phone_number_column(df):
    for col in df.columns:
        if col in PHONE_NUMBER_COLUMNS:
            return col
    return None

def validate_phone_number(number):
    number = number.replace(" ", "")  # Remove spaces
    if len(number) == 10 and number.isdigit():
        number = "+91" + number
    return number

def start_sending():
    file_path = file_entry.get()
    message = message_entry.get("1.0", tk.END).strip()

    if not file_path or not message:
        messagebox.showwarning("Input Error", "Please provide both the Excel file and the message.")
        return

    try:
        # Read only the first row to find the phone number column
        df = pd.read_excel(file_path, nrows=1)
        phone_column = find_phone_number_column(df)
        
        if not phone_column:
            messagebox.showerror("Error", "The Excel file must contain a column with phone numbers.")
            return
        
        # Read the entire file now that we know which column contains phone numbers
        df = pd.read_excel(file_path)
        phone_numbers = df[phone_column].astype(str).apply(validate_phone_number).tolist()
        send_whatsapp_message(phone_numbers, message)
        messagebox.showinfo("Success", "Messages sent successfully.")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")

# Create the main application window
root = tk.Tk()
root.title("WhatsApp Message Sender")

# Configure the grid layout for better control over the placement of widgets
root.grid_columnconfigure(0, weight=1)
root.grid_columnconfigure(1, weight=3)

# File selection
file_label = tk.Label(root, text="Select Excel File with Phone Numbers:")
file_label.grid(row=0, column=0, padx=5, pady=5, sticky="e")

file_entry = tk.Entry(root, width=50)
file_entry.grid(row=0, column=1, padx=5, pady=5)

browse_button = tk.Button(root, text="Browse", command=browse_file)
browse_button.grid(row=0, column=2, padx=5, pady=5)

# Message input
message_label = tk.Label(root, text="Enter the message to send:")
message_label.grid(row=1, column=0, padx=5, pady=5, sticky="ne")

message_entry = tk.Text(root, width=50, height=10, bg="white", fg="black")  # Set background and foreground colors
message_entry.grid(row=1, column=1, padx=5, pady=5, columnspan=2)

# Send button
send_button = tk.Button(root, text="Send Messages", command=start_sending)
send_button.grid(row=2, column=1, padx=5, pady=20)

# Run the application
root.mainloop()
