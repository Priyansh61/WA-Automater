# WhatsApp Message Sender

This application allows users to send automated WhatsApp messages to a list of phone numbers specified in an Excel file. The application provides a graphical user interface (GUI) for selecting the Excel file and inputting the message to be sent.

## Features

- Select an Excel file containing phone numbers.
- Input a custom message to be sent.
- Automatically send WhatsApp messages to the specified phone numbers.

## Requirements

- Python 3.x
- `tkinter`
- `pandas`
- `pywhatkit`
- `openpyxl`

## Installation

1. Install the required Python libraries:
    ```sh
    pip install pandas pywhatkit openpyxl
    ```

2. Save the Python script to a file, e.g., `send_whatsapp.py`.

## Usage

1. Run the script:
    ```sh
    python send_whatsapp.py
    ```

2. The application window will open.

3. Click the "Browse" button to select an Excel file containing the phone numbers. The Excel file should have a column named `Phone Number`.

4. Enter the message you want to send in the provided text area.

5. Click the "Send Messages" button to start sending messages. The application will open WhatsApp Web and send the message to each phone number listed in the Excel file.

## Running directly from Exe File

1. Navigate to `dist`:
    ```sh
    cd dist
    ```

2. Run the `wa.exe`

## Example Excel File

The Excel file should contain a column named `Phone Number` with the phone numbers formatted as strings. Here's an example:

| Phone Number  |
| ------------- |
| +918112xxxxxx |
| +918619xxxxxx |

## Notes

- Ensure you are logged into WhatsApp Web on your default browser before running the application.
- The application sends messages using the browser, so it requires an active internet connection.
- Adjust the sleep times in the script if necessary to ensure messages are sent properly.

## License

This project is licensed under the MIT License.
