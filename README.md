WhatsApp Bulk Message Sender (Python Automation)

# Description
This is a Python and Selenium-based automation bot designed to send WhatsApp messages to multiple numbers automatically. The application retrieves phone numbers from a CSV file, allowing users to send messages in bulk without manually entering each contact.

## This tool is ideal for:
✅ Business announcements and promotions
✅ Notifications for customers or community members
✅ Reminders or automated messages

# Key Features
✔ Send messages to multiple WhatsApp numbers in one execution
✔ Read contact numbers from a CSV file for convenience
✔ Automate WhatsApp Web using Selenium
✔ Implement time delay between messages to avoid being blocked by WhatsApp
✔ Easily configurable without programming expertise

# How It Works
Ketikkan di teriminal 
* sudo apt update && upgrade -y
* git clone https://github.com/namoradigital/whatsappsend.git
* cd whatsappsend
* python3 -m venv myenv
* source myenv/bin/activate
* python whatsapp.py

Reads a list of numbers from a CSV file
Opens WhatsApp Web automatically using Selenium
Sends messages to each number provided by the user
Waits 10 seconds between messages to prevent spam detection
Closes WhatsApp Web automatically once all messages are sent

# How to Use
Prepare a CSV file with the list of numbers to send messages to. The format should be:
Scan the QR Code for WhatsApp Web in the browser window.
Enter the message you want to send.
The script will automatically send messages to all numbers in the CSV file.

#Technologies Used
Python (Programming Language)
Selenium (For automating WhatsApp Web)
Pandas (For reading CSV files)
WebDriver Manager (For handling ChromeDriver)

# Benefits of This Application
🔹 Does not require WhatsApp's official API, just uses WhatsApp Web
🔹 Useful for businesses or personal use
🔹 Completely free and open-source
🔹 No advanced coding required, just run the script

