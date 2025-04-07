import os
import csv
import time
import pywhatkit
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from datetime import datetime, timedelta
from PIL import Image, ImageTk
import pyautogui
import webbrowser
import urllib.parse

class WhatsAppSenderApp:
    def __init__(self, root):
        self.root = root
        self.root.title("WhatsApp Bulk Message Sender")
        self.root.geometry("700x650")
        self.root.resizable(False, False)
        
        # Variables
        self.phone_numbers = []
        self.message = ""
        self.file_path = ""
        self.image_path = ""
        self.report_data = []
        self.delay_minutes = tk.IntVar(value=2)  # Default 2 minutes delay
        
        # Create GUI
        self.create_widgets()
        
    def create_widgets(self):
        # Notebook (Tab)
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(pady=10, padx=10, fill='both', expand=True)
        
        # Main Tab
        self.main_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.main_frame, text="Message Sender")
        
        # Phone Number Input
        ttk.Label(self.main_frame, text="Phone Numbers:").grid(row=0, column=0, padx=5, pady=5, sticky='w')
        
        # Input Method Selection
        self.input_method = tk.StringVar(value="manual")
        ttk.Radiobutton(self.main_frame, text="Manual Input", variable=self.input_method, value="manual", 
                        command=self.toggle_input_method).grid(row=1, column=0, padx=5, pady=5, sticky='w')
        ttk.Radiobutton(self.main_frame, text="Upload CSV", variable=self.input_method, value="csv", 
                        command=self.toggle_input_method).grid(row=1, column=1, padx=5, pady=5, sticky='w')
        
        # Manual Input Frame
        self.manual_frame = ttk.Frame(self.main_frame)
        self.manual_frame.grid(row=2, column=0, columnspan=2, padx=5, pady=5, sticky='ew')
        
        self.phone_entry = tk.Text(self.manual_frame, height=5, width=60)
        self.phone_entry.pack(fill='both', expand=True)
        ttk.Label(self.manual_frame, text="Enter numbers (separate with commas or new lines)").pack()
        
        # CSV Input Frame (initially hidden)
        self.csv_frame = ttk.Frame(self.main_frame)
        self.csv_btn = ttk.Button(self.csv_frame, text="Select CSV File", command=self.browse_csv)
        self.csv_btn.pack(pady=5)
        self.csv_label = ttk.Label(self.csv_frame, text="No file selected")
        self.csv_label.pack()
        
        # Message Input
        ttk.Label(self.main_frame, text="Message:").grid(row=3, column=0, padx=5, pady=5, sticky='w')
        self.message_entry = tk.Text(self.main_frame, height=8, width=60)
        self.message_entry.grid(row=4, column=0, columnspan=2, padx=5, pady=5)
        
        # Attachment Frame
        self.attachment_frame = ttk.LabelFrame(self.main_frame, text="Attachments")
        self.attachment_frame.grid(row=5, column=0, columnspan=2, padx=5, pady=5, sticky='ew')
        
        # File Attachment
        self.file_btn = ttk.Button(self.attachment_frame, text="Attach File", command=self.browse_file)
        self.file_btn.grid(row=0, column=0, padx=5, pady=5)
        self.file_label = ttk.Label(self.attachment_frame, text="No file selected", wraplength=300)
        self.file_label.grid(row=0, column=1, padx=5, pady=5, sticky='w')
        
        # Image Attachment
        self.image_btn = ttk.Button(self.attachment_frame, text="Attach Image", command=self.browse_image)
        self.image_btn.grid(row=1, column=0, padx=5, pady=5)
        self.image_label = ttk.Label(self.attachment_frame, text="No image selected")
        self.image_label.grid(row=1, column=1, padx=5, pady=5, sticky='w')
        
        # Image Preview
        self.image_preview = ttk.Label(self.attachment_frame)
        self.image_preview.grid(row=2, column=0, columnspan=2, pady=5)
        
        # Clear Attachment Button
        self.clear_btn = ttk.Button(self.attachment_frame, text="Clear Attachments", command=self.clear_attachments)
        self.clear_btn.grid(row=3, column=0, columnspan=2, pady=5)
        
        # Delay Configuration
        ttk.Label(self.main_frame, text="Delay between messages (minutes):").grid(row=6, column=0, padx=5, pady=5, sticky='w')
        ttk.Spinbox(self.main_frame, from_=1, to=10, textvariable=self.delay_minutes, width=5).grid(row=6, column=1, padx=5, pady=5, sticky='w')
        
        # Send Button
        self.send_btn = ttk.Button(self.main_frame, text="Send Messages", command=self.send_messages)
        self.send_btn.grid(row=7, column=0, columnspan=2, pady=10)
        
        # Progress Bar
        self.progress = ttk.Progressbar(self.main_frame, orient='horizontal', length=500, mode='determinate')
        self.progress.grid(row=8, column=0, columnspan=2, pady=5)
        
        # Status Label
        self.status_label = ttk.Label(self.main_frame, text="Ready to send messages")
        self.status_label.grid(row=9, column=0, columnspan=2)
        
        # Report Tab
        self.report_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.report_frame, text="Report")
        
        # Report Text
        self.report_text = tk.Text(self.report_frame, height=25, width=80)
        self.report_text.pack(pady=10, padx=10, fill='both', expand=True)
        self.report_text.insert('end', "Sending report will appear here\n")
        self.report_text.config(state='disabled')
        
        # Export Button
        self.export_btn = ttk.Button(self.report_frame, text="Export Report to CSV", command=self.export_report)
        self.export_btn.pack(pady=5)
        
        # Initially show manual input
        self.toggle_input_method()
    
    def toggle_input_method(self):
        if self.input_method.get() == "manual":
            self.manual_frame.grid()
            self.csv_frame.grid_remove()
        else:
            self.manual_frame.grid_remove()
            self.csv_frame.grid(row=2, column=0, columnspan=2, padx=5, pady=5, sticky='ew')
    
    def browse_csv(self):
        file_path = filedialog.askopenfilename(filetypes=[("CSV Files", "*.csv")])
        if file_path:
            self.csv_label.config(text=os.path.basename(file_path))
            self.process_csv(file_path)
    
    def browse_file(self):
        file_path = filedialog.askopenfilename(filetypes=[
            ("All Files", "*.*"),
            ("PDF Files", "*.pdf"),
            ("Word Documents", "*.docx"),
            ("Excel Files", "*.xlsx"),
            ("Text Files", "*.txt")
        ])
        if file_path:
            self.file_path = file_path
            self.file_label.config(text=os.path.basename(file_path))
    
    def browse_image(self):
        file_path = filedialog.askopenfilename(filetypes=[
            ("Image Files", "*.jpg *.jpeg *.png *.gif"),
            ("All Files", "*.*")
        ])
        if file_path:
            self.image_path = file_path
            self.image_label.config(text=os.path.basename(file_path))
            
            # Display image preview
            try:
                img = Image.open(file_path)
                img.thumbnail((200, 200))  # Resize for preview
                photo = ImageTk.PhotoImage(img)
                self.image_preview.config(image=photo)
                self.image_preview.image = photo  # Keep a reference
            except Exception as e:
                messagebox.showerror("Error", f"Failed to load image: {str(e)}")
    
    def clear_attachments(self):
        self.file_path = ""
        self.image_path = ""
        self.file_label.config(text="No file selected")
        self.image_label.config(text="No image selected")
        self.image_preview.config(image='')
        self.image_preview.image = None
    
    def process_csv(self, file_path):
        self.phone_numbers = []
        try:
            with open(file_path, 'r') as file:
                reader = csv.reader(file)
                for row in reader:
                    if row:  # Skip empty rows
                        # Assuming phone numbers are in the first column
                        phone = row[0].strip()
                        if phone:
                            self.phone_numbers.append(phone)
            messagebox.showinfo("Info", f"Successfully loaded {len(self.phone_numbers)} numbers from CSV file")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to process CSV file: {str(e)}")
    
    def get_phone_numbers(self):
        if self.input_method.get() == "csv":
            return self.phone_numbers
        else:
            # Process manual input
            input_text = self.phone_entry.get("1.0", "end").strip()
            numbers = []
            for item in input_text.split(','):
                for num in item.split('\n'):
                    num = num.strip()
                    if num:
                        numbers.append(num)
            return numbers
    
    def format_phone_number(self, phone):
        # Remove all non-digit characters
        digits = ''.join(c for c in phone if c.isdigit())
        
        if not digits:
            raise ValueError("Empty phone number")
            
        # Add country code if missing (assuming international format)
        if not digits.startswith('+'):
            if digits.startswith('0'):
                # Replace leading 0 with country code (adjust as needed)
                digits = '+62' + digits[1:]  # Indonesia as example
            else:
                digits = '+' + digits
                
        return digits
    
    def send_whatsapp_message(self, phone, message, image_path=None):
        try:
            formatted_phone = self.format_phone_number(phone)
            
            if image_path:
                # Send image with caption
                pywhatkit.sendwhats_image(
                    receiver=formatted_phone,
                    img_path=image_path,
                    caption=message,
                    wait_time=15,
                    tab_close=True
                )
            else:
                # Send text message
                pywhatkit.sendwhatmsg_instantly(
                    phone_no=formatted_phone,
                    message=message,
                    wait_time=15,
                    tab_close=True
                )
            
            return True, ""
        except Exception as e:
            return False, str(e)
    
    def send_messages(self):
        phone_numbers = self.get_phone_numbers()
        self.message = self.message_entry.get("1.0", "end").strip()
        
        if not phone_numbers:
            messagebox.showerror("Error", "No phone numbers entered")
            return
        
        if not self.message and not self.file_path and not self.image_path:
            messagebox.showerror("Error", "Message or attachment cannot be empty")
            return
        
        # Confirm before sending
        confirm_msg = f"Send {'message' if self.message else ''}"
        confirm_msg += f"{' and ' if self.message and (self.file_path or self.image_path) else ''}"
        confirm_msg += f"{'file' if self.file_path else ''}"
        confirm_msg += f"{' and ' if self.file_path and self.image_path else ''}"
        confirm_msg += f"{'image' if self.image_path else ''}"
        confirm_msg += f" to {len(phone_numbers)} numbers?"
        
        if not messagebox.askyesno("Confirmation", confirm_msg):
            return
        
        # Prepare for sending
        self.report_data = []
        self.progress['maximum'] = len(phone_numbers)
        self.progress['value'] = 0
        self.status_label.config(text="Sending messages...")
        self.root.update()
        
        # Calculate initial time
        now = datetime.now()
        send_time = now + timedelta(minutes=1)  # Start sending 1 minute from now
        
        # Send messages
        success_count = 0
        fail_count = 0
        
        for i, phone in enumerate(phone_numbers):
            try:
                # Format phone number
                formatted_phone = self.format_phone_number(phone)
                
                # Determine what to send
                if self.image_path:
                    # Send image with caption
                    status, error = self.send_whatsapp_message(
                        phone=formatted_phone,
                        message=self.message,
                        image_path=self.image_path
                    )
                elif self.file_path:
                    # For files, we'll use a direct WhatsApp URL
                    whatsapp_url = f"https://web.whatsapp.com/send?phone={formatted_phone[1:]}&text={urllib.parse.quote(self.message)}"
                    webbrowser.open(whatsapp_url)
                    time.sleep(10)  # Wait for page to load
                    
                    # Automate file sending with pyautogui
                    try:
                        # Click attachment button
                        pyautogui.click(x=100, y=200)  # Adjust coordinates as needed
                        time.sleep(1)
                        
                        # Click document button
                        pyautogui.click(x=100, y=250)  # Adjust coordinates as needed
                        time.sleep(1)
                        
                        # Type file path and press enter
                        pyautogui.write(self.file_path)
                        pyautogui.press('enter')
                        time.sleep(2)
                        
                        # Press enter to send
                        pyautogui.press('enter')
                        status, error = True, ""
                    except Exception as e:
                        status, error = False, f"File sending failed: {str(e)}"
                else:
                    # Send text message only
                    status, error = self.send_whatsapp_message(
                        phone=formatted_phone,
                        message=self.message
                    )
                
                # Update report
                if status:
                    self.report_data.append({
                        'phone': phone,
                        'status': 'Success',
                        'error': ''
                    })
                    success_count += 1
                else:
                    self.report_data.append({
                        'phone': phone,
                        'status': 'Failed',
                        'error': error
                    })
                    fail_count += 1
                
                # Update progress
                self.progress['value'] = i + 1
                self.status_label.config(text=f"Sending {i+1}/{len(phone_numbers)} - Success: {success_count}, Failed: {fail_count}")
                self.root.update()
                
                # Delay between messages
                if i < len(phone_numbers) - 1:  # No delay after last message
                    delay_seconds = self.delay_minutes.get() * 60
                    for remaining in range(delay_seconds, 0, -1):
                        self.status_label.config(
                            text=f"Sending {i+1}/{len(phone_numbers)} - Next in {remaining} sec - Success: {success_count}, Failed: {fail_count}"
                        )
                        self.root.update()
                        time.sleep(1)
                
            except Exception as e:
                # Update report with error
                self.report_data.append({
                    'phone': phone,
                    'status': 'Failed',
                    'error': str(e)
                })
                fail_count += 1
                
                # Update progress
                self.progress['value'] = i + 1
                self.status_label.config(text=f"Sending {i+1}/{len(phone_numbers)} - Success: {success_count}, Failed: {fail_count}")
                self.root.update()
                
                # Continue with next number
                continue
        
        # Update status
        self.status_label.config(text=f"Sending completed! Success: {success_count}, Failed: {fail_count}")
        messagebox.showinfo("Completed", f"Sending completed!\nSuccess: {success_count}\nFailed: {fail_count}")
        
        # Update report tab
        self.update_report()
    
    def update_report(self):
        self.report_text.config(state='normal')
        self.report_text.delete('1.0', 'end')
        
        # Add summary
        success_count = sum(1 for item in self.report_data if item['status'] == 'Success')
        fail_count = sum(1 for item in self.report_data if item['status'] == 'Failed')
        
        self.report_text.insert('end', f"=== SENDING REPORT ===\n")
        self.report_text.insert('end', f"Total sent: {len(self.report_data)}\n")
        self.report_text.insert('end', f"Success: {success_count}\n")
        self.report_text.insert('end', f"Failed: {fail_count}\n\n")
        self.report_text.insert('end', f"=== DETAILS ===\n")
        
        # Add details
        for item in self.report_data:
            status_icon = "✓" if item['status'] == 'Success' else "✗"
            self.report_text.insert('end', f"{status_icon} {item['phone']} - {item['status']}")
            if item['error']:
                self.report_text.insert('end', f" (Error: {item['error']})")
            self.report_text.insert('end', "\n")
        
        self.report_text.config(state='disabled')
        self.notebook.select(self.report_frame)  # Switch to report tab
    
    def export_report(self):
        if not self.report_data:
            messagebox.showwarning("Warning", "No report data to export")
            return
        
        file_path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV Files", "*.csv")],
            title="Save Report"
        )
        
        if not file_path:
            return
        
        try:
            with open(file_path, 'w', newline='', encoding='utf-8') as file:
                writer = csv.writer(file)
                # Write header
                writer.writerow(['Phone Number', 'Status', 'Error'])
                # Write data
                for item in self.report_data:
                    writer.writerow([item['phone'], item['status'], item['error']])
            
            messagebox.showinfo("Success", f"Report successfully saved to:\n{file_path}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save report: {str(e)}")

if __name__ == "__main__":
    root = tk.Tk()
    app = WhatsAppSenderApp(root)
    root.mainloop()