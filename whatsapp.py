import os
import csv
import time
import pywhatkit
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from datetime import datetime, timedelta
from PIL import Image, ImageTk
from tkinter import font as tkfont  # Baru ditambahkan

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
        
        # Create GUI
        self.create_widgets()
        
    def create_widgets(self):
        # Notebook (Tab)
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(pady=10, padx=10, fill='both', expand=True)
        
        # Main Tab
        self.main_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.main_frame, text="Pengirim Pesan")
        
        # Phone Number Input
        ttk.Label(self.main_frame, text="Nomor Telepon:").grid(row=0, column=0, padx=5, pady=5, sticky='w')
        
        # Input Method Selection
        self.input_method = tk.StringVar(value="manual")
        ttk.Radiobutton(self.main_frame, text="Input Manual", variable=self.input_method, value="manual", 
                        command=self.toggle_input_method).grid(row=1, column=0, padx=5, pady=5, sticky='w')
        ttk.Radiobutton(self.main_frame, text="Upload CSV", variable=self.input_method, value="csv", 
                        command=self.toggle_input_method).grid(row=1, column=1, padx=5, pady=5, sticky='w')
        
        # Manual Input Frame
        self.manual_frame = ttk.Frame(self.main_frame)
        self.manual_frame.grid(row=2, column=0, columnspan=2, padx=5, pady=5, sticky='ew')
        
        self.phone_entry = tk.Text(self.manual_frame, height=5, width=60)
        self.phone_entry.pack(fill='both', expand=True)
        ttk.Label(self.manual_frame, text="Masukkan nomor (pisahkan dengan koma atau baris baru)").pack()
        
        # CSV Input Frame (initially hidden)
        self.csv_frame = ttk.Frame(self.main_frame)
        self.csv_btn = ttk.Button(self.csv_frame, text="Pilih File CSV", command=self.browse_csv)
        self.csv_btn.pack(pady=5)
        self.csv_label = ttk.Label(self.csv_frame, text="Belum ada file dipilih")
        self.csv_label.pack()
        
        # Message Input
        ttk.Label(self.main_frame, text="Pesan:").grid(row=3, column=0, padx=5, pady=5, sticky='w')
        self.message_entry = tk.Text(self.main_frame, height=8, width=60)
        self.message_entry.grid(row=4, column=0, columnspan=2, padx=5, pady=5)
        
        # Attachment Frame
        self.attachment_frame = ttk.LabelFrame(self.main_frame, text="Lampiran")
        self.attachment_frame.grid(row=5, column=0, columnspan=2, padx=5, pady=5, sticky='ew')
        
        # File Attachment
        self.file_btn = ttk.Button(self.attachment_frame, text="Lampirkan File", command=self.browse_file)
        self.file_btn.grid(row=0, column=0, padx=5, pady=5)
        self.file_label = ttk.Label(self.attachment_frame, text="Tidak ada file dipilih", wraplength=300)
        self.file_label.grid(row=0, column=1, padx=5, pady=5, sticky='w')
        
        # Image Attachment
        self.image_btn = ttk.Button(self.attachment_frame, text="Lampirkan Gambar", command=self.browse_image)
        self.image_btn.grid(row=1, column=0, padx=5, pady=5)
        self.image_label = ttk.Label(self.attachment_frame, text="Tidak ada gambar dipilih")
        self.image_label.grid(row=1, column=1, padx=5, pady=5, sticky='w')
        
        # Image Preview
        self.image_preview = ttk.Label(self.attachment_frame)
        self.image_preview.grid(row=2, column=0, columnspan=2, pady=5)
        
        # Clear Attachment Button
        self.clear_btn = ttk.Button(self.attachment_frame, text="Hapus Lampiran", command=self.clear_attachments)
        self.clear_btn.grid(row=3, column=0, columnspan=2, pady=5)
        
        # Send Button
        self.send_btn = ttk.Button(self.main_frame, text="Kirim Pesan", command=self.send_messages)
        self.send_btn.grid(row=6, column=0, columnspan=2, pady=10)
        
        # Progress Bar
        self.progress = ttk.Progressbar(self.main_frame, orient='horizontal', length=500, mode='determinate')
        self.progress.grid(row=7, column=0, columnspan=2, pady=5)
        
        # Status Label
        self.status_label = ttk.Label(self.main_frame, text="Siap mengirim pesan")
        self.status_label.grid(row=8, column=0, columnspan=2)
        
        # Report Tab
        self.report_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.report_frame, text="Laporan")
        
        # Report Text
        self.report_text = tk.Text(self.report_frame, height=25, width=80)
        self.report_text.pack(pady=10, padx=10, fill='both', expand=True)
        self.report_text.insert('end', "Laporan pengiriman akan muncul di sini\n")
        self.report_text.config(state='disabled')
        
        # Configure tags for animated emojis
        self.report_text.tag_config('success', foreground='green')
        self.report_text.tag_config('warning', foreground='orange')
        self.report_text.tag_config('fail', foreground='red')
        
        # Export Button
        self.export_btn = ttk.Button(self.report_frame, text="Export Laporan ke CSV", command=self.export_report)
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
                messagebox.showerror("Error", f"Gagal memuat gambar: {str(e)}")
    
    def clear_attachments(self):
        self.file_path = ""
        self.image_path = ""
        self.file_label.config(text="Tidak ada file dipilih")
        self.image_label.config(text="Tidak ada gambar dipilih")
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
            messagebox.showinfo("Info", f"Berhasil memuat {len(self.phone_numbers)} nomor dari file CSV")
        except Exception as e:
            messagebox.showerror("Error", f"Gagal memproses file CSV: {str(e)}")
    
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
    
    def send_messages(self):
        phone_numbers = self.get_phone_numbers()
        self.message = self.message_entry.get("1.0", "end").strip()
        
        if not phone_numbers:
            messagebox.showerror("Error", "Tidak ada nomor telepon yang dimasukkan")
            return
        
        if not self.message and not self.file_path and not self.image_path:
            messagebox.showerror("Error", "Pesan atau lampiran tidak boleh kosong")
            return
        
        # Confirm before sending
        confirm_msg = f"Kirim {'pesan' if self.message else ''}"
        confirm_msg += f"{' dan ' if self.message and (self.file_path or self.image_path) else ''}"
        confirm_msg += f"{'file' if self.file_path else ''}"
        confirm_msg += f"{' dan ' if self.file_path and self.image_path else ''}"
        confirm_msg += f"{'gambar' if self.image_path else ''}"
        confirm_msg += f" ke {len(phone_numbers)} nomor?"
        
        if not messagebox.askyesno("Konfirmasi", confirm_msg):
            return
        
        # Prepare for sending
        self.report_data = []
        self.progress['maximum'] = len(phone_numbers)
        self.progress['value'] = 0
        self.status_label.config(text="Mengirim pesan...")
        self.root.update()
        
        # Calculate initial time (pywhatkit needs time to open browser)
        now = datetime.now()
        send_time = now + timedelta(minutes=1)  # Send 1 minute from now
        
        # Send messages
        success_count = 0
        fail_count = 0
        
        for i, phone in enumerate(phone_numbers):
            try:
                # Format phone number (remove any non-digit characters)
                formatted_phone = ''.join(c for c in phone if c.isdigit())
                if not formatted_phone:
                    raise ValueError("Nomor tidak valid")
                
                # Add country code if not present (assuming Indonesia +62)
                if not formatted_phone.startswith('62') and not formatted_phone.startswith('+62'):
                    if formatted_phone.startswith('0'):
                        formatted_phone = '62' + formatted_phone[1:]
                    else:
                        formatted_phone = '62' + formatted_phone
                
                # Determine what to send
                if self.image_path:
                    # Send image with optional caption
                    pywhatkit.sendwhats_image(
                        receiver=f"+{formatted_phone}",
                        img_path=self.image_path,
                        caption=self.message,
                        wait_time=15,
                        tab_close=True
                    )
                elif self.file_path:
                    # Note: pywhatkit doesn't directly support file sending, we'll use a workaround
                    # This will open the chat and you'll need to manually send the file
                    pywhatkit.sendwhatmsg(
                        phone_no=f"+{formatted_phone}",
                        message=self.message,
                        time_hour=send_time.hour,
                        time_min=send_time.minute,
                        wait_time=15,
                        tab_close=True
                    )
                    # Add delay to allow chat to open
                    time.sleep(5)
                    # Here you would normally automate file sending, but it requires additional tools
                    # For now, we'll just report that the chat was opened
                    self.report_data.append({
                        'phone': phone,
                        'status': 'Chat dibuka (file harus dikirim manual)',
                        'error': ''
                    })
                    success_count += 1
                    continue
                else:
                    # Send text message only
                    pywhatkit.sendwhatmsg(
                        phone_no=f"+{formatted_phone}",
                        message=self.message,
                        time_hour=send_time.hour,
                        time_min=send_time.minute,
                        wait_time=15,
                        tab_close=True
                    )
                
                # Update report
                self.report_data.append({
                    'phone': phone,
                    'status': 'Berhasil',
                    'error': ''
                })
                success_count += 1
                
                # Update progress
                self.progress['value'] = i + 1
                self.status_label.config(text=f"Mengirim {i+1}/{len(phone_numbers)} - Berhasil: {success_count}, Gagal: {fail_count}")
                self.root.update()
                
                # Increment time for next message
                send_time += timedelta(minutes=1)
                
                # Small delay to avoid rate limiting
                time.sleep(2)
                
            except Exception as e:
                # Update report with error
                self.report_data.append({
                    'phone': phone,
                    'status': 'Gagal',
                    'error': str(e)
                })
                fail_count += 1
                
                # Update progress
                self.progress['value'] = i + 1
                self.status_label.config(text=f"Mengirim {i+1}/{len(phone_numbers)} - Berhasil: {success_count}, Gagal: {fail_count}")
                self.root.update()
                
                # Continue with next number
                continue
        
        # Update status
        self.status_label.config(text=f"Pengiriman selesai! Berhasil: {success_count}, Gagal: {fail_count}")
        messagebox.showinfo("Selesai", f"Pengiriman selesai!\nBerhasil: {success_count}\nGagal: {fail_count}")
        
        # Update report tab
        self.update_report()
    
    def update_report(self):
        self.report_text.config(state='normal')
        self.report_text.delete('1.0', 'end')
        
        # Add summary
        success_count = sum(1 for item in self.report_data if item['status'] == 'Berhasil')
        fail_count = sum(1 for item in self.report_data if item['status'] == 'Gagal')
        manual_count = sum(1 for item in self.report_data if 'manual' in item['status'])
        
        self.report_text.insert('end', f"=== LAPORAN PENGIRIMAN ===\n\n")
        
        # Animated summary with emojis
        self.report_text.insert('end', "📊 ", 'success')
        self.report_text.insert('end', f"Total dikirim: {len(self.report_data)}\n")
        
        self.report_text.insert('end', "✅ ", 'success')
        self.report_text.insert('end', f"Berhasil: {success_count}\n")
        
        self.report_text.insert('end', "⚠️ ", 'warning')
        self.report_text.insert('end', f"Perlu tindakan manual: {manual_count}\n")
        
        self.report_text.insert('end', "❌ ", 'fail')
        self.report_text.insert('end', f"Gagal: {fail_count}\n\n")
        
        self.report_text.insert('end', f"=== DETAIL PENGIRIMAN ===\n\n")
        
        # Add details with animated emojis
        for item in self.report_data:
            if item['status'] == 'Berhasil':
                emoji = "😊"  # Happy face for success
                tag = 'success'
            elif 'manual' in item['status']:
                emoji = "😐"  # Neutral face for manual
                tag = 'warning'
            else:
                emoji = "😢"  # Sad face for failure
                tag = 'fail'
            
            self.report_text.insert('end', f"{emoji} ", tag)
            self.report_text.insert('end', f"{item['phone']} - {item['status']}")
            if item['error']:
                self.report_text.insert('end', f" (Error: {item['error']})")
            self.report_text.insert('end', "\n")
        
        # Add final animated celebration if all successful
        if fail_count == 0 and manual_count == 0 and success_count > 0:
            self.report_text.insert('end', "\n🎉 SELAMAT! Semua pesan terkirim sukses! 🎉\n", 'success')
        elif fail_count > 0:
            self.report_text.insert('end', "\n😔 Beberapa pesan gagal terkirim\n", 'fail')
        
        self.report_text.config(state='disabled')
        self.notebook.select(self.report_frame)  # Switch to report tab
        
        # Animate the emojis by cycling through similar emojis
        self.animate_emojis()

    def animate_emojis(self):
        # Only animate if the report tab is visible
        if self.notebook.index(self.notebook.select()) != 1:
            return
        
        # Get all emoji positions
        emoji_positions = []
        text = self.report_text.get("1.0", "end")
        
        # Common emojis used in the report
        target_emojis = ["😊", "😢", "😐", "✅", "❌", "⚠️", "🎉", "📊"]
        
        # Find all emoji positions
        start_idx = "1.0"
        while True:
            found = False
            for emoji in target_emojis:
                pos = self.report_text.search(emoji, start_idx, stopindex="end")
                if pos:
                    found = True
                    emoji_positions.append(pos)
                    start_idx = pos + "+1c"
                    break
            if not found:
                break
        
        # Cycle emojis for animation effect
        for pos in emoji_positions:
            emoji = self.report_text.get(pos, pos + "+1c")
            if emoji == "😊":
                new_emoji = "😄"
            elif emoji == "😄":
                new_emoji = "😁"
            elif emoji == "😁":
                new_emoji = "😊"
            elif emoji == "😢":
                new_emoji = "😭"
            elif emoji == "😭":
                new_emoji = "😢"
            elif emoji == "😐":
                new_emoji = "😶"
            elif emoji == "😶":
                new_emoji = "😐"
            elif emoji == "✅":
                new_emoji = "✔️"
            elif emoji == "✔️":
                new_emoji = "✅"
            elif emoji == "❌":
                new_emoji = "✖️"
            elif emoji == "✖️":
                new_emoji = "❌"
            elif emoji == "⚠️":
                new_emoji = "❗"
            elif emoji == "❗":
                new_emoji = "⚠️"
            elif emoji == "🎉":
                new_emoji = "✨"
            elif emoji == "✨":
                new_emoji = "🎊"
            elif emoji == "🎊":
                new_emoji = "🎉"
            elif emoji == "📊":
                new_emoji = "📈"
            elif emoji == "📈":
                new_emoji = "📊"
            else:
                continue
            
            # Replace the emoji
            self.report_text.config(state='normal')
            self.report_text.delete(pos, pos + "+1c")
            self.report_text.insert(pos, new_emoji, 'success' if new_emoji in ["😊", "😄", "😁", "✅", "✔️", "🎉", "✨", "🎊"] else 
                                  'warning' if new_emoji in ["😐", "😶", "⚠️", "❗"] else 'fail')
            self.report_text.config(state='disabled')
        
        # Schedule next animation
        self.root.after(500, self.animate_emojis)
    
    def export_report(self):
        if not self.report_data:
            messagebox.showwarning("Peringatan", "Tidak ada data laporan untuk diexport")
            return
        
        file_path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV Files", "*.csv")],
            title="Simpan Laporan"
        )
        
        if not file_path:
            return
        
        try:
            with open(file_path, 'w', newline='', encoding='utf-8') as file:
                writer = csv.writer(file)
                # Write header
                writer.writerow(['Nomor Telepon', 'Status', 'Error'])
                # Write data
                for item in self.report_data:
                    writer.writerow([item['phone'], item['status'], item['error']])
            
            messagebox.showinfo("Sukses", f"Laporan berhasil disimpan ke:\n{file_path}")
        except Exception as e:
            messagebox.showerror("Error", f"Gagal menyimpan laporan: {str(e)}")

if __name__ == "__main__":
    root = tk.Tk()
    app = WhatsAppSenderApp(root)
    root.mainloop()