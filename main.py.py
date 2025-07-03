import pandas as pd
import smtplib
import tkinter as tk
from tkinter import filedialog, messagebox
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.mime.base import MIMEBase
from email import encoders
import os
import re
import requests

# Function to get YouTube thumbnail URL from video link
def get_youtube_thumbnail_url(link):
    youtube_regex = r"(https?://)?(www\.)?(youtube|youtu|youtube-nocookie)\.(com|be)/(watch\?v=|embed/|v/|.+\?v=)?([^&=%\?]{11})"
    match = re.match(youtube_regex, link)
    if match:
        video_id = match.group(6)
        return f"https://img.youtube.com/vi/{video_id}/0.jpg", link
    return None, None

# Function to send emails with embedded YouTube thumbnail as a clickable link if present
def send_emails():
    myemail = email_entry.get()
    app_pass = password_entry.get()
    sub = subject_entry.get()
    message = body_text.get("1.0", tk.END)
    hyperlink = link_entry.get()
    image_path = image_path_var.get()
    attachments = attachment_paths_var.get().split("; ")

    if hyperlink:
        thumbnail_url, video_url = get_youtube_thumbnail_url(hyperlink)

    file_path = filedialog.askopenfilename(title="Select Excel File", filetypes=(("Excel files", "*.xlsx"),))
    if not file_path:
        messagebox.showwarning("No file selected", "Please select an Excel file with email addresses.")
        return

    data = pd.read_excel(file_path)
    emails = data["Emails"].values

    # Check for unsubscribed emails
    unsubscribed_emails = []
    if os.path.exists("unsubscribed_emails.txt"):
        with open("unsubscribed_emails.txt", "r") as file:
            unsubscribed_emails = file.read().splitlines()

    try:
        server = smtplib.SMTP("smtp.gmail.com", 587)
        server.starttls()
        server.login(myemail, app_pass)
    except Exception as e:
        messagebox.showerror("Login Failed", f"Could not log in to email. Error: {e}")
        return

    for email in emails:
        if email in unsubscribed_emails:
            continue  # Skip unsubscribed emails

        msg = MIMEMultipart()
        msg['From'] = myemail
        msg['To'] = email
        msg['Subject'] = sub

        # Add unsubscribe link to the email body at the bottom
        unsubscribe_link = f"<p style='color: #FFFFFF; font-weight: bold;'>If you no longer wish to receive emails, please click here to unsubscribe: <a href='mailto:{myemail}?subject=Unsubscribe' style='color: #007bff;'>Unsubscribe</a></p>"

        html_content = f"""
        <html>
        <body style="font-family: Arial, sans-serif; color: #333;">
            <p>{message.replace('\n', '<br>')}</p>
            {unsubscribe_link}
        """

        # Add YouTube thumbnail with a styled, interactive "Watch Video" link
        if thumbnail_url:
            try:
                response = requests.get(thumbnail_url)
                img_data = response.content
                img = MIMEImage(img_data)
                img.add_header('Content-ID', '<youtube_thumbnail>')
                msg.attach(img)
                # Embed thumbnail as a clickable link
                html_content += f'''
                <br>
                <a href="{video_url}" style="text-decoration: none;">
                    <img src="cid:youtube_thumbnail" style="border: none; width: 100%; max-width: 600px;">
                </a>
                <br>
                <p style="text-align: center; font-size: 16px; font-weight: bold; color: #007bff;">
                    <a href="{video_url}" style="color: #007bff; text-decoration: none;">Watch Video</a>
                </p>
                '''
            except Exception as e:
                messagebox.showerror("Thumbnail Error", f"Could not fetch YouTube thumbnail. Error: {e}")

        html_content += "</body></html>"
        msg.attach(MIMEText(html_content, 'html'))

        for file_path in attachments:
            if file_path:
                try:
                    attachment = MIMEBase('application', 'octet-stream')
                    with open(file_path, 'rb') as file:
                        attachment.set_payload(file.read())
                    encoders.encode_base64(attachment)
                    attachment.add_header('Content-Disposition', f'attachment; filename={os.path.basename(file_path)}')
                    msg.attach(attachment)
                except Exception as e:
                    messagebox.showerror("Attachment Error", f"Could not attach {file_path}. Error: {e}")

        try:
            server.sendmail(myemail, email, msg.as_string())
        except Exception as e:
            messagebox.showerror("Send Error", f"Failed to send email to {email}. Error: {e}")

    messagebox.showinfo("Success", "All Emails Sent Successfully!")
    server.quit()

# Function to select image
def select_image():
    image_path = filedialog.askopenfilename(title="Select an Image", filetypes=(("Image files", "*.jpg;*.jpeg;*.png"),))
    if image_path:
        image_path_var.set(image_path)

# Function to select attachments
def select_attachments():
    files = filedialog.askopenfilenames(title="Select Files to Attach")
    attachment_paths_var.set("; ".join(files))

# GUI setup
root = tk.Tk()
root.title("Business Email Sender")
root.geometry("800x600")
root.configure(bg="#e9ecef")

# Center alignment configuration
root.grid_columnconfigure(0, weight=1)
root.grid_columnconfigure(1, weight=1)
root.grid_columnconfigure(2, weight=1)

# Define Styles
label_style = {"font": ("Arial", 12), "bg": "#e9ecef"}
entry_style = {"font": ("Arial", 12), "relief": "solid", "bd": 2}

# Email entry field at the top
tk.Label(root, text="Enter Your Email:", **label_style).grid(row=0, column=1, padx=10, pady=10, sticky="w")
email_entry = tk.Entry(root, width=40, **entry_style)
email_entry.grid(row=0, column=1, padx=10, pady=10)

# App Password
tk.Label(root, text="App Password:", **label_style).grid(row=1, column=1, padx=10, pady=10, sticky="w")
password_entry = tk.Entry(root, show="*", width=40, **entry_style)
password_entry.grid(row=1, column=1, padx=10, pady=10)

# Subject
tk.Label(root, text="Email Subject:", **label_style).grid(row=2, column=1, padx=10, pady=10, sticky="w")
subject_entry = tk.Entry(root, width=40, **entry_style)
subject_entry.grid(row=2, column=1, padx=10, pady=10)

# Body
tk.Label(root, text="Email Body:", **label_style).grid(row=3, column=1, padx=10, pady=10, sticky="nw")
body_text = tk.Text(root, width=40, height=10, font=("Arial", 12), relief="solid", bd=2)
body_text.grid(row=3, column=1, padx=10, pady=10, sticky="nsew")

# Add hyperlink field
tk.Label(root, text="Add Hyperlink:", **label_style).grid(row=4, column=1, padx=10, pady=10, sticky="w")
link_entry = tk.Entry(root, width=40, **entry_style)
link_entry.grid(row=4, column=1, padx=10, pady=10)

# Attach Image
tk.Label(root, text="Attach Image:", **label_style).grid(row=5, column=1, padx=10, pady=10, sticky="w")
image_path_var = tk.StringVar()
tk.Entry(root, textvariable=image_path_var, width=40, state="readonly", **entry_style).grid(row=5, column=1, padx=10, pady=10)
tk.Button(root, text="Browse", command=select_image, font=("Arial", 10), bg="#007bff", fg="white").grid(row=5, column=2, padx=10, pady=10)

# Attach Files
tk.Label(root, text="Attach Files:", **label_style).grid(row=6, column=1, padx=10, pady=10, sticky="w")
attachment_paths_var = tk.StringVar()
tk.Entry(root, textvariable=attachment_paths_var, width=40, state="readonly", **entry_style).grid(row=6, column=1, padx=10, pady=10)
tk.Button(root, text="Browse", command=select_attachments, font=("Arial", 10), bg="#007bff", fg="white").grid(row=6, column=2, padx=10, pady=10)

# Send Button
send_button = tk.Button(root, text="Send Emails", command=send_emails, font=("Arial", 12, 'bold'), bg="#28a745", fg="white", relief="raised")
send_button.grid(row=7, column=1, pady=20)

# Run the GUI loop
root.mainloop()
