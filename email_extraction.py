import imaplib
import email
from email.header import decode_header
import re
from bs4 import BeautifulSoup
from openpyxl import Workbook

wb = Workbook()
ws = wb.active



# Login credentials
email_user = "YOUR_EMAIL@COMPANY.COM"
email_password = " EN TER YOU R APP PASS WORD"

# Connect to the Gmail IMAP server
imap = imaplib.IMAP4_SSL("imap.gmail.com")

try:
    imap.login(email_user, email_password)
    imap.select("inbox")
    sender_email = "freetier@costalerts.amazonaws.com"
    status, messages = imap.search(None, f'FROM "{sender_email}"')

    if status == "OK":
        email_ids = messages[0].split()  
        for email_id in email_ids:
            res, msg = imap.fetch(email_id, "(RFC822)")
            for response_part in msg:
                if isinstance(response_part, tuple):
                    msg = email.message_from_bytes(response_part[1])
                    match = re.search(r"<!DOCTYPE.*?</html>", str(msg), re.DOTALL | re.IGNORECASE)
                    if match:
                      extracted_html = match.group(0) 
                      soup = BeautifulSoup(extracted_html, "html.parser")
                      text = soup.get_text(separator=" ")  
                      print(text)
                      ws.append([text])
                    else:
                      print("No HTML content found.")
                    email_subject = msg["Subject"]
                    email_date = msg["Date"]
                    body = None
                    content_type = "Unknown"

                    if msg.is_multipart():
                        for part in msg.walk():
                            content_type = part.get_content_type()
                            content_disposition = str(part.get("Content-Disposition"))

                            if content_type == "text/plain" and "attachment" not in content_disposition:
                                body = part.get_payload(decode=True).decode("utf-8", errors="ignore")
                                break  
                    else:
                        content_type = msg.get_content_type()
                        body = msg.get_payload(decode=True).decode("utf-8", errors="ignore")
                    ws.append([body])

    else:
        print(f"No emails found from {sender_email}.")

except Exception as e:
    print(f"Error: {e}")

finally:
    wb.save("filename.xlsx")
    imap.logout()
