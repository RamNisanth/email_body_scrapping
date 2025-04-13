import imaplib
import email
from email.header import decode_header
import re
from bs4 import BeautifulSoup
from openpyxl import Workbook

# Login credentials
email_user = "ramnisanth1999@gmail.com"
email_password = "oglv ipeo psgb smah"

# Connect to the Gmail IMAP server
imap = imaplib.IMAP4_SSL("imap.gmail.com")

def fetch_and_process_emails(imap, email_user, email_password, sender_email_list, output_file):
    try:
        imap.login(email_user, email_password)
        imap.select("inbox")

        wb = Workbook()
        ws = wb.active

        for sender_email in sender_email_list:
            status, messages = imap.search(None, f'FROM "{sender_email}"')

            if status == "OK":
                email_ids = messages[0].split()
                for email_id in email_ids:
                    res, msg = imap.fetch(email_id, "(RFC822)")

                    for response_part in msg:
                        if isinstance(response_part, tuple):
                            msg = email.message_from_bytes(response_part[1])

                            # Extract plain text content
                            body = None
                            if msg.is_multipart():
                                for part in msg.walk():
                                    content_type = part.get_content_type()
                                    content_disposition = str(part.get("Content-Disposition"))

                                    if content_type == "text/plain" and "attachment" not in content_disposition:
                                        body = part.get_payload(decode=True).decode("utf-8", errors="ignore")
                                        break
                            else:
                                body = msg.get_payload(decode=True).decode("utf-8", errors="ignore")

                            # Extract HTML content if present
                            html_content = None
                            for part in msg.walk():
                                if part.get_content_type() == "text/html":
                                    html_content = part.get_payload(decode=True).decode("utf-8", errors="ignore")
                                    break

                            text = body if body else "No Text Content Found"

                            # Parse HTML safely
                            if html_content:
                                try:
                                    soup = BeautifulSoup(html_content, "html.parser")  # Use HTML parser
                                    text = soup.get_text(separator=" ").strip()
                                except Exception as e:
                                    print(f"Error parsing HTML: {e}")
                                    text = "HTML Parsing Error"

                            # Store extracted text in Excel (only email content in column A)
                            ws.append([text])
            else:
                print(f"No emails found from {sender_email}.")

    except Exception as e:
        print(f"Error: {e}")

    finally:
        wb.save(output_file)
        imap.logout()


########################### place all your email ids here ########################################

email_list = ["Amex_Recruiting_AXP@invalidemail.com", "orderstatus@costco.com", "discover@services.discover.com", "Jacobs.Recruitment@jacobs.global","jobalerts-noreply@linkedin.com", "eoja.fa.sender@workflow.email.ap-sydney-1.ocs.oraclecloud.com" ]
fetch_and_process_emails(imap, email_user, email_password, email_list, "data.xlsx")
