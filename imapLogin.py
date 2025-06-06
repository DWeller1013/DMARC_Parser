import email
import datetime
import imaplib
import os
from email.header import decode_header
from dotenv import load_dotenv

load_dotenv()
EMAIL = os.getenv("EMAIL")
PASSWORD = os.getenv("PASSWORD")
IMAP_SERVER = os.getenv("IMAP_SERVER")
IMAP_SSL_PORT = int(os.getenv("IMAP_SSL_PORT"))
TO_EMAIL = os.getenv("TO_EMAIL")

try:
	# Connect to the IMAP server
	mail = imaplib.IMAP4_SSL(IMAP_SERVER, IMAP_SSL_PORT)

	# Login to account
	mail.login(EMAIL, PASSWORD)
	print("IMAP login successful.")
	mail.select("DMARC")

	# status, message_count = mail.select("DMARC")
	# print(f"Number of messages in DMARC: {message_count[0].decode()}")

	# Grab emails in last 7 days
	current_date = datetime.datetime.now().strftime("%Y%m%d")
	date_since = (datetime.datetime.now() -
				  datetime.timedelta(days=7)).strftime("%d-%b-%Y")
	status, message_numbers = mail.search(None, f'SINCE {date_since}')
	ids = message_numbers[0].split()

	#print("Searching for messages since:", date_since)
	#print(f"Found {len(ids)} messages since {date_since}")

	save_dir = os.path.join("attachments", current_date)
	os.makedirs(save_dir, exist_ok=True)
	
	# Download all attachments found
	for num in ids:
		status, msg_data = mail.fetch(num, '(RFC822)')
		for response_part in msg_data:
			if isinstance(response_part, tuple):
				msg = email.message_from_bytes(response_part[1])
				for part in msg.walk():
					if part.get_content_maintype() == 'multipart':
						continue
					if part.get('Content-Disposition') is None:
						continue
					filename = part.get_filename()
					if filename:
						filepath = os.path.join(save_dir, filename)
						with open(filepath, 'wb') as f:
							f.write(part.get_payload(decode=True))
						print(f"Saved attachment: {filepath}")

	# status, mailboxes = mail.list()
	# print("Available mailboxes:")
	# for mbox in mailboxes:
	# print(mbox.decode())

	mail.logout()
except Exception as e:
	print(f"IMAP login failed: {e}")

print("")
print("EMAIL: " + EMAIL)
print("TO EMAIL: " + TO_EMAIL)
print("IMAP_SERVER: " + IMAP_SERVER)
print("IMAP_SSL_PORT: " + str(IMAP_SSL_PORT))
