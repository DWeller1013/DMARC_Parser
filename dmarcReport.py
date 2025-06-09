import email
import datetime
import imaplib
import smtplib
import zipfile
import gzip
import shutil
import mimetypes
import pandas as pd
import xml.etree.ElementTree as ET
import os
from email.header import decode_header
from email.message import EmailMessage
from dotenv import load_dotenv

# -----------------------------------------------------------------------------
def load_config():
	load_dotenv()
	config = {
		"EMAIL": os.getenv("EMAIL"),
		"PASSWORD": os.getenv("PASSWORD"),
		"IMAP_SERVER": os.getenv("IMAP_SERVER"),
		"IMAP_SSL_PORT": int(os.getenv("IMAP_SSL_PORT")),
		"TO_EMAIL": os.getenv("TO_EMAIL"),
		"FROM_EMAIL": os.getenv("FROM_EMAIL"),
	}
	return config

# -----------------------------------------------------------------------------
def connect_imap(server, port, email_addr, password):
	mail = imaplib.IMAP4_SSL(server, port)
	mail.login(email_addr, password)
	print()
	print("IMAP login successful.")
	return mail

# -----------------------------------------------------------------------------
def search_recent_emails(mail, folder="INBOX", days=7):
	mail.select("DMARC")
	date_since = (datetime.datetime.now() -
				  datetime.timedelta(days=7)).strftime("%d-%b-%Y")
	status, message_numbers = mail.search(None, f'SINCE {date_since}')
	ids = message_numbers[0].split()
	return ids

# -----------------------------------------------------------------------------
def make_save_dir(base="attachments"):
	current_date = datetime.datetime.now().strftime("%Y%m%d")
	save_dir = os.path.join("attachments", current_date)
	os.makedirs(save_dir, exist_ok=True)
	return save_dir

# -----------------------------------------------------------------------------
def save_attachments(mail, ids, save_dir):
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

# -----------------------------------------------------------------------------
def unzip_files(save_dir):
	unzipped_dir = os.path.join(save_dir, "unzipped")
	os.makedirs(unzipped_dir, exist_ok=True)
	for filename in os.listdir(save_dir):
		file_path = os.path.join(save_dir, filename)
		if os.path.isdir(file_path):
			continue
		if filename.lower().endswith(".zip"):
			try:
				with zipfile.ZipFile(file_path, 'r') as zip_ref:
					zip_ref.extractall(unzipped_dir)
				print()
				print(f"Unzipped {filename} to {unzipped_dir}")
			except Exception as e:
				print()
				print(f"Failed to unzip {filename}: {e}")
		elif filename.lower().endswith(".gz"):
			try:
				out_name = filename[:-3] # Remove .gz
				out_path = os.path.join(unzipped_dir, out_name)
				with gzip.open(file_path, 'rb') as f_in:
					with open(out_path, 'wb') as f_out:
						shutil.copyfileobj(f_in, f_out)
				print()
				print(f"Unzipped {filename} to {out_path}")
			except Exception as e:
				print()
				print(f"Failed to unzip {filename}: {e}")

# -----------------------------------------------------------------------------
def parse_dmarc_directory(unzipped_dir, report_dir, date_str):
	os.makedirs(report_dir, exist_ok=True)
	all_records = []
	for filename in os.listdir(unzipped_dir):
		file_path = os.path.join(unzipped_dir, filename)
		if os.path.isdir(file_path):
			continue
		if filename.lower().endswith(".xml"):
			try:
				tree = ET.parse(file_path)
				root = tree.getroot()
				for record in root.findall(".//record"):
					row = {
						'source_ip': record.findtext('./row/source_ip'),
						'count': record.findtext('./row/count'),
						'disposition': record.findtext('./row/policy_evaluated/disposition'),
						'dkim_result': record.findtext('./row/policy_evaluated/dkim'),
						'spf_result': record.findtext('./row/policy_evaluated/spf')
					}
					all_records.append(row)
			except Exception as e:
				print(f"Failed to parse {filename}: {e}")

		if all_records:
			df = pd.DataFrame(all_records)
			excel_path = os.path.join(report_dir, f"dmarc_report_{date_str}.xlsx")
			df.to_excel(excel_path, index=False)
			print(f"\nDMARC report written to {excel_path}")
		else:
			print("No DMARC records were found.")

# -----------------------------------------------------------------------------
def emailReport():
	# Load the config, .env file contains all data to send email
	config = load_config()
	smtp_server = os.getenv("SMTP_SERVER")
	smtp_port = int(os.getenv("SMTP_PORT", 587))
	from_email = config["FROM_EMAIL"]
	to_email = config["TO_EMAIL"]
	email = config["EMAIL"]
	password = config["PASSWORD"]

	# Grab data to populate email
	current_date = datetime.datetime.now().strftime("%Y%m%d")
	prev_date = (datetime.datetime.now() - datetime.timedelta(days=7)).strftime("%Y%m%d")
	report_dir = os.path.join(os.getcwd(), "Dmarc_Reports")
	excel_filename = f"dmarc_report_{current_date}.xlsx"
	excel_path = os.path.join(report_dir, excel_filename)

	if not os.path.exists(excel_path):
		print(f"Report file not found: {excel_path}")
		return

	# Construct email message
	msg = EmailMessage()
	msg["Subject"] = f"DMARC Report - {prev_date} - {current_date}."
	msg["To"] = to_email
	msg["From"] = from_email
	msg.set_content(f"Attached is the DMARC report for {current_date}")

	ctype, encoding = mimetypes.guess_type(excel_path)
	if ctype is None or encoding is not None:
		ctype = "application/octet-stream" # Default for binary files with unknown formats
	maintype, subtype = ctype.split("/", 1)
	with open(excel_path, "rb") as f: # rb = Read/Binary
		msg.add_attachment(f.read(), maintype=maintype, subtype=subtype, filename=excel_filename)

	try:
		with smtplib.SMTP(smtp_server, smtp_port) as server:
			server.starttls()
			#server.login(email, password)
			server.send_message(msg)
		print(f"Report emailed to {to_email} from {from_email}")
	except Exception as e:
		print(f"Failed to send email: {e}")

# -----------------------------------------------------------------------------
# -----------------------------------------------------------------------------
def main():

	# Load config to open email and download requested files
	config = load_config()
	print("EMAIL: " + config["EMAIL"])
	print("TO EMAIL: " + config["TO_EMAIL"])
	print("IMAP_SERVER: " + config["IMAP_SERVER"])
	print("IMAP_SSL_PORT: " + str(config["IMAP_SSL_PORT"]))

	try:
		mail = connect_imap(
				config["IMAP_SERVER"],
				config["IMAP_SSL_PORT"],
				config["EMAIL"],
				config["PASSWORD"]
		)

		ids = search_recent_emails(mail, folder="DMARC", days=7)
		save_dir = make_save_dir(base="attachments")
		
		# Save and unzip all attachments
		save_attachments(mail, ids, save_dir)
		unzip_files(save_dir)
		
		# Grab current date and parse directory
		current_date = datetime.datetime.now().strftime("%Y%m%d")
		unzipped_dir = os.path.join(save_dir, "unzipped")
		report_dir = os.path.join(os.getcwd(), "Dmarc_Reports")
		parse_dmarc_directory(unzipped_dir, report_dir, current_date)
		
		# Send email based on .env values
		emailReport()

		mail.logout()
		print()
		print("IMAP logout successful.")

	except Exception as e:
		print()
		print(f"IMAP login failed: {e}")

# -----------------------------------------------------------------------------

if __name__ == "__main__":
	main()

# -----------------------------------------------------------------------------
