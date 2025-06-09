# DMARC Reporter

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
# Load environment variables from .env file
def load_config():
	load_dotenv() 
	# config returns a dictionary with the following keys:
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
# Establish a secure IMAP connection to the mail server and logs in
def connect_imap(server, port, email_addr, password):
	mail = imaplib.IMAP4_SSL(server, port) # Authenticated IMAP4_SSL object
	mail.login(email_addr, password)
	print()
	print("IMAP login successful.")
	return mail

# -----------------------------------------------------------------------------
# Searches for emails in the specified IMAP folder received within the last # of days
def search_recent_emails(mail, folder="INBOX", days=7):
	mail.select("DMARC") # Select the DMARC folder inside INBOX
	date_since = (datetime.datetime.now() -
				  datetime.timedelta(days=7)).strftime("%d-%b-%Y")
	status, message_numbers = mail.search(None, f'SINCE {date_since}')
	ids = message_numbers[0].split()
	return ids # Return a list of email IDs (as bytes) received within the date range

# -----------------------------------------------------------------------------
# Create a dictionary for saving attachments.
def make_save_dir(base="attachments"):
	current_date = datetime.datetime.now().strftime("%Y%m%d")
	save_dir = os.path.join("attachments", current_date)
	os.makedirs(save_dir, exist_ok=True)
	return save_dir # Full path to the directory for today's date

# -----------------------------------------------------------------------------
# Download and save attachments from a list of email IDs in the spcified directory
def save_attachments(mail, ids, save_dir):
	# For each email ID, fetch full email message with RFC822 protocol
	for num in ids:
		status, msg_data = mail.fetch(num, '(RFC822)')
		for response_part in msg_data:
			if isinstance(response_part, tuple):
				msg = email.message_from_bytes(response_part[1])
				for part in msg.walk(): # Traverse all the branches of the msg.
					if part.get_content_maintype() == 'multipart': # skip multipart containers (not attachments)
						continue
					if part.get('Content-Disposition') is None: # Check for 'Content-Disposition' header to identify attachments
						continue
					filename = part.get_filename()
					if filename: # if fileName is present, save attachment to save_dir.
						filepath = os.path.join(save_dir, filename)
						with open(filepath, 'wb') as f:
							f.write(part.get_payload(decode=True))
						print(f"Saved attachment: {filepath}")

# -----------------------------------------------------------------------------
# Unzip all .zip and .gz files found in the specified directory (save_dir)
# Place extracted files into subdirectory called 'unzipped'
def unzip_files(save_dir):
	unzipped_dir = os.path.join(save_dir, "unzipped")
	os.makedirs(unzipped_dir, exist_ok=True)
	for filename in os.listdir(save_dir): # Iterate over all files in save_dir
		file_path = os.path.join(save_dir, filename)
		if os.path.isdir(file_path): # If file is directory, skip
			continue
		if filename.lower().endswith(".zip"): # Extract contents to unzipped_dir
			try:
				with zipfile.ZipFile(file_path, 'r') as zip_ref:
					zip_ref.extractall(unzipped_dir)
				print()
				print(f"Unzipped {filename} to {unzipped_dir}")
			except Exception as e:
				print()
				print(f"Failed to unzip {filename}: {e}")
		elif filename.lower().endswith(".gz"): # Extract contents to unzipped_dir
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
# Parse all DMARC XML files in the specified directory, extract relevant information
# and write the results to an excel file (overwrites if preexisting) 
def parse_dmarc_directory(unzipped_dir, report_dir, date_str):
	os.makedirs(report_dir, exist_ok=True)
	all_records = [] # Create dictionary 
	for filename in os.listdir(unzipped_dir): # Iterate through all files in unzipped_dir
		file_path = os.path.join(unzipped_dir, filename)
		if os.path.isdir(file_path): # Skip subdirectories 
			continue
		if filename.lower().endswith(".xml"):
			try:
				tree = ET.parse(file_path)
				root = tree.getroot()
				for record in root.findall(".//record"):
					row = {
						'source_ip': record.findtext('./row/source_ip'), # IP address source of DMARC record
						'count': record.findtext('./row/count'), # Number of messages for this record
						'disposition': record.findtext('./row/policy_evaluated/disposition'), # DMARC Policy result (None - no action, quarantine - move to spam, reject - rejected the email)
						'dkim_result': record.findtext('./row/policy_evaluated/dkim'), # DKIM Evaluation result - Check if message is signed using a valid key and if the domain in the DKIM signature
																					   # (d=) or SPF record matches the domain in the "From" address of the email
						'spf_result': record.findtext('./row/policy_evaluated/spf') # SPF Evaluation Result - Checks if the email server sending the message is authorized by the domain to send emails on its behalf
					}
					all_records.append(row) # Append each record as dictionary to all_records
			except Exception as e:
				print(f"Failed to parse {filename}: {e}")

		if all_records:
			df = pd.DataFrame(all_records) # Converts all_records to a DataFrame 
			excel_path = os.path.join(report_dir, f"dmarc_report_{date_str}.xlsx")
			df.to_excel(excel_path, index=False) # Write DataFrame to Excel file.
			print(f"\nDMARC report written to {excel_path}")
		else:
			print("No DMARC records were found.")

# -----------------------------------------------------------------------------
# Send the DMARC Excel report as an email attachment
def emailReport():
	# Load the config, .env file contains all data to send email
	config = load_config()
	smtp_server = os.getenv("SMTP_SERVER")
	smtp_port = int(os.getenv("SMTP_PORT", 587))
	from_email = config["FROM_EMAIL"]
	to_email = config["TO_EMAIL"]
	email = config["EMAIL"]
	password = config["PASSWORD"]

	# Prepare date strings for the report period 
	current_date = datetime.datetime.now().strftime("%Y%m%d")
	prev_date = (datetime.datetime.now() - datetime.timedelta(days=7)).strftime("%Y%m%d")
	report_dir = os.path.join(os.getcwd(), "Dmarc_Reports")
	excel_filename = f"dmarc_report_{current_date}.xlsx"
	excel_path = os.path.join(report_dir, excel_filename)

	# Check if the report file exists
	if not os.path.exists(excel_path):
		print(f"Report file not found: {excel_path}")
		return

	# Construct email message
	msg = EmailMessage()
	msg["Subject"] = f"DMARC Report - {prev_date} - {current_date}."
	msg["To"] = to_email
	msg["From"] = from_email
	msg.set_content(f"Attached is the DMARC report for:\n {prev_date} - {current_date}")

	# Determine the MIME type for the attachment
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
			server.quit()
		print(f"Report emailed to {to_email} from {from_email}")
	except Exception as e:
		print(f"Failed to send email: {e}")
	

# -----------------------------------------------------------------------------
# -----------------------------------------------------------------------------
# Main execution function for the DMARC processing workflow
def main():

	# Load config to open email and download requested files
	config = load_config()
	print("EMAIL: " + config["EMAIL"])
	print("TO EMAIL: " + config["TO_EMAIL"])
	print("IMAP_SERVER: " + config["IMAP_SERVER"])
	print("IMAP_SSL_PORT: " + str(config["IMAP_SSL_PORT"]))

	try:
		# Connect to the IMAP server and authenticate.
		mail = connect_imap(
				config["IMAP_SERVER"],
				config["IMAP_SSL_PORT"],
				config["EMAIL"],
				config["PASSWORD"]
		)
		
		# Check last 7 days of emails in DMARC folder
		ids = search_recent_emails(mail, folder="DMARC", days=7)
		save_dir = make_save_dir(base="attachments")
		
		# Save and unzip all attachments
		save_attachments(mail, ids, save_dir)
		unzip_files(save_dir)
		
		# Grab current date and parse directory for report generation
		current_date = datetime.datetime.now().strftime("%Y%m%d")
		unzipped_dir = os.path.join(save_dir, "unzipped")
		report_dir = os.path.join(os.getcwd(), "Dmarc_Reports")
		parse_dmarc_directory(unzipped_dir, report_dir, current_date)
		
		# Send email based on .env values
		emailReport()

		# Logout from IMAP server
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
