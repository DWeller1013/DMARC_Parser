import email
import datetime
import imaplib
import zipfile
import gzip
import shutil
import pandas as pd
import xml.etree.ElementTree as ET
import os
from email.header import decode_header
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
# -----------------------------------------------------------------------------
def main():
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
		
		save_attachments(mail, ids, save_dir)
		unzip_files(save_dir)
		
		mail.logout()
		print()
		print("IMAP logout successful.")
		
		current_date = datetime.datetime.now().strftime("%Y%m%d")
		unzipped_dir = os.path.join(save_dir, "unzipped")
		report_dir = os.path.join(os.getcwd(), "Dmarc_Reports")
		parse_dmarc_directory(unzipped_dir, report_dir, current_date)

	except Exception as e:
		print()
		print(f"IMAP login failed: {e}")

# -----------------------------------------------------------------------------

if __name__ == "__main__":
	main()

# -----------------------------------------------------------------------------
