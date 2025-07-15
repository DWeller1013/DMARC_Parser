# -----------------------------------------------------------------------------
# DMARC Reporter
# -----------------------------------------------------------------------------
# Download all attachments in emails from DMARC subfolder. Unzip and extract
# all .zip and .gz files. Parse through .xml files and display data in an excel sheet.
# Organize and format excel sheet and email back to desired user.
# Eventually will set up as Crontab to send reports every Thursday morning.
# -----------------------------------------------------------------------------
# 250620(dmw) Added SPF_Failures sheet. Investigate why SPF is failing.
# 250620(dmw) Added get_org_name function. RDAP lookup and added caching. Removed graphs/charts for now.
# 250612(dmw) Added PieChart with percentages pulling from TabularData.
# 250612(dmw) Added TabularData function.
# 250609(dmw) Added Email Reporting for parsed DMARC xml files.
# -----------------------------------------------------------------------------

import email
import datetime
from email.header import decode_header
from email.message import EmailMessage
import gzip
import imaplib
import mimetypes
import pickle
import os
import shutil
import smtplib
import zipfile
import dns.resolver
import ipaddress
import re
import xml.etree.ElementTree as ET

import pandas as pd
import plotly.express as px
from dotenv import load_dotenv
from openpyxl import load_workbook
from ipwhois import IPWhois
from tqdm import tqdm
from openpyxl.chart import BarChart, Reference, PieChart
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart.label import DataLabelList
from openpyxl.drawing.image import Image as XLImage

# -----------------------------------------------------------------------------
# Globals
CURRENT_DATE = datetime.datetime.now().strftime("%Y-%m-%d")
CACHE_FILE = "dns_cache.pkl"
tqdm.pandas()

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
	global CURRENT_DATE
	save_dir = os.path.join("attachments", CURRENT_DATE)
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
				print(f"Unzipped {filename} to {unzipped_dir}")
			except Exception as e:
				print(f"Failed to unzip {filename}: {e}")
		elif filename.lower().endswith(".gz"): # Extract contents to unzipped_dir
			try:
				out_name = filename[:-3] # Remove .gz
				out_path = os.path.join(unzipped_dir, out_name)
				with gzip.open(file_path, 'rb') as f_in:
					with open(out_path, 'wb') as f_out:
						shutil.copyfileobj(f_in, f_out)
				print(f"Unzipped {filename} to {out_path}")
			except Exception as e:
				print(f"Failed to unzip {filename}: {e}")

# -----------------------------------------------------------------------------
# Format the excel sheets 
def formatSheets(excel_path):
	wb = load_workbook(excel_path)

	fixed_width_columns = {
		'spf_record': 30,
		'spf_includes': 30,
		'spf_notes': 30
	}

	# Loop through every worksheet in the workbook
	for sheet_name in wb.sheetnames:
		ws = wb[sheet_name]
		print(f"Formatting sheet: {sheet_name}")
		header_row = [cell.value for cell in ws[1]]
		col_map ={header: get_column_letter(idx + 1) for idx, header in enumerate(header_row)}
		# Loop through all columns in the worksheet 
		for i, col in enumerate(ws.iter_cols(min_row=1, max_row=ws.max_row, max_col=ws.max_column), 1):
			#col_letter = col[0].column_letter # Get the Excel-style column letter
			col_letter = get_column_letter(i)
			header = col[0].value # Get the header value from the first row
			if header in fixed_width_columns:
				ws.column_dimensions[col_letter].width = fixed_width_columns[header]
			else: 
				max_length = 0 # Track the maximum content length for that column
				# Loop through each cell in the column
				for cell in col:
					if header == 'source_ip' or header == 'envelope_to':
						cell.alignment = Alignment(horizontal='left')
					else:
						cell.alignment = Alignment(horizontal='center')

					# Calculate content length for auto-width
					cell_value = str(cell.value) if cell.value is not None else ''
					max_length = max(max_length, len(cell_value))

				# Set column width with some padding
				ws.column_dimensions[col_letter].width = max_length + 1 

	# Save changes to file
	wb.save(excel_path)
	print(f"Formatting complete for all sheets in '{excel_path}'")

# -----------------------------------------------------------------------------
# Check if old_name sheet exists in workbook and replace with new_name.
def renameSheet(excel_path, old_name, new_name):
	wb = load_workbook(excel_path)
	if old_name in wb.sheetnames:
		ws = wb[old_name] # Access the worksheet by old_name
		ws.title = new_name # Rename the worksheet
		wb.save(excel_path) # Save the workbook and apply changes
		print(f"Renamed sheet '{old_name}' to '{new_name}'")
	else:
		print(f"Sheet '{old_name}' not found, cannot rename.")


# -----------------------------------------------------------------------------
# Parse all DMARC XML files in the specified directory, extract relevant information
# and write the results to an excel file (overwrites if preexisting) 
def parse_dmarc_directory(unzipped_dir, report_dir, date_str):
	
	os.makedirs(report_dir, exist_ok=True)
	all_records = [] # Create dictionary 
	report_metadata = [] # Store metadata about each report

	for filename in os.listdir(unzipped_dir): # Iterate through all files in unzipped_dir
		file_path = os.path.join(unzipped_dir, filename)
		if os.path.isdir(file_path): # Skip subdirectories 
			continue
		if filename.lower().endswith(".xml"):
			try:
				tree = ET.parse(file_path)
				root = tree.getroot()
				
				# Extract report metadata for better context
				report_meta = root.find('report_metadata')
				policy_published = root.find('policy_published')
				
				if report_meta is not None:
					metadata = {
						'report_filename': filename,
						'org_name': report_meta.findtext('org_name', 'Unknown'),
						'email': report_meta.findtext('email', 'Unknown'),
						'published_domain': policy_published.findtext('domain', '') if policy_published is not None else '',
						'published_policy': policy_published.findtext('p', '') if policy_published is not None else '',
						'published_subdomain_policy': policy_published.findtext('sp', '') if policy_published is not None else '',
						'published_percentage': policy_published.findtext('pct', '100') if policy_published is not None else '100',
						'published_dkim_alignment': policy_published.findtext('adkim', 'r') if policy_published is not None else 'r',
						'published_spf_alignment': policy_published.findtext('aspf', 'r') if policy_published is not None else 'r'
					}
					report_metadata.append(metadata)
				
				for record in root.findall(".//record"):
					row = {
						'report_filename': filename,
						'envelope_to': record.findtext('./identifiers/envelope_to'), # Address sent to 
						'source_ip': record.findtext('./row/source_ip'), # IP address source of DMARC record
						'count': record.findtext('./row/count'), # Number of messages for this record
						'disposition': record.findtext('./row/policy_evaluated/disposition'), # DMARC Policy result (None - no action, quarantine - move to spam, reject - rejected the email)
						'dkim_result': record.findtext('./row/policy_evaluated/dkim'), # DKIM Evaluation result
						'spf_result': record.findtext('./row/policy_evaluated/spf'), # SPF Evaluation Result
						'dkim_alignment': record.findtext('./row/policy_evaluated/dkim') if record.findtext('./row/policy_evaluated/dkim') == 'pass' else 'fail',
						'spf_alignment': record.findtext('./row/policy_evaluated/spf') if record.findtext('./row/policy_evaluated/spf') == 'pass' else 'fail',
						'header_from': record.findtext('./identifiers/header_from'),
						'envelope_from': record.findtext('./identifiers/envelope_from', ''),
						'spf_checked_domain': '',
						'spf_checked_result': '',
						'spf_for_header_from': '',
						'dkim_domain': '',
						'dkim_selector': '',
						'dkim_result_detail': '',
						'policy_override_reason': record.findtext('./row/policy_evaluated/reason/type', ''),
						'policy_override_comment': record.findtext('./row/policy_evaluated/reason/comment', ''),
					}
					
					# Extract SPF authentication results
					spf_auths = record.findall('./auth_results/spf')
					spf_for_header_from = None
					for spf in spf_auths:
						this_domain = spf.findtext('domain')
						this_result = spf.findtext('result')
						if this_domain and this_result:
							if this_domain == row['header_from']:
								spf_for_header_from = this_result
							row['spf_checked_domain'] = this_domain
							row['spf_checked_result'] = this_result
					row['spf_for_header_from'] = spf_for_header_from
					
					# Extract DKIM authentication results
					dkim_auths = record.findall('./auth_results/dkim')
					if dkim_auths:
						dkim_auth = dkim_auths[0]  # Take first DKIM result
						row['dkim_domain'] = dkim_auth.findtext('domain', '')
						row['dkim_selector'] = dkim_auth.findtext('selector', '')
						row['dkim_result_detail'] = dkim_auth.findtext('result', '')
					
					all_records.append(row) # Append each record as dictionary to all_records
			except Exception as e:
				print(f"Failed to parse {filename}: {e}")

	if all_records:
		df = pd.DataFrame(all_records) # Converts all_records to a DataFrame 
		excel_path = os.path.join(report_dir, f"dmarc_report_{date_str}.xlsx")
		
		# Create workbook with multiple sheets
		with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
			df.to_excel(writer, sheet_name="All Data", index=False)
			
			# Add report metadata sheet
			if report_metadata:
				metadata_df = pd.DataFrame(report_metadata)
				metadata_df.to_excel(writer, sheet_name="Report Metadata", index=False)
		
		print(f"\nDMARC report written to {excel_path}")
		return excel_path
	else:
		print("No DMARC records were found.")
		return None

# -----------------------------------------------------------------------------
# 250620(dmw) Check each source_ip for the DNS org's name using RDAP Lookup (WHOIS) to query the IP.
def get_org_name(ip, cache):
	if ip in cache:
		return cache[ip]
	# Try WHOIS ASN description first
	try:
		obj = IPWhois(ip)
		results = obj.lookup_rdap(depth=1)
		asn_desc = results.get('asn_description')
		if asn_desc and asn_desc.strip() and asn_desc.strip() != 'Not Announced':
			cache[ip] = asn_desc.strip()
			return asn_desc.strip()
		# Fall back to netname
		netname = results.get('network', {}).get('name', '')
		if netname:
			cache[ip] = netname.strip()
			return netname.strip()
	except Exception:
		pass

	# If all looks up fail, return Unknown for that ip.
	cache[ip] = "Unknown"
	return "Unknown"

# -----------------------------------------------------------------------------
# Get geolocation information for IP addresses to provide better context
def get_ip_geolocation(ip, cache):
	cache_key = f"geo_{ip}"
	if cache_key in cache:
		return cache[cache_key]
	
	try:
		obj = IPWhois(ip)
		results = obj.lookup_rdap(depth=1)
		
		# Extract country information
		country = "Unknown"
		if 'network' in results and 'country' in results['network']:
			country = results['network']['country']
		elif 'objects' in results:
			# Look through objects for country info
			for obj_key, obj_data in results['objects'].items():
				if isinstance(obj_data, dict) and 'contact' in obj_data:
					contact = obj_data['contact']
					if isinstance(contact, dict) and 'address' in contact:
						address = contact['address']
						if isinstance(address, list) and len(address) > 0:
							addr_info = address[0]
							if isinstance(addr_info, dict) and 'value' in addr_info:
								# Simple country extraction from address
								addr_lines = addr_info['value'].split('\n')
								if len(addr_lines) > 0:
									country = addr_lines[-1].strip()
									break
		
		cache[cache_key] = country
		return country
	except Exception:
		cache[cache_key] = "Unknown"
		return "Unknown"

# -----------------------------------------------------------------------------
# Calculate risk score based on DMARC results and provide plain English assessment
def calculate_risk_score(row):
	risk_score = 0
	risk_factors = []
	
	# High risk factors
	if row.get('disposition', '').lower() in ['quarantine', 'reject']:
		risk_score += 30
		risk_factors.append(f"Emails were {row.get('disposition', '').lower()}d by DMARC policy")
	
	if row.get('spf_result', '').lower() == 'fail':
		risk_score += 25
		risk_factors.append("SPF authentication failed")
	
	if row.get('dkim_result', '').lower() == 'fail':
		risk_score += 20
		risk_factors.append("DKIM authentication failed")
	
	# Medium risk factors
	if row.get('auth_status', '').lower() == 'failed':
		risk_score += 15
		risk_factors.append("Overall authentication failed")
	
	# Volume-based risk (higher volume failures are worse)
	count = int(row.get('count', 0))
	if count > 1000:
		risk_score += 10
		risk_factors.append("High volume of failed messages")
	elif count > 100:
		risk_score += 5
		risk_factors.append("Moderate volume of failed messages")
	
	# Cap risk score at 100
	risk_score = min(risk_score, 100)
	
	# Determine risk level
	if risk_score >= 70:
		risk_level = "Critical"
	elif risk_score >= 50:
		risk_level = "High"
	elif risk_score >= 30:
		risk_level = "Medium"
	elif risk_score >= 10:
		risk_level = "Low"
	else:
		risk_level = "Minimal"
	
	return risk_score, risk_level, risk_factors

# -----------------------------------------------------------------------------
# Read all data for each row and organize into a more readable format.
# Optimized to remove slow DNS lookups for better performance.
def organizeData(excel_path):
	
	# Load cache from file - keeping for potential future use but not using slow lookups
	if os.path.exists(CACHE_FILE):
		with open(CACHE_FILE, "rb") as f:
			dns_cache = pickle.load(f)
	else:
		dns_cache = {}

	try:
		# Read the data from the specified sheet
		df = pd.read_excel(excel_path, sheet_name="All Data")

		print(f"Processing data without slow DNS lookups for better performance...")

		# Skip the slow DNS lookups that were causing 8-9 minute execution times
		# as per user requirements - source_dns and source_country columns are not useful
		
		# Add auth_status
		df['auth_status'] = df.apply(
			lambda row: 'Authenticated' if row['dkim_result'] == 'pass' or row['spf_result'] == 'pass' else 'Failed',
			axis=1
		)

		# Add risk assessment
		risk_data = df.apply(calculate_risk_score, axis=1, result_type='expand')
		df['risk_score'] = risk_data[0]
		df['risk_level'] = risk_data[1]
		df['risk_factors'] = risk_data[2].apply(lambda x: '; '.join(x) if x else 'None')

		# Add plain English explanations for common issues
		def get_plain_english_explanation(row):
			explanations = []
			
			if row['auth_status'] == 'Failed':
				if row['spf_result'] == 'fail' and row['dkim_result'] == 'fail':
					explanations.append("Both SPF and DKIM failed - emails appear suspicious to receiving servers")
				elif row['spf_result'] == 'fail':
					explanations.append("SPF failed - the sending server is not authorized to send emails for this domain")
				elif row['dkim_result'] == 'fail':
					explanations.append("DKIM failed - the email signature is invalid or missing")
			
			if row['disposition'] in ['quarantine', 'reject']:
				explanations.append(f"Receiving servers {row['disposition']}d these emails due to DMARC policy")
			
			return '; '.join(explanations) if explanations else 'No issues detected'

		df['plain_english_explanation'] = df.apply(get_plain_english_explanation, axis=1)

		# Sort by risk score (highest first), then by count
		df['spf_result'] = df['spf_result'].fillna('')
		df = df.sort_values(by=['risk_score', 'count'], ascending=[False, False])

		# Write the organized summary to the new sheet
		with pd.ExcelWriter(excel_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
			df.to_excel(writer, sheet_name="Organized_Data", index=False)

		print(f"Optimized organized data created on sheet 'Organized_Data' in '{excel_path}'.")
		
	except FileNotFoundError:
		print(f"Error: the file '{excel_path}' was not found.")
	except Exception as e:
		print(f"An error occurred: {e}")

	# Save cache to file (persist for potential future use)
	with open(CACHE_FILE, "wb") as f:
		pickle.dump(dns_cache, f)

# -----------------------------------------------------------------------------
# Create an Executive Summary sheet with key metrics and actionable insights for non-technical users
def create_executive_summary(excel_path):
	try:
		# Read organized data
		df = pd.read_excel(excel_path, sheet_name="Organized_Data")
		
		# Calculate key metrics
		total_emails = df['count'].astype(int).sum()
		total_records = len(df)
		
		# Authentication metrics
		auth_pass = df[df['auth_status'] == 'Authenticated']['count'].astype(int).sum()
		auth_fail = df[df['auth_status'] == 'Failed']['count'].astype(int).sum()
		auth_rate = (auth_pass / total_emails * 100) if total_emails > 0 else 0
		
		# SPF metrics
		spf_pass = df[df['spf_result'] == 'pass']['count'].astype(int).sum()
		spf_fail = df[df['spf_result'] == 'fail']['count'].astype(int).sum()
		spf_rate = (spf_pass / total_emails * 100) if total_emails > 0 else 0
		
		# DKIM metrics
		dkim_pass = df[df['dkim_result'] == 'pass']['count'].astype(int).sum()
		dkim_fail = df[df['dkim_result'] == 'fail']['count'].astype(int).sum()
		dkim_rate = (dkim_pass / total_emails * 100) if total_emails > 0 else 0
		
		# Disposition metrics
		none_disp = df[df['disposition'] == 'none']['count'].astype(int).sum()
		quarantine_disp = df[df['disposition'] == 'quarantine']['count'].astype(int).sum()
		reject_disp = df[df['disposition'] == 'reject']['count'].astype(int).sum()
		
		# Risk assessment
		critical_risk = len(df[df['risk_level'] == 'Critical'])
		high_risk = len(df[df['risk_level'] == 'High'])
		medium_risk = len(df[df['risk_level'] == 'Medium'])
		
		# Top source IPs by volume
		top_sources = df.groupby(['source_ip'])['count'].sum().reset_index()
		top_sources['count'] = top_sources['count'].astype(int)
		top_sources = top_sources.sort_values('count', ascending=False).head(10)
		
		# Create enhanced summary data with targets
		summary_data = {
			'Metric': [
				'Total Email Volume',
				'Total Unique Records',
				'🎯 DMARC Authentication Rate',
				'🎯 SPF Pass Rate', 
				'🎯 DKIM Pass Rate',
				'Emails Allowed (none)',
				'⚠️  Emails Quarantined',
				'🚨 Emails Rejected',
				'🚨 Critical Risk Records',
				'⚠️  High Risk Records',
				'📊 Medium Risk Records'
			],
			'Current Value': [
				f"{total_emails:,}",
				f"{total_records:,}",
				f"{auth_rate:.1f}%",
				f"{spf_rate:.1f}%",
				f"{dkim_rate:.1f}%", 
				f"{none_disp:,}",
				f"{quarantine_disp:,}",
				f"{reject_disp:,}",
				f"{critical_risk:,}",
				f"{high_risk:,}",
				f"{medium_risk:,}"
			],
			'Target': [
				'N/A',
				'N/A',
				'100%',
				'100%',
				'80%+',
				'Current level',
				'0',
				'0',
				'0',
				'0',
				'Current level'
			],
			'Status': [
				'Info',
				'Info',
				'Good' if auth_rate >= 95 else 'Warning' if auth_rate >= 80 else 'Critical',
				'Good' if spf_rate >= 95 else 'Warning' if spf_rate >= 80 else 'Critical',
				'Good' if dkim_rate >= 80 else 'Warning' if dkim_rate >= 50 else 'Critical',
				'Info',
				'Warning' if quarantine_disp > 0 else 'Good',
				'Critical' if reject_disp > 0 else 'Good',
				'Critical' if critical_risk > 0 else 'Good',
				'Warning' if high_risk > 0 else 'Good',
				'Info'
			]
		}
		
		summary_df = pd.DataFrame(summary_data)
		
		# Generate enhanced recommendations with specific guidance
		recommendations = []
		detailed_recommendations = []
		
		if auth_rate < 100:
			if spf_rate < 100:
				recommendations.append("PRIORITY 1: Fix SPF authentication to reach 100% pass rate")
				detailed_recommendations.append("To achieve 100% SPF pass rate: 1) Review your SPF record in DNS, 2) Ensure all legitimate email servers are listed, 3) Check for email forwarding issues, 4) Verify alignment settings")
			if dkim_rate < 50:
				recommendations.append("PRIORITY 2: Implement DKIM signing for better email authentication")
				detailed_recommendations.append("To improve DKIM: 1) Enable DKIM signing on your email servers, 2) Publish DKIM public keys in DNS, 3) Test DKIM signatures, 4) Monitor DKIM pass rates")
			recommendations.append("GOAL: Achieve 100% DMARC authentication through SPF and/or DKIM")
			detailed_recommendations.append("To reach 100% DMARC authentication: Focus on SPF first (easier to implement), then add DKIM as backup. Both don't need to pass - just one for DMARC to authenticate")
		
		if quarantine_disp > 0 or reject_disp > 0:
			recommendations.append("URGENT: Investigate quarantined/rejected emails for unauthorized sending")
			detailed_recommendations.append("Review quarantined/rejected emails to identify: 1) Legitimate services not in SPF record, 2) Potential spoofing attempts, 3) Email forwarding issues")
			
		if critical_risk > 0:
			recommendations.append("CRITICAL: Address high-risk security issues immediately")
			detailed_recommendations.append("High-risk issues may indicate: 1) Email spoofing attempts, 2) Compromised email accounts, 3) Unauthorized use of your domain")
		
		# Add general guidance
		if auth_rate >= 95:
			recommendations.append("MAINTAIN: Continue monitoring for new threats and changes")
			detailed_recommendations.append("Your DMARC is performing well. Continue regular monitoring and review any new failures promptly")
		
		recommendations_df = pd.DataFrame({
			'Priority Action': recommendations,
			'Detailed Steps': detailed_recommendations[:len(recommendations)]
		})
		
		# Create the complete Excel sheet manually instead of using multiple to_excel calls
		# This avoids issues with startrow when writing to the same sheet multiple times
		from openpyxl.styles import PatternFill
		green_fill = PatternFill(start_color="90EE90", end_color="90EE90", fill_type="solid")
		yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
		red_fill = PatternFill(start_color="FF6B6B", end_color="FF6B6B", fill_type="solid")
		
		wb = load_workbook(excel_path)
		
		# Remove existing Executive_Summary sheet if it exists and create new one
		if "Executive_Summary" in wb.sheetnames:
			wb.remove(wb["Executive_Summary"])
		ws = wb.create_sheet("Executive_Summary")
		
		# Add headers manually
		ws['A1'] = "DMARC REPORT EXECUTIVE SUMMARY"
		ws['A1'].font = Font(bold=True, size=16)
		ws.merge_cells('A1:D1')
		
		# Add overview for non-technical users
		overview_text = f"📧 Email Security Status: {total_emails:,} emails analyzed | Auth Rate: {auth_rate:.1f}% | Target: 100%"
		if auth_rate >= 95:
			overview_text += " ✅ EXCELLENT"
		elif auth_rate >= 80:
			overview_text += " ⚠️ NEEDS IMPROVEMENT"  
		else:
			overview_text += " 🚨 CRITICAL - IMMEDIATE ACTION REQUIRED"
			
		ws['A2'] = overview_text
		ws['A2'].font = Font(bold=True, size=12)
		ws.merge_cells('A2:D2')
		
		# Add recommendations section
		ws['A3'] = "🔧 RECOMMENDED ACTIONS (Priority Order)"
		ws['A3'].font = Font(bold=True, size=14, color="FF0000")
		
		# Write recommendations data manually
		current_row = 4
		# Headers
		ws[f'A{current_row}'] = "Priority Action"
		ws[f'B{current_row}'] = "Detailed Steps"
		ws[f'A{current_row}'].font = Font(bold=True)
		ws[f'B{current_row}'].font = Font(bold=True)
		current_row += 1
		
		# Data
		for _, row in recommendations_df.iterrows():
			ws[f'A{current_row}'] = row['Priority Action']
			ws[f'B{current_row}'] = row['Detailed Steps']
			current_row += 1
		
		# Add metrics section
		current_row += 2
		ws[f'A{current_row}'] = "📊 DETAILED METRICS & TARGETS"
		ws[f'A{current_row}'].font = Font(bold=True, size=14)
		ws.merge_cells(f'A{current_row}:D{current_row}')
		current_row += 2
		
		# Headers
		ws[f'A{current_row}'] = "Metric"
		ws[f'B{current_row}'] = "Current Value"
		ws[f'C{current_row}'] = "Target"
		ws[f'D{current_row}'] = "Status"
		for col in ['A', 'B', 'C', 'D']:
			ws[f'{col}{current_row}'].font = Font(bold=True)
		current_row += 1
		
		# Data with conditional formatting - write each column explicitly to ensure data integrity
		for index, row in summary_df.iterrows():
			# Explicitly write each cell to ensure proper data placement
			metric_cell = ws[f'A{current_row}']
			value_cell = ws[f'B{current_row}']
			target_cell = ws[f'C{current_row}']
			status_cell = ws[f'D{current_row}']
			
			metric_cell.value = row['Metric']
			value_cell.value = row['Current Value']
			target_cell.value = row['Target']
			status_cell.value = row['Status']
			
			# Apply conditional formatting
			if status_cell.value == 'Good':
				status_cell.fill = green_fill
			elif status_cell.value == 'Warning':
				status_cell.fill = yellow_fill
			elif status_cell.value == 'Critical':
				status_cell.fill = red_fill
			current_row += 1
		
		# Add top sources section with better explanations
		current_row += 2
		ws[f'A{current_row}'] = "🌐 TOP EMAIL SOURCES BY VOLUME"
		ws[f'A{current_row}'].font = Font(bold=True, size=14)
		ws.merge_cells(f'A{current_row}:D{current_row}')
		current_row += 1
		
		# Add explanation
		ws[f'A{current_row}'] = "These are the IP addresses that sent the most emails using your domain:"
		ws[f'A{current_row}'].font = Font(italic=True)
		ws.merge_cells(f'A{current_row}:D{current_row}')
		current_row += 2
		
		# Headers with better labels
		ws[f'A{current_row}'] = "Source IP Address"
		ws[f'B{current_row}'] = "Email Volume"
		ws[f'C{current_row}'] = "Description" 
		ws[f'D{current_row}'] = "Action Required"
		for col in ['A', 'B', 'C', 'D']:
			ws[f'{col}{current_row}'].font = Font(bold=True)
		current_row += 1
		
		# Data with additional context
		for index, row in top_sources.iterrows():
			ip_address = row['source_ip']
			email_count = row['count']
			
			ws[f'A{current_row}'].value = ip_address
			ws[f'B{current_row}'].value = f"{email_count:,} emails"
			
			# Add helpful description and action guidance
			if index == 0:  # Top sender
				ws[f'C{current_row}'].value = "Primary email source"
				ws[f'D{current_row}'].value = "Verify this is your legitimate email server"
			elif email_count >= 100:
				ws[f'C{current_row}'].value = "High volume sender"
				ws[f'D{current_row}'].value = "Ensure this IP is authorized in your SPF record"
			elif email_count >= 10:
				ws[f'C{current_row}'].value = "Medium volume sender" 
				ws[f'D{current_row}'].value = "Review if this source should be in SPF record"
			else:
				ws[f'C{current_row}'].value = "Low volume sender"
				ws[f'D{current_row}'].value = "Monitor for suspicious activity"
			
			current_row += 1
		
		# Adjust column widths for better readability
		ws.column_dimensions['A'].width = 35  # Metric names and IP addresses
		ws.column_dimensions['B'].width = 20  # Current values and email counts
		ws.column_dimensions['C'].width = 25  # Targets and descriptions
		ws.column_dimensions['D'].width = 30  # Status and actions
		
		wb.save(excel_path)
		print(f"Executive Summary created in '{excel_path}'.")
		
	except Exception as e:
		print(f"Error creating executive summary: {e}")

# -----------------------------------------------------------------------------
# Send the DMARC Excel report as an email attachment
def emailReport():
	global CURRENT_DATE
	# Load the config, .env file contains all data to send email
	config = load_config()
	smtp_server = os.getenv("SMTP_SERVER")
	smtp_port = int(os.getenv("SMTP_PORT", 587)) # Default to 587 if SMTP_PORT is not present.
	from_email = config["FROM_EMAIL"]
	to_email = config["TO_EMAIL"]
	email = config["EMAIL"]
	password = config["PASSWORD"]

	# Prepare date strings for the report period 
	prev_date = (datetime.datetime.now() - datetime.timedelta(days=7)).strftime("%Y-%m-%d")
	report_dir = os.path.join(os.getcwd(), "Dmarc_Reports")
	excel_filename = f"dmarc_report_{CURRENT_DATE}.xlsx"
	excel_path = os.path.join(report_dir, excel_filename)

	# Check if the report file exists
	if not os.path.exists(excel_path):
		print(f"Report file not found: {excel_path}")
		return

	# Construct email message
	msg = EmailMessage()
	msg["Subject"] = f"DMARC Report - {prev_date} - {CURRENT_DATE}."
	msg["To"] = to_email
	msg["From"] = from_email
	msg.set_content(f"Attached is the DMARC report for:\n{prev_date} - {CURRENT_DATE}")

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
def generatePlotlyChart(df, html_path="dmarc_plotly_report.html", image_path="dmarc_plot.png"):
	
	df.columns = df.columns.str.strip()

	# Metrics
	metrics = {
		'DMARC': 'DMARC Compliance',
		'SPF': 'SPF',
		'DKIM': 'DKIM'
	}

	chart_files = []

	for label, col in metrics.items():
		# Group by pass/fail, sum email volume
		summary = df.groupby(col)['Email volume'].sum().reset_index()
		# Calculate passrate
		total = summary['Email volume'].sum()
		pass_row = summary[summary[col]].str.lower() == 'pass'


	#fig.show()
	fig.write_html(html_path)
	fig.write_image(image_path)
	print(f"Plotly chart saved as {html_path} and image as {image_path}")
	return html_path, image_path

# -----------------------------------------------------------------------------
def chartData(excel_path):

	wb = load_workbook(excel_path)
	ws = wb['Tabular Data']

	if "Charts" not in wb.sheetnames:
		chart_ws = wb.create_sheet("Charts")
	else:
		chart_ws = wb["Charts"]

	# ---------------------------------
	# Pie Chart for overall DMARC Pass/Fail
	total_pass = sum((v[0] for v in ws.iter_rows(min_row=2, min_col=3, max_col=3, values_only=True) if isinstance(v[0], (int, float))), 0)
	total_fail = sum((v[0] for v in ws.iter_rows(min_row=2, min_col=4, max_col=4, values_only=True) if isinstance(v[0], (int, float))), 0)
	
	# Summary data for pie chart
	chart_ws["A25"] = "Pass"
	chart_ws["A26"] = "Fail"
	chart_ws["B25"] = total_pass
	chart_ws["B26"] = total_fail

	total = total_pass + total_fail
	pass_pct = total_pass / total if total else 0
	fail_pct = total_fail / total if total else 0

	chart_ws["C24"] = "Percent"
	chart_ws["C25"] = pass_pct
	chart_ws["C26"] = fail_pct

	chart_ws["C25"].number_format = '0.00%'
	chart_ws["C26"].number_format = '0.00%'

	pie = PieChart()
	pie.title = "Overall DMARC Pass/Fail"
	labels = Reference(chart_ws, min_col=1, min_row=25, max_row=26)
	data = Reference(chart_ws, min_col=2, min_row=25, max_row=26)
	pie.add_data(data, titles_from_data=False)
	pie.set_categories(labels)
	chart_ws.add_chart(pie, "A2")
	# ---------------------------------

	wb.save(excel_path)

# -----------------------------------------------------------------------------
# Present data in tabular format similar to Google's DMARC Example 
def tabularData(excel_path):
	
	df = pd.read_excel(excel_path, sheet_name = "Organized_Data")
		
	# Group by source_ip
	grouped = df.groupby(['header_from', 'source_ip'])
	#grouped = df.groupby('source_ip')

	summary_data = []
	for (header_from, ip), group in grouped:
	#for ip, group in grouped:
		volume = group['count'].sum()

		# DMARC pass/fail (previously calculated 'auth_status')
		dmarc_pass = ((group['auth_status'] == 'Authenticated') * group['count']).sum()
		dmarc_fail = ((group['auth_status'] == 'Failed') * group['count']).sum()
		dmarc_rate = f"{(dmarc_pass / volume * 100):.2f}%" if volume else "0.00%"

		# SPF Pass is only for header_from alignment
		spf_for_header_from_pass = ((group['spf_for_header_from'] == 'pass' ) * group['count']).sum()
		spf_for_header_from_fail = ((group['spf_for_header_from'] != 'pass' ) * group['count']).sum()
		spf_for_header_from_rate = f"{(spf_for_header_from_pass / volume * 100):.2f}%" if volume else "0.00%"

		# SPF Counts
		#spf_pass = ((group['spf_result'] == 'pass') * group['count']).sum()
		#spf_fail = ((group['spf_result'] == 'fail') * group['count']).sum()
		#spf_rate = f"{(spf_pass / volume * 100):.2f}%" if volume else "0.00%"

		# DKIM Counts
		dkim_pass = ((group['dkim_result'] == 'pass') * group['count']).sum()
		dkim_fail = ((group['dkim_result'] == 'fail') * group['count']).sum()
		dkim_rate = f"{(dkim_pass / volume * 100):.2f}%" if volume else "0.00%"

		summary_data.append([
			header_from, ip, volume,
			dmarc_pass, dmarc_fail, dmarc_rate,
			spf_for_header_from_pass, spf_for_header_from_fail, spf_for_header_from_rate, 
			dkim_pass, dkim_fail, dkim_rate
		])

		#summary_data.append([
		#	ip, volume,
		#	dmarc_pass, dmarc_fail, dmarc_rate,
		#	spf_pass, spf_fail, spf_rate, 
		#	dkim_pass, dkim_fail, dkim_rate
		#])

	columns = [
		'Header From', 'SourceIP Address', 'Email volume',
		'DMARC Pass', 'DMARC Fail', 'DMARC Rate',
		'SPF (aligned) Pass', 'SPF (aligned) Fail', 'SPF (aligned) Rate', 
		'DKIM Pass', 'DKIM Fail', 'DKIM Rate'
	]

#	columns = [
#		'IP Address', 'Email volume',
#		'DMARC Pass', 'DMARC Fail', 'DMARC Rate',
#		'SPF Pass', 'SPF Fail', 'SPF Rate', 
#		'DKIM Pass', 'DKIM Fail', 'DKIM Rate'
#	]

	summary_df = pd.DataFrame(summary_data, columns=columns)
	summary_df = summary_df.sort_values(by='SPF (aligned) Fail', ascending=False) # Sort by highest SPF fail rate

	# Save to Excel (append as new sheet)
	with pd.ExcelWriter(excel_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
		summary_df.to_excel(writer, sheet_name="Tabular Data", index=False)

	# Style the tabular data
	wb = load_workbook(excel_path)
	ws = wb["Tabular Data"]

	# Header formatting
	yellow = PatternFill("solid", fgColor="FFF475")
	bold = Font(bold=True)
	center = Alignment(horizontal="center", vertical="center")
	border = Border(bottom=Side(style="thin"), top=Side(style="thin"),
					  left=Side(style="thin"), right=Side(style="thin"))
	
	# First header row (manually set)
	#ws.merge_cells('A1:A2')
	#ws.merge_cells('B1:B2')
	#ws['A1'] = "IP Address"
	#ws['B1'] = "Email volume"

	#ws.merge_cells('C1:E1')
	#ws['C1'] = "DMARC Compliance"

	#ws.merge_cells('F1:H1')
	#ws['F1'] = "SPF"
	#
	#ws.merge_cells('I1:K1')
	#ws['I1'] = "DKIM"

	## Second header row (sub columns)
	#ws['C2'] = "Pass"
	#ws['D2'] = "Fail"
	#ws['E2'] = "Rate"
	#
	#ws['F2'] = "Pass"
	#ws['G2'] = "Fail"
	#ws['H2'] = "Rate"
	#
	#ws['I2'] = "Pass"
	#ws['J2'] = "Fail"
	#ws['K2'] = "Rate"

	# Style headers
	#for row in ws.iter_rows(min_row=1, max_row=2, min_col=1, max_col=ws.max_column):
	#	for cell in row:
	#		cell.fill = yellow
	#		cell.font = bold
	#		cell.alignment = Alignment(horizontal='center')
	#		cell.border = border

	# Header styling
	for cell in ws[1]:
		cell.font = bold
		cell.fill = yellow
		cell.alignment = center
		cell.border = border

	# Set column widths automatically
	for i, col in enumerate(ws.columns, 1):
		max_length = 0
		col_letter = get_column_letter(i)
		#col_letter = col[0].column_letter
		for cell in col:
			try:
				if cell.value:
					max_length = max(max_length, len(str(cell.value)))
			except:
				pass
		ws.column_dimensions[col_letter].width = max_length + 4 

	# Freeze header and save
	ws.freeze_panes = "A2"
	wb.save(excel_path)
	print(f"Tabular DMARC summary added to '{excel_path}'.")

# -----------------------------------------------------------------------------
# Get the SPF Record for a given domain (first v=spf1 record found)
def get_spf_record(domain):
	try: 
		answers = dns.resolver.resolve(domain, 'TXT')
		for rdata in answers:
			# Concatenate all strings (TXT records sometimes split)
			txt = "".join([s.decode("utf-8") if isinstance(s, bytes) else s for s in rdata.strings])
			if txt.startswith('v=spf1'):
				return txt
	except Exception:
		pass
	return None

# -----------------------------------------------------------------------------
# Extract all include: return a list of domains
def extract_includes(spf_record):
	if not spf_record:
		return []
	# Find all "include:domain" entries in spf_record
	return re.findall(r'include:([^\s]+)', spf_record)

# -----------------------------------------------------------------------------
# Recursively check if IP is allowed by the SPF record of the domain.
# Limits recursion and DNS lookups to 10 (per SPF RFC)
def ip_in_spf(ip, domain, lookup_count=0, max_lookups=10, checked_domains=None):
	if checked_domains is None:
		checked_domains = set()
	# Stop if DNS lookup limit eached or if domain was already checked.
	if lookup_count >= max_lookups or domain in checked_domains:
		return False

	checked_domains.add(domain)
	spf_record = get_spf_record(domain)
	if not spf_record:
		return False

	try:
		ip_obj = ipaddress.ip_address(ip)
	except Exception:
		return False

	for part in spf_record.split():
		# Check direct IPv4 entry
		if part.startswith("ip4:") and isinstance(ip_obj, ipaddress.IPv4Address):
			try:
				if '/' in part[4:]:
					net = ipaddress.IPv4Network(part[4:], strict=False)
					if ip_obj in net:
						return True
					else:
						if ip_obj == ipaddress.IPv4Address(part[4:]):
							return True
			except Exception:
				continue
		# Check direct IPv6 entry
		elif part.startswith("ip6:") and isinstance(ip_obj, ipaddress.IPv6Address):
			try:
				if '/' in part[4:]:
					net = ipaddress.IPv6Network(part[4:], strict=False)
					if ip_obj in net:
						return True
					else:
						if ip_obj == ipaddress.IPv6Address(part[4:]):
							return True
			except Exception:
				continue
		# Recurse into include: domains
		elif part.startswith("include:"):
			included_domain = part[len("include:") :]
			# Recursively check included domain, imcrementing lookup count
			if ip_in_spf(ip, included_domain, lookup_count+1, max_lookups, checked_domains):
				return True

	return False

# -----------------------------------------------------------------------------
# Investigate an SPF failure row: get SPF_record, includes, check IP Auth and summarize findings
def investigate_spf_failure(row):
	investigation = {
		"ip_in_spf": None,
		"spf_record": None,
		"spf_includes": None,
		"spf_notes": "",
		"why_spf_failed": "",
		"recommended_action": "",
		"business_impact": ""
	}

	# Use header_from for SPF Checks
	domain = row.get("header_from", "")
	if not domain or pd.isna(domain):
		investigation["spf_notes"] = "No header_from domain found."
		investigation["why_spf_failed"] = "The email did not include a valid sender domain (header_from), so SPF could not be checked."
		investigation["recommended_action"] = "Contact your email provider to ensure proper domain configuration."
		investigation["business_impact"] = "Low - likely a configuration issue."
		return investigation

	spf_record = get_spf_record(domain)
	investigation["spf_record"] = spf_record if spf_record else "No SPF record found"
	includes = extract_includes(spf_record) if spf_record else []
	investigation["spf_includes"] = ", ".join(includes) if includes else "None"

	ip = row.get("source_ip", "")
	in_spf = ip_in_spf(ip, domain) if spf_record else False
	investigation["ip_in_spf"] = in_spf

	# Enhanced explanations for non-technical users
	if spf_record is None:
		investigation["spf_notes"] = f"No SPF record present for domain {domain}."
		investigation["why_spf_failed"] = (
			f"Your domain {domain} doesn't have an SPF record in DNS. This is like not having a list of "
			"authorized mail carriers for your business. Receiving email servers can't verify if emails "
			"claiming to be from your domain are legitimate."
		)
		investigation["recommended_action"] = (
			"Add an SPF record to your DNS settings. Consult your IT team or DNS provider to create "
			"an SPF record that lists all legitimate email servers for your domain."
		)
		investigation["business_impact"] = (
			"High - Emails may be marked as spam or rejected, affecting business communications."
		)
	elif in_spf:
		investigation["spf_notes"] = f"Source IP is directly allowed in SPF record for {domain}."
		investigation["why_spf_failed"] = (
			f"The sending IP ({ip}) is authorized in your SPF record, but SPF still failed. "
			"This could be due to email forwarding, strict alignment settings, or other technical issues."
		)
		investigation["recommended_action"] = (
			"Check if this email was forwarded through another service. Consider implementing DKIM "
			"as a backup authentication method, or review DMARC alignment settings."
		)
		investigation["business_impact"] = (
			"Medium - Authentication is working but there may be forwarding or alignment issues."
		)
	elif includes:
		investigation["spf_notes"] = (
			f"Source IP may be authorized via include(s) for {domain}: {investigation['spf_includes']}. "
			"Manual review required to verify authorization."
		)
		investigation["why_spf_failed"] = (
			f"Your SPF record includes other domains ({investigation['spf_includes']}) that should authorize "
			f"the sending IP ({ip}). However, the SPF check still failed. This could mean the included "
			"domains don't actually authorize this IP, or there's a lookup limit issue."
		)
		investigation["recommended_action"] = (
			"Review the included domains in your SPF record. Verify that each included service "
			"actually authorizes the failing IP address. Consider flattening your SPF record if "
			"there are too many includes (limit is 10 DNS lookups)."
		)
		investigation["business_impact"] = (
			"Medium - May indicate issues with third-party email services or SPF record complexity."
		)
	elif row.get("spf_checked_domain","") and row.get("spf_checked_domain", "") != domain and row.get("spf_checked_result", "") == "pass":
		investigation["spf_notes"] += f" SPF passed for non-aligned domain ({row.get('spf_checked_domain','')}), indicating email forwarding."
		investigation["why_spf_failed"] = (
			f"This email was forwarded by another service. SPF passed for the forwarding service's domain "
			f"({row.get('spf_checked_domain','')}) but failed for your domain ({domain}). This is normal "
			"behavior for email forwarding services like mailing lists or forwarders."
		)
		investigation["recommended_action"] = (
			"If this forwarding is expected (e.g., mailing lists, email forwarders), consider implementing "
			"DKIM signing and setting DMARC policy to relaxed alignment. Document legitimate forwarding services."
		)
		investigation["business_impact"] = (
			"Low - This is normal for forwarded emails, but may affect deliverability if not properly configured."
		)
	else:
		investigation["spf_notes"] = (
			f"Source IP ({ip}) is not authorized in SPF record for {domain}. "
			"This could indicate unauthorized email use or missing SPF configuration."
		)
		investigation["why_spf_failed"] = (
			f"The email server at IP {ip} is not listed as an authorized sender for {domain}. "
			"This is like someone using your company letterhead without permission. It could be "
			"a legitimate service you forgot to authorize, or potentially fraudulent email."
		)
		investigation["recommended_action"] = (
			f"Investigate if {ip} belongs to a legitimate email service you use. If legitimate, "
			"add it to your SPF record. If not legitimate, this could be email spoofing - "
			"consider tightening your DMARC policy and monitoring for abuse."
		)
		investigation["business_impact"] = (
			"High - Could indicate email spoofing/fraud, or legitimate emails being blocked."
		)

	return investigation

# -----------------------------------------------------------------------------
# Main SPF Failure investigation routine. Create new sheet with all SPF failures and investigation results.
def spfFailures(excel_path):

	# Read Organized_Data for parsed DMARC records and filter for SPF failures only
	df = pd.read_excel(excel_path, sheet_name="Organized_Data")
	spf_failures = df[df['spf_result'].str.lower() == 'fail'].copy()
	if spf_failures.empty:
		print("No SPF failures found.")
		return
	
	print(f"Investigating {len(spf_failures)} SPF failures...")

	# Apply investigation to each row, adding new columns for SPF info
	investigation_results = spf_failures.progress_apply(investigate_spf_failure, axis=1, result_type='expand')
	# Combine original SPF failure data with investigation columns
	spf_failures_subset = spf_failures[["envelope_to", "source_ip", "spf_result", "count"]].reset_index(drop=True)
	result_df = pd.concat([spf_failures_subset, investigation_results.reset_index(drop=True)], axis=1)

	# Sort by ip_in_spf True
	result_df['ip_in_spf'] = result_df['ip_in_spf'].astype(bool)
	result_df = result_df.sort_values('ip_in_spf', ascending=False).reset_index(drop=True)

	# Save the investigation results to a new worksheet 'SPF_Failures'.
	with pd.ExcelWriter(excel_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
		result_df.to_excel(writer, sheet_name="SPF_Failures", index=False)

	# Adjust column widths for spf_record, spf_includes, spf_notes
	wb = load_workbook(excel_path)
	ws = wb["SPF_Failures"]
	col_names = {cell.value: idx+1 for idx, cell in enumerate(ws[1])}
	for col in ["spf_record", "spf_includes", "spf_notes"]:
		if col in col_names:
			ws.column_dimensions[get_column_letter(col_names[col])].width = 40
	wb.save(excel_path)

	print(f"SPF investigation sheet created as 'SPF_Failures' in {excel_path}.")

# -----------------------------------------------------------------------------
# -----------------------------------------------------------------------------
# Main execution function for the DMARC processing workflow
def main():

	global CURRENT_DATE

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
		unzipped_dir = os.path.join(save_dir, "unzipped")
		report_dir = os.path.join(os.getcwd(), "Dmarc_Reports")
		excel_path = parse_dmarc_directory(unzipped_dir, report_dir, CURRENT_DATE)
		
		organizeData(excel_path)
		create_executive_summary(excel_path)
		tabularData(excel_path)
		spfFailures(excel_path)
		formatSheets(excel_path)
		
		df_tabular = pd.read_excel(excel_path, sheet_name="Tabular Data")

		# Send email based on .env values
		emailReport()

		# Logout from IMAP server
		mail.logout()
		print("IMAP logout successful.")

	except Exception as e:
		print(f"IMAP login failed: {e}")

# -----------------------------------------------------------------------------

if __name__ == "__main__":
	main()

# -----------------------------------------------------------------------------
