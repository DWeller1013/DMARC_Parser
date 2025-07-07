# DMARC_Parser

## Overview

**DMARC_Parser** is a Python-based tool designed to automate the collection, parsing, analysis, and reporting of DMARC (Domain-based Message Authentication, Reporting, and Conformance) aggregate reports. The parser downloads email attachments containing DMARC XML files, extracts and processes their contents, and generates comprehensive Excel reports for review. It also investigates SPF failures and can automatically email the results to a specified user.

## Features

- **Automated Email Processing**: Connects securely to an IMAP server, searches the DMARC folder for relevant emails from the last 7 days, downloads all attachments, and extracts `.zip` and `.gz` archives containing DMARC XML reports.
- **DMARC XML Parsing**: Parses XML files to extract critical DMARC data (such as envelope_to, source_ip, disposition, DKIM and SPF results, header_from, and more).
- **Excel Report Generation**: Converts all parsed records into structured Excel files, including detailed data sheets, summary tables, and visual charts (e.g., DMARC pass/fail rates).
- **SPF Failure Investigation**: Analyzes SPF failures, retrieves and checks SPF records, and summarizes findings in a dedicated Excel sheet.
- **Tabular Summary**: Groups results by sending domain and source IP, calculating DMARC, SPF, and DKIM pass/fail rates for fast assessment.
- **Automated Email Reporting**: Optionally sends the final Excel report as an email attachment to a configured recipient.
- **Configurable and Extensible**: Loads configuration (email credentials, server info) from environment variables for secure and flexible deployment.

## How It Works

1. **Setup**: Configure your email and IMAP server details in a `.env` file.
2. **Run the Script**: Execute the main script (`dmarcReport.py`). The main routine will:
    - Log in to your email account and collect all DMARC report emails.
    - Download attachments, extract and parse all included XML files.
    - Organize the data and generate Excel files with multiple sheets:
        - All Data
        - Organized Data
        - Tabular Data (with DMARC, SPF, DKIM summary)
        - SPF Failures
        - Charts
    - Optionally, email the resulting report to a specified address.

## Usage

1. Clone this repository:
    ```bash
    git clone https://github.com/DWeller1013/DMARC_Parser.git
    cd DMARC_Parser
    ```

2. Install dependencies:
    ```bash
    pip install -r requirements.txt
    ```

3. Create and fill out a `.env` file in the project root with:
    ```
    EMAIL=your_email@example.com
    PASSWORD=your_password
    IMAP_SERVER=imap.example.com
    IMAP_SSL_PORT=993
    TO_EMAIL=recipient@example.com
    FROM_EMAIL=your_email@example.com
    ```

4. Run the parser:
    ```bash
    python dmarcReport.py
    ```

5. Check the `Dmarc_Reports` directory for Excel output, or your email for the report.

## Notable Implementation Details

- Uses `imaplib` for IMAP email access and `openpyxl`/`pandas` for Excel file creation and manipulation.
- Performs RDAP (WHOIS) lookups on source IPs for organization name enrichment.
- SPF record resolution and DNS handling are included for advanced SPF failure investigation.
- Progress bars and print statements provide feedback during processing.

## Requirements

- Python 3.x
- See `requirements.txt` for a full list of dependencies.

## Author

[DWeller1013](https://github.com/DWeller1013)

---
*This project automates the end-to-end process of DMARC report collection and analysis, making DMARC monitoring and compliance much easier for email administrators and security teams.*
