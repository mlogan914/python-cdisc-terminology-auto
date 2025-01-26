# Automating CDISC Terminology Downloads and Maintenance

**Author**: Melanie Logan  
**Version**: 1.0  

## Overview

This project is designed for professionals in the biotech and clinical research fields, specifically those involved in implementing CDISC (Clinical Data Interchange Standards Consortium) data standards and managing CDISC controlled terminology. If you are responsible for ensuring your datasets adhere to CDISC standards, this tool will help you automate the download, maintenance, and archiving of CDISC terminology files.

CDISC controlled terminology refers to the standardized list of terms used across clinical trial data, ensuring consistency and clarity in the reporting of clinical data. These terms include codes for medical conditions, interventions, and other clinical trial-related information. The terminology is updated regularly by CDISC to reflect new medical insights, regulatory requirements, and industry best practices.

This solution is designed to:  
- Download CDISC SDTM and ADaM terminology files.  
- Archive old files and update them with the latest versions.  
- Automate execution using Windows Task Scheduler for periodic maintenance.  

By following this project, you can eliminate repetitive tasks and focus on more impactful work.

---

## Features
- **Automated Downloads**: Uses Python's Selenium package to access and download files directly from the [NCI EVS FTP Server](https://evs.nci.nih.gov/ftp1/CDISC/).
- **File Archiving**: Automatically archives old terminology files to maintain version control.
- **Conversion to Binary Format**: Converts `.xls` files to `.xlsb` using PowerShell for better efficiency and storage.
- **Scheduled Execution**: Supports automation via Windows Task Scheduler for regular updates.
- **Logging and Notifications**: Creates detailed logs of the process and sends email notifications.

---

## Requirements
### Software
- Python 3.x
- PowerShell (Windows)
- Google Chrome
- ChromeDriver compatible with your Chrome version

### Python Libraries
- `selenium`
- `logging`
- `os`
- `shutil`
- `itertools`
- `time`
- `pyxlsb`
- `subprocess`
- `smtplib`

Install missing libraries using `pip install <library-name>`.

---

## Setup Instructions

### 1. Clone the Repository
```bash
git clone https://github.com/mlogan914/cdisc-terminology-automation.git
```
### 2. Install Dependencies

Ensure all required Python libraries are installed. Run:
```bash
pip install <library-names>
```
### 3. Configure File Paths
Update the file paths in the script (cdisc_terminology.py) to match your environment:

- Downloads directory: Where Chrome downloads files.
- Terminology directory: Location to store updated terminology files.
- Archive directory: Location for archived files.
- Log directory: Directory for process logs.

### 4. ChromeDriver Setup

Ensure that the ChromeDriver executable matches your Chrome version. Update the following line in the script:
```bash
browser = webdriver.Chrome(r'path_to_chromedriver')
```

### 5. Automate with Windows Task Scheduler

- Create a new task in Windows Task Scheduler.
- Set the task to run python `cdisc_terminology.py` at your preferred frequency.

### Execution

Run the script manually or through Windows Task Scheduler:
```bash
python cdisc_terminology.py
```
### Logging

Logs are generated in the specified log directory, including:
- Process details
- Errors encountered during execution
- Timestamps of file downloads

### Notifications

- The script includes email functionality to notify stakeholders of successful or failed operations. 
- Update the sender_email and receiver_email fields in the script with your email credentials.

### File Outputs

- Updated Terminology Files: Saved to the terminology directory.
- Archived Files: Stored in the archive directory.
- Change Files: Copied to the changes directory.
- Log Files: Created in the log directory with timestamps.

### Notes

- If files require additional time to download, you can adjust the wait time by modifying time.sleep(n) in the script (n = seconds).
- Ensure Chrome and ChromeDriver versions are up to date for seamless execution.
