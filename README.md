# Remote Software Inventory Tool

## Overview
This project is a Python-based automation tool developed during my internship to remotely collect installed software information from multiple Windows machines in a network and generate a structured Excel report.

It is designed for IT asset tracking, software auditing, and internal system inventory management.



## Features
- Scans installed software from remote PCs/laptops/servers
- Multi-threaded execution for faster processing
- Ping-based online/offline check before execution
- Extracts software details (Name, Version, Publisher, Install Date)
- Generates structured Excel report using `openpyxl`
- Tracks failed/offline machines with reasons
- Works via PowerShell remote execution


## Requirements

- Windows OS
- Python 3.10+
- Administrator privileges
- PowerShell enabled with remote execution access
- Network access to all target machines

## Python Dependencies

Install required packages:
#### - ```bash
pip install -r requirements.txt

## Setup Instructions

### 1. Add target machines
Create a file named:

pc_namelist.txt

Add one machine per line:


PC01
PC02
LAPTOP-DEV
SERVER01
...

### 2. Configure Output Location
The Excel report will automatically be saved at:


/output/inventory_report.xlsx

### 3. Run the Script

Open Command Prompt as Administrator and run:

#### ```bash
python inventory_system.py

## Output Details

The generated Excel file contains:

### Sheet 1: Inventory
| Computer | Application | Version | Publisher | Install Date |

### Sheet 2: Report
- Total machines scanned
- Successful scans
- Failed/offline machines
- List of failed machines with reasons



## Important Notes

- Administrator privileges are required to run the script
- All target machines must be reachable over the network
- PowerShell remoting must be enabled
- Machines must be powered on during execution
- Firewall settings may affect connectivity



## Use Cases

- IT asset management
- Software compliance auditing
- Corporate device inventory tracking
- System administration automation



## Disclaimer

This project was developed as part of an internship experience.  
Any sensitive or company-specific configurations have been removed for public sharing.



## Author

Developed by: MUHAMMED NASEER 
Internship Project – Internal IT Automation Tool
