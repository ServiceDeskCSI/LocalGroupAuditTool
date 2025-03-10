![image](https://github.com/user-attachments/assets/3fb9f6b8-3527-4914-91ec-fba5bb72e73a)

# Local Groups Audit Tool

Local Groups Audit Tool is a PowerShell-based GUI application designed to audit local group memberships across multiple remote computers. It retrieves computer names from a configuration file or Active Directory, scans the specified local group (defaulting to "Administrators") using ADSI, and logs the results into CSV files for both successful and failed scans.

## Features

- **GUI Interface:**  
  User-friendly Windows Forms application for managing and auditing computers.

- **Active Directory Integration:**  
  Retrieve computer names directly from Active Directory with the "Get AD Computers" button.

- **Local Group Audit:**  
  Scans a specified local group on remote computers and extracts member details.

- **Rescan Failed Machines:**  
  Automatically rescan machines that fail during the initial audit until all are processed or the user stops the scan.

- **CSV Logging:**  
  Outputs detailed results to separate CSV files:
  - `LocalGroupMembers_Success.csv`
  - `LocalGroupMembers_Failure.csv`

- **Real-Time Feedback:**  
  Displays progress and log messages via a console text box and progress bar.

- **Machine List Editing:**  
  Edit the list of target machines using "Up", "Down", "Delete", and "Save List" controls.

## Prerequisites

- **PowerShell 3.0 or later**  
- **Windows OS** (the script uses Windows Forms)
- **Active Directory Module** (if using the AD retrieval feature)
- **Appropriate Permissions:**  
  Ensure you have the necessary rights to access remote computers and query local groups.

## Installation and Setup

1. **Download the Script:**  
   Clone or download the script file into a dedicated folder.

2. **Configuration File:**  
   Place the script in a folder along with (or without) the `computers.config` file. If it does not exist, the script will create it during runtime.

3. **Dependencies:**  
   The script dynamically loads the required assemblies (`System.Windows.Forms` and `System.Drawing`), so no additional installation is necessary.

## How to Use

1. **Launch the Script:**  
   Open PowerShell, navigate to the script folder, and run:
   ```powershell
   .\LocalGroupAuditToolGUI.ps1
