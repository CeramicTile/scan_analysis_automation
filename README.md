# Automation Tool for Analyzing and Scoring Reportable ACAS IAVM and STIG Results
## About:
This tool is used to calculate the IAVM and STIG scores using raw data from ACAS scans and from them, generate the overall CCRI Score for the system.  It can be also be used by security engineers to target low hanging fruit for patching and STIG compliance audits.

## User Guide:
1. Copy the latest version of the script into the assessment’s root directory; ensure the Excel macros are also placed in the correct location

2. Extract all STIG xccdf zips so that the containing folder is in the assessment’s root directory

3. After applying the necessary filters, download the IAVM file from ACAS.  Do NOT rename the file, ensure it is still named “vulns.csv”

4. Right Click the “All-In-One” PowerShell script and choose “Run with PowerShell”.

5. Choose the appropriate option when prompted, based on assessment requirements.

## Excel Macro Addins

Download and place both IAVM_Score.xlam and STIG_Score.xlam are stored in the following location:

C:\Users\%username%\AppData\Roaming\Microsoft\AddIns
