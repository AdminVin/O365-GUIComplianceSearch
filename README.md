# O365-GUICompliance Search
<img src="https://github.com/AdminVin/O365-GUIComplianceSearch/blob/main/GUIComplianceSearch-Screenshot.png?raw=true"">

## What is "O365-GUICompliance Search"?
O365-GUICompliance Search was designed to easily run content searches with the primary purpose of purging spam messages. 

**Credentials are never stored, and authentication is required every search. (MSP FRIENDLY)**

## Why create this and have a compiled version?
**Question:** Why have a mini program for this?
1. Office 365's website is slow to navigate/load. 
2. I wanted a lightweight, fast, taskbar-pinnable program to create content searches quickly.

**Question:** How was it compiled?  
- It was compiled with **PS2EXE** using the following command:  
- Invoke-PS2EXE -inputFile '.\O365-ComplianceSearch.ps1' -outputFile '.\O365-ComplianceSearch.exe' -iconFile '.\search_icon.ico' -noConsole -noOutput
- Source: **PS2EXE** https://github.com/MScholtes/PS2EXE

## Usage
1. Run either `O365-ComplianceSearch.ps1` or `O365-ComplianceSearch.exe`.
2. Fill out fields > Search > Authenticate into Tenant
    - 1st Authentication Prompt: Connect-ExchangeOnline
    - 2nd Authentication Prompt: Connect-IPPSSession
3. Monitor if you like, but the rest of the process is fully automated.

## Donate
Saved you time? Great! --- Sponsor my next coffee? [PayPal](https://www.paypal.com/donate/?hosted_button_id=EZU78ZANFT24C)