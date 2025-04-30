# O365-GUICompliance Search  
<img src="https://github.com/AdminVin/O365-GUIComplianceSearch/blob/main/GUIComplianceSearch-Screenshot.png?raw=true" />

## What is "O365-GUICompliance Search"?  
O365-GUICompliance Search was designed to streamline running content searches, with a primary focus on purging spam messages.

> **Note:** Credentials are never stored. Authentication is required for every search. (MSP-friendly)

## Why create this and compile it?  

**Q: Why build a mini program for this?**  
1. Office 365’s website is slow to load and navigate.  
2. I wanted a lightweight, fast, taskbar-pinnable program to create content searches quickly.

**Q: How was it compiled?**  
- It was compiled with **PS2EXE** using the following command:  
  `Invoke-PS2EXE -inputFile '.\O365-ComplianceSearch.ps1' -outputFile '.\O365-ComplianceSearch.exe' -iconFile '.\search_icon.ico' -noConsole -noOutput`  
- Source: [PS2EXE GitHub](https://github.com/MScholtes/PS2EXE)

## Usage  
1. Run either `O365-ComplianceSearch.ps1` or `O365-ComplianceSearch.exe`.  
2. Fill out fields → click **Search** → authenticate into your tenant.  
   - **1st Authentication Prompt:** `Connect-ExchangeOnline`  
   - **2nd Authentication Prompt:** `Connect-IPPSSession`  
3. Optionally monitor progress; the rest of the process is fully automated.

## Donate  
Saved you time? Consider sponsoring my next coffee: [PayPal](https://www.paypal.com/donate/?hosted_button_id=EZU78ZANFT24C)
