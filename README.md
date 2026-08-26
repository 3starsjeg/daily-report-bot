# Daily Report Bot

A Python desktop automation bot built with PyAutoGUI that fetches GOOG stock information from Google Finance and creates a daily Excel report automatically.

## What This Project Does

The bot automates the following workflow:

1. Opens Google Chrome.
2. Opens the Google Finance page for GOOG (Alphabet Inc. Class C).
3. Copies information from the webpage.
4. Extracts the current GOOG stock price from the copied webpage information.
5. Generates the current date and time automatically at runtime.
6. Opens Microsoft Excel.
7. Creates a new Excel workbook.
8. Enters the date/time, GOOG price, and a comment into Excel.
9. Saves the Excel report with the current date in the filename.
10. Takes a screenshot of the completed Excel report.

## Automation Flow

Google Chrome  
↓  
Google Finance - GOOG  
↓  
Copy webpage information  
↓  
Extract GOOG price  
↓  
Generate current date and time  
↓  
Microsoft Excel  
↓  
Create daily report  
↓  
Save Excel file  
↓  
Take screenshot

## Technologies Used

- Python
- PyAutoGUI
- Pyperclip
- Google Chrome
- Microsoft Excel

## Requirements

- Windows
- Python 3
- Google Chrome
- Microsoft Excel

## Installation

Clone the repository:

```bash
git clone https://github.com/YOUR_USERNAME/daily-report-bot.git
cd daily-report-bot