# Vendor Invoice Scraper 

This is a Python utility that extracts data from structured PDF forms (invoices) and compiles them into a clean Excel spreadsheet.

## Features

- GUI folder picker to select and process multiple PDFs at once  
- Extracts header fields (like Payee, Date, Fund, Department, and Account)  
- Parses multiple invoice rows, even with multi-line descriptions  
- Cleans and formats dates for consistency  
- Outputs to an `Invoice_Output.xlsx` file, appending new data if the file already exists  

## Requirements

Install dependencies using:

```bash
pip install -r requirements.txt
```

## License

This project is ©Megan Nowak and is not licensed for redistribution or modification. You may download and use the compiled version for personal use only.
