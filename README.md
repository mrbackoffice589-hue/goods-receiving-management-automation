# Goods Receiving Management Automation

An automated goods receiving and management workflow built with Google Apps Script.

## Overview

This project demonstrates how Google Workspace services can be connected into an event-driven workflow for recording received goods and generating structured daily reports.

## Features

- Create a Google Form for goods receiving
- Capture supplier, receiving officer, item, quantity, unit, condition, temperature, notes, and photo references
- Create a daily menu form and spreadsheet database
- Store system document IDs in Apps Script Properties instead of hard-coding private IDs
- Install a Google Sheets form-submit trigger
- Generate a daily goods receiving report from submitted records
- Organize generated files by year, month, and date
- Export reports to PDF and DOCX
- Build a 3×3 photo documentation gallery
- Provide a manual function for testing the latest submission

## Technology Stack

- Google Apps Script
- Google Forms
- Google Sheets
- Google Docs
- Google Drive
- Apps Script Triggers
- Google Drive API export endpoints

## Workflow

```text
Goods Receiving Form
        |
        v
Google Sheets Response Database
        |
        v
Form Submit Trigger
        |
        +----> Daily Menu Database
        |
        +----> Report Template
        |
        v
Daily Report Generation
        |
        +----> DOCX
        +----> PDF
        +----> Photo Documentation
```

## Setup

1. Create a Google Apps Script project.
2. Add the contents of `goods-receiving-management-automation.gs`.
3. Run `createGoodsReceivingForm()` to create the goods receiving form.
4. Run `createDailyMenuForm()` to create the daily menu form and database.
5. Run `createReportTemplate()` to create the report template.
6. Run `installAutomation()` to store the generated file IDs and install the form-submit trigger.
7. Submit a test goods receiving entry and use `testLatestReport()` if a manual test is required.

The functions require Google Workspace authorization because they access Forms, Sheets, Docs, Drive, and Apps Script services.

## Portfolio / Security Notes

This repository is intended as a portfolio example. Do not commit passwords, API keys, OAuth credentials, private spreadsheet IDs, private URLs, personal information, or confidential operational data.

The source code uses Apps Script Properties for document IDs rather than hard-coding private IDs.

## License

No license is currently specified. Contact the repository owner before reusing the code in another project.
