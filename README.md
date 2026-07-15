# Shipment Tracking Automation

An automation pipeline built for a logistics planning team to track inbound shipments. It consolidates data from multiple sources daily (carrier websites/APIs, customs broker reports, warehouse inbound schedules, etc.) into a single PO tracking spreadsheet, giving planners real-time visibility into shipment status, estimated arrival times, and inbound priority.

> This project originated as an internal company logistics tracking system and is published here with company authorization as a portfolio piece. All credentials, API keys, internal network paths, and real names have been removed or replaced with environment variables / generic placeholders. The program logic itself is unchanged.

## Purpose

- Improve planners' visibility into shipment status and cut down time spent manually consolidating reports
- Automatically calculate inbound priority/urgency for each shipment based on stock levels and shipment status

## Architecture

A main script (`Main.py`) runs seven core modules in sequence, plus five carrier-tracking submodules and two shared utility modules:

| # | Module | Description |
|---|--------|-------------|
| 1 | `sales_data_sync.py` | Syncs the latest PO sales data into the master sheet |
| 2 | `cargoo.py` | Uses browser automation (Playwright) to log into the Cargoo platform and download/import shipment reports |
| 3 | `carrier_tracking_main.py` | Calls five carrier-specific modules to fetch real-time shipment status, cross-referencing the Cargoo data |
| 4 | `customs.py` | Consolidates reports from multiple customs brokers to get clearance progress and ETA |
| 5 | `warehouse.py` | Integrates warehouse inbound schedules to project next-day inbound plans |
| 6 | `compare_last_file.py` | Compares the current and previous report versions to track ETA change history |
| 7 | `sort_query_sap.py` | Refreshes Excel Power Query and SAP-linked data used for inbound priority calculations |

### Carrier tracking submodules (`carrier_tracking/`)

Fetches real-time shipment status via official APIs (e.g., Maersk), unofficial APIs, or web scraping (Evergreen, Wan Hai, Yang Ming, ONE), covering roughly 60-70% of orders.

### Shared utilities (`support_module/`)

- `support_function.py`: shared helpers for file I/O and copying data between Excel sheets
- `email_automation.py`: automatically fetches email attachments (e.g., inbound notices) via the Outlook COM interface

## Tech Stack

- Python 3.12
- Playwright (browser automation)
- pandas / openpyxl / xlwings (Excel data processing)
- requests (carrier API integration)
- pywin32 (Outlook automation, Windows only)

## Setup

```bash
pip install -r requirements.txt
playwright install chromium
```

The program expects the following environment variables (the real credentials/keys have been removed from this codebase and must be set up independently):

```bash
CARGOO_EMAIL=your_email@example.com
CARGOO_PASSWORD=your_password
MAERSK_CLIENT_ID=xxx
MAERSK_CLIENT_SECRET=xxx
MAERSK_CLIENT_ID_2=xxx        # if you have a second set of API credentials
MAERSK_CLIENT_SECRET_2=xxx
```

Path constants at the top of each module (e.g. `YOUR_NETWORK_DRIVE\ShipmentTracking\...`) are generic placeholders for the original internal network drive paths — replace them with your own folder paths before running.

## Usage

1. Open `Main.py` and confirm the source/temp folder paths
2. Choose whether to create a new file or use an existing one (comment out the unused option)
3. Choose which of the seven modules to run (comment out any you want to skip)
4. Run `Main.py` — it will update each section of the master Excel sheet in sequence

A full run takes about 16 minutes; the carrier tracking module (`carrier_tracking_main`) accounts for over half of that, so it can be skipped if not needed.

## Disclaimer

This project is shared as a technical portfolio piece. Third-party service names mentioned in the code (Cargoo, Maersk, customs brokers, etc.) are included only to describe the data sources being integrated and do not imply any business partnership. All original sensitive information was removed prior to publication.
