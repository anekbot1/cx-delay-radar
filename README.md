# CX Delay Radar

A client-side web tool for analyzing FOP (Flight Operations Performance) delay data.

## Usage

1. Open `index.html` in a browser (or use GitHub Pages)
2. Drop an FOP delay Excel spreadsheet onto the page
3. Explore the dashboard

**No server required** — everything runs in the browser. Your data never leaves your machine.

## Features

- **KPIs**: Total flights, departure/arrival OTP, average variance, total delay hours
- **Delay codes**: Top delay codes ranked by total minutes and frequency
- **Delay groups**: Doughnut chart of delay categories
- **OTP by day of week**: Spot weekday vs weekend patterns
- **Monthly trend**: OTP and average variance over time
- **Worst routes**: Routes with lowest OTP (min 5 operations)
- **Worst flights**: Flight numbers with highest average delays
- **Worst equipment**: Aircraft (DN) with persistent delays
- **Reactionary delays**: Flights delayed due to late inbound
- **Delay responsibility**: Breakdown by responsibility code

## Filters

Filter by departure, arrival, flight number, aircraft type, equipment ID, and season. Adjustable delay threshold (default 15 min).

## Data Format

Expects the standard FOP delay dataset schema with columns like:
- `SDD (UTC)`, `FN`, `DA`, `AA`, `STD (UTC)`, `ATD (UTC)`, `DTV`, `ATV`
- `DC1`-`DC5`, `DG1`-`DG5`, `DM1`-`DM5` (delay attribution)
- `DN` (equipment), `DOW`, `MTH`, `SSN`, etc.

## Tech

- [SheetJS](https://sheetjs.com/) for Excel parsing
- [Chart.js](https://www.chartjs.org/) for visualizations
- Zero dependencies to install — CDN-loaded
