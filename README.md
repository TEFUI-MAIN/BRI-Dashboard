# TEFUI Dashboard

Executive dashboard for the Sydney Outstanding Orders Report.

The dashboard is designed for Business Directors and General Manager visibility across:

- Outstanding customer backlog
- Orders due today, overdue orders, extremely overdue orders, and future due load
- Client-specific exposure by orders, lines, and units
- Service severity and operational pressure
- Drilldown views for client backlog and future load dates

## Live Site

Once GitHub Pages is enabled, the dashboard will usually be available at:

```text
https://tefui-main.github.io/BRI-Dashboard
```

## How To Use

1. Open the dashboard link.
2. Click **Choose File**.
3. Select the latest Outstanding Orders Excel file.
4. The dashboard will process the file locally in your browser.
5. Click any client box to open the detailed order drilldown.
6. Click any future load date to review the orders due on that date.
7. Use **Export PDF** if you need a presentation copy.

## Data Privacy

The Excel file is processed locally in the browser.

The dashboard does not upload the spreadsheet to a server, database, or third-party service.

## Updating The Dashboard On GitHub

To update the app:

1. Replace the existing `index.html` file in the GitHub repository.
2. Commit the change to the `main` branch.
3. Wait for GitHub Pages to refresh.
4. Open the dashboard and hard refresh with `Cmd + Shift + R`.

## Excel File Requirements

The dashboard tries to automatically detect the required columns. It supports common column names such as:

- `Order Date`
- `Despatch Due Date` or `Due Date`
- `Status` or `Order Status`
- `Customer`
- `Receiver Name`
- `Client Reference`, `Order Number`, or `Sales Order`
- `Order Units`, `Units`, `Qty`, or `Outstanding Qty`
- `Order Line Count` or `Lines`

If the workbook contains both a pivot sheet and a raw data sheet, the dashboard scans the workbook and uses the sheet that looks like the raw outstanding orders data.

## Current Dashboard Sections

- **KPI Cards**: Total outstanding orders, extremely overdue orders, overdue orders, due today, future due orders, and oldest order age.
- **Client Backlog Control**: Client-specific boxes showing orders, lines, units, oldest order age, and the worst service-severity colour for that client.
- **Future Load Risk**: Calendar-style view of future due dates with drilldown overlays.
- **Detail Overlays**: Clickable client and future-load views showing row-based order-level information without horizontal scrolling.

## Due Date Logic

If an uploaded order has a blank due date, the dashboard calculates a due date as:

```text
Calculated Due Date = Order Date + 2 days
```

That calculated due date is then used for all KPI cards, client colours, future load, and drilldown views.

## Colour And Severity Logic

The dashboard uses due date severity to colour client cards and order rows:

- **Blue**: Future due date.
- **Green**: Due today.
- **Amber**: Overdue by 1-2 days.
- **Red**: Extremely overdue, meaning overdue by more than 2 days.

A client card uses the worst order severity inside that client. For example, if a client has mostly future orders but one extremely overdue order, the client card will be red.

## At-Risk Logic

An order is marked **At Risk** when:

- The due date is overdue, and
- The status does not appear to be completed, closed, delivered, invoiced, or cancelled.

## Suggested Next Enhancements

- Combine with DIFOT data.
- Add truck capacity and driver roster data.
- Add customer priority tiers.
- Add postcode and route data.
- Add automated alerts for overdue or high-risk orders.
