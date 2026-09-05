# Excel Sales Dashboard

An interactive sales performance dashboard built in Microsoft Excel, tracking
130 sales executives across 8 regions with slicer-driven filtering and VBA
controlled pivot connections.

![Dashboard]  DashBoard Pic.png

## Features

- **Region slicer** filtering across Chennai, Delhi, Mumbai, Nagpur, Patna,
  Pune, Ranchi and Surat
- **Four toggleable dashboard panels**, each controlled by a checkbox that
  connects or disconnects the slicer from that panel's pivot table
- **VBA macro** (`SlicerConnection`) driving the checkbox logic
- **Top / bottom performer tables** ranked by Total Sales
- **Target achievement analysis** — Target Hit % and Away From Target %
- **Bar, pie and 3D line charts** linked to the underlying pivots

## Dataset

| Column | Description |
|--------|-------------|
| Emp Code | Unique employee identifier |
| Sales Executive | Name of the executive |
| Region | One of 8 city regions |
| Day1 – Day5 | Daily sales figures |
| Total Sales | Sum of Day1–Day5 |
| Target | Sales target (500) |
| Target Hit % | Total Sales / Target |
| Away From Target % | 1 − Target Hit % |

## Files

| File | Description |
|------|-------------|
| `Excel_DashBoard_Project.xlsm` | Main macro-enabled workbook |
| `SlicerConnection.bas` | Exported VBA module |
| `DashBoard_Pic.png` | Dashboard preview |

## How to use

1. Download `Excel_DashBoard_Project.xlsm`
2. Open in Excel and click **Enable Content** to allow macros
3. Select a region from the top slicer
4. Tick or untick the Dashboard 1–4 checkboxes to control which panels
   respond to the slicer

## Tools used

Microsoft Excel · Pivot Tables · Slicers · VBA · Charts · Conditional Formatting
