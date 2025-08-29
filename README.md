# Flexible Daily Timecard

A production-quality Excel timecard generator built with Python and openpyxl.

## Features

- **Flexible Time Tracking**: Track multiple clock-in/out intervals throughout the day
- **Real-time Calculations**: Automatically calculates hours worked, remaining time, and projections
- **Goal Setting**: Configurable daily work goal (default: 8 hours)
- **Smart Projections**: Shows projected end time based on current work pace
- **Professional Formatting**: Clean, professional appearance with proper borders and colors
- **Formula Safety**: All formulas use absolute references to prevent accidental breakage

## Requirements

- Python 3.6+
- openpyxl library

## Installation

```bash
pip install openpyxl
```

## Usage

```bash
python flexible_timecard.py
```

This will generate `Flexible_Timecard.xlsx` in the current directory.

## How It Works

1. **Set Daily Goal**: Enter your target hours in cell B3
2. **Log Work Intervals**: Enter clock-in and clock-out times for each work period
3. **Automatic Calculations**: The spreadsheet calculates:
   - Duration of each interval
   - Total hours worked
   - Time remaining to reach goal
   - Projected completion time
4. **Real-time Updates**: All calculations update automatically as you enter data

## File Structure

- **Header Section**: Date, date goal, current time
- **Time Tracking**: 10 rows for work intervals with automatic duration calculation
- **Summary Section**: Real-time totals and projections
- **Instructions**: Clear usage guidelines

## Technical Details

- Uses Excel formulas for all calculations
- Locale-agnostic (comma separators)
- Absolute cell references ($) for stability
- Professional styling with borders and colors
- Data validation for time entries
