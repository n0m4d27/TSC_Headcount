# TSC Headcount

Desktop application for managing employee project allocations, forecasting hours, and planned staffing changes.

This project is a renamed, independently maintained fork of [Forecast-Tool](https://github.com/jbishop216/Forecast-Tool) (original author: jbishop216).

## Features

- **Employee Management**: Add, edit, and remove employees with their details
- **GA01 Week Configuration**: Set up GA01 weeks for accurate time calculations
- **Forecast Calculation**: Calculate and export monthly forecasts based on project allocations
- **Project Time Allocation**: Manage project time allocations by manager code and cost center
- **Planned Employee Changes**: Schedule and manage future employee changes
  - New hires
  - Terminations
  - Employment type conversions
- **Visualizations**: Generate charts and graphs for:
  - Monthly forecast hours
  - Manager allocations
  - Employee type distribution
- **Settings**: Configure FTE and contractor weekly hours
- **Excel Integration**: Export forecast data to Excel

## Tabs

1. **Employees**: Manage employee information (name, manager code, employment type, start/end dates)
2. **GA01 Weeks**: Set the number of working weeks for each month of the year
3. **Forecast**: Calculate and view forecasts based on project allocations and GA01 weeks
4. **Project Allocation**: Manage project allocations (cost center, work code, monthly hours)
5. **Planned Changes**: Track employee changes (new hires, conversions, terminations)
6. **Visualizations**: View various charts and analytics

## Technical Details

- Built with Python 3.11 and Tkinter
- SQLite database for data storage (`forecast_tool.db` in the working directory)
- SQLAlchemy ORM for database interaction
- Matplotlib for data visualization
- OpenPyXL for Excel export

## Requirements

See `requirements.txt`. Install with:

```bash
pip install -r requirements.txt
```

## Installation

1. Clone this repository:
   ```bash
   git clone https://github.com/n0m4d27/TSC_Headcount.git
   cd TSC_Headcount
   ```

2. Install required packages:
   ```bash
   pip install -r requirements.txt
   ```

3. Run the application (creates the SQLite database on first run if needed):
   ```bash
   python app_tkinter.py
   ```

## Publish your own copy to GitHub (n0m4d27)

On [github.com/new](https://github.com/new), create a repository named **`TSC_Headcount`** under account **n0m4d27**. Do **not** add a README, `.gitignore`, or license (keep the repo empty).

From this project folder, with [Git for Windows](https://git-scm.com/download/win) available in your terminal:

```powershell
cd path\to\TSC_Headcount
git remote rename origin upstream
git remote add origin https://github.com/n0m4d27/TSC_Headcount.git
git push -u origin main
```

- **`upstream`** keeps a reference to `https://github.com/jbishop216/Forecast-Tool` so you can pull upstream changes if you want.
- If you prefer not to keep the old remote, skip `rename` and use `git remote set-url origin https://github.com/n0m4d27/TSC_Headcount.git` instead.

Authenticate with GitHub when prompted (HTTPS personal access token, not account password), or use SSH: `git@github.com:n0m4d27/TSC_Headcount.git`.

If this working copy already has a remote named **`tsc-headcount`** pointing at that URL, you can push with:

```powershell
git push -u tsc-headcount main
```

## Usage

1. Start by adding employees in the Employees tab
2. Configure GA01 weeks for accurate time calculations
3. Add project allocations for each manager
4. Use the Forecast tab to calculate and view forecasts
5. Export data to Excel as needed
6. Track planned changes in the Planned Changes tab
7. View analytics in the Visualizations tab

## Mid-Month Employee Changes

The tool handles mid-month employee changes based on GA01 weeks:

- If changes (hiring, termination, conversion) occur after the 2nd GA01 week, employees are counted for the entire month
- If changes occur before the end of the 2nd GA01 week, they are prorated accordingly

## Settings

Access the Settings dialog from the File menu to configure:

- FTE weekly hours (default: 34.5)
- Contractor weekly hours (default: 39.0)
