# 📊 Student Behavior Tracking Dashboard

An automated Excel VBA dashboard built for real-world use at the Flushing YMCA. The tool aggregates student behavior data across multiple tracking sheets, calculates performance metrics, and generates dynamic charts — all triggered from a single Master Sheet interface.

---

## 🔍 Overview

Staff manually log student behavior across daily Excel sheets. This dashboard was built to eliminate the manual work of compiling that data — automatically pulling records by student name, calculating compliance rates across 5 behavioral criteria, identifying the area most needing improvement, and rendering a live chart. Two analysis modes are supported: **all-time summary** and **week-filtered summary**.

---

## 📸 Demo

<img width="1752" height="882" alt="image" src="https://github.com/user-attachments/assets/4cb77af8-189b-4fb8-a28c-a87b594f9323" />

<img width="1388" height="909" alt="image" src="https://github.com/user-attachments/assets/46109608-e154-4b6c-862a-1d3de6f55ee3" />


---

## ⚙️ Features

### 📋 Two Analysis Modes
| Mode | Description |
|------|-------------|
| **All-Time Summary** | Aggregates all available data for a student across every sheet |
| **Weekly Summary** | Filters data to a selected 7-day window via dropdown |

### 🤖 Automation
- **Dynamic week dropdown** — auto-generates 52 upcoming Mondays as selectable dates using a named range and Excel data validation — no manual list maintenance
- **Auto-generated charts** — programmatically creates and updates a clustered column chart on every run, positioned dynamically relative to existing dashboard elements
- **Focus behavior detection** — automatically identifies and flags the behavioral criterion with the highest failure rate as the priority area for improvement
- **Multi-sheet aggregation** — loops across all sheets matching the student's name to compile a complete behavioral record

### 📊 Behavioral Criteria Tracked
1. Calm Body
2. Listening Ears
3. Kind Words
4. Stay in Area
5. Finished Work

Each criterion is reported as **% Met** and **% Needs Improvement** across the selected time range.

---

## 🛠️ Tech Stack

| Tool | Purpose |
|------|---------|
| Excel VBA | Macro automation, data processing, chart generation |
| Excel Data Validation | Dynamic dropdown for week selection |
| Excel Charts | Clustered column chart rendered programmatically |
| Named Ranges | Dynamic reference for dropdown list |

---

## 📂 Project Structure

```
behavior-dashboard/
│
├── BehaviorTracker.xlsm        # Main workbook with embedded macros
├── vba_code.vb                 # Exported VBA source code (readable without Excel)
└── README.md
```

---

## 🛠️ VBA Modules

| Subroutine | Purpose |
|-----------|---------|
| `BehaviorSummary_PercentOnly_WithFocus` | All-time behavior summary — aggregates all sheets for a student |
| `BehaviorSummary_Weekly_WithFocus` | Weekly behavior summary — filters by selected date range |
| `SetupWeekDropdown_Limited` | Generates 52-week dropdown in F2 using dynamic named range |
| `CreateOrUpdateWeeklyChart` | Creates or refreshes clustered column chart from summary data |
| `AddRunButton` | Programmatically adds the Run button to the Master Sheet |
| `AddUpdateWeeklyChartButton` | Programmatically adds the Update Weekly Chart button |

---

## 🚀 Getting Started

### Prerequisites
- Microsoft Excel with macros enabled (`.xlsm`)

### 1. Clone or download the repository
```bash
git clone https://github.com/yourusername/behavior-dashboard.git
```

### 2. Open the workbook
Open `BehaviorTracker.xlsm` in Excel and **enable macros** when prompted.

### 3. Set up the dashboard
1. Run `SetupWeekDropdown_Limited` to generate the week dropdown in cell `F2`
2. Run `AddRunButton` to add the **Run Behavior Summary** button
3. Run `AddUpdateWeeklyChartButton` to add the **Update Weekly Chart** button

### 4. Run a summary
1. Enter a student name in cell `F1` on the Master Sheet
2. Select a week from the `F2` dropdown (for weekly analysis)
3. Click **Run Behavior Summary** for all-time data, or **Update Weekly Chart** for weekly data
4. View the generated summary table and chart on the Master Sheet

---

## 💡 Engineering Highlights

- **Dynamic named range** — `WeekStartList` is rebuilt programmatically each run, always reflecting the current week forward — no hardcoded dates
- **Chart create-or-update pattern** — checks for an existing chart by name before creating a new one, preventing duplicates across repeated runs
- **Focus behavior algorithm** — iterates all 5 criteria to find the maximum false percentage, then surfaces that label as the priority recommendation
- **Sheet name matching** — uses `InStr` for case-insensitive partial matching, allowing flexible student sheet naming conventions
- **Boolean type checking** — uses `VarType` to safely distinguish boolean cell values from empty or non-boolean entries before counting

---

## 💡 Future Improvements

- [ ] Export summary report to PDF with one click
- [ ] Add trend line showing improvement over multiple weeks
- [ ] Support filtering by date range (not just a single week)
- [ ] Email summary report directly to staff via Outlook VBA integration
- [ ] Add class-wide summary across all students

---

## 📄 License

MIT License — free to use and modify.
