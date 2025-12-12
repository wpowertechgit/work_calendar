# Work Hours Tracker (Tkinter)

A desktop application built with **Python + Tkinter** to track daily work hours, automatically calculate weekly and monthly balances, and export professional Excel reports.

The app is designed for real-world work schedules, including **compensation of missed days**, **holidays**, and **weekend handling**.

---

## ✨ Features

- 📅 Calendar-based daily input
- 💾 Automatic saving (JSON persistence)
- 🧮 Correct weekly & monthly calculations
- 🏖 Holiday support (reduces weekly obligation)
- 📊 Excel export with tables & summaries
- 🖥 Live daily details in the GUI

---

## ⌨️ Input Symbols

| Input | Meaning | Effect |
|------|--------|------|
| `9-17` | Worked interval | Counts real worked hours |
| `0` or empty | No work | No hours worked |
| `-` | Missed workday | Requires compensation within the week |
| `X` / `x` | Holiday | Reduces weekly required hours by 4 |
| Weekend | Non-working day | Ignored unless hours are entered |

---

## 🧠 Calculation Logic (Important)

### Daily
- **Normal workday** = 4 hours
- Extra hours are any hours worked beyond 4
- Missed days do **not subtract worked hours**
- Holidays remove the daily obligation

### Weekly
- Base requirement: **5 × 4 = 20 hours**
- Each holiday reduces the requirement by 4 hours
- Week is **OK** if:
total_worked_hours ≥ total_required_hours

### Monthly
- Monthly extra hours are calculated as:
max(0, total_worked - total_required)

This ensures that:
- Extra hours automatically compensate missed days
- Holidays are handled correctly
- Weekends do not inflate requirements

---

## 📊 Excel Export

The Excel report includes:

### Daily Table
- Date
- Input
- Worked hours
- Daily program (4h)
- Difference
- Holiday flag
- Notes

### Weekly Summary
- Worked hours per week
- Required hours per week
- Difference
- Status (OK / INCOMPLETE)

### Monthly Summary
- Total worked hours
- Total required hours
- Total extra hours for the month

---

## 🛠 Requirements

- Python **3.10+**
- Required packages:
```bash
pip install tkcalendar openpyxl
▶️ How to Run

Clone or download the project

Make sure the file is NOT named calendar.py

Run:
```bash
python work_calendar.py
```
All data is automatically saved to:
work_hours.json

