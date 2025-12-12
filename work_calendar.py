import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from tkcalendar import Calendar
import json
import os
from datetime import datetime
from collections import defaultdict
from openpyxl import Workbook

DATA_FILE = "work_hours.json"
NORMAL_HOURS = 4
BASE_WEEKLY_MIN = 20


def is_weekend(date_str):
    d = datetime.strptime(date_str, "%Y-%m-%d")
    return d.weekday() >= 5


def load_data():
    if os.path.exists(DATA_FILE):
        with open(DATA_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}


def save_data(data):
    with open(DATA_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)


def calculate_hours(entry, date_str):
    entry = entry.strip().lower()

    if is_weekend(date_str):
        if entry in ("", "0", "-", "x"):
            return 0, 0, 0, "Weekend (nelucrător)"

    if entry == "x":
        return 0, 0, 1, "Sărbătoare legală"

    if entry == "-":
        return 0, 0, 0, "Zi lucrătoare lipsă (−4h)"

    if entry in ("", "0"):
        return 0, 0, 0, "Zi liberă"

    try:
        start, end = entry.split("-")
        total = int(end) - int(start)
        if total <= 0:
            return 0, 0, 0, "Interval invalid"

        diff = total - NORMAL_HOURS
        note = "Lucrat în weekend" if is_weekend(date_str) else "Zi lucrătoare"
        return total, diff, 0, note
    except:
        return None, None, None, None


class WorkCalendarApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Evidență Ore Lucrate")

        self.data = load_data()

        self.calendar = Calendar(root, selectmode="day", date_pattern="yyyy-mm-dd")
        self.calendar.grid(row=0, column=0, columnspan=2, padx=10, pady=10)

        ttk.Label(
            root,
            text="Interval: 9-17 | 0 = nimic | - = lipsă | X = sărbătoare"
        ).grid(row=1, column=0, columnspan=2, sticky="w", padx=10)

        self.entry = ttk.Entry(root, width=20)
        self.entry.grid(row=2, column=1)
        ttk.Label(root, text="Input:").grid(row=2, column=0)
        self.entry.bind("<Return>", self.save_entry)

        ttk.Button(root, text="Salvează ziua", command=self.save_entry).grid(row=3, column=0, pady=5)
        ttk.Button(root, text="Export Excel", command=self.export_excel).grid(row=3, column=1, pady=5)

        self.info = ttk.Label(root, text="", justify="left")
        self.info.grid(row=4, column=0, columnspan=2, pady=10)

        self.calendar.bind("<<CalendarSelected>>", self.load_selected_date)

    def load_selected_date(self, event=None):
        date = self.calendar.get_date()
        self.entry.delete(0, tk.END)

        if date in self.data:
            d = self.data[date]
            self.entry.insert(0, d["raw"])
            self.info.config(
                text=(
                    f"Data: {date}\n"
                    f"Input: {d['raw']}\n"
                    f"Ore lucrate: {d['total']}\n"
                    f"Diferență: {d['diff']}\n"
                    f"Sărbătoare: {d['holiday']}\n"
                    f"Observații: {d['note']}"
                )
            )
        else:
            self.info.config(text=f"Data: {date}\nNicio înregistrare")

    def save_entry(self, event=None):
        date = self.calendar.get_date()
        raw = self.entry.get()

        total, diff, holiday, note = calculate_hours(raw, date)
        if total is None:
            messagebox.showerror("Eroare", "Format invalid")
            return

        self.data[date] = {
            "raw": raw,
            "total": total,
            "diff": diff,
            "holiday": holiday,
            "note": note
        }

        save_data(self.data)
        self.load_selected_date()

    def export_excel(self):
        selected = datetime.strptime(self.calendar.get_date(), "%Y-%m-%d")
        month_name = selected.strftime("%B")
        year = selected.year

        wb = Workbook()
        ws = wb.active
        ws.title = f"Raport Pontaj {month_name} {year}"

        ws.append([
            "Data", "Input", "Ore lucrate",
            "Program (4h)", "Diferență", "Sărbătoare", "Observații"
        ])

        weekly_worked = defaultdict(int)
        weekly_required = defaultdict(int)

        monthly_worked = 0
        monthly_required = 0

        for date in sorted(self.data):
            d = datetime.strptime(date, "%Y-%m-%d")
            if d.month != selected.month or d.year != selected.year:
                continue

            week = d.isocalendar()[1]
            entry = self.data[date]

            worked = entry["total"]
            holiday = entry["holiday"]

            # daily obligation
            required = 0
            if not is_weekend(date) and not holiday:
                required = NORMAL_HOURS

            weekly_worked[week] += worked
            weekly_required[week] += required

            monthly_worked += worked
            monthly_required += required

            ws.append([
                date,
                entry["raw"],
                worked,
                NORMAL_HOURS,
                worked - required,
                "DA" if holiday else "",
                entry["note"]
            ])

        # Weekly report
        ws.append([])
        ws.append(["=== RAPORT SĂPTĂMÂNAL ==="])

        for week in sorted(weekly_worked):
            diff = weekly_worked[week] - weekly_required[week]
            ws.append([
                f"Săptămâna {week}",
                "",
                weekly_worked[week],
                weekly_required[week],
                diff,
                "",
                "OK" if diff >= 0 else "INCOMPLET"
            ])

        # Monthly summary
        ws.append([])
        ws.append(["=== SUMAR LUNAR ==="])

        monthly_extra = max(0, monthly_worked - monthly_required)

        ws.append(["Ore lucrate total", "", monthly_worked])
        ws.append(["Ore obligatorii total", "", monthly_required])
        ws.append(["Ore suplimentare lună", "", monthly_extra])

        path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            initialfile=f"raport_pontaj_{month_name.lower()}_{year}.xlsx",
            filetypes=[("Excel", "*.xlsx")]
        )

        if path:
            wb.save(path)
            messagebox.showinfo("Export", "Fișier Excel creat cu succes")



if __name__ == "__main__":
    root = tk.Tk()
    WorkCalendarApp(root)
    root.mainloop()
