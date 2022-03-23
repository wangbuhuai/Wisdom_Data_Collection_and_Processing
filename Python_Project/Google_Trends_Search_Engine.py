# Created by Dayu Wang (dwang@stchas.edu) on 2022-03-15

# Last updated by Dayu Wang (dwang@stchas.edu) on 2022-03-17


import tkcalendar
import tkinter as tk
from ctypes import windll
from tkinter import ttk


def main():
    app = tk.Tk()
    app.title("Wisdom Google Trends Data Collector")
    app.attributes("-topmost", 1)
    app.iconbitmap("./images/Wisdom_Google_Trends_Data_Collector_-_Icon_-_ICO.ico")

    # Label of app name
    ttk.Label(app,
              text="Wisdom Google Trends Data Collector",
              font="Verdana 16 bold",
              foreground="#7f1428") \
        .grid(row=0, column=0, columnspan=2, padx=12, pady=12)

    # Label of developer
    ttk.Label(app,
              text="Application developed by Dayu Wang (dwang@stchas.edu)",
              font="Verdana 12 italic",
              foreground="darkcyan") \
        .grid(row=1, column=0, columnspan=2, padx=12, pady=12)

    # Date picker of start date
    start_date = ttk.Frame(app)
    ttk.Label(start_date, text="Start Date", font="Verdana 14") \
        .grid(row=0, column=0, padx=12, pady=12)
    tkcalendar.DateEntry(start_date,
                         selectmode="day",
                         year=2004,
                         month=1,
                         day=1,
                         firstweekday="sunday",
                         locale="en_US",
                         date_pattern="mm/dd/y",
                         showweeknumbers=False,
                         font="Verdana 12") \
        .grid(row=1, column=0, columnspan=2)
    start_date.grid(row=2, column=0, padx=16, pady=16)

    # Date picker of end date
    end_date = ttk.Frame(app)
    ttk.Label(end_date, text="End Date", font="Verdana 14") \
        .grid(row=0, column=0, padx=12, pady=12)
    tkcalendar.DateEntry(end_date,
                         selectmode="day",
                         year=2021,
                         month=12,
                         day=31,
                         firstweekday="sunday",
                         locale="en_US",
                         date_pattern="mm/dd/y",
                         showweeknumbers=False,
                         font="Verdana 12") \
        .grid(row=1, column=0, columnspan=2)
    end_date.grid(row=2, column=1, padx=16, pady=16)

    # Search terms
    keywords = ttk.Frame(app)
    ttk.Label(keywords,
              text="Search terms (one search term per row)",
              font="Verdana 14") \
        .grid(row=0, column=0, padx=12, pady=12)
    tk.Text(keywords,
            font="Verdana 14",
            height=8) \
        .grid(row=1, column=0, padx=12, pady=12)
    keywords.grid(row=3, column=0, columnspan=2, padx=12, pady=12)

    # Suggestions
    suggestions = ttk.Frame(app)
    on_off = tk.StringVar()

    # Style settings
    s = ttk.Style()
    s.configure("TRadiobutton", font="Verdana 14")
    s.configure("TButton", font="Verdana 14 bold")

    ttk.Label(suggestions,
              text="Turn on suggestions?",
              font="Verdana 14") \
        .grid(row=0, column=0, padx=12, pady=12)
    ttk.Radiobutton(suggestions,
                    text="Yes",
                    style="TRadiobutton",
                    value="Yes",
                    variable=on_off) \
        .grid(row=0, column=1, padx=12, pady=12)
    ttk.Radiobutton(suggestions,
                    text="No",
                    style="TRadiobutton",
                    value="No",
                    variable=on_off) \
        .grid(row=0, column=2, padx=12, pady=12)
    on_off.set("No")
    tk.Text(suggestions,
            font="Verdana 14",
            height=8) \
        .grid(row=1, column=0, columnspan=3, padx=12, pady=12)
    suggestions.grid(row=4, column=0, columnspan=2, padx=12, pady=12)

    # Index
    index = ttk.Frame(app)
    index_value = tk.StringVar()
    ttk.Label(index,
              text="Index:",
              font="Verdana 14") \
        .grid(row=0, column=0, padx=12, pady=12)
    ttk.Entry(index,
              textvariable=index_value,
              font="Verdana 14") \
        .grid(row=0, column=1, padx=0, pady=12)
    index_value.set("1")
    index.grid(row=5, column=0, columnspan=2, padx=12, pady=12, sticky=tk.W)

    # Locale
    locale = ttk.Frame(app)
    locale_value = tk.StringVar()
    locale_value.set("US")
    ttk.Label(locale,
              text="Locale:",
              font="Verdana 14") \
        .grid(row=0, column=0, padx=12, pady=12)
    ttk.Entry(locale,
              textvariable=locale_value,
              font="Verdana 14",
              state="readonly") \
        .grid(row=0, column=1, padx=0, pady=12)
    ttk.Label(locale,
              text="(cannot change)",
              font="Verdana 14") \
        .grid(row=0, column=2, padx=12, pady=12)
    locale.grid(row=6, column=0, columnspan=2, padx=12, pady=12, sticky=tk.W)

    # Output directory
    output = ttk.Frame(app)
    dir_name = tk.StringVar()
    ttk.Label(output,
              text="Output Directory:",
              font="Verdana 14") \
        .grid(row=0, column=0, padx=12, pady=12)
    ttk.Entry(output, textvariable=dir_name, font="Verdana 14") \
        .grid(row=0, column=1, padx=0, pady=12)
    ttk.Button(output, text="Browse", style="TButton") \
        .grid(row=0, column=2, padx=12, pady=12)
    output.grid(row=7, column=0, columnspan=2, padx=12, pady=12, sticky=tk.W)

    # Search button
    ttk.Button(app, text="Start Search Process", style="TButton") \
        .grid(row=8, column=0, columnspan=2, padx=12, pady=12, ipadx=12, ipady=12)

    try:
        windll.shcore.SetProcessDpiAwareness(1)
    finally:
        app.mainloop()


if __name__ == "__main__":
    main()
