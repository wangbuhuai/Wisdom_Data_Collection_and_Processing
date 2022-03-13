# Created by Dayu Wang (dwang@stchas.edu) on 2022-03-11

# Last updated by Dayu Wang (dwang@stchas.edu) on 2022-03-12


from csv import reader, writer
from os import listdir
from pandas import read_excel
from re import findall, search
from time import time
from tkinter import Tk, filedialog


ROW_BEGIN = 0
ROW_END = float("inf")


def extract_year_and_month(date_str):
    """ Extracts the year and month from a date string.
        :param date_str: a string representing a date
        :type date_str: str
        :return: extracted year and month in a dictionary
        :rtype: Dict
    """
    year = search(r"^[0-9]{4}|[0-9]{4}$", date_str)
    month = findall(r"(?<![0-9])[0-9]{1,2}(?![0-9])", date_str)[0]
    return {
        "year": int(year.group()),
        "month": int(month)
    }


def main():
    # Record the time.
    time_start = time()

    # Open the database containing the split names.
    Tk().withdraw()
    pathname_1 = filedialog.askopenfilename(
        filetypes=[("MS Excel Spreasheet", ".xlsx .xls .xlsm .xlm")],
        title="Open the database containing extracted first names and last names:")
    name_tokens = read_excel(pathname_1)

    # Open the huge database to begin data merge.
    pathname_2 = filedialog.askopenfilename(
        filetypes=[("CSV File", ".csv")],
        title="Open the huge database:")
    huge_database = open(pathname_2, encoding='utf-8', errors='ignore')
    database_handler = reader(huge_database)

    # Open the directory that contains reliable SVI data.
    pathname_3 = filedialog.askdirectory(title="Open the directory with reliable SVI data:")

    # Open the output database.
    if ROW_BEGIN == 0:
        output_database = open("Output_Database_2.csv", 'w', newline='', encoding="utf-8")
    else:
        output_database = open("Output_Database_2.csv", 'a', newline='', encoding="utf-8")
    output_handler = writer(output_database)

    first_row = ROW_BEGIN == 0
    row_index = 0
    try:
        for row_index, row in enumerate(database_handler):
            if row_index == 706615:
                continue
            if first_row:
                output_row = row
                output_row.extend(["CEO_INDEX", "GOOGLE_TRENDS_KEY", "SVI", "SVI_DATE"])
                output_handler.writerow(output_row)
                first_row = False
                continue

            if row_index < ROW_BEGIN:
                continue
            if row_index > ROW_END:
                break

            owner = row[6].lower()
            for index, ceo in name_tokens.iterrows():
                first_name = ceo["Extracted First Name"].lower()
                last_name = ceo["Extracted Last Name"].lower()

                if search(r"(?<![a-z])" + first_name + r"(?![a-z])", owner) is not None \
                        and search(r"(?<![a-z])" + last_name + r"(?![a-z])", owner) is not None:
                    output_row = row
                    for filename in listdir(pathname_3):
                        if int(filename[:3]) == int(ceo["Number Index"]):
                            svi = open("%s/%s" % (pathname_3, filename))
                            svi_handler = reader(svi)

                            header_row = True
                            for svi_record in svi_handler:
                                if header_row:
                                    header_row = False
                                    continue

                                svi_date = extract_year_and_month(svi_record[0])
                                t_date = extract_year_and_month(row[18])  # Transaction date

                                if svi_date["year"] == t_date["year"] and svi_date["month"] == t_date["month"]:
                                    ceo_index = int(ceo["Number Index"])
                                    google_trends_key = ceo["Full Search Name"]
                                    svi_value = int(svi_record[1])
                                    svi_exact_date = svi_record[0]

                                    output_row.extend([ceo_index, google_trends_key, svi_value, svi_exact_date])
                                    output_handler.writerow(output_row)
                                    break
                            svi.close()
                            break
                    break
    except UnicodeDecodeError:

        print("UnicodeDecodeError occurred in row %d." % row_index)
    finally:
        # Close the input database file.
        huge_database.close()

        # Close the output database file.
        output_database.close()

        # Record the time again.
        time_end = time()

        hours = (time_end - time_start) / 3600
        minutes = (time_end - time_start) % 3600 / 60
        seconds = (time_end - time_start) % 3600 % 60
        print("Data processing took %d hour(s) %d minute(s) %d second(s)." % (hours, minutes, seconds))


if __name__ == "__main__":
    main()
