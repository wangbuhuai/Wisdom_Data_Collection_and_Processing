# Created by Dayu Wang (dwang@stchas.edu) on 2022-03-10

# Last updated by Dayu Wang (dwang@stchas.edu) on 2022-03-11


from csv import reader, writer
from tkinter import Tk, filedialog


def main():
    # Open the "huge" CSV database.
    Tk().withdraw()
    pathname = filedialog.askopenfilename(filetypes=[("CSV File", ".csv")])

    output_database = open("Test_Database.csv", 'w', newline='')  # Output file
    output_handler = writer(output_database)

    with open(pathname) as input_database:
        input_handler = reader(input_database)
        row_index = 0
        for row in input_handler:
            output_handler.writerow(row)
            if row_index == 10000:
                break
            row_index += 1

    # Close the output file.
    output_database.close()


if __name__ == "__main__":
    main()
