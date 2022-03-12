# Created by Dayu Wang (dwang@stchas.edu) on 2022-03-11

# Last updated by Dayu Wang (dwang@stchas.edu) on 2022-03-11


from re import findall, search


def extract_year_and_month(date):
    """ Extracts the year and month from a date string.
        :param date: a string representing a date
        :type date: str
        :return: extracted year and month in a dictionary
        :rtype: Dict
    """
    year = search(r"^[0-9]{4}|[0-9]{4}$", date)
    month = findall(r"(?<![0-9])[0-9]{1,2}(?![0-9])", date)[0]
    return {
        "year": int(year.group()),
        "month": int(month)
    }


def main():
    date_1 = "2022-03-11"
    date_2 = "03/11/2022"
    date_3 = "03-11-2022"
    date_4 = "2022/03/11"

    print(extract_year_and_month(date_1))
    print(extract_year_and_month(date_2))
    print(extract_year_and_month(date_3))
    print(extract_year_and_month(date_4))


if __name__ == "__main__":
    main()
