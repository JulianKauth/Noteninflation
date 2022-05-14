import os

import pandas as pd
import openpyxl


def load_data() -> pd.DataFrame:
    """
    returns a dataframe that contains _all_ the data.
    Each row has a Year, State, Participants, Passed, Failed, Grade10, Grade11, Grade{12-40}
    """
    pickle_file = "/share/Programming/PycharmProjects/Noteninflation/dataframe.pkl"
    if os.path.exists(pickle_file):
        return pd.read_pickle(pickle_file)

    columns = ["Year", "State", "Participants", "Passed", "Failed", "Grade", "Count"]

    rows = []  # cols in the Excel sheet, but let's move to the dataframe state of mind as soon as possible
    for year in range(2006, 2021 + 1):
        filename = f"/share/Programming/PycharmProjects/Noteninflation/Aus_Abiturnoten_{year}.xlsx"
        print(f"Reading Workbook {filename}")
        ws = openpyxl.load_workbook(filename=filename, read_only=True)["Noten"]
        for col in "BCDEFGHIJKLMNOPQ":
            row = [year, ws[col + "5"].value]
            row.extend([int(cell[0].value) for cell in ws[f"{col}6:{col}8"]])
            grade_counts = [int(cell[0].value) for cell in ws[f"{col}12:{col}42"]]
            for i, count in enumerate(grade_counts):
                grade = (i + 10) / 10
                rows.append(row + [grade, count])

    re = pd.DataFrame(rows, columns=columns)
    pd.to_pickle(re, pickle_file)
    return re


if __name__ == '__main__':
    df = load_data()
    print(df)
