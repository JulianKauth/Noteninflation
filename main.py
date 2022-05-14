import os

import matplotlib.pyplot as plt
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

    columns = ["Year", "State", "Participants", "Passed", "Failed", "FailedPercent", "AvgGrade", "Grade", "Count"]

    rows = []  # cols in the Excel sheet, but let's move to the dataframe state of mind as soon as possible
    for year in range(2006, 2021 + 1):
        filename = f"/share/Programming/PycharmProjects/Noteninflation/Aus_Abiturnoten_{year}.xlsx"
        print(f"Reading Workbook {filename}")
        ws = openpyxl.load_workbook(filename=filename, read_only=True)["Noten"]
        for col in "BCDEFGHIJKLMNOPQ":
            row = [year, ws[col + "5"].value]
            row.extend([cell[0].value for cell in ws[f"{col}6:{col}10"]])
            grade_counts = [int(cell[0].value) for cell in ws[f"{col}12:{col}42"]]
            for i, count in enumerate(grade_counts):
                grade = (i + 10) / 10
                rows.append(row + [grade, count])

    re = pd.DataFrame(rows, columns=columns)
    pd.to_pickle(re, pickle_file)
    return re


def simple_plot(df, kind, x, y, title="Test"):
    print(df)
    df.reset_index().plot(kind=kind, x=x, y=y)
    plt.title(title)
    plt.show()


def plot(df):
    kind = "line"
    x = "Year"
    y = "AvgGrade"
    title = "Test"

    ax = plt.gca()
    for state, state_df in df.groupby("State"):
        state_df.groupby(["Year"]).mean().reset_index().plot(kind=kind, x=x, y=y, ax=ax)

    plt.title(title)
    plt.show()


if __name__ == '__main__':
    dataframe = load_data()

    # simple_plot(dataframe.groupby(["Year"]).sum().reset_index(), "bar", "Year", "Participants")
    # simple_plot(dataframe.groupby(["Grade", "Year"]).sum(), "bar", "Grade", "Count")
    # simple_plot(dataframe.groupby(["Year"]).mean(), "line", "Year", "AvgGrade")
    # simple_plot(dataframe.groupby(["Year"]).mean(), "line", "Year", "FailedPercent")
    # simple_plot(dataframe[dataframe.State == "BY"].groupby(["Year"]).mean(), "line", "Year", "AvgGrade")
    plot(dataframe)
