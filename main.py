import pandas as pd
import openpyxl


def load_data() -> pd.DataFrame:
    """
    returns a dataframe that contains _all_ the data.
    Each row has a Year, State, Participants, Passed, Failed, Grade10, Grade11, Grade{12-40}
    """

    columns = ["Year", "State", "Participants", "Passed", "Failed",
               "Grade10", "Grade11", "Grade12", "Grade13", "Grade14",
               "Grade15", "Grade16", "Grade17", "Grade18", "Grade19",
               "Grade20", "Grade21", "Grade22", "Grade23", "Grade24",
               "Grade25", "Grade26", "Grade27", "Grade28", "Grade29",
               "Grade30", "Grade31", "Grade32", "Grade33", "Grade34",
               "Grade35", "Grade36", "Grade37", "Grade38", "Grade39",
               "Grade40"]
    re_df = pd.DataFrame(columns=columns)

    rows = []  # cols in the Excel sheet, but let's move to the dataframe state of mind as soon as possible
    for year in range(2006, 2021 + 1):
        filename = f"/share/Programming/PycharmProjects/Noteninflation/Aus_Abiturnoten_{year}.xlsx"
        print(f"Reading Workbook {filename}")
        ws = openpyxl.load_workbook(filename=filename, read_only=True)["Noten"]
        for col in "BCDEFGHIJKLMNOPQ":
            row = [year]
            row.extend([cell[0].value for cell in ws[f"{col}5:{col}8"]])
            row.extend([cell[0].value for cell in ws[f"{col}12:{col}42"]])
            rows.append(row)

    return pd.DataFrame(rows, columns=columns, dtype=int)


if __name__ == '__main__':
    df = load_data()
    print(df)
