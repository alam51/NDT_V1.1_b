import os
from dateutil.parser import parse as date_parse
import pandas as pd
from openpyxl.utils import get_column_letter, column_index_from_string
from openpyxl.utils.cell import coordinate_from_string
from typing import Union


class Excel:
    def __init__(self, file_path: str, sheet_name: str):
        self.file_path = file_path
        self.sheet_name = sheet_name

    def slice(self, top_left_cell: str, lower_right_cell: str) -> pd.DataFrame:
        """
        Let, top_left_cell = B3, lower_right_cell = F8
        slices B3:F8 and returns as dataframe
        """
        left_col, top_row = coordinate_from_string(top_left_cell)
        right_col, bottom_row = coordinate_from_string(lower_right_cell)
        df = pd.read_excel(self.file_path,
                           sheet_name=self.sheet_name,
                           header=None,
                           index_col=None,
                           usecols=f'{left_col}:{right_col}',
                           skiprows=top_row - 1,
                           nrows=bottom_row - top_row,
                           )
        return df

    def vlookup(self, lookup_value_series: pd.Series, lookup_array_top_left_cell: str,
                lookup_array_lower_right_cell: str,
                x_shift: int) -> pd.Series:
        """
        x_shift = cells to shift along x-axis
        """
        lookup_df = self.slice(lookup_array_top_left_cell, lookup_array_lower_right_cell)
        output_series = pd.Series()
        for value in lookup_value_series:
            matched_row = lookup_df[lookup_df.iloc[:, 0] == value]
            if not matched_row.empty:
                output_series.loc[len(output_series)] = matched_row.iloc[0, x_shift]
            else:
                output_series.loc[len(output_series)] = pd.NA
        return output_series


file_path = r'I:\My Drive\IMD\Analysis\Generation\daily report P2 with fuel type.xlsm'
xl1 = Excel(file_path=file_path, sheet_name='P2')
# df1 = xl1.slice('C9', 'C200')
# s1 = df1.iloc[:, 0]  # for converting df to series
# op = xl1.vlookup(s1, 'C9', 'I150', 6)
# op.to_excel('lol.xlsx')
# b = 4

reference_frame = xl1.slice('C9', 'H174')
reference_series = xl1.slice('C9', 'C174').iloc[:, 0]

# vlookup_result_series = xl1.vlookup(reference_series, 'C9', 'I150', 6)
# reference_frame.loc[:, reference_frame.columns[-1] + 1] = vlookup_result_series
b = 0

folder_directory = r'C:\Users\hE\Downloads\July2022\July2022'
files = os.listdir(folder_directory)
excel_files = [file for file in files if file.endswith('.xlsx') or file.endswith('.xlsm')]
for file in excel_files:
    full_file_path = os.path.join(folder_directory, file)
    xl = Excel(file_path=full_file_path, sheet_name='P2')
    vlookup_result_series = xl.vlookup(reference_series, 'C9', 'I182', 6)
    date = date_parse(file, dayfirst=True, yearfirst=False, fuzzy=True)
    # reference_frame.loc[:, f'{date.year}.{date.month}/{date.day}'] = vlookup_result_series
    reference_frame.loc[:, f'{date.day - 1} {date.strftime("%b")}'] = vlookup_result_series

reference_frame.to_excel('op.xlsx')
b = 0