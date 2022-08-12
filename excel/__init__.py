import pandas as pd
from openpyxl.utils import get_column_letter, column_index_from_string
from openpyxl.utils.cell import coordinate_from_string


class Excel:
    def __init__(self, file_path: str, sheet_name=0):
        self.file_path = file_path
        self.sheet_name = sheet_name

    def slice(self, top_left_cell:str, lower_right_cell:str) -> pd.DataFrame:
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

    def vlookup(self, lookup_value_list, lookup_array_top_left_cell:str, lookup_array_lower_right_cell:str,
                x_shift:int) -> pd.Series:
        """
        x_shift = cells to shift along x-axis
        """
        lookup_array = self.slice(lookup_array_top_left_cell, lookup_array_lower_right_cell)
        output_list = pd.Series()
        for value in lookup_value_list:
            matched_row = lookup_array[lookup_array.iloc[:, 0] == value]
            if matched_row:
                output_list.loc[len(output_list)] = matched_row[x_shift]
            else:
                output_list.loc[len(output_list)] = pd.NA
        return output_list
