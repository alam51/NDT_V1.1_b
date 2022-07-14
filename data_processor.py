import pandas as pd
import matplotlib.pyplot as plt
# from scipy import stats


class DfProcessor:
    """
    raw_df = pandas df with column=None and index=None
    source =
        1. 's' for SCADA (after last update)
        2. 'sb' for SCADA (before last update)
        3. to be updated....
    """

    def __init__(self, file_path: str, source=''):
        source = input(f"Source?\n"
                       f"'' (blank): Formatted\n"
                       f"'s': SCADA Archive\n"
                       f"'e': EMS\n")

        if source.lower() == 's':
            if file_path.endswith('.csv'):
                raw_df = pd.read_csv(file_path, skiprows=[1], parse_dates=[0],
                                     infer_datetime_format=True, index_col=[0], dayfirst=False)
                self.df = raw_df.dropna(axis=1, thresh=1)
                # for j in self.df.columns:
                #     self.df[str(j)].replace(to_replace=0, method='bfill', inplace=True)

            else:
                raw_df = pd.read_excel(file_path, skiprows=[1], parse_dates=[0],
                                       index_col=[0])
                self.df = raw_df.dropna(axis=1, thresh=1)
                # for j in self.df.columns:
                #     self.df[str(j)].replace(to_replace=0, method='bfill', inplace=True)

        elif source.lower() == 'ems':
            raw_df = pd.read_csv(file_path, skiprows=None, parse_dates=[0, 4],
                                 infer_datetime_format=True, index_col=None, dayfirst=False)
            raw_df1 = raw_df.dropna(axis=1, thresh=5)
            raw_df = raw_df1.dropna(axis=0, how='any')
            raw_df1 = raw_df.set_index('Time')
            self.df = raw_df1.loc[:, ['Value', 'Value.1']]
            self.df.columns = ['Freq', 'MW']

        elif source == '':
            df = pd.read_csv(file_path, skiprows=None, parse_dates=[0],
                             infer_datetime_format=True, index_col=[0], dayfirst=False)
            self.df = df.dropna(axis=1, thresh=1)
            # for j in self.df.columns:
            #     self.df[str(j)].replace(to_replace=0, method='bfill', inplace=True)

        self.df = self.df[self.df >= 2]
