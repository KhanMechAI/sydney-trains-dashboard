from pathlib import Path
from typing import Any, Dict, List

import pandas as pd
from pathlib import Path
from typing import Any, Dict, List

import pandas as pd


class DataLoader:
    EXCEL_FILES = (".xlsx", ".xlsm")

    def __init__(self,
                 data_path: [str, Path],
                 exclusions: Dict[str, List],  # {column string: [list of exclusions]}
                 column_order: List[str],
                 date_cols: [None, List[str]] = None):

        self.exclusions = exclusions
        self.column_order = column_order
        self.date_cols: [None, List[str]] = date_cols
        self.df: pd.DataFrame

        self.load_data(Path(data_path))

    def add_missing_columns(self):
        for col in self.column_order:
            if col not in self.df.columns:
                self.df[col] = ""

    def date_time_handler(self):
        """
        Converts datetimes to string representation to handle non-standard datetime input.
        """
        if self.date_cols is None:
            return
        for col in self.date_cols:
            if col in self.df.columns:
                self.df[col] = self.df[col].astype(str)

    def exclude(self):

        for column, exclusions in self.exclusions.items():
            self.df = self.df[~self.df[column].isin(exclusions)]

    def clean_data(self):
        self.df = self.df.loc[:, ~self.df.columns.str.contains("^Unnamed")]
        self.date_time_handler()
        self.df = drop_empty_rows(self.df)

    def load_data(self, path: [str, Path]):
        data_path = Path(path)
        if not data_path.exists():
            raise FileNotFoundError
        elif data_path.suffix not in self.EXCEL_FILES:
            raise ValueError(f"Input file type not one of {self.EXCEL_FILES}")

        self.df: pd.DataFrame = pd.read_excel(data_path, index_col=None)
        self.clean_data()
        return

    def sort(self, sort_col, ascending=False):
        self.df.sort_values(by=sort_col, ascending=ascending, inplace=True)

    def filter(self, filter_col: str, filter_val: Any):
        self.df = self.df[self.df[filter_col] == filter_val]

    def remove_duplicates(self, duplicate_col: [str, list, None] = None):
        self.df.drop_duplicates(subset=duplicate_col, keep="first", inplace=True, ignore_index=True)

    def select_columns(self, columns: List[str]):
        self.df = self.df[columns]

    def rename_columns(self, column_mapping: dict):
        self.df.rename(columns=column_mapping, inplace=True, )

    def rename_index(self, new_name: str):
        self.df.index.rename(new_name, inplace=True)

    def project_managers(self, project_manager_column):
        return set(self.df[project_manager_column].unique())

    def projects(self, project_column):
        return set(self.df[project_column].unique())

    def set_index(self, index_column: [str, List]):
        self.df.set_index(keys=index_column, inplace=True)

    def select_in(self, column: str, values: [list]) -> pd.DataFrame:
        return self.df[self.df[column].isin(values)]

    def clear_data(self, columns: List[str]):
        self.df[columns] = ""

    def append(self, df_to_append: pd.DataFrame) -> pd.DataFrame:
        return self.df.append(df_to_append)

    def set_column_order(self):
        self.df = self.df[self.column_order]


class BSTLoader(DataLoader):
    dup_filter_col: str
    DEFAULT_SHEET_IDX = 0

    def __init__(self, path: [str, Path], exclusions: Dict[str, List], column_order: List[str],
                 date_cols: [None, List[str]] = None):
        super().__init__(path, exclusions, column_order, date_cols)

    def remove_proposals(self,
                         proposal_col: str,
                         proposal_str: str,
                         prop_number_lb: int = 210900000,
                         prop_number_ub: int = 210999999):
        self.df = self.df[self.df[proposal_col] != proposal_str]  # Removes BST10 props
        self.df.drop([proposal_col], inplace=True, axis=1)  # Remove the task code column
        self.df = self.df[
            (self.df.index < prop_number_ub) | (
                    self.df.index > prop_number_lb)]  # remove projects in proposal number range to cover historic
        # proposals

        return self.df
