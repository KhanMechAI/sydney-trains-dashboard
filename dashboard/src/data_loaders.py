from collections import defaultdict
from pathlib import Path
from typing import Any, Dict, List

import pandas as pd

from utils import DynamicResolver, drop_empty_rows, reset_index


class DataLoader:
    EXCEL_FILES = (".xlsx", ".xlsm")

    def __init__(self,
                 data_path: [str, Path],
                 exclusions: Dict[str, List],  # {column string: [list of exclusions]}
                 column_order: List[str],
                 date_cols: [None, List[str]] = None,
                 ):

        self.exclusions = exclusions
        self.column_order = column_order
        self.date_cols: [None, List[str]] = date_cols
        self.df: pd.DataFrame
        self.resolver = DynamicResolver()
        self.data_path = data_path
        self.load_data(Path(self.data_path))

    def __repr__(self):
        return self.data_path.name

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
            self.df = self.df[~self.df[column].isin(self.resolver.resolve(exclusions))]

        self.df = reset_index(self.df)

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

    def filter_to(self, filter_col: str, filter_val: Any):
        self.df = self.df[self.df[filter_col] == filter_val]

    def filter_out(self, filter_col: str, filter_val: Any):
        self.df = self.df[self.df[filter_col] != filter_val]

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
        return self.df.append(df_to_append, verify_integrity=True, ignore_index=True)

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

    def remove_inactive(self, activity_col: str, activity_value: Any):
        self.filter_out(activity_col, activity_value)


class DataConsolidator:

    def __init__(self,
                 bst_loader: BSTLoader,
                 project_manager_col: str,
                 project_col: str,
                 column_order: List[str],
                 previous_dashboard: [None, DataLoader] = None,
                 project_manager_sheets: [None, List[DataLoader]] = None,
                 ):
        self.project_manager_col = project_manager_col
        self.project_col = project_col
        self.column_order = column_order
        self.new_data = defaultdict()
        self.bst: BSTLoader = bst_loader
        self.previous_dashboard: [DataLoader, None] = previous_dashboard
        self.project_manager_sheets: [List[DataLoader], None] = project_manager_sheets
        self.master: pd.DataFrame = pd.DataFrame()
        self.pm_df: pd.DataFrame = pd.DataFrame()

    @property
    def new_project_managers(self):
        return self.new_data["project_managers"]

    @property
    def new_projects(self):
        return self.new_data["projects"]

    @property
    def project_managers(self):
        return self.master[self.project_manager_col].unique()

    @property
    def projects(self):
        return self.master[self.project_col].unique()

    def get_new_project_managers(self) -> set:
        new_pms = self.bst.projects(self.project_manager_col)
        if self.previous_dashboard is not None:
            prev_pms = self.previous_dashboard.projects(self.project_manager_col)
            new_pms = new_pms.intersection(prev_pms)
        return new_pms

    def get_new_projects(self) -> set:
        new_projects = self.bst.projects(self.project_col)
        if self.previous_dashboard is not None:
            prev_projects = self.previous_dashboard.projects(self.project_col)
            new_projects = new_projects.intersection(prev_projects)
        return new_projects

    def find_new_data(self):
        self.new_data["project_managers"] = self.get_new_project_managers()
        self.new_data["projects"] = self.get_new_projects()

    def create_master(self) -> pd.DataFrame:
        if "projects" not in self.new_data:
            raise AttributeError("New test_data not searched. Please run 'find_new_data()' first")

        filtered_bst: pd.DataFrame = self.bst.select_in(self.project_col, self.new_data["projects"])

        if self.previous_dashboard is not None:
            self.master = self.previous_dashboard.append(filtered_bst)
        else:
            self.master = filtered_bst

        self.master = drop_empty_rows(self.master)
        return self.master

    def join_project_manager_sheets(self):
        if self.project_manager_sheets is None:
            raise ValueError("No project manager sheets passed")

        for pm_sheet in self.project_manager_sheets:
            self.pm_df = pm_sheet.append(self.pm_df)

        self.pm_df.drop_duplicates(inplace=True)

        self.pm_df = drop_empty_rows(self.pm_df)

        return

    def load_pm_sheets_to_master(self):
        self.join_project_manager_sheets()
        master = self.master.copy(deep=True)
        master.set_index(keys=self.project_col, inplace=True)

        pm_df = self.pm_df.copy(deep=True)
        pm_df.set_index(keys=self.project_col, inplace=True)

        master.update(pm_df, overwrite=True)
        master.reset_index(inplace=True, drop=False)

        master = drop_empty_rows(master)

        self.master = master[self.column_order]

    def filter(self, filter_col: str, filter_val: Any):
        return self.master[self.master[filter_col] == filter_val]

    def filter_by_project_manager(self, project_manager: str):
        return self.filter(self.project_manager_col, project_manager)

    def filter_by_project(self, project: str):
        return self.filter(self.project_col, project)