from collections import defaultdict
from pathlib import Path
from typing import Any, Dict, List

import dynamic_yaml
import numpy as np
import pandas as pd

# Directory constants
from dashboard.excel import ExcelFormats, ExcelGenerator
from dashboard.utils import resolve_list

# TODO make it be able to read from a URL.

DASHBOARD_DIRECTORY = r"C:\Users\kschroder-turner\Documents\TEMP\Monthly Dashboards"


def reset_index(df, index_start: int = 1):
    df.reset_index(inplace=True, drop=True)

    # make index start from the new index start point, default index is 0, new default is 1
    df.index = df.index + index_start
    return df


def drop_empty_rows(df) -> pd.DataFrame:
    """
    Pandas doesnt recognise empty string as an empty value. So change all empty strings to nan, then drop and
    replace all nans with empty strings.
    """
    df.replace(["", " "], np.nan, inplace=True)
    df.dropna(inplace=True, how="all")
    df.replace(np.nan, "", inplace=True)

    # need to reset index after dropped rows.
    df = reset_index(df)

    return df


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


class DashboardGenerator:

    def __init__(self, config_path: [str, Path], exclusions_file_path: [str, Path], issue: int):

        with open(config_path) as file:
            self.config = dynamic_yaml.load(file, recursive=True)

        with open(exclusions_file_path) as file:
            self.exclusions = dynamic_yaml.load(file, recursive=True)
        self.ghd_logo_path = Path().cwd() / self.config.logos.ghd
        self.client_logo_path = Path().cwd() / self.config.logos.client
        self.issue = issue
        self.bst: BSTLoader
        self.previous_dashboard: DataLoader
        self.project_manager_loaders: List[DataLoader] = []
        self.consolidator: DataConsolidator

        self.margins = self.config.margins

    def _loader(self, data_path: [str, Path], is_pm: bool = False) -> DataLoader:
        loader = DataLoader(
            data_path=data_path,
            exclusions=self.exclusions,
            column_order=resolve_list(self.config.column.lists.col_order),
            date_cols=resolve_list(self.config.column.lists.date_cols),
        )
        loader.rename_columns(self.config.mapping.legacy)
        loader.remove_duplicates(duplicate_col=self.config.column.names.project_number)
        loader.add_missing_columns()
        if not is_pm:
            loader.clear_data(self.config.column.lists.mandatory)
        loader.set_column_order()
        return loader

    def _dashboard_maker(self, sheet_name: str, workbook, df: pd.DataFrame, formatter: ExcelFormats) -> ExcelGenerator:
        dashboard = ExcelGenerator(
            workbook=workbook,
            margins=self.margins["inch"],
            issue=self.issue,
            df=df,
            formats=formatter,
            pm_col=self.config.column.names.pm,
            project_col=self.config.column.names.project_number,
            column_widths=resolve_list(self.config.column.widths),
            ghd_image_path=self.ghd_logo_path,
            client_image_path=self.client_logo_path,
            schedule_col=self.config.column.names.sch,
        )
        dashboard.add_worksheet(sheet_name)
        dashboard.setup_worksheet(sheet_name)
        dashboard.setup_header(sheet_name)
        return dashboard

    def load_bst(self, bst_path: [str, Path]):
        self.bst = BSTLoader(
            path=bst_path,
            exclusions=self.exclusions,
            column_order=resolve_list(self.config.column.lists.col_order),
            date_cols=None,
        )
        self.bst.rename_columns(self.config.mapping.bst)
        self.bst.remove_duplicates(duplicate_col=self.config.column.names.project_number)
        self.bst.remove_proposals(
            proposal_col=self.config.column.names.task_code,
            proposal_str=self.config.proposal.code,
            prop_number_lb=self.config.proposal.lower_bound,
            prop_number_ub=self.config.proposal.upper_bound,
        )
        self.bst.select_columns(resolve_list(self.config.column.lists.bst_cols))
        self.bst.add_missing_columns()
        self.bst.set_column_order()

    def load_prev(self, dashboard_path: [str, Path]):
        self.previous_dashboard = self._loader(dashboard_path)

    def load_pm_sheets(self, pm_sheet_folder: [str, Path]):
        pm_sheet_folder = Path(pm_sheet_folder)
        for file_path in pm_sheet_folder.glob("*.xlsx"):
            self.project_manager_loaders.append(self._loader(file_path))

    def create_consolidator(self):
        self.consolidator = DataConsolidator(
            bst_loader=self.bst,
            project_col=self.config.column.names.project_number,
            project_manager_col=self.config.column.names.pm,
            column_order=resolve_list(self.config.column.lists.col_order),
            project_manager_sheets=self.project_manager_loaders
        )
        self.consolidator.find_new_data()
        self.consolidator.create_master()

    def update_master_with_pm(self):
        self.consolidator.load_pm_sheets_to_master()

    def create_formatter(self, workbook, new_pms, new_projects) -> ExcelFormats:
        formatter = ExcelFormats(
            workbook=workbook,
            base=self.config.formats.base,
            on_hold=self.config.formats.on_hold,
            at_risk=self.config.formats.at_risk,
            behind_schedule=self.config.formats.behind_schedule,
            on_track=self.config.formats.on_track,
            mandatory_input=self.config.formats.mandatory_input,
            new_project=self.config.formats.new_project,
            new_pm=self.config.formats.new_pm,
            client_project_number_error=self.config.formats.client_project_number_error,
            header=self.config.formats.header,
            mandatory_cols=resolve_list(self.config.column.lists.mandatory),
            new_pms=new_pms,
            new_projects=new_projects,
            client_project_number_errors=[],  # Ignoring for now as I dont think we need any more
        )
        return formatter

    def _create_dashboard(self, out_path: [str, Path], df, sheet_name: str = "Dashboard"):
        writer = pd.ExcelWriter(out_path, engine="xlsxwriter")
        workbook = writer.book
        formatter = self.create_formatter(
            workbook=workbook,
            new_pms=self.consolidator.new_project_managers,
            new_projects=self.consolidator.new_project_managers,
        )

        dashboard = self._dashboard_maker(
            sheet_name=sheet_name,
            workbook=workbook,
            df=df,
            formatter=formatter
        )
        dashboard.write_to_sheet(sheet_name)
        workbook.close()

    def create_master_dashboard(self, out_path: [str, Path], sheet_name: str = "Dashboard", ):

        self._create_dashboard(out_path, self.consolidator.master, sheet_name, )

    def get_pm_out_path(self, out_path: [str, Path], name, new_folder: str = None):
        if new_folder:
            out_path = out_path / new_folder / f"{name}.xlsx"
        else:
            out_path = out_path / f"{name}.xlsx"

        out_path.parent.mkdir(exist_ok=True, parents=True)
        return out_path

    def create_project_manager_sheets(self,
                                      out_path: [str, Path],
                                      new_folder: str = None,
                                      sheet_name: str = "Dashboard", ):

        for pm in self.consolidator.project_managers:
            pm_df = self.consolidator.filter_by_project_manager(pm)

            # Index's will be the same as in the master, so need to reset it so it starts at row 1 again
            pm_df = reset_index(pm_df)
            pm_out_path = self.get_pm_out_path(out_path, name=pm, new_folder=new_folder)

            self._create_dashboard(pm_out_path, pm_df, sheet_name, )


if __name__ == "__main__":

    #     #TODO: Sydney trains here

    config_path = Path().cwd() / "config_excel.yaml"
    exclusions_path = Path().cwd() / "exclusions.yml"

    test_out = Path().cwd() / "test"
    test_out.mkdir(exist_ok=True)

    bst_path = Path().cwd() / "test_data" / "Project Detail.xlsx"
    prev_dash = Path().cwd() / "test_data" / "Dashboard.xlsx"
    pm_sheets_path = Path().cwd() / "test_data" / "Project Manager Sheets"

    dashboard = DashboardGenerator(
        config_path=config_path,
        exclusions_file_path=exclusions_path,
        issue=6
    )

    dashboard.load_bst(bst_path)

    dashboard.load_pm_sheets(pm_sheets_path)

    dashboard.create_consolidator()

    dashboard.update_master_with_pm()

    dashboard.create_master_dashboard(test_out / "test.xlsx")

    # dashboard.create_project_manager_sheets(test_out, new_folder="Project Manager Sheets" )
