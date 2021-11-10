import datetime
import os
import warnings
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Tuple, Union
from collections import defaultdict

import numpy as np
import pandas as pd
import dynamic_yaml


# TODO make it be able to read from a URL.


def _cm_to_inch(length):
    return np.divide(length, 2.54)


def resolve_list(config_list):
    return [config_list[x] for x in range(len(config_list))]


def resolve_dictionary(config_dict):
    return {k: v for k, v in config_dict.items()}


# Directory constants

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


class ExcelFormats:
    """
    Container class for all the formats
    """

    def __init__(self,
                 workbook,
                 base: Dict,
                 on_hold: Dict,
                 at_risk: Dict,
                 behind_schedule: Dict,
                 on_track: Dict,
                 mandatory_input: Dict,
                 new_project: Dict,
                 new_pm: Dict,
                 client_project_number_error: Dict,
                 header: Dict,
                 mandatory_cols: List[str],
                 new_pms: List,
                 new_projects: List,
                 client_project_number_errors: List,
                 ):
        self.workbook = workbook
        self.base = base
        self.header = header
        self.on_hold = on_hold
        self.at_risk = at_risk
        self.behind_schedule = behind_schedule
        self.on_track = on_track
        self.mandatory_input = mandatory_input
        self.new_project = new_project
        self.new_pm = new_pm
        self.client_project_number_error = client_project_number_error
        self.mandatory_cols: List = mandatory_cols
        self.new_pms = new_pms
        self.new_projects = new_projects
        self.client_project_number_errors = client_project_number_errors

    def check_new_pm(self, pm) -> bool:
        return True if pm in self.new_pms else False

    def check_new_project(self, project) -> bool:
        return True if project in self.new_projects else False

    def check_st_project_number_error(self, project) -> bool:
        return True if project in self.client_project_number_errors else False

    def check_mandatory_column(self, column, value) -> bool:
        check_mandatory_col = True if column in self.mandatory_cols else False
        check_empty_value = True if value in ["", " ", None] else False
        return check_empty_value and check_mandatory_col

    def get_base_format(self):
        return self.workbook.add_format(self.base)

    def get_schedule_format(self,
                            value: str,
                            column: str,
                            pm: str,
                            project: int
                            ):
        cell_format = dict(self.base)
        if value == "On Hold":
            cell_format.update(self.on_hold)
        elif value == "Behind Schedule":
            cell_format.update(self.behind_schedule)
        elif value == "On Track":
            cell_format.update(self.on_track)
        elif value == "At risk of being delayed":
            cell_format.update(self.at_risk)
        elif value == "":
            pass
        else:
            raise ValueError(f"Schedule not defined: {value}")

        if self.check_mandatory_column(column, value):
            cell_format.update(self.mandatory_input)

        if self.check_new_pm(pm):
            cell_format.update(self.new_pm)

        if self.check_new_project(project):
            cell_format.update(self.new_project)

        if self.check_st_project_number_error(project):
            cell_format.update(self.client_project_number_error)

        return self.workbook.add_format(cell_format)

    def get_header_format(self):
        return self.workbook.add_format(self.header)

    def get_format(self, format_type: str, column: str, pm: str, project: int, value=None):
        if format_type == "schedule":
            return self.get_schedule_format(
                value=value,
                column=column,
                pm=pm,
                project=project
            )

        if format_type == "header":
            return self.get_header_format()

        return self.get_base_format()


class ExcelGenerator:
    protected_cell_modifier: tuple = (0, -1, 0, -1)  # [row_start, row_finish, col_start, col_finish
    footer_format_str: str = "&CPage &P of &N"
    zoom_percentage = 60
    a3_paper = 8

    def __init__(self,
                 workbook,
                 margins: Dict[str, float],
                 issue: int,
                 df: pd.DataFrame,
                 formats: ExcelFormats,
                 pm_col: str,
                 project_col: str,
                 column_widths: List[Union[float, int]],
                 ghd_image_path: str = None,
                 client_image_path: str = None,
                 n_repeat_rows: int = 0,
                 schedule_col: str = "Schedule"
                 ):
        self.pm_col = pm_col
        self.project_col = project_col
        self.schedule_col = schedule_col
        self.workbook = workbook
        self.worksheets: Dict[str,] = {}
        self.margins: Dict[str, float] = margins  # expects dict of 4 strings int/floats in units inches
        self.issue = issue
        self.ghd_image_path = ghd_image_path
        self.client_image_path = client_image_path
        self.n_repeat_rows = n_repeat_rows
        self.column_widths = column_widths
        self.df = df
        self.formats: ExcelFormats = formats

    @property
    def data_width(self):
        return self.df.shape[0]

    @property
    def data_height(self):
        return self.df.shape[1]

    @property
    def header_string(self):
        return f'&L&[Picture]&C&14&"Arial,Bold"GHD Quarterly Dashboard\nIssue {self.issue}: ({self.get_month})&R&[' \
               f'Picture]'

    @property
    def protected_cells(self):
        # [row 1 (headers take 1 row), last row of test_data, column A, last column of test_data
        return [1, self.data_height + 1, 0, self.data_width + 1]

    @property
    def printable_cells(self):
        # I think this needs to be modified bc excel referes to zero index on the backend, vs the 1 index in the UI
        return [sum(x) for x in zip(self.protected_cells, self.protected_cell_modifier)]

    def set_print_area(self, worksheet):
        worksheet.print_area(*self.printable_cells)
        return worksheet

    @property
    def get_month(self) -> str:
        """
        returns this month as a string representation. i.e. October.
        """
        return datetime.today().strftime("%B")

    def get_worksheet(self, sheet_name):
        if sheet_name in self.worksheets:
            return self.worksheets[sheet_name]

        raise ValueError(f"Sheet: {sheet_name} already exists in the workbook")

    def add_worksheet(self, sheet_name: str):
        if sheet_name in self.worksheets:
            raise ValueError(f"Sheet: {sheet_name} already exists in the workbook")
        worksheet = self.workbook.add_worksheet(sheet_name)
        self.worksheets[sheet_name] = worksheet

    def setup_worksheet(self, sheet_name: str):

        worksheet = self.get_worksheet(sheet_name)
        worksheet.set_page_view()
        worksheet.set_landscape()
        worksheet.set_zoom(self.zoom_percentage)
        worksheet.hide_gridlines()  # default = 1 = hide when printed.
        worksheet.set_header(
            self.header_string,
            {
                "image_left": self.ghd_image_path,
                "image_right": self.client_image_path,
            }
        )
        worksheet.set_footer(self.footer_format_str)
        worksheet.set_margins(
            left=self.margins["left"],
            right=self.margins["right"],
            top=self.margins["top"],
            bottom=self.margins["bottom"],
        )
        worksheet.repeat_rows(self.n_repeat_rows)
        worksheet = self.set_print_area(worksheet)
        worksheet.set_paper(self.a3_paper)

    def add_data_validation(self,
                            range_reference: Tuple[int, int, int, int],
                            data_validation: Dict[str, str],
                            sheet_name: str):
        worksheet = self.get_worksheet(sheet_name)
        worksheet.data_validation(*range_reference, data_validation)

    def write_row(self, worksheet, row, row_idx):
        # To format the entire row, we need to set some of the format arguments before the cell formatting. We pass
        # these down so that the cell can get the right formatting to match the row + cell specific,
        # i.e. for formatting for mandatory fields
        format_args = dict(
            format_type="schedule",
            pm=row[self.pm_col],
            project=row[self.project_col],
            value=row[self.schedule_col],
        )
        for col_idx, (column, value) in enumerate(zip(row.index, row.values)):
            # set the remainder of the args for the cell formatting
            format_args.update({"column": column})
            cell_format = self.formats.get_format(**format_args)
            self.write_cell(worksheet, row_idx, col_idx, value, cell_format)

    def write_cell(self, worksheet, row_idx: int, col_idx: int, value, cell_format):

        worksheet.write(row_idx, col_idx, value, cell_format)

    def write_to_sheet(self, sheet_name):
        worksheet = self.get_worksheet(sheet_name)
        for row_idx, row in self.df.iterrows():
            self.write_row(worksheet, row, row_idx)

    def setup_header(self, sheet_name):
        header_row_idx = 0
        header_format = self.formats.get_header_format()
        sheet = self.get_worksheet(sheet_name)
        for col_num, (col_width, col_name) in enumerate(zip(self.column_widths, self.df.columns)):
            sheet.write(header_row_idx, col_num, col_name, header_format)
            sheet.set_column(col_num, col_num, col_width)


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
