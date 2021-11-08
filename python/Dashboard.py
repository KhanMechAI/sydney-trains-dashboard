import datetime
import os
import warnings
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Tuple

import numpy as np
import pandas as pd


# TODO make it be able to read from a URL.


def _cm_to_inch(length):
    return np.divide(length, 2.54)


# Directory constants

DASHBOARD_DIRECTORY = r"C:\Users\kschroder-turner\Documents\TEMP\Monthly Dashboards"


class DataLoader:
    EXCEL_FILES = (".xlsx", ".xlsm")

    def __init__(self,
                 exclusions: Dict[str, List], # {column string: [list of exclusions]}
                 column_order: List,
                 index_col: [str, None]=None,
                 date_cols: [None, List[str]] = None):

        self.index_col = index_col
        self.exclusions = exclusions
        self.column_order = column_order
        self.date_cols: [None, List[str]] = date_cols

    def _index_handler(self, df) -> pd.DataFrame:
        df = df[~df[self.index_col].duplicated(keep="first")]
        df = df.loc[df.index.dropna()]
        df.index = df.index.astype("uint64")
        df.sort_index(inplace=True)
        return df

    def _add_missing_col(self, df) -> pd.DataFrame:
        for col in self.column_order:
            if col not in df.columns:
                df[col] = ""

        return df

    def _date_time_handler(self, df) -> pd.DataFrame:
        """
        Converts datetimes to string representation to handle non-standard datetime input.
        """
        df[self.date_cols] = df[self.date_cols].astype(str)
        return df

    def _exclude(self, df) -> pd.DataFrame:

        for column, exclusions in self.exclusions.items():
            df = df[~df[column].isin(exclusions)]

        return df

    def _clean_data(self, df) -> pd.DataFrame:
        df = df.loc[:, ~df.columns.str.contains("^Unnamed")]
        self._index_handler(df)
        self._add_missing_col(df)
        self._date_time_handler(df)
        self._exclude(df)
        return df

    def load_data(self, path: [str, Path]) -> pd.DataFrame:
        data_path = Path(path)
        if not data_path.exists():
            raise FileNotFoundError
        elif data_path.suffix not in self.EXCEL_FILES:
            raise ValueError(f"Input file type not one of {self.EXCEL_FILES}")

        df: pd.DataFrame = pd.read_excel(data_path, index_col=self.index_col)
        return self._clean_data(df)


class BSTLoader(DataLoader):
    DEFAULT_SHEET_IDX = 0

    def __init__(self, path: [str, Path], exclusions: Dict[str, List], column_order: List, index_col: int,):
        super().__init__(exclusions, column_order, index_col)

        self.index_col = index_col

        self.df: pd.DataFrame = self.load_data(path)

    def project_managers(self, project_manager_column):
        return set(self.df[project_manager_column].unique())

    def projects(self, project_column):
        return set(self.df[project_column].unique())

    def rename_index(self, new_name: str):
        self.df.index.rename(new_name, inplace=True)

    def remove_proposals(self,
                         proposal_col: str,
                         proposal_str: str,
                         prop_number_lb: int = 210999999,
                         prop_number_ub: int = 210900000):
        self.df = self.df[self.df[proposal_col] != proposal_str]  # Removes BST10 props
        self.df.drop([proposal_col], inplace=True, axis=1)  # Remove the task code column
        self.df = self.df[(self.df.index < prop_number_ub) | (
                self.df.index > prop_number_lb)]  # remove projects in proposal number range to cover historic
        # proposals

        return self.df

    def sort(self, sort_col):
        self.df.sort_values(by=sort_col, ascending=False, inplace=True)

    def filter(self, filter_col: str, filter_val: Any):
        self.df = self.df[self.df[filter_col] == filter_val]

    def select_columns(self, bst_columns: List[str]):
        self.df = self.df[bst_columns]


    def rename_columns(self, column_mapping: dict):
        self.df.rename(columns=column_mapping, inplace=True)



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

    def check_mandatory_column(self, column) -> bool:
        return True if column in self.mandatory_cols else False

    def get_base_format(self):
        return self.workbook.add_format(self.base)

    def get_schedule_format(self,
                            schedule: str,
                            column: str,
                            pm: str,
                            project: int
                            ):
        cell_format = self.base
        if schedule == "on_hold":
            cell_format.update(self.on_hold)
        elif schedule == "behind_schedule":
            cell_format.update(self.behind_schedule)
        elif schedule == "on_track":
            cell_format.update(self.on_track)
        elif schedule == "at_risk":
            cell_format.update(self.at_risk)
        else:
            raise ValueError(f"Schedule not defined: {schedule}")

        if self.check_mandatory_column(column):
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
                schedule=value,
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
                 out_path: [str, Path],
                 margins: Dict[str, float],
                 issue: int,
                 df: pd.DataFrame,
                 formats: ExcelFormats,
                 pm_col: str,
                 project_col: str,
                 column_widths: List[float, int],
                 ghd_image_path: str = None,
                 client_image_path: str = None,
                 n_repeat_rows: int = 1,
                 schedule_col: str = "Schedule"
                 ):
        self.pm_col = pm_col
        self.project_col = project_col
        self.schedule_col = schedule_col
        self.out_path = Path(out_path)
        self.writer = pd.ExcelWriter(self.out_path, engine="xlsxwriter")
        self.workbook = self.writer.book
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
        return f"&L&[Picture]&C&14&'Arial,Bold'GHD Quarterly Dashboard\nIssue {self.issue}: ({self.get_month})&R&[" \
               f"Picture]"

    @property
    def protected_cells(self):
        # [row 1 (headers take 1 row), last row of data, column A, last column of data
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
            project=row[self.project_col]
        )
        for col_idx, column, value in enumerate(zip(row.index, row.values)):
            #set the remainder of the args for the cell formatting
            format_args.update({"column": column, "value": value})
            self.write_cell(worksheet, row_idx, col_idx, value, format_args)

    def write_cell(self, worksheet, row_idx: int, col_idx: int, value, format_args: dict):

        cell_format = self.formats.get_format(**format_args)

        worksheet.write(row_idx, col_idx, value, cell_format)

    def write_to_sheet(self, sheet_name):
        worksheet = self.get_worksheet(sheet_name)
        for row_idx, row in self.df.iterrows():
            self.write_row(worksheet, row, row_idx)

    def setup_header(self, sheet):
        header_row = 0
        header_format = self.formats.get_header_format()
        header_format_obj = self.workbook.add_format(header_format)
        for col_num, col_width, col_name in enumerate(zip(self.column_widths, self.df.columns)):
            sheet.write(header_row, col_num, col_name, header_format_obj)
            sheet.set_column(col_num, col_num, col_width)


class DataConsolidator:
    num_cols = 10
    PM_SUB_DIR = "Project Manager Sheets"

    def __init__(self, config, bst_loader: BSTLoader, previous_dashboard: [None, DataLoader]=None, project_manager_sheets: [None, List[DataLoader]]=None, client=None):
        self.client = client
        self.config = config
        self.projects = set()
        self.project_managers = set()
        self._new_data = {}
        self._editable_cells = []
        self.new_data = {
            "pm": set(),
            "job": set(),
        }
        self.bst: BSTLoader = bst_loader
        self.previous_dashboard: DataLoader = previous_dashboard
        self.project_manager_sheets: DataLoader = project_manager_sheets

    def c

    def _load_conflict_handler(self, bst=True, other_df=None):
        if bst:
            # Keep jobs that are in both BST and the Prev Dashboard. = intersection of master with bst.
            # Add new jobs from BST to dashboard. = set difference of bst to master then append to master
            # Remove old jobs from dashboard. = negate the master bst intersection
            intersect_mask = np.in1d(self.df.index.values, self.bst._df.index.values, assume_unique=True)
            append_mask = np.in1d(self.bst._df.index.values, self.df.index.values, assume_unique=True, invert=True)
            # Remove old jobs
            # temp_bst = self.bst.df.loc[intersect_mask]
            self.bst._df.sort_index(inplace=True)
            self.df = self.df[intersect_mask]
            self.df.sort_index(inplace=True)

            current_pms = self.df[PM].unique()
            # Append new jobs
            self.df = self.df.append(self.bst._df[append_mask])
            self.df.sort_index(inplace=True)
            self.df[PM] = self.bst._df[PM]
            # Keep track of the new projects and project managers
            new_pms = np.in1d(self.bst._df[PM], current_pms, invert=True)
            self.new_data["job"] = set(self.bst._df.index.values[append_mask])
            self.new_data["pm"] = set(self.bst._df.loc[new_pms, PM])
        else:
            # TODO: This code is outdated, fix to match above

            if not other_df:
                raise ValueError("If bst=False, other_df must be specified")
            other_df_projects = set(other_df.index.values)
            proj_to_keep = self.projects.intersection(other_df_projects)
            other_df = other_df.loc[proj_to_keep]

            # ToDo: Check what BST_COLS used to be
            non_bst_cols = [x for x in self.config.lists.col_order if x not in BST_COLS]
            self.df.loc[proj_to_keep, non_bst_cols] = other_df[non_bst_cols]

    def load_prev_dashboard(self, path_to_dashboard):
        df = pd.read_excel(path_to_dashboard, index_col=0)
        df = self._load_helper(df)
        if self.df.empty:
            self.df = df
        else:
            self._load_conflict_handler(bst=False, other_df=df)

        self.projects = set(self.df.index.values)
        self.project_managers = set(self.df[PM].unique())

        self.df[[SCH, CUR_STAT, NEXT_ACTION, ACTION_BY]] = ""

    def load_pm(self, path, all_in_path=True, ):
        if all_in_path:
            for f_name in os.listdir(path):
                file = path / f_name
                if XLSX in f_name and TMP_FILE not in f_name:
                    self._load_pm(file)
        else:
            self._load_pm(path)

    def _load_pm(self, path):
        df = pd.read_excel(path, index_col=0)
        dups = self._index_dup_check(df, path)
        if dups:
            df = self._load_helper(df)
            df = self._check_pm_error(df)
            overwrite_mask = df.index.values
            self.df.loc[overwrite_mask] = df.loc[overwrite_mask]

    def _check_pm_error(self, df):
        # TODO: Print Error to logging file. So make a logging file as well....
        mask = np.isin(df.index.values, self.df.index.values)
        if len(df[~mask].index) > 0:
            msg0 = f"Project manager {df[PM].unique().astype(str)} might have errors. Please check the PM Sheet."
            msg1 = f"Project(s) found not in master: {df[~mask].index.values.astype(str)}"
            warnings.warn(msg0, Warning)
            warnings.warn(msg1, Warning)
        return df[mask]

    def _index_dup_check(self, df, path):
        len_idx_init = len(df.index)
        len_idx_fin = len(set(df.index))
        if len_idx_init != len_idx_fin:
            msg1 = f"Duplicate index values. Unique indices: {len_idx_init}, Total indicies: {len_idx_fin}. Skipping " \
                   f"{path.name}"
            warnings.warn(msg1, Warning)
            return False
        return True

    def export(self, path, pm=True, to_excel=True):
        if to_excel:
            path.mkdir(parents=True, exist_ok=True)
            if pm:
                pm_path = path / self.PM_SUB_DIR
                pm_path.mkdir(parents=True, exist_ok=True)
                for pm in self.df[PM].unique():
                    self._export_to_excel(pm_path, pm)
            self._export_to_excel(path, pm=False)

    def _export_to_excel(self, path, pm):
        def _setup_excel():
            nonlocal path
            path = self._get_output_name(path, pm=pm)
            writer = pd.ExcelWriter(path, engine="xlsxwriter")  # Create new workbook for PM
            workbook = writer.book
            worksheet = workbook.add_worksheet(self.sheet_name)  # Add a named worksheet to the workbook
            worksheet = _sheet_setup(worksheet)
            # writer.save()
            return writer, workbook, worksheet

        def _sheet_setup(worksheet):
            month = str(datetime.today().strftime("%B"))
            worksheet.set_page_view()
            worksheet.set_landscape()
            worksheet.set_zoom(60)
            worksheet.hide_gridlines(1)
            worksheet.set_header(
                f"&L&[Picture]&C&14&'Arial,Bold'GHD Monthly Dashboard\nIssue {ISSUE}: ({month})&R&[Picture]",
                {
                    "image_left": GHD_LOGO,
                    "image_right": ST_LOGO,
                }
            )
            worksheet.set_footer("&CPage &P of &N")
            worksheet.set_margins(
                left=MARGINS["left"],
                right=MARGINS["right"],
                top=MARGINS["top"],
                bottom=MARGINS["bottom"],
            )
            worksheet.repeat_rows(0)
            _row_start, _row_finish, _col_start, _col_finish = self.printable_cells
            worksheet.print_area(_row_start, _row_finish, _col_start, _col_finish)
            worksheet.set_paper(8)
            return worksheet

        def _header_format(workbook, sheet):
            header_format = {**BASE_FORMAT, **HEADER_FORMAT}
            header_format = workbook.add_format(header_format)
            for col_num, value in enumerate(HEADERS):
                sheet.write(0, col_num, value, header_format)
                sheet.set_column(col_num, col_num, COL_WIDTH[col_num])

        def _data_validation(sheet, col, val_fmt):
            return sheet.data_validation(*self._data_val_range(col), val_fmt)

        def _format_cells(workbook, sheet, df=pd.DataFrame()):
            _row_start, _row_finish, _col_start, _col_finish = self.protected_cells(df)

            def _get_format(schedule=None, contains_data=True, col=None):
                cell_format = BASE_FORMAT
                if df.index.values[row - 1] in self.new_data["job"]:
                    cell_format = {**cell_format, **NEW_JOB_FORMAT}

                if df.iloc[row - 1, df.columns.get_loc(PM)] in self.new_data["pm"]:
                    cell_format = {**cell_format, **NEW_PM_FORMAT}

                if _st_pn_regex_check(
                        df.iloc[row - 1, df.columns.get_loc(ST_REF_PO)],
                        df.iloc[row - 1, df.columns.get_loc(ST_P_NUM)]
                ):
                    cell_format = {**cell_format, **ST_PNUM_ERROR_FORMAT}

                if schedule:
                    if schedule.lower() == SCHEDULE_D_VAL["source"][0].lower():
                        cell_format = {**cell_format, **ON_TRACK_FORMAT}

                    elif schedule.lower() == SCHEDULE_D_VAL["source"][1].lower():
                        cell_format = {**cell_format, **AT_RISK_FORMAT}

                    elif schedule.lower() == SCHEDULE_D_VAL["source"][2].lower():
                        cell_format = {**cell_format, **BEHIND_SCHEDULE_FORMAT}

                    elif schedule.lower() == SCHEDULE_D_VAL["source"][3].lower():
                        cell_format = {**cell_format, **ON_HOLD_FORMAT}

                if not contains_data:
                    cell_format = {**cell_format, **MANDATORY_INPUT_FORMAT}

                if col:
                    cell_format = {**cell_format, **DATE_FORMAT}

                return cell_format

            if not df.empty:
                for row in range(_row_start, _row_finish):
                    schedule = _check_not_nan(df.iloc[row - 1, df.columns.get_loc(SCH)])
                    base_cell_format = _get_format(schedule=schedule)
                    for col in range(_col_start, _col_finish):
                        cell_format = base_cell_format
                        if col == 0:
                            value = df.index.values[row - 1]
                        else:
                            value = _check_not_nan(df.iloc[row - 1, col - 1])
                            if pm and (col in MANDATORY_COL_IDX) and not value:
                                cell_format = {**cell_format, **_get_format(contains_data=value)}
                            if col in DATE_COLS:
                                cell_format = {**cell_format, **_get_format(col=col)}
                        write_format = workbook.add_format(cell_format)
                        sheet.write(row, col, value, write_format)

        _wr, _wb, _ws = _setup_excel()  # Specify the header format

        if pm:
            _ws.protect()  # Lock all the cells

        _header_format(_wb, _ws)  # Format the header cells
        _data_validation(_ws, PHASE, PHASE_D_VAL)  # Set up data validation
        _data_validation(_ws, SCH, SCHEDULE_D_VAL)
        _data_validation(_ws, ACTION_BY, ACTION_D_VAL)

        self.df.sort_values(by=PM, axis=0, inplace=True)

        if pm:
            _format_cells(
                _wb,
                _ws,
                df=self.df[self.df[PM] == pm]
            )  # Unlock the desired range of editable cells and paste in data
        else:
            _format_cells(_wb, _ws, df=self.df)

        # _sheet_setup(_ws)
        _wr.save()  # Save the workbook

    def _get_output_name(self, path, pm=True):
        if pm:
            path = path / (pm + XLSX)
        else:
            path = path / (self.workbook_name + XLSX)
        return path

    def _data_val_range(self, col):
        _col = self.df.columns.get_loc(col) + 1
        return 1, _col, self.df.shape[0] + 1, _col

    def protected_cells(self, df):
        return [1, df.shape[0] + 1, 0, df.shape[1] + 1]

    @property
    def printable_cells(self):
        modifier = [0, -1, 0, -1, ]
        return [sum(x) for x in zip(self.protected_cells(self.df), modifier)]

    def show_new(self):
        print(f"\nNew project managers:\n")
        print(f"{self.new_data['pm']}")
        print(f"\nNew projects:\n")
        print(f"{self.new_data['job']}")


def _check_not_nan(value):
    if not value:
        return None
    elif str(value) == "nan":
        return None
    elif type(value) == str:
        return value
    else:
        return value


def _st_pn_regex_check(purchase_order_col, project_number_col):
    match1 = ST_PN_REGEX.match(str(purchase_order_col))
    if match1:
        match2 = ST_PN_REGEX.match(str(project_number_col))
        if match2:
            po_nums = [match1.groupdict()["st_pn"]]
            p_nums = [match2.groupdict()["st_pn"]]
            if len(p_nums) == len(po_nums):
                if po_nums == p_nums:
                    return False
                else:
                    return True
            else:
                return True
        else:
            return False
    else:
        return False


class MakeDashboard():
    def __init__(self, prev_dash_path, pm_sheets_path, out_path, bst_path):
        self.prev_dash_path = prev_dash_path
        self.pm_sheets_path = pm_sheets_path
        self.out_path = out_path
        self.bst_path = bst_path
        BSTLoader(
            path=bst_path,
            exclusions=config.exclusions,

        )

    def run(self, client):
        new_dash = Dashboard(client="Sydney Trains", )

        new_dash.load_prev_dashboard(self.prev_dash_path)

        new_dash.load_bst(self.bst_path)

        new_dash.load_pm(self.pm_sheets_path, all_in_path=True)

        new_dash.export(self.out_path)


if __name__ == "__main__":

    #     #TODO: Test code here
    #     # bst_path = Path(r"C:\Users\kschroder-turner\Documents\TEMP\tmp\bst\sydney_water_bst.xlsx")

    #     # output_path = Path(r"C:\Users\kschroder-turner\Documents\TEMP\tmp\bst")

    #     # new_dash = Dashboard(client="Sydney Water", workbook_name="Sydney Water Dashboard")

    #     # new_dash.load_bst(bst_path)

    #     # new_dash.export(output_path)

    #     #TODO: Sydney trains here

    dash_month = "July 2021"

    # parent = Path(r"\\gis010495\c$\Users\kschroder-turner\OneDrive - GHD\Projects\Misc\st_dashboard\data\Monthly
    # Dashboards\July 2020")

    parent = Path(
        r"C:\Users\kschroder-turner\OneDrive - GHD\Projects\Misc\st_dashboard\data\Monthly Dashboards"
    ) / dash_month

    prev_dash_path = parent / "PREV" / "Dashboard.xlsx"

    # output_path = Path(r"\\teams.ghd.com@SSL\DavWWWRoot\operations\SOCSydneyTrainsPanel\Documents\Monthly
    # Dashboards") / dash_month
    output_path = parent  # / "test"
    output_path.mkdir(parents=True, exist_ok=True)

    # output_path = Path(r"C:\Users\kschroder-turner\Documents\TEMP\Monthly Dashboards\November 2019")

    pm_sheets_path = parent / "PM"

    bst_path = parent / "BST" / "Project Detail.xlsx"

    new_dash = Dashboard(client="Sydney Trains", )

    new_dash.load_prev_dashboard(prev_dash_path)

    new_dash.load_bst(bst_path)

    #     new_dash.show_new()
    #
    new_dash.load_pm(pm_sheets_path, all_in_path=True)

    new_dash.export(output_path, pm=False)
