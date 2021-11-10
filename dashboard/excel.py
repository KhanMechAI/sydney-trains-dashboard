import datetime
from datetime import datetime
from typing import Dict, List, Tuple, Union

import pandas as pd


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

