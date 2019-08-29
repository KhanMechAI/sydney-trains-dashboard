import os
from pathlib import Path
import pandas as pd
import numpy as np
import xlsxwriter
import datetime

#TODO make it be able to read from a URL.


def _cm_to_inch(length):
    return np.divide(length,2.54)

#Directory constants
ISSUE = 3
MONTH = 'August 2019'
SUB_DIR = 'Job Managers'
DASH = 'Dashboard'
DASHBOARD_DIRECTORY = r"C:\Users\kschroder-turner\Documents\TEMP\Monthly Dashboards"
XLSX = '.xlsx'
TMP_FILE = "~$"
GHD_LOGO = r'C:\Users\kschroder-turner\Documents\TEMP\tmp\logo\ghd_logo.png'
ST_LOGO = r'C:\Users\kschroder-turner\Documents\TEMP\tmp\logo\st_logo.png'
MASTER_FNAME = 'OUTPUT.xlsx'

#Dataframe constants
TASK_CODE = "Task Code" #BST constant
JOB_NUM = 'GHD Job Number'
C_C_DATE = "Contractual Completion Date"
CUR_STAT = 'Current Status'
F_C_DATE = "Forecast Completion Date"
PM = "GHD Project Manager"
NEXT_ACTION = 'Next Actions'
PHASE = 'Phase'
PROJECT = "Project Name"
ST_DES_MAN = 'ST Design Manager'
ST_REF_PO = 'ST Reference No. / Purchase Order Number'
SCH = 'Schedule'
COMMENTS = 'Comments'
ACTION_BY = 'Action By'
COL_ORDER = [
    ST_REF_PO,
    PROJECT,
    PM,
    ST_DES_MAN,
    PHASE,
    SCH,
    C_C_DATE, 
    F_C_DATE,
    CUR_STAT, 
    NEXT_ACTION,
    ACTION_BY,
    COMMENTS,
]
# MANDATORY_COL_IDX = [0, 3, 4, 5, 6, 7, 8, 9, 10,]
MANDATORY_COL_IDX = [1, 4, 5, 6, 7, 8, 9, 10, 11,]
BST_COLS = [
    PROJECT,
    PM,
]


HEADERS = [JOB_NUM, *COL_ORDER]


DATE_COLS = (7,8)
DATE_FORMAT = '%d-%m-%Y'
DATETIME_TYPE_STRING = {'datetime64', "datetime"}

# XLSXWRITER constants
COL_WIDTH = [13, 15, 14, 13, 15, 14, 10, 14, 14, 29, 29, 16, 110,]
MARGINS = {
    'left':_cm_to_inch(0.6),
    'right':_cm_to_inch(0.6),
    'top':_cm_to_inch(3),
    'bottom':_cm_to_inch(1.9),
}
SHEET1_NAME = 'ST Dashboard'
## Colours
GHD_BLUE = '#006DA3'
WHITE = '#FFFFFF'
BEHIND_SCHEDULE_TEXT_COLOUR = '#9c0006'
AT_RISK_TEXT_COLOUR = '#9c6500'
ON_TRACK_TEXT_COLOUR = '#375623'
BEHIND_SCHEDULE_CELL_FILL = '#ffc7ce'
AT_RISK_CELL_FILL = '#ffeb9c'
ON_TRACK_CELL_FILL = '#c6efce'
MANDATORY_INPUT_CELL_FILL = '#ff6d4b'
## Formats
BASE_FORMAT = {
    'bold': False,
    'text_wrap': True,
    'valign': 'vcenter',
    'align': 'center',
    'border': 1,
    'font_name': 'arial',
    'font_size': 10,
    'locked': 0,
}
HEADER_FORMAT = {
    'bg_color': GHD_BLUE,
    'border_color': WHITE,
    'font_color': WHITE,
    'font_size': 11,
}
BEHIND_SCHEDULE_FORMAT = {
    'bg_color': BEHIND_SCHEDULE_CELL_FILL,
    'font_color': BEHIND_SCHEDULE_TEXT_COLOUR,
}
AT_RISK_FORMAT = {
    'bg_color': AT_RISK_CELL_FILL,
    'font_color': AT_RISK_TEXT_COLOUR,
}
ON_TRACK_FORMAT = {
    'bg_color': ON_TRACK_CELL_FILL,
    'font_color': ON_TRACK_TEXT_COLOUR,
}
MANDATORY_INPUT_FORMAT = {
    'bg_color': MANDATORY_INPUT_CELL_FILL,
}
NEW_JOB_FORMAT = {
    'bold': True,
    # 'border': 12,
}
NEW_PM_FORMAT = {
    'bold': True,
    # 'border': 13,
}
DATE_FORMAT = {
    'num_format': 'DD-MM-YYYY'
}
## Data valdation
PHASE_D_VAL = {
    'validate': 'list',
    'source': [
        'Proposal',
        'Condition Assessment',
        'Preliminary Investigation',
        'Options Assessment',
        'Concept Design',
        'Detailed Design',        
        'Construction Support',
        'Approve for Construction',
        'Construction Phase Services',
    ],
    'input_title': 'Select a Project Phase',
    'input_message': 'Select a project phase from the list.',
}
SCHEDULE_D_VAL = {
    'validate': 'list',
    'source': [
        'On Track',
        'At risk of being delayed',
        'Behind Schedule',
    ],
    'input_title': 'Select a schedule desciption',
    'input_message': 'Please be realistic when selecting a schedule status. Risks and issues can\'t be mitigated or resolved unless they\'re communicated.',
}
ACTION_D_VAL = {
    'validate': 'list',
    'source': [
        'GHD',
        'ST',
    ],
    'input_title': 'Select A Stakeholder',
    'input_message': 'Select either GHD or ST from the drop down',
}
## Exclusions
EXCLUSIONS = {
    JOB_NUM: [
        2127653,
    ],
    PM: [
        'Winston Wang',
        'Ruevern Barritt',
    ]
}

#Dataframe constants
TASK_CODE = "Task Code" #BST constant
JOB_NUM = 'GHD Job Number'
C_C_DATE = "Contractual Completion Date"
CUR_STAT = 'Current Status'
F_C_DATE = "Forecast Completion Date"
PM = "GHD Project Manager"
NEXT_ACTION = 'Next Actions'
PHASE = 'Phase'
PROJECT = "Project Name"
ST_DES_MAN = 'ST Design Manager'
ST_REF_PO = 'ST Reference No. / Purchase Order Number'
SCH = 'Schedule'
COMMENTS = 'Comments'
ACTION_BY = 'Action By'

DEFAULT_SHEET = "Sheet1"
DEFAULT_NAME = "Dashboard"

class Dashboard():
    num_cols = 10
    PM_SUB_DIR = "Project Manager Sheets"
    BST_RAW_COLS = [
        "Project Code",
        "Project Manager Name",
        "Project Name",
        TASK_CODE,
    ]
    def __init__(self, client=None, sheet_name=DEFAULT_SHEET, workbook_name=DEFAULT_NAME):
        self.client = client
        self.sheet_name = sheet_name
        self.workbook_name = workbook_name
        self._projects = set()
        self._project_managers = set()
        self._df = pd.DataFrame()
        self._new_data = {}
        self._editable_cells = []
        self.new_data = {
            'pm':set(),
            'job':set(),
            }

    def load_bst(self, path_to_bst):
        self.bst = Bst10(path_to_bst, dashboard=self)
        self.bst.load()
        if self._df.empty:
            self.projects = self.bst.projects
            self.project_managers = self.bst.project_managers
            self._df = self.bst.df
        else:
            self._load_conflict_handler()
    
    def _load_conflict_handler(self, bst=True, other_df=None):
        if bst:
            proj_to_drop = self.projects.difference(self.bst.projects)
            proj_to_append = self.bst.projects.difference(self.projects)
            self._df.drop(labels=proj_to_drop)
            self._df.append(self.bst.df.loc[proj_to_append])
            self.new_data['job'] = proj_to_append
            self.new_data['pm'] = set(self._df.loc[proj_to_append, PM].unique())
        else:
            if not other_df:
                raise ValueError("If bst=False, other_df must be specified")
            other_df_projects = set(other_df.index.values)
            proj_to_keep = self.projects.intersection(other_df_projects)
            other_df = other_df.loc[proj_to_keep]
            non_bst_cols = [x for x in COL_ORDER if x not in BST_COLS]
            self._df.loc[proj_to_keep, non_bst_cols] = other_df[non_bst_cols]



    def load_prev_dashboard(self, path_to_dashboard):
        df = pd.read_excel(path_to_dashboard, index_col=0)
        df = self._load_helper(df)
        if self._df.empty:
            self._df = df
        else:
            self._load_conflict_handler(bst=False, other_df=df)

    
    def _load_helper(self, df):
        df = df.loc[df.index.dropna()]
        self._index_handler(df)
        df = df.loc[:, ~df.columns.str.contains('^Unnamed')]
        df.sort_index(inplace=True)
        df.sort_index(axis=1, inplace=True)
        self._add_missing_col(df)
        self._date_time_handler(df)
        return df[COL_ORDER]

    def load_pm(self, path, all_in_path=True,):
        if all_in_path:
            for f_name in os.listdir(path):
                file = path / f_name
                if XLSX in f_name and TMP_FILE not in f_name:
                    self._load_pm(file)
        else:
            self._load_pm(file)
    
    def exclude(self, exclusions):
        for key, val in exclusions.items():
            if key == JOB_NUM:
                continue
            self._df = self._df[~self._df[key].isin(val)]

    def export(self, path, pm=True, to_excel=True):
        if to_excel:
            path.mkdir(parents=True, exist_ok=True)
            if pm:
                pm_path = path / self.PM_SUB_DIR
                pm_path.mkdir(parents=True, exist_ok=True)
                for pm in self.project_managers:
                    self._export_to_excel(pm_path, pm)
            self._export_to_excel(path, pm=False)

    def _export_to_excel(self, path, pm):
        def _setup_excel():
            nonlocal path
            path = self._get_output_name(path, pm=pm)
            writer = pd.ExcelWriter(path, engine='xlsxwriter')#Create new workbook for PM
            workbook = writer.book 
            worksheet = workbook.add_worksheet(self.sheet_name)#Add a named worksheet to the workbook
            return writer, workbook, worksheet
        
        def _sheet_setup(worksheet):
            worksheet.set_paper(8)
            worksheet.set_page_view()
            worksheet.set_landscape()
            worksheet.set_zoom(60)
            worksheet.hide_gridlines(1)
            worksheet.set_header(
                f'&L&[Picture]&C&14&"Arial,Bold"GHD Monthly Dashboard\nIssue {ISSUE}: ({MONTH})&R&[Picture]', 
                {
                    'image_left': GHD_LOGO,
                    'image_right': ST_LOGO,
                    }
            )
            worksheet.set_footer('&CPage &P of &N')
            worksheet.set_margins(
                left=MARGINS['left'],
                right=MARGINS['right'],
                top=MARGINS['top'],
                bottom=MARGINS['bottom'],
            )
            worksheet.repeat_rows(0)
            _row_start, _row_finish, _col_start, _col_finish = self.printable_cells
            worksheet.print_area(_row_start, _row_finish, _col_start, _col_finish)
            return

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
                if df.index.values[row-1] in self.new_data['job']:
                    cell_format = {**cell_format, **NEW_JOB_FORMAT}
                
                if df.iloc[row-1, 2] in self.new_data['pm']:
                    cell_format = {**cell_format, **NEW_PM_FORMAT}

                if schedule:
                    if schedule.lower() ==  SCHEDULE_D_VAL['source'][0].lower():
                        cell_format = {**cell_format, **ON_TRACK_FORMAT}
                    
                    elif schedule.lower() ==  SCHEDULE_D_VAL['source'][1].lower():
                        cell_format = {**cell_format, **AT_RISK_FORMAT}
                    
                    elif schedule.lower() ==  SCHEDULE_D_VAL['source'][2].lower():
                        cell_format = {**cell_format, **BEHIND_SCHEDULE_FORMAT}
                
                if not contains_data:
                    cell_format = {**cell_format, **MANDATORY_INPUT_FORMAT}

                if col:
                    cell_format = {**cell_format, **DATE_FORMAT}

                return cell_format

            if not df.empty:
                for row in range(_row_start, _row_finish):
                    schedule = _check_not_nan(df.iloc[row-1, 5])
                    base_cell_format = _get_format(schedule=schedule)
                    for col in range(_col_start, _col_finish):
                        cell_format = base_cell_format
                        if col == 0:
                            value = df.index.values[row-1]
                        else:
                            value = _check_not_nan(df.iloc[row-1, col-1])
                            if pm and (col in MANDATORY_COL_IDX) and not value:
                                cell_format = {**cell_format, **_get_format(contains_data=value)}
                            if col in DATE_COLS:
                                cell_format = {**cell_format, **_get_format(col=col)}
                        write_format = workbook.add_format(cell_format)
                        sheet.write(row, col, value, write_format)

        _wr, _wb, _ws = _setup_excel()#Specify the header format
        _ws.protect() #Lock all the cells
        _header_format(_wb, _ws) #Format the header cells
        _data_validation(_ws, PHASE, PHASE_D_VAL)#Set up data validation
        _data_validation(_ws, SCH, SCHEDULE_D_VAL) 
        _data_validation(_ws, ACTION_BY, ACTION_D_VAL)
        if pm:
            if "Gordon" in pm:
                suck_eggs = True
            _format_cells(_wb, _ws, df=self._df[self._df[PM]==pm]) #Unlock the desired range of editable cells and paste in data
        else:
            _format_cells(_wb, _ws, df=self._df)
        _sheet_setup(_ws)
        _wr.save()#Save the workbook
        
    def _get_output_name(self, path, pm=True):
        trailingExtension = self.workbook_name + XLSX
        if pm:
            path = path / (pm + ' ' + trailingExtension)
        else:
            path = path / trailingExtension
        return path

    def _data_val_range(self, col):
        _col = self._df.columns.get_loc(col) + 1
        return 1, _col, self._df.shape[0] + 1, _col
    
    
    def protected_cells(self, df):
        return [1, df.shape[0] + 1, 0, df.shape[1] + 1]

    @property
    def printable_cells(self): 
        modifier = [0, -1, 0, -1,]
        return [sum(x) for x in zip(self.protected_cells(self._df), modifier)]

    def _load_pm(self, path):
        df = pd.read_excel(path, index_col=0)
        df = self._load_helper(df)
        self._df.update(df,overwrite=True, errors='ignore')

    def _index_handler(self, df):
        df = df[~df.index.duplicated(keep='first')]
        df = df.loc[df.index.dropna()]
        df.index = df.index.astype('uint64')
        df.sort_index(inplace=True)
        return df

    def _date_time_handler(self, df):
        df[F_C_DATE].to_string(na_rep='')
        df[C_C_DATE].to_string(na_rep='')
        return df

    def _add_missing_col(self, df):
        for col in COL_ORDER:
            if col not in df.columns:
                df[col] = ""
        return df

    @property
    def project_managers(self):
        if self._df.empty:
            return set()
        else:
            return set(self._df[PM].unique())
    
    @property
    def projects(self):
        if self._df.empty:
            return set()
        else:
            return set(self._df.index.values)
    # @property
    # def new_data(self):
    #     return {
    #             'pm':self.bst.project_managers.difference(self.project_managers),
    #             'job':self.bst.projects.difference(self.projects),
    #         }

class Bst10(Dashboard):
    DEFAULT_SHEET_IDX = 0
    def __init__(self, path_to_bst, sheet_name=DEFAULT_SHEET, dashboard=Dashboard()):
        self.path = Path(path_to_bst)
        self.sheet_name = sheet_name
        self.index_col = 0
        self.dashboard = dashboard
        self._df = pd.DataFrame()

    def load(self, cols_to_keep=Dashboard.BST_RAW_COLS):
        self.df = pd.read_excel(self.path, sheet_name=self.DEFAULT_SHEET_IDX, index_col=self.index_col, usecols=cols_to_keep)
        self._clean()
        self.df = self._load_helper(self.df)
        return

    def _clean(self, drop_proposals=True):
        if drop_proposals:
            self.df = self.df[self.df[TASK_CODE] != "PP"]
            self.df.drop([TASK_CODE], inplace=True, axis=1) 
        self.df.columns = [PROJECT, PM]
        self.df = self.dashboard._load_helper(self.df)
        return 

    @property
    def df(self):
        return self._df

    @df.setter
    def df(self, value):
        self._df = value

def _check_not_nan(value):
    if not value:
        return None
    elif str(value) == 'nan':
        return None
    elif type(value) == str:
        return value
    else:
        return value

if __name__ == "__main__":

    prev_dash_path = Path(DASHBOARD_DIRECTORY) / "July 2019" / "July 19 Dashboard.xlsx"

    pm_sheets_path = Path(r"C:\Users\kschroder-turner\Documents\TEMP\Monthly Dashboards\August 2019\Job Managers")

    bst_path = Path(r"C:\Users\kschroder-turner\Documents\TEMP\tmp\BST10 Output.xlsx")

    output_path = Path(DASHBOARD_DIRECTORY) / MONTH

    new_dash = Dashboard(client="Sydney Trains",)

    new_dash.load_prev_dashboard(prev_dash_path)

    new_dash.load_bst(bst_path)

    new_dash.load_pm(pm_sheets_path)

    new_dash.export(output_path)