import os
from pathlib import Path
import pandas as pd
import numpy as np
import xlsxwriter
import datetime
import string
import warnings
import re
from datetime import datetime
import math

#TODO make it be able to read from a URL.


def _cm_to_inch(length):
    return np.divide(length,2.54)

#Directory constants
ISSUE = 4
# MONTH = 'October 2019'
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
ST_REF_PO = 'ST Purchase Order Number'
ST_P_NUM = "ST Project Number"
SCH = 'Schedule'
COMMENTS = 'Comments'
ACTION_BY = 'Action By'
COL_ORDER = [
    ST_REF_PO,
    ST_P_NUM,
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
MANDATORY_COL_IDX = [1, 4, 5, 6, 7, 8, 9, 10, 11, 12,]
BST_COLS = [
    PROJECT,
    PM,
]

ST_PN_REGEX = re.compile("(?:P\.(?P<st_pn>\d+))")

HEADERS = [JOB_NUM, *COL_ORDER]


DATE_COLS = (7,8)
DATE_FORMAT = '%d-%m-%Y'
DATETIME_TYPE_STRING = {'datetime64', "datetime"}

# XLSXWRITER constants
13,	12,	12,	24,	17,	15,	14,	10,	14,	14,	19,	24,	6,

COL_WIDTH = [13, 12, 12, 24, 17, 15, 14, 10, 14, 14, 19, 24, 7, 110,]
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

ON_HOLD_TEXT_COLOUR = '#9c0006'
BEHIND_SCHEDULE_TEXT_COLOUR = '#000000'
AT_RISK_TEXT_COLOUR = '#9c6500'
ON_TRACK_TEXT_COLOUR = '#375623'

ON_HOLD_CELL_FILL = '#ffc7ce'
BEHIND_SCHEDULE_CELL_FILL = '#FF0000'
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
    'bold': True,
    'bg_color': GHD_BLUE,
    'border_color': WHITE,
    'font_color': WHITE,
    'font_size': 11,
}
ON_HOLD_FORMAT = {
    'bg_color': ON_HOLD_CELL_FILL,
    'font_color': ON_HOLD_TEXT_COLOUR,
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
ST_PNUM_ERROR_FORMAT = {
    'italic': True,
    'bold': True,
}
DATE_FORMAT = {
    # 'num_format': 'DD-MM-YYYY'
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
        'On Hold',
    ],
    # 'input_title': 'Select a schedule desciption',
    # 'input_message': 'Please be realistic when selecting a schedule status.',
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
        2125276,
        210921566,
        2127943,
        210921575,
        12515936,
        210921633,
        210921586,
        210921339,
        2128388,
        2127791,
        2126491,
        2128181,
        2128198,
        2127638,
        12517416,
        2127322,
        2126614,
        2127996,
        2126623,
        2127998,
        2128168,
        2220070,
        2127971,
        2127994,
        2128118,
        2128272,
        2126566,
        12518172,
        12511240,
        12511262,
        2128269,
        2128267,
        12526272,
        2316680,
        12510799,
        12540664,
        12510429
    ],
    PM: [
        'Winston Wang',
        'Ruevern Barritt',
        'Michael Hastings',
        'Elena Bullo',
        'Brodie Hayter',
    ]
}

DEFAULT_SHEET = "Dashboard"
DEFAULT_NAME = "Dashboard"

BST_MAPPING = {
    "Project Manager Name": PM,
    "Project Name": PROJECT,
}

class Dashboard():
    num_cols = 10
    PM_SUB_DIR = "Project Manager Sheets"
    BST_RAW_COLS = [
        "Project Manager Name",
        "Project Name",
        TASK_CODE,
    ]
    def __init__(self, client=None, sheet_name=DEFAULT_SHEET, workbook_name=DEFAULT_NAME):
        self.client = client
        self.sheet_name = sheet_name
        self.workbook_name = workbook_name
        self.projects = set()
        self.project_managers = set()
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
            self._df = self.bst._df
            self.projects = self.bst.projects
            self.project_managers = self.bst.project_managers
        else:
            self._load_conflict_handler()
        
    def _load_conflict_handler(self, bst=True, other_df=None):
        if bst:
            #Keep jobs that are in both BST and the Prev Dashboard. = intersection of master with bst.
            #Add new jobs from BST to dashboard. = set difference of bst to master then append to master
            #Remove old jobs from dashboard. = negate the master bst intersection
            intersect_mask = np.in1d(self._df.index.values, self.bst._df.index.values, assume_unique=True)
            append_mask = np.in1d(self.bst._df.index.values, self._df.index.values, assume_unique=True, invert=True)
            #Remove old jobs
            # temp_bst = self.bst._df.loc[intersect_mask]
            self.bst._df.sort_index(inplace=True)
            self._df = self._df[intersect_mask]
            self._df.sort_index(inplace=True)
            
            current_pms = self._df[PM].unique()
            #Append new jobs
            self._df = self._df.append(self.bst._df[append_mask])
            self._df.sort_index(inplace=True)
            self._df[PM] = self.bst._df[PM]
            #Keep track of the new projects and project managers
            new_pms = np.in1d(self.bst._df[PM], current_pms, invert=True)
            self.new_data['job'] = set(self.bst._df.index.values[append_mask])
            self.new_data['pm'] = set(self.bst._df.loc[new_pms, PM])
        else:
            #TODO: This code is outdated, fix to match above

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
        
        self.projects = set(self._df.index.values)
        self.project_managers = set(self._df[PM].unique())

        self._df[[SCH, CUR_STAT, NEXT_ACTION, ACTION_BY]] = ''
    
    def _load_helper(self, df):
        df = self._index_handler(df)
        df = df.loc[:, ~df.columns.str.contains('^Unnamed')]
        df = self._add_missing_col(df)
        df = self._date_time_handler(df)
        df = self.exclude(df)
        return df[COL_ORDER]

    def load_pm(self, path, all_in_path=True,):
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
            self._df.loc[overwrite_mask] = df.loc[overwrite_mask]

    def _check_pm_error(self, df):
        #TODO: Print Error to logging file. So make a logging file as well....
        mask = np.isin(df.index.values, self._df.index.values)
        if len(df[~mask].index) > 0:
            msg0 = f'Project manager {df[PM].unique().astype(str)} might have errors. Please check the PM Sheet.'
            msg1 = f'Project(s) found not in master: {df[~mask].index.values.astype(str)}'
            warnings.warn(msg0, Warning) 
            warnings.warn(msg1, Warning) 
        return df[mask]

    def _index_dup_check(self, df, path):
        len_idx_init = len(df.index)
        len_idx_fin = len(set(df.index))
        if len_idx_init != len_idx_fin:
            msg1 = f'Duplicate index values. Unique indices: {len_idx_init}, Total indicies: {len_idx_fin}. Skipping {path.name}'
            warnings.warn(msg1, Warning) 
            return False
        return True

    def exclude(self, df):
        for key, val in EXCLUSIONS.items():
            if key == JOB_NUM:
                df = df[~df.index.isin(val)]
            else:
                df = df[~df[key].isin(val)]
        return df

    def export(self, path, pm=True, to_excel=True):
        if to_excel:
            path.mkdir(parents=True, exist_ok=True)
            if pm:
                pm_path = path / self.PM_SUB_DIR
                pm_path.mkdir(parents=True, exist_ok=True)
                for pm in self._df[PM].unique():
                    self._export_to_excel(pm_path, pm)
            self._export_to_excel(path, pm=False)

    def _export_to_excel(self, path, pm):
        def _setup_excel():
            nonlocal path
            path = self._get_output_name(path, pm=pm)
            writer = pd.ExcelWriter(path, engine='xlsxwriter')#Create new workbook for PM
            workbook = writer.book 
            worksheet = workbook.add_worksheet(self.sheet_name)#Add a named worksheet to the workbook
            worksheet = _sheet_setup(worksheet)
            # writer.save()
            return writer, workbook, worksheet
        
        def _sheet_setup(worksheet):
            month = str(datetime.today().strftime('%B'))
            worksheet.set_page_view()
            worksheet.set_landscape()
            worksheet.set_zoom(60)
            worksheet.hide_gridlines(1)
            worksheet.set_header(
                f'&L&[Picture]&C&14&"Arial,Bold"GHD Monthly Dashboard\nIssue {ISSUE}: ({month})&R&[Picture]', 
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
                if df.index.values[row-1] in self.new_data['job']:
                    cell_format = {**cell_format, **NEW_JOB_FORMAT}
                
                if df.iloc[row-1, df.columns.get_loc(PM)] in self.new_data['pm']:
                    cell_format = {**cell_format, **NEW_PM_FORMAT}

                if _st_pn_regex_check(df.iloc[row-1, df.columns.get_loc(ST_REF_PO)], df.iloc[row-1, df.columns.get_loc(ST_P_NUM)]):
                    cell_format = {**cell_format, **ST_PNUM_ERROR_FORMAT}

                if schedule:
                    if schedule.lower() ==  SCHEDULE_D_VAL['source'][0].lower():
                        cell_format = {**cell_format, **ON_TRACK_FORMAT}
                    
                    elif schedule.lower() ==  SCHEDULE_D_VAL['source'][1].lower():
                        cell_format = {**cell_format, **AT_RISK_FORMAT}
                    
                    elif schedule.lower() ==  SCHEDULE_D_VAL['source'][2].lower():
                        cell_format = {**cell_format, **BEHIND_SCHEDULE_FORMAT}
                    
                    elif schedule.lower() ==  SCHEDULE_D_VAL['source'][3].lower():
                        cell_format = {**cell_format, **ON_HOLD_FORMAT}
                
                if not contains_data:
                    cell_format = {**cell_format, **MANDATORY_INPUT_FORMAT}

                if col:
                    cell_format = {**cell_format, **DATE_FORMAT}

                return cell_format

            if not df.empty:
                for row in range(_row_start, _row_finish):
                    schedule = _check_not_nan(df.iloc[row-1, df.columns.get_loc(SCH)])
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

        if pm:
            _ws.protect() #Lock all the cells

        _header_format(_wb, _ws) #Format the header cells
        _data_validation(_ws, PHASE, PHASE_D_VAL)#Set up data validation
        _data_validation(_ws, SCH, SCHEDULE_D_VAL) 
        _data_validation(_ws, ACTION_BY, ACTION_D_VAL)

        self._df.sort_values(by=PM, axis=0, inplace=True)

        if pm:
            _format_cells(_wb, _ws, df=self._df[self._df[PM]==pm]) #Unlock the desired range of editable cells and paste in data
        else:
            _format_cells(_wb, _ws, df=self._df)

        # _sheet_setup(_ws)
        _wr.save()#Save the workbook
        
    def _get_output_name(self, path, pm=True):
        if pm:
            path = path / (pm + XLSX)
        else:
            path = path / (self.workbook_name + XLSX)
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

    def _index_handler(self, df):
        df = df[~df.index.duplicated(keep='first')]
        df = df.loc[df.index.dropna()]
        df.index = df.index.astype('uint64')
        df.sort_index(inplace=True)
        return df

    def _date_time_handler(self, df):
        df[[C_C_DATE,F_C_DATE]] = df[[C_C_DATE,F_C_DATE]].astype(str)
        return df

    def _add_missing_col(self, df):
        for col in COL_ORDER:
            if col not in df.columns:
                df[col] = ""
        return df
    def show_new(self):
        print(f'\nNew project managers:\n')
        print(f'{self.new_data["pm"]}')
        print(f'\nNew projects:\n')
        print(f'{self.new_data["job"]}')


class Bst10(Dashboard):
    DEFAULT_SHEET_IDX = 0
    def __init__(self, path_to_bst, sheet_name=DEFAULT_SHEET, dashboard=Dashboard()):
        self.path = Path(path_to_bst)
        self.sheet_name = sheet_name
        self.index_col = 0
        self.dashboard = dashboard
        self._df = pd.DataFrame()
        self.projects = set()

    def load(self, cols_to_keep=Dashboard.BST_RAW_COLS):
        self._df = pd.read_excel(self.path, sheet_name=self.DEFAULT_SHEET_IDX, index_col=self.index_col)
        self._clean()
        return

    def _clean(self, drop_proposals=True):
        # import unicodedata
        # cols = self._df.columns.to_list()
        # cols = [unicodedata.normalize('NFKD', x).encode('ascii','ignore') for x in cols]
        # self._df.columns = [x.decode("UTF-8") for x in cols]
        self._df.sort_values(by='Transaction Date', ascending=False, inplace=True)
        self._df = self._df[~self._df.index.duplicated(keep='first')]
        self._df = self._df[~self._df.index.isna()]
        self._df = self._df[self._df['Project Status'] == 'Active']
        self._df.index = self._df.index.astype(int)
        self._df = self._df[Dashboard.BST_RAW_COLS]
        self._df.rename(columns=BST_MAPPING, inplace=True)
        if drop_proposals:
            mask = self._df[TASK_CODE] != "PP" 
            self._df = self._df[mask]#Removes BST10 props
            self._df.drop([TASK_CODE], inplace=True, axis=1) #Remove the task code column
            self._df = self._df[(self._df.index < 210900000) | (self._df.index > 210999999)] #Removes old MIS Props
        self._df = self._df[[PROJECT, PM]]
        self._df.index.rename(JOB_NUM, inplace=True)
        self._df = self._load_helper(self._df)
        self.projects = set(self._df.index.values)
        self.project_managers = set(self._df[PM].unique())
        return 

def _check_not_nan(value):
    if not value:
        return None
    elif str(value) == 'nan':
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

    def run(self, client):
        new_dash = Dashboard(client="Sydney Trains",)

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

    # parent = Path(r'\\gis010495\c$\Users\kschroder-turner\OneDrive - GHD\Projects\Misc\st_dashboard\data\Monthly Dashboards\July 2020')
    
    parent = Path(r'C:\Users\kschroder-turner\OneDrive - GHD\Projects\Misc\st_dashboard\data\Monthly Dashboards') / dash_month

    prev_dash_path = parent / 'PREV' / 'Dashboard.xlsx'
    
    # output_path = Path(r"\\teams.ghd.com@SSL\DavWWWRoot\operations\SOCSydneyTrainsPanel\Documents\Monthly Dashboards") / dash_month
    output_path = parent #/ 'test'
    output_path.mkdir(parents=True, exist_ok=True)

    # output_path = Path(r"C:\Users\kschroder-turner\Documents\TEMP\Monthly Dashboards\November 2019")

    pm_sheets_path = parent / "PM"

    bst_path = parent / "BST" / "Project Detail.xlsx"

    new_dash = Dashboard(client="Sydney Trains",)

    new_dash.load_prev_dashboard(prev_dash_path)

    new_dash.load_bst(bst_path)

#     new_dash.show_new()
#
    new_dash.load_pm(pm_sheets_path, all_in_path=True)

    new_dash.export(output_path, pm=False)
