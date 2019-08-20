import os
from pathlib import Path
import pandas as pd
import numpy as np
import xlsxwriter


#Directory constants
MONTH = 'July 2019'
SUB_DIR = 'Job Managers'
DASH = 'Dashboard'
DASHBOARD_DIRECTORY = r"C:\Users\kschroder-turner\Documents\TEMP\Monthly Dashboards"
XLSX = '.xlsx'
TMP_FILE = "~$"


#Dataframe constants
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
]
HEADERS = [JOB_NUM, *COL_ORDER]

DATE_FORMAT = '%d-%m-%Y'
DATETIME_TYPE_STRING = 'datetime64[ns]'

#XLSXWRITER constants
SHEET1_NAME = 'ST Dashboard'
GHD_BLUE = '#006DA3'
WHITE = '#FFFFFF'
HEADER_FORMAT = {
    'bold': False,
    'text_wrap': True,
    'valign': 'vcenter',
    'align': 'center',
    'bg_color': GHD_BLUE,
    'border': 1,
    'border_color': WHITE,
    'font_name': 'arial',
    'font_color': WHITE,
    'font_size': 11,
}
COL_WIDTH = 25
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
LOCKED_FMT = {
    'locked': 0,
}



#TODO: Make .exe version
#TODO: Fix the path and directory bug. Access the sharepoint directly.
#TODO: Need to be able to customise the month based on current month. Annette could possibly just use the script to change the month
#TODO: Copy last month's data into new sheet for current month
#TODO: Comment the code
#TODO: Add to github then create a markdown user guide.
#TODO: Fix dates copying over as numbers


def main():
    writer = pd.ExcelWriter(Path(DASHBOARD_DIRECTORY) / 'OUTPUT.xlsx', engine='xlsxwriter')
    [columns, master_df] = _get_master(MONTH)
    new_df = _get_projects(master_df, columns, MONTH)
    _write_master_sheet(new_df, writer)

def _write_pm_sheets(df):
    _all_pms = df[PM].unique()
    _all_pms.sort()
    _output_dir =  Path(DASHBOARD_DIRECTORY) / Path(MONTH) / Path(SUB_DIR + '_' + MONTH)
    if not _output_dir.exists():
        _output_dir.mkdir()
    for name in _all_pms:
        #Get name of PM
        _name = _output_dir / Path(str(name)+XLSX)
        #Extract PM data from main dataframe
        _df = df[df[PM]==name]
        #Create new workbook for PM
        _wr = pd.ExcelWriter(_name, engine='xlsxwriter')
        _wb = _wr.book
        #Add a named worksheet to the workbook
        _ws = _wb.add_worksheet(SHEET1_NAME)
        #Get the unlocked cell format
        _ul = _wb.add_format(LOCKED_FMT)
        #Specify the header format
        header_format = _wb.add_format(HEADER_FORMAT)
        #Lock all the cells
        _ws.protect()
        #Format the header cells
        _header_format(_ws, header_format)
        #Get the range of editable cells
        _cell_range = _editable_cell_range(_df)
        #Set up data validation
        _data_validation(_ws, _df, PHASE, PHASE_D_VAL)
        _data_validation(_ws, _df, SCH, SCHEDULE_D_VAL)
        #Unlock the desired range of editable cells and paste in data
        _format_cells(_ws, _ul, _cell_range, df=_df)
        #Save the workbook
        _wr.save()
    return

def _get_col_idx(df, col):
    return df.columns.get_loc(col)

def _data_val_range(df, col):
    _col = _get_col_idx(df, col) +1
    return 1, _col, df.shape[0]+1, _col

def _data_validation(sheet, df, col, val_fmt):
    return sheet.data_validation(*_data_val_range(df, col), val_fmt)

def _editable_cell_range(df):
    return [1, df.shape[0]+1, 0, df.shape[1]+1]

def _format_cells(sheet, cell_format, cell_range, df=pd.DataFrame(), value=None):
    _row_start, _row_finish, _col_start, _col_finish = cell_range

    if not value:
        for row in range(_row_start, _row_finish):
            for col in range(_col_start, _col_finish):
                sheet.write(row, col, value, cell_format)
    
    if not df.empty:
        for row in range(_row_start, _row_finish):
            for col in range(_col_start, _col_finish):
                if col == 0:
                    value = df.index.values[row-1]
                else:
                    value = _check_not_nan(df.iloc[row-1, col-1])
                sheet.write(row, col, value, cell_format)
                
    
def _check_not_nan(value):
    if not value:
        return None
    elif str(value) == 'nan':
        return None
    elif type(value) == str:
        return value
    else:
        return value

def _header_format(sheet, header_format):
    for col_num, value in enumerate(HEADERS):
        sheet.write(0, col_num, value, header_format)
        sheet.set_column(col_num, col_num, COL_WIDTH)

def _write_master_sheet(df, writer):
    df.to_excel(writer, sheet_name=SHEET1_NAME, startrow=1, header=False)
    workbook = writer.book
    worksheet = writer.sheets[SHEET1_NAME]
    print(df.head())
    header_format = workbook.add_format(HEADER_FORMAT)
    _write_pm_sheets(df)
    _header_format(worksheet, header_format)
    writer.save()
    return

def _get_master(month):
    '''This function opens the master spreadsheet to extract the header columns.'''
    _dash_dir = Path(DASHBOARD_DIRECTORY)
    _month_dir = _dash_dir / month
    columns = []
    for f_name in os.listdir(_month_dir):
        if XLSX in f_name and DASH in f_name and TMP_FILE not in f_name:
            df = pd.read_excel(_month_dir / f_name, index_col=1)
            df.index.astype(int, copy=False)
    df = df.loc[:, ~df.columns.str.contains('^Unnamed')]
    df.sort_index(inplace=True)
    df.sort_index(axis=1, inplace=True)
    columns = df.columns
    return [columns, df]

def _date_time_handler(df):
    if not df[C_C_DATE].isnull().values.all():
        if df[C_C_DATE].dtype == DATETIME_TYPE_STRING:
            df[C_C_DATE] = df[C_C_DATE].dt.strftime(DATE_FORMAT)
        if df[F_C_DATE].dtype == DATETIME_TYPE_STRING:
            df[F_C_DATE] = df[F_C_DATE].dt.strftime(DATE_FORMAT)
    df.replace(r'NaT', '',regex=True, inplace=True)
    return df

def _get_projects(master, columns, month):
    _dash_dir = Path(DASHBOARD_DIRECTORY)
    _month_dir = _dash_dir / month
    _jm_sheets = _month_dir / SUB_DIR
    for f_name in os.listdir(_jm_sheets):
        file = _jm_sheets / f_name
        if XLSX in f_name and TMP_FILE not in f_name:
            df = pd.read_excel(file, index_col=0)
            df = df[columns]
            df.index.astype(int, copy=False)
            df.sort_index(inplace=True)
            df.sort_index(axis=1, inplace=True)
            # print(df.head())
            df = _date_time_handler(df)
            master.update(df,overwrite=True, errors='ignore')
    return master[COL_ORDER]

if __name__ == "__main__":
    main()