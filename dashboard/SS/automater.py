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


#TODO: Fix the data_path and directory bug. Access the sharepoint directly.
#TODO: Add dropdown menu to the phase column
#TODO: Need to be able to customise the month based on current month. Annette could possibly just use the script to change the month
#TODO: Copy last month's test_data into new sheet for current month
#TODO: Comment the code
#TODO: Add to github then create a markdown user guide.


def main():
    writer = pd.ExcelWriter(Path(DASHBOARD_DIRECTORY) / 'OUTPUT.xlsx', engine='xlsxwriter')
    [columns, master_df] = _get_master(MONTH)
    new_df = _get_projects(master_df, columns, MONTH)
    _output(new_df, writer)

def _pm_sheets(df):
    _all_pms = df[PM].unique()
    _all_pms.sort()
    _output_dir =  Path(DASHBOARD_DIRECTORY) / Path(MONTH) / Path(SUB_DIR + '_' + MONTH)
    if not _output_dir.exists():
        _output_dir.mkdir()
    for name in _all_pms:
        _name = _output_dir / Path(str(name)+XLSX)
        _wr = pd.ExcelWriter(_name, engine='xlsxwriter')
        _df = df[df[PM]==name]
        _df.to_excel(_wr, sheet_name=SHEET1_NAME, startrow=1, header=False)
        _wb = _wr.book
        _ws = _wr.sheets[SHEET1_NAME]
        header_format = _wb.add_format(HEADER_FORMAT)
        _header_format(_ws, header_format)
        _wr.save()
    return

def _header_format(sheet, header_format):
    for col_num, value in enumerate(HEADERS):
        sheet.write(0, col_num, value, header_format)
        sheet.set_column(col_num, col_num, COL_WIDTH)

def _output(df, writer):
    df.to_excel(writer, sheet_name=SHEET1_NAME, startrow=1, header=False)
    workbook = writer.book
    worksheet = writer.sheets[SHEET1_NAME]
    header_format = workbook.add_format(HEADER_FORMAT)
    _pm_sheets(df)
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
            df = _date_time_handler(df)
            master.update(df,overwrite=True, errors='ignore')
    return master[COL_ORDER]

if __name__ == "__main__":
    main()