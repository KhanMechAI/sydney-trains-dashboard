import os
from pathlib import Path
import pandas as pd
import numpy as np
import xlsxwriter

def _cm_to_inch(length):
    return np.divide(length,2.54)

#Directory constants
ISSUE = 3
MONTH = 'July 2019'
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
    ST_REF_PO,
    ST_DES_MAN,
    PHASE,
    SCH,
    C_C_DATE, 
    F_C_DATE,
    CUR_STAT, 
    NEXT_ACTION, 
]

BST_RAW_COLS = [
    "Project Code",
    "Project Manager Name",
    "Project Name",
    TASK_CODE,
]
HEADERS = [JOB_NUM, *COL_ORDER]



DATE_FORMAT = '%d-%m-%Y'
DATETIME_TYPE_STRING = {'datetime64', "datetime"}

# XLSXWRITER constants
COL_WIDTH = [13, 15, 14, 13, 15, 14, 10, 14, 14, 29, 30, 16, 110,]
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

#TODO: Make .exe version
#TODO: Fix the data_path and directory bug. Access the sharepoint directly.
#TODO: Need to be able to customise the month based on current month. Annette could possibly just use the script to change the month
#TODO: Copy last month's data into new sheet for current month
#TODO: Comment the code
#TODO: Add to github then create a markdown user guide.
#TODO: Fix dates copying over as numbers





def main():
    # writer = pd.ExcelWriter(, engine='xlsxwriter')
    master_file_path = Path(DASHBOARD_DIRECTORY) / MASTER_FNAME
    [columns, master_df] = _get_master(MONTH)
    bst_df = _get_bst(0, sheet=0)
    master_df, new_data = _update_bst(master_df, bst_df)
    master_df = _copy_to_master(master_df, columns, MONTH)
    master_df =_exclude(master_df, EXCLUSIONS)
    _export_sheet(master_file_path, master_df, sheet_name=SHEET1_NAME, new_data=new_data, is_pm=False)
    _export_pm_sheets(master_df)
    return

def _exclude(df, exclusions):
    for key, val in exclusions.items():
        if key == JOB_NUM:
            continue
        df = df[~df[key].isin(val)]
    return df

def _get_bst(fileId, sheet=0):
    fileId = Path(r"C:\Users\kschroder-turner\Documents\TEMP\tmp\BST10 Output.xlsx")
    df = pd.read_excel(fileId, sheet_name=sheet, index_col=0, usecols=BST_RAW_COLS)
    df = df[df[TASK_CODE] != "PP"]
    df.drop([TASK_CODE], inplace=True, axis=1) 
    df = _handle_index(df)
    df.columns = [PROJECT, PM]
    return df

def _handle_index(df):
    df = df[~df.index.duplicated(keep='first')]
    df = df.loc[df.index.dropna()]
    df.index = df.index.astype('uint64')
    df.sort_index(inplace=True)
    return df

def _update_bst(master, bst):
    idx1 = bst.index
    idx2 = master.index
    idx_diff = idx1.difference(idx2)
    bst_pm = set(bst[PM].unique())
    master_pm = set(bst[PM].unique())
    new_data = {
        'pm':bst_pm.difference(master_pm),
        'jobs':idx_diff,
    }
    return master.append(bst.loc[idx_diff]), new_data


def _export_pm_sheets(df):
    _all_pms = df[PM].unique()
    _all_pms.sort()
    _output_dir =  Path(DASHBOARD_DIRECTORY) / Path(MONTH) / Path(SUB_DIR + '_' + MONTH)
    if not _output_dir.exists():
        _output_dir.mkdir()
    for name in _all_pms: 
        _name = _output_dir / Path(str(name)+XLSX) #Generate file data_path and name for PM
        _df = df[df[PM]==name] #Extract PM data from main dataframe
        _export_sheet(_name, _df, sheet_name=SHEET1_NAME)
    return

def _export_sheet(file_path, df, sheet_name, new_data={}, is_pm=True):
    _wr, _wb, _ws = _excel_setup(file_path, sheet_name)#Specify the header format
    _ws.protect() #Lock all the cells
    _header_format(_wb, _ws) #Format the header cells
    _cell_range = _editable_cell_range(df) #Get the range of editable cells
    _data_validation(_ws, df, PHASE, PHASE_D_VAL)#Set up data validation
    _data_validation(_ws, df, SCH, SCHEDULE_D_VAL) 
    _data_validation(_ws, df, ACTION_BY, ACTION_D_VAL)
    _format_cells(_wb, _ws, _cell_range, df=df, is_pm=is_pm, new_data=new_data) #Unlock the desired range of editable cells and paste in data
    _sheet_setup(_ws, df)
    _wr.save()#Save the workbook
    return

def _excel_setup(file_path, sheet_name):
    writer = pd.ExcelWriter(file_path, engine='xlsxwriter')#Create new workbook for PM
    workbook = writer.book 
    worksheet = workbook.add_worksheet(sheet_name)#Add a named worksheet to the workbook
    return writer, workbook, worksheet

def _sheet_setup(worksheet, df):
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
    _row_start, _row_finish, _col_start, _col_finish =_editable_cell_range(df, printable=True)
    worksheet.print_area(_row_start, _row_finish, _col_start, _col_finish)
    return

def _get_col_idx(df, col):
    return df.columns.get_loc(col)

def _data_val_range(df, col):
    _col = _get_col_idx(df, col) + 1
    return 1, _col, df.shape[0] + 1, _col

def _data_validation(sheet, df, col, val_fmt):
    return sheet.data_validation(*_data_val_range(df, col), val_fmt)

def _editable_cell_range(df, printable=False):
    offset = 1
    if printable:
        offset = 0
    return [1, df.shape[0] + offset, 0, df.shape[1] + offset]

def _format_cells(workbook, sheet, cell_range, is_pm=True, df=pd.DataFrame(), new_data={}):
    _row_start, _row_finish, _col_start, _col_finish = cell_range
    
    new_pms = new_data.get('pm',[])
    new_jobs = new_data.get('jobs',[]) 

    def _get_format(schedule=None, contains_data=True):
        
        cell_format = BASE_FORMAT
        
        if df.index.values[row-1] in new_jobs:
            cell_format = {**cell_format, **NEW_JOB_FORMAT}
        
        if df.iloc[row-1, 2] in new_pms:
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
        
        return workbook.add_format(cell_format)

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
                    if is_pm and (col in MANDATORY_COL_IDX) and not value:
                        cell_format = _get_format(contains_data=value)
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

def _header_format(workbook, sheet):
    header_format = {**BASE_FORMAT, **HEADER_FORMAT}
    header_format = workbook.add_format(header_format)
    for col_num, value in enumerate(HEADERS):
        sheet.write(0, col_num, value, header_format)
        sheet.set_column(col_num, col_num, COL_WIDTH[col_num])

def _get_master(month):
    """ This function opens the master spreadsheet to extract the header columns.
    
    Arguments:
        month {string} -- String representation of the folder, usually is a month.
    
    Returns:
        columns {list} -- List of the column names from the master sheet.
        df
    """
    _dash_dir = Path(DASHBOARD_DIRECTORY)
    _month_dir = _dash_dir / month
    columns = []
    for f_name in os.listdir(_month_dir):
        if XLSX in f_name and DASH in f_name and TMP_FILE not in f_name:
            df = pd.read_excel(_month_dir / f_name, index_col=0)
    df = _handle_index(df)
    df = df.loc[:, ~df.columns.str.contains('^Unnamed')]
    df.sort_index(inplace=True)
    df.sort_index(axis=1, inplace=True)
    df = _date_time_handler(df)
    df = _add_missing_col(df)
    columns = df.columns
    return [columns, df]

def _add_missing_col(df):
    for col in COL_ORDER:
        if col not in df.columns:
            df[col] = ""
    return df

def _date_time_handler(df):
    df[F_C_DATE].apply(lambda x: _date_time_converter(x))
    df[C_C_DATE].apply(lambda x: _date_time_converter(x))
    df.replace(r'(NaT|NaN|nan|nat)', '',regex=True, inplace=True)
    return df

def _date_time_converter(elem):
    if type(elem).__name__ in DATETIME_TYPE_STRING:
        return elem.strftime(DATE_FORMAT)
    else:
        return elem

def _copy_to_master(master, columns, month):
    _dash_dir = Path(DASHBOARD_DIRECTORY)
    _month_dir = _dash_dir / month
    _jm_sheets = _month_dir / SUB_DIR
    for f_name in os.listdir(_jm_sheets):
        file = _jm_sheets / f_name
        if XLSX in f_name and TMP_FILE not in f_name:
            df = pd.read_excel(file, index_col=0)
            df = df.iloc[:, :10]
            df = _handle_index(df)
            df.sort_index(axis=1, inplace=True)
            df = _date_time_handler(df)
            master.update(df,overwrite=True, errors='ignore')
    return master[COL_ORDER]

if __name__ == "__main__":
    main()