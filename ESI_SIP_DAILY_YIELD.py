import xlrd
import pyodbc
from openpyxl import Workbook
import pandas as pd
from xlutils.copy import copy
from xlwt import Workbook, XFStyle, Borders
from openpyxl.styles import Border, Side, Alignment, Font
from datetime import datetime
import pytz
import base64

Vietnam_time = pytz.timezone('Asia/Ho_Chi_Minh')
datetime_str = str(datetime.now(Vietnam_time))
year = datetime_str.split('-')[0]
month = datetime_str.split('-')[1]
date = datetime_str.split('-')[2][:2]
Curr_Date = year+month+date

def get_data_group(cursor, device_no, INACTIVE, Cur_Date):
    get_group_data = f"EXEC [GetGroupStationData] @Device_Type_No = '{device_no}', @Status = '{INACTIVE}', @CurrentDate = '{Cur_Date}'"
    cursor.execute(get_group_data)
    group_data = cursor.fetchall()
    return group_data

def get_hitter(cursor, device_no, cur_date, cus_no):
    get_hitter_date = f"EXEC [Get_Hitter_Assy_SP] @DEVICE_TYPE_NO = '{device_no}', @CURRENT_DATE = '{cur_date}', @CUS_NO = '{cus_no}'"
    cursor.execute(get_hitter_date)
    hitter_data = cursor.fetchall()
    return hitter_data

def Get_Yield(Yield):
    Yield = str(Yield).split('.')[0] + '.' + str(Yield).split('.')[1][:2] + '%'
    if int(Yield.split('.')[0]) >= 100:
        Yield = '100.00%'
        return Yield
    return Yield

# Define a dictionary to store the data
def generate_report_daily(cursor, device_no, INACTIVE, Cur_Date):
    data_INACTIVE = get_data_group(cursor, device_no, INACTIVE, Cur_Date)
    if data_INACTIVE == []:
        return 0

    data_dict = {}
    # print("INACTIVE")
    # for index in data_INACTIVE:
    #     print(index)

    for index in data_INACTIVE:
        data_dict[index[0]] = {'In': index[1], 'Out': index[2], 'Yield': Get_Yield(index[3])}

    # Load the workbook and select the first sheet and define list of the keys
    if 'QM' not in device_no:   #For ESI
        rb = xlrd.open_workbook(r"C:\Workplace\Task\Support_Assy\Auto_Mail_Yield\sample_format\ESI_SIP_Sample_input.xls", formatting_info=True)
        keys = ['SUB/L', 'SMT1', 'MOLD1', 'SMT2', 'MOLD2', 'SMT3', 'LASER', 'PKG Saw', 'SPUTTER1', 'SPUTTER2', 'DMZ &FVI', 'SLT0', 'SLT1', 'SLT2', 'SLT3', 'AVI/TNR']
    else:                       #For QORVO
        rb = xlrd.open_workbook(r"C:\Workplace\Task\Support_Assy\Auto_Mail_Yield\sample_format\QORVO_Sample_input.xls", formatting_info=True)
        keys = ['2DSM', 'TOP SMT', 'TOP MOLD', 'BTM SMT', 'BTM MOLD', 'LASER', 'SMT Reball', 'PKG Saw', 'SPUTTER1', 'DMZ &FVI', 'SLT0', 'SLT1', 'SLT2', 'SLT3', 'AVI/TNR']

    wb = copy(rb)
    sheet = wb.get_sheet(0)

    # Define the border style
    borders = Borders()
    borders.left = Borders.THIN
    borders.right = Borders.THIN
    borders.top = Borders.THIN
    borders.bottom = Borders.THIN
    style = XFStyle()
    style.borders = borders

    sheet.write(1, 0, '2277', style)
    sheet.write(1, 1, device_no, style)
    
    # Initialize the overall In and Out
    Overall_In = 0
    Overall_Out = 0

    # Modify cells from column D to H and row 4 to 17
    for row in range(3, len(keys) + 3):  # Rows 4 to 17 (0-indexed)
        key = keys[row - 3]
        if key in data_dict:
            sheet.write(row, 8, data_dict[key]['Yield'], style)
            sheet.write(row, 9, data_dict[key]['In'], style)
            sheet.write(row, 10, data_dict[key]['Out'], style)
            Overall_In += data_dict[key]['In']
            Overall_Out += data_dict[key]['Out']

    # Calculate the overall yield
    Overall_Yield = round((Overall_Out/Overall_In)*100,2)
    if int(Overall_Yield) >= 100:
        Overall_Yield = '100.00%'
    else:
        Overall_Yield = str(Overall_Yield) + '%'

    # Write the overall yield, in and out
    sheet.write(len(keys) + 4, 8, Overall_Yield, style)
    sheet.write(len(keys) + 4, 9, Overall_In, style)
    sheet.write(len(keys) + 4, 10, Overall_Out, style)
    fileName = f'{device_no}_{Cur_Date}_IO_DAILY_YIELD.xls'
    # Save the workbook
    wb.save(fileName)
    print(f"Exported -> {fileName}")
    return fileName

def generate_data_yield_summary(cursor, device_no, INACTIVE, Cur_Date, cus_no):
    data_dict = {}
    data_dict_hitter = {}
    data_INACTIVE = get_data_group(cursor, device_no, INACTIVE, Cur_Date)
    if data_INACTIVE == []:
        return 0
    data_Hitter = get_hitter(cursor, device_no, Cur_Date, cus_no )
    for index in data_INACTIVE:
        failQty = index[1] - index[2]
        hit_type = index[0]
        hitter_info = {'In': int(index[1]), 'Fail' : int(failQty),'Yield': Get_Yield(index[3])}    
    # Append the hitter info to a list under the hit type key
        if hit_type not in data_INACTIVE:
            data_dict[hit_type] = hitter_info
        else:
            data_dict[hit_type].append(hitter_info)

    # for index in data_Hitter:
    flag = 0
    for ele in data_Hitter:
        station = ele[-1]
        for index in data_INACTIVE:
            if index[0] == station:
                flag = 1
        if flag != 1:
            continue
        failDefect = int(ele[-2])
        failStation = int(data_dict[station]['Fail'])
        if failStation == 0:
            rateDefect = '100.00%'
        else:
            rateDefect = round((failDefect/failStation)*100,2)
            rateDefect = str(rateDefect) + '%'
        dat_hitter = {
            'Hitter' : {
                'Des' : ele[2],
                'failQty' : int(ele[-2]),
                'Rate' : rateDefect}}
        if station not in data_dict_hitter:
            data_dict_hitter[station] = []
        data_dict_hitter[station].append(dat_hitter)
            
    # print(data_dict_hitter)
    if 'QM' not in device_no:
        stations = ['SUB/L','SMT1', 'Mold1', 'SMT2','Mold2', 'SMT3', 'LASER', 'PKG Saw', 'SPUTTER1', 'SPUTTER2', 'DMZ &FVI', 'SLT0', 'SLT1', 'SLT2', 'SLT3', 'AVI/TNR']
    else:
        stations = ['2DSM', 'TOP SMT', 'TOP MOLD', 'BTM SMT', 'BTM MOLD', 'LASER', 'SMT Reball', 'PKG Saw', 'SPUTTER1', 'DMZ &FVI', 'SLT0', 'SLT1', 'SLT2', 'SLT3', 'AVI/TNR']
    for station in stations:
        if station not in data_dict_hitter:
            continue
        data_dict[station]['Hitter'] = data_dict_hitter[station]
    return data_dict

def all_data_build(data, device_type):
    if 'QM' not in device_type:
        data_all = {
        'FOL': {key: "" for key in ['SUB/L', 'SMT1', 'Mold1', 'SMT2']},
        'EOL': {key: "" for key in ['Mold2', 'SMT3', 'LASER', 'PKG Saw', 'SPUTTER1', 'SPUTTER2', 'DMZ &FVI']},
        'TEST': {key: "" for key in ['SLT0', 'SLT1', 'SLT2', 'SLT3', 'AVI/TNR']}
    }
    else:
        data_all = {
        'FOL': {key: "" for key in ['2DSM', 'TOP SMT', 'TOP MOLD', 'BTM SMT']},
        'EOL': {key: "" for key in ['BTM MOLD', 'LASER', 'SMT Reball', 'PKG Saw', 'SPUTTER1', 'DMZ &FVI']},
        'TEST': {key: "" for key in ['SLT0', 'SLT1', 'SLT2', 'SLT3', 'AVI/TNR']}
    }

    for key in data:
        if key in data_all['FOL']:
            data_all['FOL'][key] = data[key]
        elif key in data_all['EOL']:
            data_all['EOL'][key] = data[key]
        elif key in data_all['TEST']:
            data_all['TEST'][key] = data[key]

    return data_all

def generate_yield_hitter_report(data_all, device_no, Cur_Date, cus_no):
    # Tạo workbook và active worksheet
    from openpyxl import Workbook
    wb = Workbook()
    ws = wb.active

    thin_border = Border(left=Side(style='thin'), 
                        right=Side(style='thin'), 
                        top=Side(style='thin'), 
                        bottom=Side(style='thin'))

    # Set the title of the worksheet
    ws.title = "Yield & Hitters Review"

    # Merge cells for the main title and set its value
    ws.merge_cells('A2:D3')
    main_title = ws['A2']
    main_title.value = "Yield & Hitters Review"
    main_title.alignment = Alignment(horizontal='left')
    main_title.font = Font(bold=True)

    ws.merge_cells('I1:J2')
    second_main_title = ws['I1']
    second_main_title.value = f"M6 / Z6 / 050 - {cus_no}"
    second_main_title.alignment = Alignment(horizontal='left')
    second_main_title.font = Font(bold=True)

    A5cell = ws['A5']
    A5cell.value = "Yield Review Table"
    A5cell.alignment = Alignment(horizontal='left')
    A5cell.font = Font(bold=True)

    ws.merge_cells('B5:E5')
    B5cell_title = ws['B5']
    B5cell_title.value = f"{device_no} / {Cur_Date}"
    B5cell_title.alignment = Alignment(horizontal='left')
    B5cell_title.font = Font(bold=True)

    ws.merge_cells('F5:J5')
    B5cell_title = ws['F5']
    B5cell_title.value = "24Hrs Top5 Hitters"
    B5cell_title.alignment = Alignment(horizontal='left')
    B5cell_title.font = Font(bold=True)

    # Add column headers
    headers = ['Dept./Group', 'Opr', 'Input', 'Fail Q\'ty', 'Engineering Yield', 'Hitters', 'Q\'ty', 'Defect Rate', 'Root Cause', 'Action']
    ws.append(headers)

    # Apply formatting to header row
    for cell in ws[6]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal='center')

    # Set column widths (optional, adjust as needed)
    column_widths = [18, 12, 10, 10, 18, 30, 10, 12, 12, 12]
    for i, column_width in enumerate(column_widths, start=1):
        ws.column_dimensions[chr(64+i)].width = column_width

    # Initialize current_row
    current_row = 7

    for group, stations in data_all.items():
        group_start_row = current_row
        ws.cell(row=current_row, column=1, value=group)
        for station, data in stations.items():
            station_start_row = current_row
            ws.cell(row=current_row, column=2, value=station)
            if data == "":
                ws.cell(row=current_row, column=3, value=0)
                ws.cell(row=current_row, column=4, value=0)
                ws.cell(row=current_row, column=5, value="00.00%")
                ws.cell(row=current_row, column=7, value=0)
                ws.cell(row=current_row, column=8, value="00.00%")
                current_row += 1
                continue
            detail_column = 3
            for _, detail in data.items():
                data_start_row = current_row
                if len(str(detail)) < 10: 
                    ws.cell(row=current_row, column=detail_column, value=f"{detail}")
                    detail_column += 1
                    if detail_column == 6:
                        ws.cell(row=current_row, column=7, value=0)
                        ws.cell(row=current_row, column=8, value="00.00%")
                else:
                    for hitter_no in detail:
                        detail_start_row = current_row
                        for hitter, hitter_info in hitter_no.items():
                            hitter_column = 6
                            for key, value in hitter_info.items():
                                ws.cell(row=current_row, column=hitter_column, value=f'{value}')
                                hitter_column += 1
                        if hitter_no != detail[-1]:
                            current_row += 1
            
            ws.merge_cells(start_row=data_start_row, start_column=2, end_row=current_row, end_column=2)
            ws.merge_cells(start_row=data_start_row, start_column=3, end_row=current_row, end_column=3)
            ws.merge_cells(start_row=data_start_row, start_column=4, end_row=current_row, end_column=4)
            ws.merge_cells(start_row=data_start_row, start_column=5, end_row=current_row, end_column=5)
            current_row += 1
        ws.merge_cells(start_row=group_start_row, start_column=1, end_row=current_row-1, end_column=1)

    for row in ws.iter_rows(min_row=5, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
        for cell in row:
            cell.border = thin_border    

    # Save the workbook
    fileName = f"{device_no}_{Cur_Date}_Yield_Hitter_Summary.xls"
    wb.save(fileName)
    print(f"Exported -> {fileName}")
    return fileName

def convert_file_to_base64(file_path):
    with open(file_path, "rb") as file:
        encoded_string = base64.b64encode(file.read())
    return encoded_string.decode('utf-8')

def sending_email(list_attached):
    toList = ['Hiep.Letien@amkor.com','Tuan.Vuongmanh@amkor.com']
    ccList = ['Hoan.Nguyenvan@amkor.com']
    dictionary_email = {
        "sender": "testit@amkor.com",
        "subject": "test_email",
        "body": "<h1>This is a test email</h1>",
        "toMailList": toList,
        "ccMailList": ccList,
        "bccMailList": [""],
        "attachmentList": list_attached
    }
    request_API(dictionary_email)

def request_API(payload):
    import requests
    import json
    headers = {'Content-Type': 'application/json'}
    # Send the files to the API
    response = requests.post("http://10.201.54.56:5067/Common/Send_Email", data=json.dumps(payload), headers=headers)
    print(response.text)

def main():
    cnxn = pyodbc.connect("DRIVER={ODBC Driver 17 for SQL Server};SERVER=10.201.21.84,50150;DATABASE=MCSDB;UID=cimitar2;PWD=TFAtest1!2!")
    cursor = cnxn.cursor()
    device_no_list = ['639-18807', '639-18808', 'QM76300', 'QM76309', 'QM76095']
    device_no_1 = ['QM76300']
    INACTIVE = 'INACTIVE'
    ACTIVE = 'OTHERSTATUS'
    cus_no = '2277'
    Cur_Date='202408'
    list_attached = []
    for device_no in device_no_1:
        report_daily = generate_report_daily(cursor, device_no, INACTIVE, Cur_Date)
        if report_daily != 0:
            data = generate_data_yield_summary(cursor, device_no, INACTIVE, Cur_Date, cus_no)
            all_data = all_data_build(data, device_no)
            yield_hitter_report = generate_yield_hitter_report(all_data, device_no, Cur_Date, cus_no)
            base64_report_daily = convert_file_to_base64(report_daily)
            base64_hitter_report = convert_file_to_base64(yield_hitter_report)
            list_attached.append({
                "base64File": base64_report_daily,
                "fileName": report_daily,
                "mimeType" : "application/vnd.ms-excel"
            })
            list_attached.append({
                "base64File": base64_hitter_report,
                "fileName": yield_hitter_report,
                "mimeType" : "application/vnd.ms-excel"
            })
        else:
            print(f"Cannot generate report because {device_no} has no data on {Cur_Date}")
    cnxn.close()
    # sending_email(list_attached)
    
if __name__ == '__main__':

    main()



