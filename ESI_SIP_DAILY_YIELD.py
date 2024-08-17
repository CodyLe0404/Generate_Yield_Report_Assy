import xlrd
import pyodbc
import pandas as pd
from xlutils.copy import copy
from xlwt import Workbook, XFStyle, Borders
import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.styles import Border, Side, Alignment

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
    return Yield

# Define a dictionary to store the data
def generate_report_daily(cursor, device_no, INACTIVE, Cur_Date, report_type):
    data_INACTIVE = get_data_group(cursor, device_no, INACTIVE, Cur_Date)
    if data_INACTIVE == []:
        raise Exception("There is no data")

    data_dict = {}
    print("INACTIVE")
    for index in data_INACTIVE:
        print(index)

    for index in data_INACTIVE:
        data_dict[index[0]] = {'In': index[1], 'Out': index[2], 'Yield': Get_Yield(index[3])}

    # Load the workbook and select the first sheet and define list of the keys
    if report_type == 'ESI':
        rb = xlrd.open_workbook(r"C:\Workplace\Task\Support_Assy\Auto_Mail_Yield\ESI_SIP_Sample_input.xls", formatting_info=True)
        keys = ['SUB/L', 'SMT1', 'MOLD1', 'SMT2', 'MOLD2', 'SMT3', 'LASER', 'PKG Saw', 'SPUTTER1', 'SPUTTER2', 'DMZ &FVI', 'SLT0', 'SLT1', 'SLT2', 'SLT3', 'AVI/TNR']
    elif report_type == 'QORVO':
        rb = xlrd.open_workbook(r"C:\Workplace\Task\Support_Assy\Auto_Mail_Yield\QORVO_Sample_input.xls", formatting_info=True)
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
    Overall_Yield = str(Overall_Yield) + '%'

    # Write the overall yield, in and out
    sheet.write(len(keys) + 4, 8, Overall_Yield, style)
    sheet.write(len(keys) + 4, 9, Overall_In, style)
    sheet.write(len(keys) + 4, 10, Overall_Out, style)

    # Save the workbook
    wb.save(f'ESI_IO_YIELD_{device_no}.xls')
    print("Exported")

def generate_yield_summary(cursor, device_no, INACTIVE, Cur_Date, cus_no):
    data_dict = {}
    data_dict_hitter = {}
    data_INACTIVE = get_data_group(cursor, device_no, INACTIVE, Cur_Date)
    data_Hitter = get_hitter(cursor, device_no, Cur_Date, cus_no )
    for index in data_INACTIVE:
        failQty = index[1] - index[2]
        hit_type = index[0]
        hitter_info = {'In': index[1], 'Fail' : failQty,'Yield': Get_Yield(index[3])}    
    # Append the hitter info to a list under the hit type key
        if hit_type not in data_INACTIVE:
            data_dict[hit_type] = hitter_info
        else:
            data_dict[hit_type].append(hitter_info)

    # for index in data_Hitter:
    for ele in data_Hitter:
        station = ele[-1]
        failDefect = int(ele[-2])
        failStation = int(data_dict[station]['Fail'])
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
    stations = ['SUB/L','SMT1', 'Mold1', 'SMT2','Mold2', 'SMT3', 'LASER', 'PKG Saw', 'SPUTTER1', 'SPUTTER2', 'DMZ &FVI', 'SLT0', 'SLT1', 'SLT2', 'SLT3', 'AVI/TNR']
    for station in stations:
        if station not in data_dict_hitter:
            continue
        data_dict[station]['Hitter'] = data_dict_hitter[station]
    return data_dict

def all_data_build(data):
    data_all = {
            'FOL' : {
                'SUB/L' : "",
                'SMT1' : "", 
                'Mold1' : "", 
                'SMT2' : ""},
            'EOL' : {
                'Mold2' : "", 
                'SMT3' : "", 
                'LASER' : "", 
                'PKG Saw' : "", 
                'SPUTTER1' : "", 
                'SPUTTER2' : "", 
                'DMZ &FVI' : ""}, 
            'TEST' : {
                'SLT0' : "", 
                'SLT1' : "", 
                'SLT2' : "", 
                'SLT3' : "", 
                'AVI/TNR' : ""}}
    stations_EOL = ['SUB/L','SMT1', 'Mold1', 'SMT2']
    stations_FOL = ['Mold2', 'SMT3', 'LASER', 'PKG Saw', 'SPUTTER1', 'SPUTTER2', 'DMZ &FVI']
    stations_TEST = ['SLT0', 'SLT1', 'SLT2', 'SLT3', 'AVI/TNR']
    for station in stations_EOL:
        if station not in data:
            continue
        data_all['EOL'][station] = data[station]
    for station in stations_FOL:
        if station not in data:
            continue
        data_all['FOL'][station] = data[station]
    for station in stations_TEST:
        if station not in data:
            continue
        data_all['TEST'][station] = data[station]
    # print(data_all)
    return data_all

def generate_yield_hitter_report(data_all):
    # Tạo workbook và active worksheet
    wb = openpyxl.Workbook()
    ws = wb.active
    start_row = 5
    current_row = start_row

    # Thiết lập các đường viền
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

    # Ghi dữ liệu vào file Excel
    for group, stations in data_all.items():
        group_start_row = current_row
        print(f"{group} : ")
        for station, data in stations.items():
            station_start_row = current_row
            print(f" {station} : ")
            if data == "": continue
            for info, detail in data.items():
                data_start_row = current_row
                if len(str(detail)) < 10:
                    print(f"  {info} : {detail}")
                else:
                    print(f"  {info} : ")
                    for hitter_no in detail:

                        detail_start_row = current_row
                        for hitter, hitter_info in hitter_no.items():
                            print(f'     {hitter} : ')
                            for key, value in hitter_info.items():
                                print(f'      {key} : {value}')
                                current_row += 1

    # Lưu file Excel
    wb.save(r'yield_summary.xlsx')
    print(f"exported")

def main():
    cnxn = pyodbc.connect("DRIVER={ODBC Driver 17 for SQL Server};SERVER=10.201.21.84,50150;DATABASE=MCSDB;UID=cimitar2;PWD=TFAtest1!2!")
    cursor = cnxn.cursor()
    ESI_18807 = '639-18807'
    ESI_18808 = '639-18808'
    QM76300 = 'QM76300'
    QM76309 = 'QM76309'
    QM76095 = 'QM76095'
    INACTIVE = 'INACTIVE'
    ACTIVE = 'OTHERSTATUS'
    Cur_Date = '202408'
    cus_no = '2277'
    
    # generate_report_daily(cursor, ESI_18808, INACTIVE, Cur_Date, report_type='ESI')
    data = generate_yield_summary(cursor, ESI_18808, INACTIVE, Cur_Date, cus_no)
    all_data = all_data_build(data)
    generate_yield_hitter_report(all_data)


if __name__ == '__main__':
    main()






