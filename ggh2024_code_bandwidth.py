'''
1. ggh2024_code_bandwidth.py is a python file that is resposible for calculating the bandwidth of the simulator provided monitor output, 
   a sample of which is given in this repository.
2. This code is applicable to any such monitor output, provided the column names remain the same.
'''
import openpyxl
def calculate_bandwidth(excel_path):
    try:
        workbook = openpyxl.load_workbook(excel_path)
        sheet = workbook.active
    except FileNotFoundError:
        print("Error: File not found.")
        return None, None, None
    except Exception as e:
        print("Error:", e)
        return None, None, None
    total_bytes_transferred = 0
    last_transaction_time = None
    for row in sheet.iter_rows(min_row=2, values_only=True):
        timestamp, txn_type, data = row
        if txn_type.startswith('Rd'):
            continue
        total_bytes_transferred += 32
        last_transaction_time = timestamp
    total_time_duration = last_transaction_time
    if total_time_duration:
        bandwidth = total_bytes_transferred / total_time_duration
    else:
        bandwidth = 0
    return total_bytes_transferred, total_time_duration, bandwidth
# Enter the path of your own Excel file here.
excel_path = "C:\\Users\\mnsle\\OneDrive\\Desktop\\ggh2024\\monitor_output.xlsx"
total_bytes, total_time, bandwidth = calculate_bandwidth(excel_path)
if total_bytes is not None and total_time is not None and bandwidth is not None:
    print("Total Bytes Transferred:", total_bytes)
    print("Total Time Duration:", total_time, "cycles")
    print("Bandwidth:", bandwidth, "bytes/cycle")
