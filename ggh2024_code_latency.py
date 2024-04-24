'''
1. ggh2024_code_latency.py is a python file that is resposible for calculating the average latency of the simulator provided monitor output, 
   a sample of which is given in this repository.
2. This code is applicable to any such monitor output, provided the column names remain the same.
'''
import openpyxl
def calculate_latency(sheet):
    read_timestamps = {}
    total_write_latency = 0
    total_read_latency = 0
    total_transactions = 0
    for row in sheet.iter_rows(min_row=2, values_only=True):
        timestamp, txn_type, data = row
        try:
            address = txn_type.split()[-1] if txn_type.startswith(('Rd', 'Data')) else None
        except IndexError:
            print(f"Error extracting address from transaction type: {txn_type}")
            continue
        if txn_type.startswith('Rd'):
            read_timestamps[address] = timestamp
            total_transactions += 1
        elif txn_type.startswith('Wr'):
            total_write_latency += 1
            total_transactions += 1
        elif txn_type.startswith('Data'):
            if address in read_timestamps:
                read_timestamp = read_timestamps.pop(address)
                total_read_latency += timestamp - read_timestamp
                total_transactions += 1
    return total_read_latency, total_write_latency, total_transactions
# Enter the path of your own Excel file here
file_path = 'C:\\Users\\mnsle\\OneDrive\\Desktop\\ggh2024\\monitor_output.xlsx'
workbook = openpyxl.load_workbook(file_path)
sheet = workbook.active
total_read_latency, total_write_latency, total_transactions = calculate_latency(sheet)
average_latency = (total_read_latency + total_write_latency) / total_transactions
print(f"Total read latency: {total_read_latency} cycles")
print(f"Total write latency: {total_write_latency} cycles")
print(f"Total transactions: {total_transactions}")
print(f"Average latency: {average_latency} cycles")
