import openpyxl

def calculate_bandwidth(excel_path):
    # Open the Excel file
    try:
        workbook = openpyxl.load_workbook(excel_path)
        sheet = workbook.active
    except FileNotFoundError:
        print("Error: File not found.")
        return None, None, None
    except Exception as e:
        print("Error:", e)
        return None, None, None
    
    # Initialize variables
    total_bytes_transferred = 0
    last_transaction_time = None
    
    # Iterate through transactions
    for row in sheet.iter_rows(min_row=2, values_only=True):
        timestamp, txn_type, data = row

        # If it's a read transaction, skip
        if txn_type.startswith('Rd'):
            continue
        
        # If it's a write transaction, add 32 bytes to total_bytes_transferred
        total_bytes_transferred += 32
        
        # Update last transaction time
        last_transaction_time = timestamp
    
    # Calculate total time duration
    total_time_duration = last_transaction_time
    
    # Calculate bandwidth
    if total_time_duration:
        bandwidth = total_bytes_transferred / total_time_duration
    else:
        bandwidth = 0
    
    return total_bytes_transferred, total_time_duration, bandwidth

# Excel file path
excel_path = "C:\\Users\\mnsle\\OneDrive\\Desktop\\ggh2024\\monitor_output.xlsx"

# Calculate bandwidth
total_bytes, total_time, bandwidth = calculate_bandwidth(excel_path)
if total_bytes is not None and total_time is not None and bandwidth is not None:
    print("Total Bytes Transferred:", total_bytes)
    print("Total Time Duration:", total_time, "cycles")
    print("Bandwidth:", bandwidth, "bytes/cycle")
