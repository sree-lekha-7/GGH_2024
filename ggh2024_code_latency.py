import openpyxl

def calculate_latency(sheet):
    read_timestamps = {}
    total_write_latency = 0
    total_read_latency = 0
    total_transactions = 0
    
    for row in sheet.iter_rows(min_row=2, values_only=True):
        timestamp, txn_type, data = row
        
        # Extract address
        try:
            address = txn_type.split()[-1] if txn_type.startswith(('Rd', 'Data')) else None
        except IndexError:
            print(f"Error extracting address from transaction type: {txn_type}")
            continue
        
        # Handle read operation
        if txn_type.startswith('Rd'):
            read_timestamps[address] = timestamp
            total_transactions += 1
        
        # Handle write operation
        elif txn_type.startswith('Wr'):
            total_write_latency += 1
            total_transactions += 1
        
        # Handle data operation
        elif txn_type.startswith('Data'):
            if address in read_timestamps:
                read_timestamp = read_timestamps.pop(address)
                total_read_latency += timestamp - read_timestamp
                total_transactions += 1
    
    return total_read_latency, total_write_latency, total_transactions

# Load the Excel file
file_path = 'C:\\Users\\mnsle\\OneDrive\\Desktop\\ggh2024\\monitor_output.xlsx'
workbook = openpyxl.load_workbook(file_path)
sheet = workbook.active

# Calculate latency
total_read_latency, total_write_latency, total_transactions = calculate_latency(sheet)

# Calculate average latency
average_latency = (total_read_latency + total_write_latency) / total_transactions

# Print results
print(f"Total read latency: {total_read_latency} cycles")
print(f"Total write latency: {total_write_latency} cycles")
print(f"Total transactions: {total_transactions}")
print(f"Average latency: {average_latency} cycles")
