import pandas as pd
import json

# Function to read VM size from Excel file
def read_vm_size_from_excel(excel_file_path, sheet_name, column_name, row_index):
    # Read the Excel file
    df = pd.read_excel(excel_file_path, sheet_name=sheet_name)
    
    # Get the VM size from the specified column and row
    vm_size = df.loc[row_index, column_name]
    return vm_size

# Function to read the Azure VM template (JSON)
def read_azure_vm_template(json_file_path):
    with open(json_file_path, 'r') as file:
        template = json.load(file)
    return template

# Function to verify VM size in the template
def verify_vm_size(template, key_path, expected_vm_size):
    # Traverse the template using the key path
    keys = key_path.split('.')
    current_value = template
    for key in keys:
        if '[' in key and ']' in key:  # Handle array indices in key path
            key, index = key.split('[')
            index = int(index[:-1])
            current_value = current_value.get(key, [])[index]
        else:
            current_value = current_value.get(key, None)
        if current_value is None:
            return False, f"Key '{key}' not found in the template."
    
    # Compare the extracted value with the expected value
    if current_value == expected_vm_size:
        return True, "VM size matches."
    else:
        return False, f"VM size mismatch. Expected: {expected_vm_size}, Found: {current_value}"

# Main execution
if __name__ == "__main__":
    # File paths
    excel_file_path = "C:\\WORK\\python\\vm_parameters.xlsx"
    json_file_path = "C:\\WORK\python\\azure_vm_template.json"
    
    # Excel parameters
    sheet_name = "Sheet1"  # Adjust as per your Excel sheet name
    column_name = "VM Size"  # Adjust as per your Excel column name
    row_index = 0  # The row number where the VM size is located
    
    # Key path in the JSON template for VM size
    key_path = "properties.hardwareProfile.vmSize"  # Adjust based on your template structure
    
    # Step 1: Read the VM size from the Excel file
    vm_size_from_excel = read_vm_size_from_excel(excel_file_path, sheet_name, column_name, row_index)
    print(f"VM size from Excel: {vm_size_from_excel}")
    
    # Step 2: Read the Azure VM template (JSON)
    azure_template = read_azure_vm_template(json_file_path)
    
    # Step 3: Verify the VM size in the template
    is_match, message = verify_vm_size(azure_template, key_path, vm_size_from_excel)
    print(message)
