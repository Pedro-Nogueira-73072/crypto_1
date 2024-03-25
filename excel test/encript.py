import xlrd
import xlwt
from Crypto.Cipher import AES
from Crypto.Random import get_random_bytes
from Crypto.Util.Padding import pad

def encrypt_content(content, cipher):
    # Create an AES cipher object with the provided key
    
    # Pad the content to match the block size of AES
    padded_content = pad(content.encode('utf-8'), AES.block_size)
    
    # Encrypt the padded content
    encrypted_content = cipher.encrypt(padded_content)
    
    # Return the encrypted content
    return encrypted_content

def encrypt_excel_file(input_file_path, output_file_path, cipher):
    # Open the original Excel file
    workbook = xlrd.open_workbook(input_file_path)
    
    # Create a new workbook to write the encrypted content
    new_workbook = xlwt.Workbook()
    new_sheet = new_workbook.add_sheet('Sheet1')
    
    # Iterate over each sheet in the original workbook
    for sheet_index in range(workbook.nsheets):
        sheet = workbook.sheet_by_index(sheet_index)
        
        # Iterate over each row in the sheet
        for row_index in range(sheet.nrows):
            # Iterate over each cell in the row
            for col_index in range(sheet.ncols):
                # Read the content of the cell
                cell_content = sheet.cell_value(row_index, col_index)
                
                # Encrypt the content
                encrypted_content = encrypt_content(str(cell_content), cipher)
                
                # Write the encrypted content to the new workbook
                new_sheet.write(row_index, col_index, encrypted_content.hex())
    
    # Save the new workbook
    new_workbook.save(output_file_path)

# Example usage
input_file_path = "C:\\Users\\Pedro\\Documents\\GitHub\\crypto_1\\test.xls"
output_file_path = "C:\\Users\\Pedro\\Documents\\GitHub\\crypto_1\\test_output.xls"
encryption_key = b'mysecretpassword'  # Example key (must be 16, 24, or 32 bytes long)

cipher = AES.new(encryption_key, AES.MODE_CBC)
with open('cipher_file', 'wb')as c_file:
    c_file.write(cipher.iv)
    c_file.write(encryption_key)

encrypt_excel_file(input_file_path, output_file_path, cipher)

print("Content encrypted and saved to", output_file_path)
