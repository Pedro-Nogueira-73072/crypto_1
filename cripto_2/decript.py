import xlrd
import xlwt
from Crypto.Cipher import AES
from Crypto.Util.Padding import unpad

def decrypt_content(encrypted_content, cipher):
    
    # Decrypt the content
    decrypted_content = unpad(cipher.decrypt(encrypted_content), AES.block_size)
    
    # Decode the decrypted content to string
    return decrypted_content.decode()

def decrypt_excel_file(input_file_path, output_file_path, cipher):
    # Open the encrypted Excel file
    workbook = xlrd.open_workbook(input_file_path)
    
    # Create a new workbook to write the decrypted content
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
                
                # Decrypt the content
                decrypted_content = decrypt_content(bytes.fromhex(cell_content), cipher)
                
                # Write the decrypted content to the new workbook
                new_sheet.write(row_index, col_index, decrypted_content)
    
    # Save the new workbook
    new_workbook.save(output_file_path)

# Example usage
input_file_path = "C:\\Users\\Pedro\\Documents\\GitHub\\crypto_1\\test_output.xls"
output_file_path = "C:\\Users\\Pedro\\Documents\\GitHub\\crypto_1\\test_decrypt.xls"
encryption_key = b'mysecretpassword'  # Example key (must be 16, 24, or 32 bytes long)

with open("C:\\Users\\Pedro\\Documents\\GitHub\\crypto_1\\cipher_file", 'rb')as c_file:
    iv = c_file.read(16)
    encryption_key = c_file.read()
# Create an AES cipher object with the provided key
cipher = AES.new(encryption_key, AES.MODE_CBC, iv)

decrypt_excel_file(input_file_path, output_file_path, cipher)

print("Content decrypted and saved to", output_file_path)
