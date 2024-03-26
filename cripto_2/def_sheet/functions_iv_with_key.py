import xlrd
import xlwt
import secrets
import os
from Crypto.Cipher import AES
from Crypto.Random import get_random_bytes
from Crypto.Util.Padding import pad, unpad

def encrypt_content(content, cipher):               #create AES cipher object with random key
    
    padded_content = pad(content.encode('utf-8'), AES.block_size)
    encrypted_content = cipher.encrypt(padded_content)
    
    return encrypted_content

def decrypt_content(content, cipher):

    decrypted_content = unpad(cipher.decrypt(content), AES.block_size)
    
    return decrypted_content.decode()

def encrypt_excel_file(input_file_path):

    encryption_key = secrets.token_bytes(16)            #16 random bytes
    cipher = AES.new(encryption_key, AES.MODE_CBC)
    with open('cipher_file', 'wb')as c_file:
        c_file.write(cipher.iv)
        c_file.write(encryption_key)
    
    workbook = xlrd.open_workbook(input_file_path)

    new_workbook = xlwt.Workbook()
    
    for sheet_index in range(workbook.nsheets):         #each sheet
        sheet = workbook.sheet_by_index(sheet_index)
        new_sheet = new_workbook.add_sheet(sheet.name)
        for row_index in range(sheet.nrows):            #each row
            for col_index in range(sheet.ncols):        #each cell
                
                cell_content = sheet.cell_value(row_index, col_index)
                encrypted_content = encrypt_content(str(cell_content), cipher)
                
                new_sheet.write(row_index, col_index, encrypted_content.hex())      #write encrypted values
    
    new_workbook.save(input_file_path)

def decrypt_excel_file(input_file_path, cipher_file_path):

    with open(cipher_file_path, 'rb')as c_file:
        iv = c_file.read(16)
        encryption_key = c_file.read()
    cipher = AES.new(encryption_key, AES.MODE_CBC, iv)

    workbook = xlrd.open_workbook(input_file_path)
    
    new_workbook = xlwt.Workbook()
    
    for sheet_index in range(workbook.nsheets):         #each sheet
        sheet = workbook.sheet_by_index(sheet_index)
        new_sheet = new_workbook.add_sheet(sheet.name)
        for row_index in range(sheet.nrows):            #each row
            for col_index in range(sheet.ncols):        #each cell

                cell_content = sheet.cell_value(row_index, col_index)
                decrypted_content = decrypt_content(bytes.fromhex(cell_content), cipher)
                
                new_sheet.write(row_index, col_index, decrypted_content)
    
    new_workbook.save(input_file_path)


input_file_path = "C:\\Users\\Pedro\\Documents\\GitHub\\crypto_1\\test.xls"
cipher_file_path = "C:\\Users\\Pedro\\Documents\\GitHub\\crypto_1\\cipher_file"


directory_path, file_name_with_extension = os.path.split(input_file_path)
cipher_directory = directory_path + "\\" + "cipher_file"
print(cipher_directory)

#encrypt_excel_file(input_file_path)
#decrypt_excel_file(input_file_path, cipher_file_path)