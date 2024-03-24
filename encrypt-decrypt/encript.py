import xlrd
import xlwt
from Crypto.Cipher import AES
from Crypto.Util.Padding import unpad, pad

key = b'mysecretpassword'
cipher = AES.new(key, AES.MODE_CBC) #especificar iv e usar sempre
with open('cipher_file.txt', 'wb')as file:
    file.write(cipher.iv)

# Pad the content to match the block size of AES
padded_content = pad("hello".encode('utf-8'), AES.block_size)

# Encrypt the padded content
encrypted_content = cipher.encrypt(padded_content)

print("encrypted content:")
print(encrypted_content)
