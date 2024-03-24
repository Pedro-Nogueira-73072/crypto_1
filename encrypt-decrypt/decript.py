import xlrd
import xlwt
from Crypto.Cipher import AES
from Crypto.Util.Padding import unpad, pad


with open("C:\\Users\\Pedro\\Documents\\GitHub\\crypto_1\\cipher_file.txt", 'rb')as file:
    iv = file.read(16)
key = b'mysecretpassword'
# Create an AES cipher object with the provided key
cipher = AES.new(key, AES.MODE_CBC, iv)

# Decrypt the content
decrypted_content = unpad(cipher.decrypt(b'Gn`cR^)\x82\x83\xc5\xe3G\xf8\xcd\x95O'), AES.block_size)

# Decode the decrypted content to string
print("decrypted content:")
print(decrypted_content.decode())
