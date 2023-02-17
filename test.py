import binascii
str1 = 'прайс_whatsapp3.xlsx'
# str2 = "D:\YandexDisk\WORC\Python\text_to_img\dist\main.exe.notanexecutable\n"

# print(binascii.crc32(str1))
#
# # hash_value = hex(hash(str1))
# # hash_value = str(hash_value)
# # hash_value = hash_value.replace("-", "")
# #
# # print(hash_value)
# #
# # hash_value = hex(hash(str2))
# # hash_value = str(hash_value)
# # hash_value = hash_value.replace("-", "")
# #
# # print(hash_value)
#
#
# exit(-1)

file = 'temp_image' + ".xlsx"

with open(str1, "rb") as file_img:
    a = file_img.read()
with open(file, "wb") as file_temp:
    file_temp.write(a)
