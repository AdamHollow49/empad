import cryptography
import csv
from cryptography.fernet import Fernet

key = ###ENCRYPTIONKEY###
r = Fernet(key)
lst2 = []

with open(r'\\PATH_TO_CSV_FOLDER\Emp_Info.csv', 'r') as csvfile:
    readCSV = csv.reader(csvfile, delimiter=',')
    for row in readCSV:
        lst = row
        for x in range(11):
            lst[x] = r.decrypt(str.encode(lst[x][2:-1]))
            row[x] = str(lst[x])[2:-1]
        lst2.append(row)

with open(
        r'C:\PATH_TO_DECRYPT_FOLDER\Decrypted.csv', 'a', newline='') as csvfile:
    for listz in lst2:
        writer = csv.writer(csvfile)
        writer.writerow(listz)