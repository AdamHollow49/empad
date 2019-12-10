"""
Authored by: Adam Hollow - DevOps Technician, Devices.
For: South Ayrshire Council

"""

from cryptography.fernet import Fernet

key = ###ENCRYPTION KEY###  # Best practice to store this outside of the application, in a file.
f = Fernet(key)

def decrKey():
    encrypted = ### ENCRYPTED PASSWORD ###
    decrypted = f.decrypt(encrypted)
    return str(decrypted)[2:-1]


