import os
import sys
from fernet import Fernet

key = b"MnI-3q_UH-1RIHSVUVQvSmtvBsvEKv8U88QVH7hkxDU="
crypter = Fernet(key)

execute_path = os.path.abspath(sys.argv[0])
config_file = os.path.dirname(execute_path) + r"\config.bin"

def readfile_value(row):
    # Read and decrypt the config file
    with open(config_file, "rb") as file:
        encrypted_data = file.read()
    decrypted_data = crypter.decrypt(encrypted_data).decode()

    # Split the decrypted data into lines
    config_lines = decrypted_data.splitlines()

    # Try to get the specific row's value
    try:
        parameter_text_configvalue = config_lines[row - 1].split(": ")[1].strip()
    except IndexError:
        parameter_text_configvalue = ""
    return parameter_text_configvalue


def text_config(row, name):
    # Read and decrypt the config file
    with open(config_file, "rb") as read:
        encrypted_data = read.read()
    decrypted_data = crypter.decrypt(encrypted_data).decode()
    config_lines = decrypted_data.splitlines()

    # Update the specific row
    try:
        parameter_text_configname = config_lines[row - 1].split(": ")[0].strip()
        if hasattr(name, 'get'):
            config_new_value = name.get()
        else:
            config_new_value = name
        updated_text_row = f"{parameter_text_configname}: {config_new_value}"
        config_lines[row - 1] = updated_text_row
    except IndexError:
        print("Row index out of range.")
        return

    # Join the lines back together and encrypt
    updated_config_text = "\n".join(config_lines)
    encrypted_data = crypter.encrypt(updated_config_text.encode())

    # Write the updated and encrypted config file
    with open(config_file, "wb") as write:
        write.write(encrypted_data)