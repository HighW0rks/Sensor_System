config_file = "config.txt"

def readfile_value(row):
    with open(config_file, "r") as read:
        config_text_value = read.readlines()
    try:
        parameter_text_configvalue = str(config_text_value[row - 1]).split(": ")[1].strip()
    except IndexError:
        parameter_text_configvalue = ""
    return parameter_text_configvalue


def text_config(row, name):
    with open(config_file, "r") as read:
        config_text_value = read.readlines()
    parameter_text_configname = str(config_text_value[row - 1]).split(": ")[0].strip()
    try:
        config_new_value = name.get()
    except Exception:
        config_new_value = name
    updated_text_row = f"{parameter_text_configname}: {config_new_value}\n"
    config_text_value[row - 1] = updated_text_row
    with open(config_file, "w") as write:
        write.writelines(config_text_value)