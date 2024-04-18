import customtkinter as ctk
from Config import readfile_value
import time
import Chart
import os
import sys

execute_path = os.path.abspath(sys.argv[0])
icon = os.path.dirname(execute_path) + r"\skalar_analytical_bv_logo_Zoy_icon.ico"
default_script = os.path.dirname(execute_path) + "/default_script.txt"

class app(ctk.CTk):

    def __init__(self, artikel, channel):
        super().__init__()
        self.title(channel)
        self.iconbitmap(icon)
        self.resizable(width=False, height=False)
        self.time_entry_array = []
        self.sensor_type = artikel
        self.script_channel = channel
        self.setpoint_value_entry_array = []
        self.channel_value_combobox_array = []
        self.folder = readfile_value(8)
        self.entries = []
        self.i = 0
        self.x = 0
        self.test = 0
        self.load()

    def load(self):
        self.Scroll_Frame()
        self.file_check()

    def Scroll_Frame(self):
        self.scrollable_frame = ctk.CTkScrollableFrame(self, width=210, label_text="Script Config")
        self.scrollable_frame.grid(row=0, column=0, columnspan=2, padx=30, pady=30)
        self.scrollable_frame_entry = []

    def file_check(self):
        try:
            with open(fr"{self.folder}\{self.sensor_type}\Scripts\Script {self.script_channel}.txt", "r") as file:
                self.entries = file.readlines()
                self.entries = [line.strip() for line in self.entries if line.strip()]
                self.Entry()
        except FileNotFoundError:
            with open(fr"{self.folder}\{self.sensor_type}\Scripts\Script {self.script_channel}.txt", "a") as file:
                for i in range(10):
                    file.write(f"{i}. 0; 0; 0\n")
                    self.Entry()

    def Entry(self):
        for self.i in range(len(self.entries)):
            if self.i < len(self.entries):
                self.time_value = self.entries[self.i].split("; ")[0].split(". ")[1].strip()
                self.setpoint_value = self.entries[self.i].split("; ")[1].strip()
                self.combobox_value = self.entries[self.i].split("; ")[2].strip()
            else:
                self.time_value = ""
                self.setpoint_value = ""
                self.combobox_value = ""
            self.Entry_add()
        self.time_value = ""
        self.setpoint_value = ""
        self.combobox_value = ""
        self.i += 1
        self.add_entry_button = ctk.CTkButton(self, text="Add a new row", command= self.new_entry)
        self.add_entry_button.grid(row=1, column=0, columnspan=2, pady=(0,5))
        self.delete_button = ctk.CTkButton(self, text="Delete script", command= lambda: self.delete_script(True))
        self.delete_button.grid(row=2, column=0, pady=(0,5))
        self.reset_button = ctk.CTkButton(self, text="Reset script", command=self.confirm_reset)
        self.reset_button.grid(row=2, column=1, pady=(0,5))

    def new_entry(self):
        self.Entry_add()
        with open(fr"{self.folder}\{self.sensor_type}\Scripts\Script {self.script_channel}.txt", "r") as read:
            read_value = read.readlines()
        with open(fr"{self.folder}\{self.sensor_type}\Scripts\Script {self.script_channel}.txt", "a") as write:
            line_number = len(read_value) + 1
            write.write(f"{line_number}. 0; 0; 0\n")
        self.i += 1

    def Entry_add(self):
        self.time_entry = ctk.CTkEntry(self.scrollable_frame, width=60, height=30, justify="center")
        self.time_entry.grid(row=self.i, column=1)
        self.time_entry.insert(0, self.time_value)
        self.time_entry_array.append(self.time_entry)
        self.setpoint_value_entry = ctk.CTkEntry(self.scrollable_frame, width=60, height=30, justify="center")
        self.setpoint_value_entry.grid(row=self.i, column=2)
        self.setpoint_value_entry.insert(0, self.setpoint_value)
        self.setpoint_value_entry_array.append(self.setpoint_value_entry)
        self.channel_value_combobox = ctk.CTkComboBox(self.scrollable_frame, width=60, height=30, justify="center", values=["1","2","3","All"], command= lambda event: self.update_script_file())
        self.channel_value_combobox.grid(row=self.i, column=3)
        self.channel_value_combobox.set(self.combobox_value)
        self.channel_value_combobox_array.append(self.channel_value_combobox)
        self.time_entry.bind("<KeyRelease>", lambda event: self.update_script_file())
        self.setpoint_value_entry.bind("<KeyRelease>", lambda event: self.update_script_file())
        self.scrollable_frame_entry.append(self.time_entry)

    def update_script_file(self):
        with open(fr"{self.folder}\{self.sensor_type}\Scripts\Script {self.script_channel}.txt", "r") as file:
            lines = file.readlines()
        for i in range(len(lines)):
            lines[i] = f"{i + 1}. {self.time_entry_array[i].get()}; {self.setpoint_value_entry_array[i].get()}; {self.channel_value_combobox_array[i].get()}\n"
        with open(fr"{self.folder}\{self.sensor_type}\Scripts\Script {self.script_channel}.txt", "w") as file:
            file.writelines(lines)

    def delete_script(self, x=False):
        with open(fr"{self.folder}\{self.sensor_type}\Scripts\Script {self.script_channel}.txt", "r") as file:
            lines = file.readlines()
        for index in range(10000):
            for widget in self.scrollable_frame.grid_slaves(row=index):
                widget.destroy()
        if x == True:
            with open(fr"{self.folder}\{self.sensor_type}\Scripts\Script {self.script_channel}.txt", "w") as write:
                write.truncate()
        self.time_entry_array = []
        self.setpoint_value_entry_array = []
        self.channel_value_combobox_array = []
        self.file_check()

    def confirm_reset(self):
        Confirm(self).mainloop()

    def reset_script(self):
        with open(default_script,"r") as read:
            default = read.readlines()
        with open(fr"{self.folder}\{self.sensor_type}\Scripts\Script {self.script_channel}.txt","w") as write:
            write.writelines(default)
        self.delete_script()

    def deletemessage(self, entry):
        entry.delete(0, ctk.END)

    def destroy(self):
        super().destroy()


class Confirm(ctk.CTk):
    def __init__(self, parent):
        super().__init__()
        self.geometry("200x200")
        self.title("Confirm")
        self.resizable(width=True, height=True)
        self.parent = parent
        self.Confirm()

    def Confirm(self):
        self.note_text = ctk.CTkLabel(self, text="All current data will be lost")
        self.note_text.grid(row=0, column=0)
        self.confirm_text = ctk.CTkLabel(self, text="Are you sure?")
        self.confirm_text.grid(row=1, column=0)
        self.confirm_button = ctk.CTkButton(self, text="Confirm", command=self.confirm_reset)
        self.confirm_button.grid(row=2, column=0)

    def confirm_reset(self):
        self.destroy()
        self.parent.reset_script()

    def destroy(self):
        super().destroy()


if __name__ == '__main__':
    channel = "Ch1"
    app(channel).mainloop()