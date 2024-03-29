import customtkinter as ctk
import time

class app(ctk.CTk):

    def __init__(self, channel):
        super().__init__()
        self.title(channel)
        self.iconbitmap("skalar_analytical_bv_logo_Zoy_icon.ico")
        self.resizable(width=False, height=False)
        self.time_entry_array = []
        self.script_channel = channel
        self.setpoint_value_entry_array = []
        self.channel_value_combobox_array = []
        self.entries = []
        self.i = 0
        self.x = 0
        self.test = 0
        self.load()


    def load(self):
        self.Scroll_Frame()
        self.file_check()
        self.reset_button()

    def Scroll_Frame(self):
        self.scrollable_frame = ctk.CTkScrollableFrame(self, width=210, label_text="Script Config")
        self.scrollable_frame.grid(row=0, column=0, padx=30, pady=30)
        self.scrollable_frame_entry = []

    def file_check(self):
        try:
            with open(fr"Sensor\Script {self.script_channel}.txt", "r") as file:
                self.entries = file.readlines()
                self.entries = [line.strip() for line in self.entries if line.strip()]
                self.Entry()
        except FileNotFoundError:
            with open(fr"Sensor\Script {self.script_channel}.txt", "a") as file:
                for i in range(10):
                    file.write(f"{i}. 0, 0, 0\n")
                    self.Entry()

    def Entry(self):
        print(len(self.entries))
        for self.i in range(len(self.entries)):
            if self.i < len(self.entries):
                self.time_value = self.entries[self.i].split(", ")[0].strip()
                self.setpoint_value = self.entries[self.i].split(", ")[1].strip()
                self.combobox_value = self.entries[self.i].split(", ")[2].strip()
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
        self.add_entry_button.grid(row=1, column=0, pady=(0, 30))

    def new_entry(self):
        self.Entry_add()
        with open(fr"Sensor\Script {self.script_channel}.txt", "r") as read:
            read_value = read.readlines()
            print(len(read_value))

        with open(fr"Sensor\Script {self.script_channel}.txt", "a") as write:
            line_number = len(read_value) + 1
            write.write(f"{line_number}. 0, 0, 0\n")

    def Entry_add(self):
        self.time_entry = ctk.CTkEntry(self.scrollable_frame, width=60, height=30, justify="center")
        self.time_entry.grid(row=self.i, column=2)
        self.time_entry.insert(0, self.time_value)
        self.time_entry_array.append(self.time_entry)
        self.setpoint_value_entry = ctk.CTkEntry(self.scrollable_frame, width=60, height=30, justify="center")
        self.setpoint_value_entry.grid(row=self.i, column=3)
        self.setpoint_value_entry.insert(0, self.setpoint_value)
        self.setpoint_value_entry_array.append(self.setpoint_value_entry)
        self.channel_value_combobox = ctk.CTkComboBox(self.scrollable_frame, width=60, height=30, justify="center", values=["1","2","3","All"], command= lambda event: self.update_script_file())
        self.channel_value_combobox.grid(row=self.i, column=4)
        self.channel_value_combobox.set(self.combobox_value)
        self.channel_value_combobox_array.append(self.channel_value_combobox)
        self.delete_button = ctk.CTkButton(self.scrollable_frame, text="X", width=27,command=lambda row=self.i: self.delete_row(row), fg_color="red")
        self.delete_button.grid(row=self.i, column=1)
        self.time_entry.bind("<Return>", lambda event: self.update_script_file())
        self.setpoint_value_entry.bind("<Return>", lambda event: self.update_script_file())
        self.scrollable_frame_entry.append(self.time_entry)


    def delete_row(self, index):
        with open(fr"Sensor\Script {self.script_channel}.txt", "r") as file:
            print(index)
            self.value_entry = file.readlines()
            print("Before: ",self.value_entry, " | ", self.x)
            del self.value_entry[index-self.x]
            print("After: ", self.value_entry)
            self.x += 1
        self.update_script_file(index)

    def update_script_file(self,index):
        for widget in self.scrollable_frame.grid_slaves(row=index):
            widget.destroy()
        with open(fr"Sensor\Script {self.script_channel}.txt", "w") as file:
            file.writelines(self.value_entry)

    def reset_button(self):
        self.reset_button = ctk.CTkButton(self, text="Reset script", command=self.confirm_reset)
        self.reset_button.grid(row=2, column=0, pady=(0,30))

    def confirm_reset(self):
        confirm_dialog = Confirm(self)
        confirm_dialog.mainloop()

    def reset_script(self):
        with open("default_script.txt","r") as read:
            default = read.readlines()
        with open(fr"Sensor\Script {self.script_channel}.txt","w") as write:
            write.writelines(default)
        self.file_check()

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