import customtkinter as ctk

class app(ctk.CTk):
    
    def __init__(self):
        super().__init__()
        self.geometry("500x500")
        self.resizable(width=False, height=False)
        self.time_entry_array = []
        self.setpoint_value_entry_array = []
        self.channel_value_combobox_array = []
        self.entries = []
        self.load()
        

    def load(self):
        self.Scroll_Frame()
        self.file_check()
        self.reset_button()

    def Scroll_Frame(self):
        self.scrollable_frame = ctk.CTkScrollableFrame(self, label_text="Script Config")
        self.scrollable_frame.grid(row=0, column=0)
        self.scrollable_frame_entry = []

    def file_check(self):
        try:
            with open("script.txt", "r") as file:
                self.entries = file.readlines()
                self.Entry()
        except FileNotFoundError:
            with open("script.txt", "a") as file:
                for i in range(10):
                    file.write(f"{i}. \n")
                    self.Entry()

    def Entry(self):
        for i in range(10):
            if i < len(self.entries):
                self.time_value = self.entries[i].split(", ")[0].split(".")[1].strip()
                self.setpoint_value = self.entries[i].split(", ")[1].strip()
                self.combobox_value = self.entries[i].split(", ")[2].strip()
            else:
                self.time_value = ""
                self.setpoint_value = ""
                self.combobox_value = ""

            self.time_entry = ctk.CTkEntry(self.scrollable_frame, width=60, height=30, justify="center")
            self.time_entry.grid(row=i, column=1)
            self.time_entry.insert(0, self.time_value)
            self.time_entry_array.append(self.time_entry)
            self.setpoint_value_entry = ctk.CTkEntry(self.scrollable_frame, width=60, height=30, justify="center")
            self.setpoint_value_entry.grid(row=i, column=2)
            self.setpoint_value_entry.insert(0, self.setpoint_value)
            self.setpoint_value_entry_array.append(self.setpoint_value_entry)
            self.channel_value_combobox = ctk.CTkComboBox(self.scrollable_frame, width=60, height=30, justify="center", values=["1","2","3","All"], command= lambda event: self.update_script_file())
            self.channel_value_combobox.grid(row=i, column=3)
            self.channel_value_combobox.set(self.combobox_value)
            self.channel_value_combobox_array.append(self.channel_value_combobox)
            self.time_entry.bind("<Return>", lambda event: self.update_script_file())
            self.setpoint_value_entry.bind("<Return>", lambda event: self.update_script_file())
            self.scrollable_frame_entry.append(self.time_entry)

    def update_script_file(self):
        with open("script.txt", "r") as file:
            lines = file.readlines()
        for i in range(10):
            lines[i] = f"{i + 1}. {self.time_entry_array[i].get()}, {self.setpoint_value_entry_array[i].get()}, {self.channel_value_combobox_array[i].get()}\n"
        with open("script.txt", "w") as file:
            file.writelines(lines)

    def reset_button(self):
        self.reset_button = ctk.CTkButton(self, text="Reset script", command=self.confirm_reset)
        self.reset_button.grid(row=1, column=0)

    def confirm_reset(self):
        confirm_dialog = Confirm(self)
        confirm_dialog.mainloop()

    def reset_script(self):
        with open("default_script.txt","r") as read:
            default = read.readlines()
        with open("script.txt","w") as write:
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
    app().mainloop()
