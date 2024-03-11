import customtkinter as ctk
import Chart

class app(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.geometry("500x500")
        self.iconbitmap("skalar_analytical_bv_logo_Zoy_icon.ico")
        self.resizable(width=False, height=False)
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
                self.entry_value = self.entries[i].split(". ")[1].strip()
            else:
                self.entry_value = ""
            self.entry = ctk.CTkEntry(self.scrollable_frame, width=140, height=30)
            self.entry.grid(row=i, column=1)
            self.entry.insert(0, self.entry_value)
            self.entry.bind("<Return>", lambda event, index=i, entry=self.entry: self.update_script_file(index, entry))
            self.scrollable_frame_entry.append(self.entry)

    def update_script_file(self, index, entry):
        try:
            self.entry_content = int(entry.get())
            with open("script.txt", "r") as file:
                lines = file.readlines()
            lines[index] = f"{index + 1}. {self.entry_content}\n"
            with open("script.txt", "w") as file:
                file.writelines(lines)

        except ValueError:
            entry.delete(0, ctk.END)
            entry.insert(0, "Invalid value")
            self.after(2000, lambda: self.deletemessage(entry))

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
        app = Chart.App()
        app.deiconify()


class Confirm(ctk.CTk):
    def __init__(self, parent):
        super().__init__()
        self.geometry("200x200")
        self.title("Confirm")
        self.resizable(width=True, height=True)
        self.iconbitmap("skalar_analytical_bv_logo_Zoy_icon.ico")
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
