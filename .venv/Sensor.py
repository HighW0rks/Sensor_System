import _tkinter
import customtkinter as ctk
import serial
import main
import threading
import time

class App(ctk.CTk):
    def __init__(self, con):
        super().__init__()
        self.title("PPM Meter")
        self.iconbitmap("skalar_analytical_bv_logo_Zoy_icon.ico")
        self.value = 0
        self.c = con
        self.running = None
        self.command = b"\x02\x30\x35\x30\x30\x30\x37\x03"
        self.ppm_meter()

    def ppm_meter(self):
        self.ppm_meter_text = ctk.CTkLabel(self, text=self.value)
        self.ppm_meter_text.grid(row=0, column=0, padx=50)
        self.update()
        self.thread()

    def thread(self):
        self.connect_thread = threading.Thread(target=self.connection)
        self.connect_thread.daemon = True
        self.connect_thread.start()

    def connection(self):
        self.status = self.c.status_sensor
        if self.status == True:
            self.serial = self.c.sensor
            self.ppm_meter_text.configure(text="Starting connection")
            self.running = True
        else:
            self.ppm_meter_text.configure(text="No sensor was detected 1")
            self.running = False

        if self.running:
            self.sensor_data()
        else:
            self.after(1000, self.thread)

    def sensor_data(self):
        try:
            print(self.serial)
            self.serial.write(self.command)
            self.response = self.serial.read(1024).decode().split("   ")[5]
            self.ppm_meter_text.configure(text=self.response)
            self.update()
            time.sleep(0.5)
        except _tkinter.TclError:
            pass
        except Exception as e:
            print(e)
            time.sleep(0.5)
            self.ppm_meter_text.configure(text=f"No sensor was detected 2")
            self.c.close_sensor()
            self.c.initialize_sensor()
            self.running = False
            self.thread()

        if self.running == True:
            try:
                self.sensor_data()
            except Exception:
                pass

    def destroy(self):
        super().destroy()
        self.quit()

if __name__ == '__main__':
    mainapp = App()
    mainapp.mainloop()
