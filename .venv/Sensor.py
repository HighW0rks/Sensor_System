import _tkinter
import customtkinter as ctk
import serial
import main
import threading
import time

class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("PPM Meter")
        self.iconbitmap("skalar_analytical_bv_logo_Zoy_icon.ico")
        self.value = 0
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
        try:
            self.serial = serial.Serial("COM5", 19200, timeout=1)
            self.ppm_meter_text.configure(text="Starting connection")
            self.running = True
        except serial.serialutil.SerialException as e:
            self.ppm_meter_text.configure(text=f"Error connecting to serial portt: {e}")
            self.running = False

        except IndexError as e:
            self.ppm_meter_text.configure(text="Sensor not detected")
            self.running = False
        if self.running:
            self.sensor_data()
        else:
            self.thread()

    def sensor_data(self):
        try:
            self.serial.write(self.command)
            self.response = self.serial.read(1024).decode().split("   ")[5]
            self.ppm_meter_text.configure(text=self.response)
            self.update()
            time.sleep(0.5)
        except Exception as e:
            self.ppm_meter_text.configure(text=f"Error communicating with serial port: {e}")
            self.running = False
            self.thread()
        if self.running:
            try:
                self.sensor_data()
            except Exception:
                pass

    def close_serial(self):
        try:
            self.serial.close()
        except Exception:
            pass
        except serial.serialutil.PortNotOpenError:
            pass

    def destroy(self):
        try:
            self.serial.close()
        except Exception:
            pass
        super().destroy()
        main.MainApp().mainloop()

if __name__ == '__main__':
    mainapp = App()
    mainapp.mainloop()
