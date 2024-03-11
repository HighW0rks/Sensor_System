import customtkinter as ctk
import serial
import main
import threading
import time

class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.main_app = main.MainApp()
        self.title("PPM Meter")
        self.iconbitmap("skalar_analytical_bv_logo_Zoy_icon.ico")
        self.value = 0
        self.running = True
        self.command = b"\x02\x30\x35\x30\x30\x30\x37\x03"
        self.ppm_meter()

    def ppm_meter(self):
        self.ppm_meter_text = ctk.CTkLabel(self, text=self.value)
        self.ppm_meter_text.grid(row=0, column=0, sticky="nsew")
        self.reset_button = ctk.CTkButton(self, text="Restart", command=self.reset_button)
        self.reset_button.grid(row=1, column=0, sticky="nsew")
        self.update()
        self.threat()
    def threat(self):
        self.connect_thread = threading.Thread(target=self.connection)
        self.connect_thread.start()

    def reset_button(self):
        self.running = True
        print(self.running)
        self.threat()

    def connection(self):
        try:
            self.serial = serial.Serial("COM3", 19200, timeout=1)
            while self.running:
                self.sensor_data()
        except serial.serialutil.SerialException as e:
            self.ppm_meter_text.configure(text=f"Error connecting to serial port: {e}")
            self.close_serial()
            self.running = False
        except IndexError as e:
            self.ppm_meter_text.configure(text="Sensor not detected")
            self.close_serial()
            self.running = False

    def sensor_data(self):
        try:
            self.serial.write(self.command)
            self.response = self.serial.read(1024)
            self.response_str = self.response.decode()
            self.value = self.response_str.split("   ")[5]
            print(self.value)
            self.ppm_meter_text.configure(text=self.value)
            self.update()
            time.sleep(0.5)
        except serial.serialutil.SerialException as e:
            self.ppm_meter_text.configure(text=f"Error communicating with serial port: {e}")
            self.close_serial()
            self.running = False
    def close_serial(self):
        try:
            self.serial.close()
        except Exception:
            pass

    def destroy(self):
        self.close_serial()
        self.main_app.deiconify()
        super().destroy()
        main.MainApp()

if __name__ == '__main__':
    mainapp = App()
    mainapp.mainloop()
