import propar
import serial
import time
import threading
import Chart
from Config import readfile_value


class Connection:
    _instance = None
    _lock = threading.Lock()

    def __new__(cls, *args, **kwargs):
        with cls._lock:
            if not cls._instance:
                cls._instance = super().__new__(cls)
        return cls._instance

    def __init__(self):
        self.status_flow = None
        self.status_sensor = None
        self.sensor = None
        self.last_error_time = None
        self.file_check()
        self.initialize_connections()

    def file_check(self):
        self.sensor_comport = readfile_value(9)
        self.flow_comport = readfile_value(10)

    def initialize_connections(self):
        threading.Thread(target=self.initialize_flow).start()
        threading.Thread(target=self.initialize_sensor).start()

    def initialize_flow(self):
        try:
            self.flow_1 = propar.instrument(f"{self.flow_comport}", 5)
            self.flow_2 = propar.instrument(f"{self.flow_comport}", 6)
            self.flow_3 = propar.instrument(f"{self.flow_comport}", 7)
            self.status_flow = True
        except Exception as e:
            self.status_flow = False
            print(f"Failed to initialize flow controllers: {e}")
            time.sleep(1)
            self.initialize_flow()

    def initialize_sensor(self):
        try:
            self.sensor = serial.Serial(f"{self.sensor_comport}", 19200, timeout=1)
            print("Sensor is connected")
            self.status_sensor = True
        except PermissionError:
            self.close_sensor()
            self.initialize_sensor()
        except Exception as e:
            if self.last_error_time is None or time.time() - self.last_error_time > 1:
                self.status_sensor = False
                print(f"Failed to initialize sensor: {e}")
                self.last_error_time = time.time()
            time.sleep(1)
            if self.sensor is None:
                self.initialize_sensor()

    def close_sensor(self):
        if self.sensor is not None:
            self.sensor.close()
            self.status_sensor = False
            print("Sensor closed successfully")
