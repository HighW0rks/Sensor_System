import propar
import serial
import time
import threading

class Connection:
    _instance = None
    _lock = threading.Lock()

    def __new__(cls, *args, **kwargs):
        with cls._lock:
            if not cls._instance:
                cls._instance = super().__new__(cls)
        return cls._instance

    def __init__(self):
        self.status_flow_1 = None
        self.status_flow_2 = None
        self.status_flow_3 = None
        self.status_sensor = None
        self.sensor = None
        self.initialize_connections()

    def initialize_connections(self):
        threading.Thread(target=self.initialize_flow_1).start()
        threading.Thread(target=self.initialize_flow_2).start()
        threading.Thread(target=self.initialize_flow_3).start()
        threading.Thread(target=self.initialize_sensor).start()

    def initialize_flow_1(self):
        try:
            self.flow_1 = propar.instrument("COM3", channel=1)
            self.status_flow_1 = True
        except Exception as e:
            self.status_flow_1 = False
            print(f"Failed to initialize flow_1: {e}")
            time.sleep(1)
            self.initialize_flow_1()

    def initialize_flow_2(self):
        try:
            self.flow_2 = propar.instrument("COM3", channel=2)
            self.status_flow_2 = True
        except Exception as e:
            self.status_flow_2 = False
            print(f"Failed to initialize flow_2: {e}")
            time.sleep(1)
            self.initialize_flow_2()

    def initialize_flow_3(self):
        try:
            self.flow_3 = propar.instrument("COM3", channel=3)
            self.status_flow_3 = True
        except Exception as e:
            self.status_flow_3 = False
            print(f"Failed to initialize flow_3: {e}")
            time.sleep(1)
            self.initialize_flow_3()

    def initialize_sensor(self):
        try:
            self.sensor = serial.Serial("COM5", 19200, timeout=1)
            self.status_sensor = True
        except PermissionError:
            self.close_sensor()
            self.initialize_sensor()
        except Exception as e:
            self.status_sensor = False
            print(f"Failed to initialize sensor: {e}")
            time.sleep(1)
            if self.sensor == None:
                self.initialize_sensor()

    def close_sensor(self):
        if self.sensor is not None:
            self.sensor.close()
            self.status_sensor = False
            print("Sensor closed successfully")