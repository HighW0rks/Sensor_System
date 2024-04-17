# Sensor Script Configuration GUI

The Saxon Sensor GUI is a graphical user interface (GUI) application designed for use with Saxon sensors and Bronkhorst flow controllers. This application enables users to monitor and control various parameters such as parts per million (PPM), temperature, and air pressure. It provides real-time data visualization and interactive controls for efficient management of the sensor and flow controller.

## Features
**- PPM Measurement:** The GUI accurately measures and displays the PPM value detected by the Saxon sensor.

**- Real-time Data Visualization:** Users can visualize sensor data through intuitive charts and graphs.

**- Flow Control:** The application allows users to control the flow of air or CO2 using Bronkhorst flow controllers. Sliders and entry boxes facilitate easy adjustment of flow rates.

**- Custom Configuration:** Customizable configurations are available for both the Saxon sensor and the charts, providing flexibility to adapt to different environments and requirements.

**- Script Execution:** Users can initiate scripts via a prominent start button within the main application. These scripts are editable using a built-in custom script editor, enabling users to tailor functionality to their specific needs.


## Usage
1. Launch the application by running main.py.
2. Configure the Saxon sensor and Bronkhorst flow controller settings as required.
3. Use the sliders or entry boxes to adjust flow rates manually.
4. Monitor PPM values and other sensor data in real-time through the GUI.
5. Execute custom scripts by clicking the start button and edit scripts using the built-in script editor.

## Library
```
pip install tkchart
```
```
pip install tkinter
```
```
pip install customtkinter
```
```
pip install pyserial
```
```
pip install openpyxl
```
```
pip install xlsxwriter
```
```
pip install bronkhorst_propar
```
