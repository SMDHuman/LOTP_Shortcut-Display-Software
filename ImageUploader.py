from PIL import Image
from  PIL import ImageOps
import pyperclip
from serial import Serial, SerialException
from time import sleep
import win32com.client
import tkinter as tk
import customtkinter as Ctk
from tkinterdnd2 import TkinterDnD, DND_ALL

def get_path(event):
    print(event.data)

class Tk(Ctk.CTk, TkinterDnD.DnDWrapper):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.TkdndVersion = TkinterDnD._require(self)

class App(Tk):
	def __init__(self):
		super().__init__()

		self.width = 1000
		self.height = round(self.width * 9/16)
		self.title("BitBoard Logic IDE")
		self.geometry(f"{self.width}x{self.height}")
		self.protocol("WM_DELETE_WINDOW", self.on_closing)

		label = Ctk.CTkLabel(self, text="➕ \nDrag & Drop Here", corner_radius=10, fg_color="blue", wraplength=300)
		label.pack(expand=True, fill="both", padx=40, pady=40)

		entryWidget = tk.Entry(self)
		entryWidget.pack(padx=5, pady=5)

		label.drop_target_register(DND_ALL)
		label.dnd_bind("<<Drop>>", get_path)

	def loadSettings(self):
		pass

	def saveSettings(self):
		pass

	def on_closing(self):
	    self.saveSettings()
	    self.destroy()

KEY_F1 = 0xC2
KEY_F2 = 0xC3
KEY_F3 = 0xC4
KEY_F4 = 0xC5
KEY_F5 = 0xC6
KEY_F6 = 0xC7
KEY_F7 = 0xC8
KEY_F8 = 0xC9
KEY_F9 = 0xCA
KEY_F10 = 0xCB
KEY_F11 = 0xCC
KEY_F12 = 0xCD
KEY_F13 = 0xF0
KEY_F14 = 0xF1
KEY_F15 = 0xF2
KEY_F16 = 0xF3
KEY_F17 = 0xF4
KEY_F18 = 0xF5
KEY_F19 = 0xF6
KEY_F20 = 0xF7
KEY_F21 = 0xF8
KEY_F22 = 0xF9
KEY_F23 = 0xFA
KEY_F24 = 0xFB

port = []
ports = {}
wmi = win32com.client.GetObject("winmgmts:")
for serial in wmi.InstancesOf("Win32_SerialPort"):
	port = serial.Name.split(" (")
	port[1] = port[1][:-1]
	ports[port[0]] = port[1]

if('USB Seri Cihaz' in ports):
	com = Serial(ports['USB Seri Cihaz'], 9600);

def color565(red, green, blue, *args):
	if(len(args) > 0 and args[0] < 127):
		red, green, blue = 255-args[0], 255-args[0], 255-args[0]
	return ((red & 0xF8) << 8) | ((green & 0xFC) << 3) | (blue >> 3)

def sendImage(path, sector):
	img = Image.open(path)
	fileName = img.filename.split(".")[0]
	img = img.resize([16, 16])     
	img = ImageOps.flip(img)      
	img = ImageOps.mirror(img)                         

	buffer = []
	for y in range(img.height):
		for x in range(img.width):
			color = color565(*img.getpixel((x, y)))
			buffer.append(color >> 8)
			buffer.append(color & 0xff)

	com.write(bytearray([1]))
	com.write(bytearray([sector]))
	com.write(bytearray(buffer))
	com.read(1)

def sendKey(sector, key):
	if(type(key[0]) == str): key[0] = ord(key[0])
	if(type(key[1]) == str): key[1] = ord(key[1])

	com.write(bytearray([2]))
	sleep(0.1)
	com.write(bytearray([sector]))
	sleep(0.1)
	com.write(bytearray([key[0]]))
	sleep(0.1)
	com.write(bytearray([key[1]]))

ALT = 130
CTRL = 128
SHIFT = 129

keys = [ALT, "1", 
		ALT, "2",
		ALT, "3",
		ALT, "4",
		ALT, "5", 
		ALT, "6", 
		ALT, "7",
		ALT, "8", 
		0, 0, 
		0, 0, 
		0, 0, 
		0, 0]

if __name__ == "__main__":
	app = App()
	app.mainloop()

for i in range(12):
	sendKey(i, keys[i*2:i*2+2])
	print("Keys:", i)

for i in range(12):        
	sendImage(f"images/{i+1}.png", i)
	print("Images:", i)

com.close()