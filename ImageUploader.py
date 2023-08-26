import os

try:
	from PIL import Image, ImageTk
	from  PIL import ImageOps
	import pyperclip
	from serial import Serial, SerialException
	from time import sleep
	import win32com.client
	import tkinter as tk
	import customtkinter as Ctk
	from tkinterdnd2 import TkinterDnD, DND_ALL
	import webcolors
except:
	os.system("pip install -r requirements.txt")
	from PIL import Image, ImageTk
	from  PIL import ImageOps
	import pyperclip
	from serial import Serial, SerialException
	from time import sleep
	import win32com.client
	import tkinter as tk
	import customtkinter as Ctk
	from tkinterdnd2 import TkinterDnD, DND_ALL
	import webcolors

class Tk(Ctk.CTk, TkinterDnD.DnDWrapper):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.TkdndVersion = TkinterDnD._require(self)

class App(Tk):
	def __init__(self):
		super().__init__()

		self.after(1000, self.update)
		self.ports = {}
		self.updatePorts()

		self.hoverButton = 0
		self.selectedButton = 0
		self.selectedBgColor = webcolors.name_to_hex("cyan")
		self.selectedBtColor = webcolors.name_to_hex("pink")

		self.width = 1000
		self.height = round(self.width * 9/16)
		self.title("BitBoard Logic IDE")
		self.geometry(f"{self.width}x{self.height}")
		self.protocol("WM_DELETE_WINDOW", self.on_closing)

		self.columnconfigure(0, weight=1)
		self.columnconfigure(1, weight=0)

		self.rowconfigure(0, weight=1)
		self.rowconfigure(1, weight=0)
		self.rowconfigure(2, weight=0)

		#---------------------------------------------------------
		self.buttonOptionsFrame = Ctk.CTkFrame(self)
		self.buttonOptionsFrame.grid(row=0, column=1, sticky="wens", padx = (5, 10), pady = (10, 5))

		#---------------------------------------------------------
		self.optionsFrame = Ctk.CTkFrame(self)
		self.optionsFrame.grid(row=1, column=1, sticky="wens", padx = (5, 10), pady = 5)

		self.backgroundLabel = Ctk.CTkLabel(self.optionsFrame, text="Background Color")
		self.backgroundLabel.grid(row = 0, column=0, padx=(10, 5))
		self.backgroundEntry = Ctk.CTkEntry(self.optionsFrame)
		self.backgroundEntry.bind("<Key>", self.backgroundEntered)
		self.backgroundEntry.insert(0, self.selectedBgColor)
		self.backgroundEntry.grid(row = 0, column=1, padx=(0, 10), sticky="ew")

		self.buttonColorLabel = Ctk.CTkLabel(self.optionsFrame, text="Buttons Color")
		self.buttonColorLabel.grid(row = 1, column=0, padx=(10, 5))
		self.buttonColorEntry = Ctk.CTkEntry(self.optionsFrame)
		self.buttonColorEntry.bind("<Key>", self.buttonColorEntered)
		self.buttonColorEntry.insert(0, self.selectedBtColor)
		self.buttonColorEntry.grid(row = 1, column=1, padx=(0, 10), sticky="ew")

		self.devicePortLabel = Ctk.CTkLabel(self.optionsFrame, text="Device Port")
		self.devicePortLabel.grid(row = 2, column=0, padx=(10, 5))
		self.devicePortEntry = Ctk.CTkOptionMenu(self.optionsFrame, values = [name + " " + self.ports[name] for name in self.ports.keys()])
		self.devicePortEntry.grid(row = 2, column=1, padx=(0, 10))

		#---------------------------------------------------------
		self.buttonsFrame = Ctk.CTkFrame(self, fg_color=self.selectedBgColor)
		self.buttonsFrame.grid(row=0, column=0, sticky="wens", padx = (10, 5), pady = (10, 5), rowspan=2)

		self.buttons = []
		self.buttonsOptionFrame = []
		self.keys1Combo = []
		self.keys2Combo = []
		self.keys3Combo = []
		self.delays1Entry = []
		self.delays2Entry = []
		self.delays3Entry = []
		for y in range(3):
			self.buttonsFrame.rowconfigure(y, weight=1)
			for x in range(4):
				self.buttonsFrame.columnconfigure(x, weight=1)
				png = Image.open(f"images/{x+y*4}.png").convert('RGBA')

				png = Image.open(f"images/{x+y*4}.png").convert('RGBA').resize((16, 16), 0)

				background = Image.new('RGBA', (16, 16), self.selectedBtColor)
				alpha_composite = Image.alpha_composite(background, png).resize((100, 100), 0)


				button = Ctk.CTkButton(self.buttonsFrame, text="", command=getattr(self, f"button{x+y*4}press"), fg_color=self.selectedBgColor, image=ImageTk.PhotoImage(alpha_composite))

				button.grid(row=y, column=x)
				button.drop_target_register(DND_ALL)
				button.dnd_bind("<<Drop>>", getattr(self, f"button{x+y*4}dnd"))
				self.buttonsOptionFrame.append(Ctk.CTkScrollableFrame(self.buttonOptionsFrame, label_text=f'Buton {x+y*4}'))

				self.keys1Combo.append(Ctk.CTkComboBox(self.buttonsOptionFrame[-1]))
				self.keys2Combo.append(Ctk.CTkComboBox(self.buttonsOptionFrame[-1]))
				self.keys3Combo.append(Ctk.CTkComboBox(self.buttonsOptionFrame[-1]))
				self.keys1Combo[-1].set("")
				self.keys2Combo[-1].set("")
				self.keys3Combo[-1].set("")
				self.keys1Combo[-1].grid(row = 0, column = 0)
				self.keys2Combo[-1].grid(row = 2, column = 0)
				self.keys3Combo[-1].grid(row = 4, column = 0)

				self.delays1Entry.append(Ctk.CTkEntry(self.buttonsOptionFrame[-1]))
				self.delays2Entry.append(Ctk.CTkEntry(self.buttonsOptionFrame[-1]))
				self.delays3Entry.append(Ctk.CTkEntry(self.buttonsOptionFrame[-1]))
				self.delays1Entry[-1].insert(0, "0")
				self.delays2Entry[-1].insert(0, "0")
				self.delays3Entry[-1].insert(0, "0")
				self.delays1Entry[-1].grid(row = 1, column = 1)
				self.delays2Entry[-1].grid(row = 3, column = 1)
				self.delays3Entry[-1].grid(row = 5, column = 1)

				self.buttons.append(button)

		#---------------------------------------------------------
		self.functionsFrame = Ctk.CTkFrame(self)
		self.functionsFrame.grid(row=2, column=0, sticky="wens", padx = (5, 10), pady = 5, columnspan=2)

		for i in range(4):
			self.functionsFrame.columnconfigure(i, weight=1)

		self.updateKeysButton = Ctk.CTkButton(self.functionsFrame, command=self.uploadKeys, text = "Update Keys", height=100, font = Ctk.CTkFont(family='Helvetica', size=36, weight='bold'))
		self.updateKeysButton.grid(row=0, column=0, padx=10, pady=10, rowspan = 3)
		self.updateImagesButton = Ctk.CTkButton(self.functionsFrame, command=self.uploadImages, text = "Update Images", height=100, font = Ctk.CTkFont(family='Helvetica', size=36, weight='bold'))
		self.updateImagesButton.grid(row=0, column=1, padx=10, pady=10, rowspan = 3)
		self.updateAllButton = Ctk.CTkButton(self.functionsFrame, command=self.uploadAll, text="Update All", height=100, font = Ctk.CTkFont(family='Helvetica', size=36, weight='bold'))
		self.updateAllButton.grid(row=0, column = 2, padx=10, pady=10, rowspan = 3)

		self.uploadLabel = Ctk.CTkLabel(self.functionsFrame, text="Upload In Progress", font = Ctk.CTkFont(family='Helvetica', size=20))
		#self.uploadLabel.grid(row = 1, column = 3, padx=10, pady=(10, 0))
		self.uploadProgressBar = Ctk.CTkProgressBar(self.functionsFrame)
		self.uploadProgressBar.set(0)
		#self.uploadProgressBar.grid(row = 2, column=3, padx=10, pady=(0, 10))

	def backgroundEntered(self, event):

		if(event.keycode == 13):
			entry = self.backgroundEntry.get()
			oldColor = self.selectedBgColor
			try:
				if(entry[0] == "#"):
					self.selectedBgColor = entry
				elif("," in entry):
					self.selectedBgColor = webcolors.rgb_to_hex([int(i) for i in entry.replace(" ", "").split(",")])
				else:
					self.selectedBgColor = webcolors.name_to_hex(entry)

				self.buttonsFrame.configure(fg_color = self.selectedBgColor)
				for button in self.buttons:
					button.configure(fg_color = self.selectedBgColor)
			except:
				self.selectedBgColor = oldColor
				self.backgroundEntry.delete(0, tk.END)

				self.buttonsFrame.configure(fg_color = self.selectedBgColor)
				for button in self.buttons:
					button.configure(fg_color = self.selectedBgColor)


	def buttonColorEntered(self, event):
		if(event.keycode == 13):
			entry = self.buttonColorEntry.get()
			oldColor = self.selectedBtColor
			try:
				if(entry[0] == "#"):
					self.selectedBtColor = entry
				elif("," in entry):
					self.selectedBtColor = webcolors.rgb_to_hex([int(i) for i in entry.replace(" ", "").split(",")])
				else:
					self.selectedBtColor = webcolors.name_to_hex(entry)

				for i, button in enumerate(self.buttons):
					png = Image.open(f"images/{i}.png").convert('RGBA').resize((16, 16), 0)

					background = Image.new('RGBA', (16, 16), self.selectedBtColor)
					alpha_composite = Image.alpha_composite(background, png).resize((100, 100), 0)

					button.configure(image=ImageTk.PhotoImage(alpha_composite))
			except:
				self.selectedBtColor = oldColor
				self.buttonColorEntry.delete(0, tk.END)

				for i, button in enumerate(self.buttons):
					png = Image.open(f"images/{i}.png").convert('RGBA').resize((16, 16), 0)

					background = Image.new('RGBA', (16, 16), self.selectedBtColor)
					alpha_composite = Image.alpha_composite(background, png).resize((100, 100), 0)

					button.configure(image=ImageTk.PhotoImage(alpha_composite))

	def forgetLastButtonOptions(self):
		try:
			self.buttonsOptionFrame[self.selectedButton].pack_forget()
		except:
			print("hayat")

	def packCurrentSelectedButton(self):
		self.buttonsOptionFrame[self.selectedButton].pack(expand = True, fill = Ctk.BOTH, padx=10, pady=10)

	def button0press(self):
		self.forgetLastButtonOptions()
		self.selectedButton = 0
		self.packCurrentSelectedButton()
	def button1press(self):
		self.forgetLastButtonOptions()
		self.selectedButton = 1
		self.packCurrentSelectedButton()
	def button2press(self):
		self.forgetLastButtonOptions()
		self.selectedButton = 2
		self.packCurrentSelectedButton()
	def button3press(self):
		self.forgetLastButtonOptions()
		self.selectedButton = 3
		self.packCurrentSelectedButton()
	def button4press(self):
		self.forgetLastButtonOptions()
		self.selectedButton = 4
		self.packCurrentSelectedButton()
	def button5press(self):
		self.forgetLastButtonOptions()
		self.selectedButton = 5
		self.packCurrentSelectedButton()
	def button6press(self):
		self.forgetLastButtonOptions()
		self.selectedButton = 6
		self.packCurrentSelectedButton()
	def button7press(self):
		self.forgetLastButtonOptions()
		self.selectedButton = 7
		self.packCurrentSelectedButton()
	def button8press(self):
		self.forgetLastButtonOptions()
		self.selectedButton = 8
		self.packCurrentSelectedButton()
	def button9press(self):
		self.forgetLastButtonOptions()
		self.selectedButton = 9
		self.packCurrentSelectedButton()
	def button10press(self):
		self.forgetLastButtonOptions()
		self.selectedButton = 10
		self.packCurrentSelectedButton()
	def button11press(self):
		self.forgetLastButtonOptions()
		self.selectedButton = 11
		self.packCurrentSelectedButton()

	def button0dnd(self, event):
		self.hoverButton = 0
		self.getPathForButton(event)
	def button1dnd(self, event):
		self.hoverButton = 1
		self.getPathForButton(event)
	def button2dnd(self, event):
		self.hoverButton = 2
		self.getPathForButton(event)
	def button3dnd(self, event):
		self.hoverButton = 3
		self.getPathForButton(event)
	def button4dnd(self, event):
		self.hoverButton = 4
		self.getPathForButton(event)
	def button5dnd(self, event):
		self.hoverButton = 5
		self.getPathForButton(event)
	def button6dnd(self, event):
		self.hoverButton = 6
		self.getPathForButton(event)
	def button7dnd(self, event):
		self.hoverButton = 7
		self.getPathForButton(event)
	def button8dnd(self, event):
		self.hoverButton = 8
		self.getPathForButton(event)
	def button9dnd(self, event):
		self.hoverButton = 9
		self.getPathForButton(event)
	def button10dnd(self, event):
		self.hoverButton = 10
		self.getPathForButton(event)
	def button11dnd(self, event):
		self.hoverButton = 11
		self.getPathForButton(event)

	def getPathForButton(self, event):
		print(event.data)
		print(self.hoverButton)

		png = Image.open(event.data).convert('RGBA').resize((16, 16), 0)

		background = Image.new('RGBA', (16, 16), self.selectedBtColor)
		alpha_composite = Image.alpha_composite(background, png).resize((100, 100), 0)

		self.buttons[self.hoverButton].configure(image=ImageTk.PhotoImage(alpha_composite))
		png.save(f"images/{self.hoverButton}.png")

	def loadSettings(self):
		pass

	def saveSettings(self):
		pass

	def uploadKeys(self):
		pass

	def uploadImages(self):

		com = Serial(self.ports['USB Seri Cihaz'], 9600)
		for i in range(12):   
			png = Image.open(f"images/{11-i}.png").convert('RGBA').resize((16, 16), 0)   
			png = ImageOps.flip(png)      
			png = ImageOps.mirror(png) 

			background = Image.new('RGBA', (16, 16), self.selectedBtColor)
			alpha_composite = Image.alpha_composite(background, png)
			buffer = []
			for y in range(alpha_composite.height):
				for x in range(alpha_composite.width):
					color = color565(*alpha_composite.getpixel((x, y)))
					buffer.append(color >> 8)
					buffer.append(color & 0xff)

			com.write(bytearray([1]))
			com.write(bytearray([i]))
			com.write(bytearray(buffer))
			com.read(1)     
			print("Images:", i)

	def uploadAll(self):
		pass

	def updatePorts(self):
		ports = {}
		wmi = win32com.client.GetObject("winmgmts:")
		for serial in wmi.InstancesOf("Win32_SerialPort"):
			port = serial.Name.split(" (")
			port[1] = port[1][:-1]
			ports[port[0]] = port[1]
		if(self.ports != ports):
			self.ports = ports.copy()
			return(True)
		else:
			self.ports = ports.copy()
			return(False)

	def update(self):
		if(self.updatePorts()):
			print(1)
			self.devicePortEntry.configure(values = [name + " " + self.ports[name] for name in self.ports.keys()])

		self.after(500, self.update)

	def on_closing(self):
	    self.saveSettings()
	    self.destroy()
	    exit()

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