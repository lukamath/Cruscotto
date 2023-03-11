from pywinauto.application import Application
import keyboard

def salvacode():
	app = Application().connect(path=r"C:/Program Files/GCTI/CCPulse+/CallCenter.exe")

	dlg_spec = app.window(title='CCPulse+ - [Covisian_Code_REP_00] - [InfoQueueIsolaEl_ABR - View 1]')
	dlg_spec.set_focus()  #this is working!!!
	keyboard.press_and_release("alt+f")
	keyboard.press_and_release("h")
	keyboard.write("code.html")
	keyboard.press_and_release("alt+s")
	keyboard.press_and_release("y")

def salvaprod():
	print("sti kaz prod")