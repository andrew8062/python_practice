import pyautogui
import time

def open_RADR():
	secs_between_keys = 0.05
	pyautogui.hotkey('win', '2')
	time.sleep(1)
	pyautogui.hotkey('ctrl', 'f')

	pyautogui.typewrite('loki amd radr xls', interval=secs_between_keys)  # useful for entering text, newline is Enter
	pyautogui.hotkey('down')
	pyautogui.hotkey('enter')
def download_RADR():
	# time.sleep(10)
	x,y = downlaod_button_position = pyautogui.locateCenterOnScreen('RADR_download_button.png')
	time.sleep(0.5)
	pyautogui.doubleClick(x=x, y=y)
	print downlaod_button_position

if __name__ == '__main__':
	# open_RADR()
	download_RADR()

