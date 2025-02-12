from config import *

import subprocess
import pyautogui
import pyotp
import time
from pynput.keyboard import Controller, Key
from time import sleep
import sys

sys.stdout.reconfigure(encoding="utf-8")

pyautogui.FAILSAFE = False

def get_otp(secret):
    totp = pyotp.TOTP(secret)
    return totp.now()


def connect_vpn(config_file, secret):

    command = [OPEN_VPN_PATH, "--config", config_file]
    try:
        subprocess.run(command, shell=True, check=True)
        print("Đã mở OpenVPN thành công")
    except subprocess.CalledProcessError as e:
        print(f"Lỗi khi kết nối OpenVPN: {e}")


def find_elemont_by_image(button_location):
    if button_location:
        # Lấy vị trí trung tâm của nút
        button_center = pyautogui.center(button_location)
        return button_center
    else:
        print("Không tìm thấy nút trên màn hình.")


def on_openvpn():
    try:
        connect_vpn(OPEN_VPN_CONFIG_PATH, OTP_SECRET)
        sleep(3)
        # Locate the "ON" button on the screen
        on_button_path = str(IMAGE_PATH / "on_button.png")
        try:
            on_button = pyautogui.locateOnScreen(on_button_path, confidence=0.8)
        except pyautogui.ImageNotFoundException:
            print("Không tìm thấy hình ảnh On button của OPVN")
            return False

        if on_button:
            try: 
                pyautogui.click(pyautogui.center(on_button))  # Click the button
                
                try:
                    pyautogui.typewrite(get_otp(OTP_SECRET))  # Enter OTP

                except Exception as e:
                    print(f"Lỗi khi cố gắng nhập password: {e}")    
                    return False

            except Exception as e:
                print(f"Lỗi khi cố gắng nhấn On button: {e}")
                return False

            keyboard = Controller()
            keyboard.press(Key.enter)  # Press Enter
            keyboard.release(Key.enter)
            sleep(2)
            return True  # Indicate success
        else:
            print("ON button not found.")
            return False

    except Exception as e:
        print(f"Error in on_openvpn: {e}")
        return False


def off_openvpn():
    try:
        connect_vpn(OPEN_VPN_CONFIG_PATH, OTP_SECRET)

        # Locate the "OFF" button on the screen
        off_button_path = str(IMAGE_PATH / "off_button.png")
        
        try:
            off_button = pyautogui.locateOnScreen(off_button_path, confidence=0.8)
        except pyautogui.ImageNotFoundException:
            print("Không tìm thấy hình ảnh Off button của OPVN")
            return False

        if off_button:
            try:
                pyautogui.click(pyautogui.center(off_button))  # Click the button
                return True  # Indicate success
            except Exception as e:
                print(f"Lỗi khi cố gắng nhấp vào Off button: {e}")
                return False

        else:
            print("OFF button not found.")
            return False

    except Exception as e:
        print(f"Error in off_openvpn: {e}")
        return False

def end_opvn_application():
    opvn_logo_img = str(IMAGE_PATH / "opvn_logo.png")
    x_tich_path = str(IMAGE_PATH / "X_opvn.png")

    opvn_application = None
    try:
        opvn_application = pyautogui.locateOnScreen(opvn_logo_img, confidence=0.8)
    except pyautogui.ImageNotFoundException:
        print("Không tìm thấy ứng dụng OPVN đang bật")
    
    if opvn_application is not None:
        try:
            pyautogui.click(pyautogui.center(opvn_application))  # Click the button
            
            sleep(2)

            X_tich = None
            try:
                X_tich = pyautogui.locateOnScreen(x_tich_path, confidence=0.8)
            except pyautogui.ImageNotFoundException:
                print("Không tìm thấy ứng dụng OPVN đang bật")

            if X_tich is not None:
                try:
                    x, y, width, height = X_tich  # Lấy tọa độ và kích thước của vùng tìm thấy
                    offset_x = 10  # Dịch sang phải 20 pixel (điều chỉnh giá trị này nếu cần)
                    pyautogui.click(x + width // 2 + offset_x, y + height // 2)

                except Exception as e:
                    print(f"Lỗi khi cố gắng ngừng (X) OPVN đang bật: {e}")  

        except Exception as e:
            print(f"Lỗi khi cố gắng nhấp vào OPVN đang bật: {e}")
    

