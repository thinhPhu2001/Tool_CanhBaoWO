from config import *

import subprocess
import pyautogui
import pyotp
import time
from pynput.keyboard import Controller, Key
from time import sleep
import sys

sys.stdout.reconfigure(encoding="utf-8")


def get_otp(secret):
    totp = pyotp.TOTP(secret)
    return totp.now()


def connect_vpn(config_file, secret):

    command = [OPEN_VPN_PATH, "--config", config_file]
    try:
        subprocess.run(command, check=True)
        return True

    except subprocess.CalledProcessError as e:
        print(f"Lỗi khi kết nối OpenVPN: {e}")
        return False

def connect_vpn_on(config_file, secret):

    command = [OPEN_VPN_PATH, "--config", config_file]
    try:
        process = subprocess.Popen(command, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        stdout, stderr = process.communicate()
        if process.returncode == 0:
            return True
        else:
            print(f"Lỗi khi kết nối OpenVPN: {stderr.decode()}")
            return False
    except Exception as e:
        print(f"Lỗi khi khởi động OpenVPN: {e}")
        return False
    
def find_elemont_by_image(button_location):
    if button_location:
        # Lấy vị trí trung tâm của nút
        button_center = pyautogui.center(button_location)
        return button_center
    else:
        print("Không tìm thấy nút trên màn hình.")


def stop_OVPN():
    try:
        # subprocess.run(["taskkill", "/F", "/IM", "openvpn.exe"], check=True)
        subprocess.run(["taskkill", "/F", "/IM", "OpenVPNConnect.exe"], check=True)
        print("OpenVPN đã tắt thành công.")

    except subprocess.CalledProcessError:
        print("Không tìm thấy OpenVPN đang chạy.")


def on_openvpn():
    try:
        if not connect_vpn(OPEN_VPN_CONFIG_PATH, OTP_SECRET):
            print("Mở app OVPN thất bại")
            return False
        print("Mở app OVPN thành công")

        sleep(5)
        # Locate the "ON" button on the screen
        on_button_path = str(IMAGE_PATH / "on_button.png")
        on_button = pyautogui.locateOnScreen(on_button_path, confidence=0.8)

        if on_button:
            pyautogui.click(pyautogui.center(on_button))  # Click the button
            # totp = pyotp.TOTP(OTP_SECRET)
            # print(
            #     f"OTP hiện tại: {totp.now()}",
            # )

            sleep(2)
            pyautogui.typewrite(get_otp(OTP_SECRET))  # Enter OTP
            keyboard = Controller()
            sleep(2)
            keyboard.press(Key.enter)  # Press Enter
            keyboard.release(Key.enter)
            sleep(2)

            off_button_path = str(IMAGE_PATH / "off_button.png")
            ovpn_status = False

            for _ in range(5):
                try:
                    off_button = pyautogui.locateOnScreen(
                        off_button_path, confidence=0.8
                    )
                    if off_button:
                        print("    ✔ Đăng nhập openvpn thành công!!!")
                        ovpn_status = True
                        break
                except pyautogui.ImageNotFoundException:
                    sleep(3)
                    pass

            if not ovpn_status:
                print("    ❌ Đăng nhập openvpn thất bại!!!")
                stop_OVPN()

            return ovpn_status

        else:
            print("     ❌ ON button not found.")
            return False

    except Exception as e:
        print(f"    ❌ Loi khi co dang nhap OVPN: {e}")
        stop_OVPN()
        return False


def off_openvpn():
    try:
        connect_vpn(OPEN_VPN_PROFILE_PATH, OTP_SECRET)

        # Locate the "OFF" button on the screen
        off_button_path = str(IMAGE_PATH / "off_button.png")
        off_button = pyautogui.locateOnScreen(off_button_path, confidence=0.8)

        if off_button:
            pyautogui.click(pyautogui.center(off_button))  # Click the button

            on_button_path = str(IMAGE_PATH / "on_button.png")
            ovpn_status = True

            for _ in range(5):
                try:
                    on_button = pyautogui.locateOnScreen(on_button_path, confidence=0.8)
                    if on_button:
                        print("    Thoát openvpn thành công!!!\n")
                        ovpn_status = True
                        break

                except pyautogui.ImageNotFoundException:
                    sleep(3)
                    pass

            if not ovpn_status:
                print("    ❌ Thoát openvpn thất bại!!!")

            return ovpn_status

        else:
            print("    ❌ OFF button not found.")
            return False

    except Exception as e:
        print(f"    ❌ Loi khi co dang nhap OVPN: {e}")
        return False
