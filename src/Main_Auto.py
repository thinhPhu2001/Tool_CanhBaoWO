from config import *
from database import *
from utils import *
from openVPN import *
from excel_handler import *
from browser import *

from CDBR_process import *
from Di_dong import *

from pynput.keyboard import Controller, Key
import schedule
from openpyxl import load_workbook
import sys


# thay đôi môi trường tiếng Việt
sys.stdout.reconfigure(encoding="utf-8")


# WhatsApp
def auto_combines():
    """
    quá trình chạy di động và CĐBR (WhatsApp)
    """
    try:
        auto_process_CDBR()
        print("")
        print("!!!!!!!!!!!!!!!!!!!!ĐÃ CHẠY CẢNH BÁO CĐBR !!!!!!!!!!!!!!")
        print("")

        try:
            auto_process_diDong()
            print()
            print("")
            print("!!!!!!!!!!!!!!!!!!!!ĐÃ CHẠY CẢNH BÁO DI ĐỘNG !!!!!!!!!!!!!!")
            print("")
            print("Đang chờ đến thời gian chạy tác vụ tiếp theo")

        except Exception as e:
            print(f"Lỗi ở di động: {e}")

    except Exception as e:
        print(f"Lỗi ở CDBR: {e}")


if __name__ == "__main__":

    # # Lên lịch chạy
    # schedule.every().day.at("06:30").do(auto_combines)
    # schedule.every().day.at("08:30").do(auto_process_CDBR)
    # schedule.every().day.at("10:30").do(auto_process_CDBR)
    # schedule.every().day.at("12:30").do(auto_combines)
    # schedule.every().day.at("14:00").do(auto_process_CDBR)
    # schedule.every().day.at("16:20").do(auto_combines)

    # print("Đang chờ đến thời gian chạy tác vụ tiếp theo")
    # while True:
    #     schedule.run_pending()
    #     sleep(3)

    # browser.start_browser(CHROME_PROFILE_CDBR_PATH)
    # whatsapp.driver = browser.driver
    # whatsapp.access_whatsapp()
    # whatsapp.find_name("test 4")
    # whatsapp.send_img(
    #     r"D:\2-Job\Viettel\project_thu_viec\A_Tool_GGSheet\data\excel\img\tth\tinh.jpg",
    # )

    # auto_process_diDong()
    # browser.start_browser(CHROME_PROFILE_DI_DONG_PATH)
    # whatsapp.driver = browser.driver
    # sleep(2)
    # phone_number = "0964741020"
    # message = "hello"
    # whatsapp.driver.get(
    #     f"https://web.whatsapp.com/send?phone={phone_number}&text={message}"
    # )
    # if whatsapp.send_Error_Notification(PHONE_NUMBER, "hello"):
    #     print("đã gửi nội dung")
    # else:
    #     print("chưa gửi nội dung")

    getDB_to_excel(DATA_GNOC_RAW_PATH)
