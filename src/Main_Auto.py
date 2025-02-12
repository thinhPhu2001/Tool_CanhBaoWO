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

    # Lên lịch chạy
    schedule.every().day.at("06:30").do(auto_combines)
    schedule.every().day.at("08:20").do(auto_process_CDBR)
    schedule.every().day.at("10:20").do(auto_process_CDBR)
    schedule.every().day.at("12:30").do(auto_combines)
    schedule.every().day.at("14:20").do(auto_process_CDBR)
    schedule.every().day.at("16:25").do(auto_combines)
    schedule.every().day.at("19:30").do(auto_combines)

    print("Đang chờ đến thời gian chạy tác vụ tiếp theo")
    while True:
        schedule.run_pending()
        sleep(3)

    # auto_combines()
    # browser.start_browser(CHROME_PROFILE_CDBR_PATH)
    # zalo.driver = browser.driver
    # zalo.find_group_name("https://chat.zalo.me/?g=mfgjms759")
    # sleep(10000)
    # whatsapp.driver = browser.driver
    # whatsapp.access_whatsapp()
    # whatsapp.find_group_name("https://chat.whatsapp.com/K8nUGdLAhKC9u3aqvwTauH")
    # sleep(10)

    # on_openvpn()
    # sleep(10)

    # browser.start_browser(CHROME_PROFILE_CDBR_PATH)
    # sleep(5000)

    # data_tool_manager.open_file()
    # data_tool_manager.run_macro("Module2.len_cum_doi")
    # data_tool_manager.save_file()
    # data_tool_manager.close_all_file()
    # auto_combines()

    # lấy dữ liệu di động
    # getDB_to_excel(DATA_GNOC_RAW_PATH)
    # excel_transition_and_run_macro(
    #                 data_gnoc_raw_manager, data_tool_manager
    #             )
    

