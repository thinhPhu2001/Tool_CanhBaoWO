from config import *
from database import *
from utils import *
from openVPN import *
from excel_handler import *
from browser import *
from GG_Sheet import *
from CDBR_process import *
from Di_dong import *

from pynput.keyboard import Controller, Key
import schedule
from openpyxl import load_workbook
import sys


# thay đôi môi trường tiếng Việt
sys.stdout.reconfigure(encoding="utf-8")


def title_end_process(name_process):
    print("")
    print(f"!!!!!!!!!!!!!!!!!!!! {name_process} !!!!!!!!!!!!!!")
    print("")


# WhatsApp
def auto_combines():
    """
    quá trình chạy di động và CĐBR (WhatsApp)
    """
    try:
        date_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        auto_process_diDong()
        title_end_process(f"CHẠY CẢNH BÁO DI ĐỘNG KẾT THÚC VÀO {date_time}")

        auto_process_CDBR()
        title_end_process(f"CHẠY CẢNH BÁO CDBR KẾT THÚC VÀO {date_time}")

    except Exception as e:
        print(f"Lỗi chạy: {e}")


if __name__ == "__main__":

    # Lên lịch chạy
    # schedule.every().day.at(time_shift["First"]["start"]).do(auto_process_CDBR) #6:30
    # schedule.every().day.at(time_shift["Second"]["start"]).do(auto_combines) #8:30
    # schedule.every().day.at(time_shift["Third"]["start"]).do(auto_process_CDBR) #10:30
    # schedule.every().day.at(time_shift["Fourth"]["start"]).do(auto_combines) #12:10
    # schedule.every().day.at(time_shift["Fifth"]["start"]).do(auto_process_CDBR) #14:20
    # schedule.every().day.at(time_shift["Sixth"]["start"]).do(auto_process_CDBR) #16:20
    # schedule.every().day.at(time_shift["Seventh"]["start"]).do(auto_process_CDBR) #19:20

    # print("Đang chờ đến thời gian chạy tác vụ tiếp theo")
    # while True:
    #     schedule.run_pending()
    #     sleep(3)
    
    # browser.start_browser(CHROME_PROFILE_DI_DONG_PATH)
    # gnoc.driver = browser.driver
    # sleep(1000)
    # on_openvpn()

    # auto_process_CDBR()
    # auto_process_diDong()
    # excel_transition_and_run_macro(data_tool_manager)
    push_data_GGsheet()


    # sleep(900)
    # data_tool_manager.open_file()
    # sleep(80)
    # data_tool_manager.save_file()
    # data_tool_manager.close_all_file()


    
   