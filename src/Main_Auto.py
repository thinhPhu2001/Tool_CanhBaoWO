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
        auto_process_CDBR()
        title_end_process("ĐÃ CHẠY CẢNH BÁO CĐBR")

        auto_process_diDong()
        title_end_process("ĐÃ CHẠY CẢNH BÁO DI ĐỘNG")

    except Exception as e:
        print(f"Lỗi chạy: {e}")


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

    # Bước 3: Gửi tin nhắn qua WhatsApp hoặc Zalo
    if SENDBY.upper() in ["WHATSAPP", "ZALO"]:
        if not run_macro_and_send_message_CDBR(SENDBY.lower()):
            print(f"Gửi tin nhắn qua {SENDBY} thất bại!!!")
        print(f"Gửi tin nhắn qua {SENDBY} thành công!!!")
    else:
        print(f"CĐBR: Phương thức gửi '{SENDBY}' không được hỗ trợ")
