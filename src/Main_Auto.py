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


# def shilf_time()
if __name__ == "__main__":

    # Lên lịch chạy
    schedule.every().day.at(time_shift["First"]["start"]).do(auto_combines)
    schedule.every().day.at(time_shift["Second"]["start"]).do(auto_process_CDBR)
    schedule.every().day.at(time_shift["Third"]["start"]).do(auto_process_CDBR)
    schedule.every().day.at(time_shift["Fourth"]["start"]).do(auto_combines)
    schedule.every().day.at(time_shift["Fifth"]["start"]).do(auto_process_CDBR)
    schedule.every().day.at(time_shift["Sixth"]["start"]).do(auto_combines)

    print("Đang chờ đến thời gian chạy tác vụ tiếp theo")
    while True:
        schedule.run_pending()
        sleep(3)

    # getDB_to_excel(DATA_GNOC_RAW_PATH)
    # on_openvpn()

    # auto_process_diDong()

