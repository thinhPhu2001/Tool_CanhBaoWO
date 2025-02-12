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

def test():
    on_openvpn()
    sleep(5)
    off_openvpn()

    browser.start_browser(CHROME_PROFILE_CDBR_PATH)

if __name__ == "__main__":

    # Lên lịch chạy
    # schedule.every().day.at("06:20").do(auto_combines)
    # schedule.every().day.at("08:26").do(auto_process_CDBR)
    # schedule.every().day.at("10:20").do(auto_process_CDBR)
    # schedule.every().day.at("12:00").do(auto_combines)
    # schedule.every().day.at("14:20").do(auto_process_CDBR)
    # schedule.every().day.at("16:20").do(auto_combines)
    # schedule.every().day.at("19:30").do(auto_process_CDBR)

    # print("Đang chờ đến thời gian chạy tác vụ tiếp theo")
    # while True:
    #     schedule.run_pending()
    #     sleep(10)

    # auto_combines()
    # browser.start_browser(CHROME_PROFILE_CDBR_PATH)
    # zalo.driver = browser.driver
    # sleep(10000)
    send_message_cnct_zalo()

    # whatsapp.driver = browser.driver
    # whatsapp.access_whatsapp()
    # whatsapp.find_group_name("https://chat.whatsapp.com/K8nUGdLAhKC9u3aqvwTauH")
    # sleep(10)

    # on_openvpn()
    # sleep(5)
    # off_openvpn()

    # browser.start_browser(CHROME_PROFILE_DI_DONG_PATH)
    # zalo.driver = browser.driver
    # # zalo.find_group_name('https://chat.zalo.me/?g=kywtxh405')
    # sleep(1000)

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
    #             data_gnoc_raw_manager, data_tool_manager
    # )

    # auto_combines()

    # run_macro_and_send_message_CDBR_ZALO()
    # data_CDBR_tool_manager.close_all_file()
    # link_KT = 'https://chat.zalo.me/?g=hlcrpb220'

   
    # temp = zalo.find_group_name(link_KT)
    # sleep(1000)
    # data_CDBR_tool_manager.open_file()

    # data_CDBR_tool_manager.run_macro("Handle_data.pic_General")
    # print("Run macro thành công")
    # sleep(5)
    # zalo.send_message_CDBR("Cảnh báo tồn tác vụ mức cụm - huyện")
    # data_CDBR_tool_manager.close_all_file()

    # end_opvn_application()
    # on_openvpn()
#    getDB_to_excel_CDBR()
#    excel_process_CDBR()

    # getDB_to_excel(DATA_GNOC_RAW_PATH)
#    excel_transition_and_run_macro(
#                     data_gnoc_raw_manager, data_tool_manager
#                 )

#    browser.start_browser(CHROME_PROFILE_CDBR_PATH)
#    zalo.driver = browser.driver
#    zalo.access_zalo()
#    zalo.find_group_name('https://chat.zalo.me/?g=hlcrpb220')
#    zalo.send_file_zalo(DATA_GNOC_RAW_PATH)

   