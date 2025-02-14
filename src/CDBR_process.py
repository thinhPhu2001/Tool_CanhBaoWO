from config import *
from database import *
from utils import *
from openVPN import *
from excel_handler import *
from browser import *

from pynput.keyboard import Controller, Key
import schedule
import sys

# thay đôi môi trường tiếng Việt
sys.stdout.reconfigure(encoding="utf-8")

# excel variable
# |-- CĐBR
data_CDBR_tool_manager = ExcelManager(DATA_TOOL_MANAGEMENT_CDBR_PATH)

module_CDBR_ZALO = [
    (
        "BDG-CĐBR - Bến Cát - Bàu Bàng",
        '//*[@id="group-item-g3178603850816266117"]',
        "Handle_Data.pic_TKM_BCT_BBG",
        "Handle_Data.pic_PAKH_BCT_BBG",
        "Handle_Data.pic_NV_BCT_BBG",
    ),
    (
        "BDG-CĐBR - Dầu Tiếng",
        '//*[@id="group-item-g7285665710449611138"]',
        "Handle_Data.pic_TKM_DTG",
        "Handle_Data.pic_PAKH_DTG",
        "Handle_Data.pic_NV_DTG",
    ),
    (
        "BDG-CĐBR - Dĩ An",
        '//*[@id="group-item-g5411862601516187267"]',
        "Handle_Data.pic_TKM_DAN",
        "Handle_Data.pic_PAKH_DAN",
        "Handle_Data.pic_NV_DAN",
    ),
    (
        "BDG-CĐBR - Phú Giáo",
        '//*[@id="group-item-g1989074657098575363"]',
        "Handle_Data.pic_TKM_PGO",
        "Handle_Data.pic_PAKH_PGO",
        "Handle_Data.pic_NV_PGO",
    ),
    (
        "BDG-CĐBR - Tân Uyên - Bắc Tân Uyên",
        '//*[@id="group-item-g7717459319638711680"]',
        "Handle_Data.pic_TKM_TUN_BTU",
        "Handle_Data.pic_PAKH_TUN_BTU",
        "Handle_Data.pic_NV_TUN_BTU",
    ),
    (
        "BDG-CĐBR - Thủ Dầu Một",
        '//*[@id="group-item-g8365703888020947841"]',
        "Handle_Data.pic_TKM_TDM",
        "Handle_Data.pic_PAKH_TDM",
        "Handle_Data.pic_NV_TDM",
    ),
    (
        "BDG-CĐBR - Thuận An",
        '//*[@id="group-item-g7464963621321691137"]',
        "Handle_Data.pic_TKM_TAN",
        "Handle_Data.pic_PAKH_TAN",
        "Handle_Data.pic_NV_TAN",
    ),
]

module_CDBR_ZALO_test = [
    (
        "BDG-CĐBR - Bến Cát - Bàu Bàng",
        "https://chat.zalo.me/?g=pgrplj867",
        "Handle_Data.pic_TKM_BRA_LDN",
        "Handle_Data.pic_PAKH_BRA_LDN",
        "Handle_Data.pic_NV_BRA_LDN",
    ),
    (
        "BDG-CĐBR - Dầu Tiếng",
        "https://chat.zalo.me/?g=hlcrpb220",
        "Handle_Data.pic_TKM_BRA_LDN",
        "Handle_Data.pic_PAKH_BRA_LDN",
        "Handle_Data.pic_NV_BRA_LDN",
    ),
    (
        "BDG-CĐBR - Dĩ An",
        "https://chat.zalo.me/?g=pgrplj867",
        "Handle_Data.pic_TKM_CDC",
        "Handle_Data.pic_PAKH_CDC",
        "Handle_Data.pic_NV_BRA_LDN",
    ),
]

# Module cho WhatsApp
module = [
    (
        "Dây máy CTH-CTA",
        "https://chat.whatsapp.com/DvxLt1cfq5w4mafYT1YB5n",
        "Handle_Data.pic_TKM_CTH_CTHA",
        "Handle_Data.pic_PAKH_CTH_CTHA",
        "Handle_Data.pic_SOC2_CTH_CTHA",
        "Handle_Data.pic_NV_CTH_CTHA",
        "Cảnh báo tồn TKM",
        "Cảnh báo tồn PAKH",
        "Cảnh báo tồn SOC2",
        "Cảnh báo tồn tác vụ mức FT",
    ),
    (
        "Dây máy LMY",
        "https://chat.whatsapp.com/C2Xjm2UN6DS6TgkvvR6jCp",
        "Handle_Data.pic_TKM_LMY",
        "Handle_Data.pic_PAKH_LMY",
        "Handle_Data.pic_SOC2_LMY",
        "Handle_Data.pic_NV_LMY",
        "Cảnh báo tồn TKM",
        "Cảnh báo tồn PAKH",
        "Cảnh báo tồn SOC2",
        "Cảnh báo tồn tác vụ mức FT",
    ),
    (
        "KH-BC COĐINH PHP-NBY",
        "https://chat.whatsapp.com/BRCuJqieBgg2QChRxEEkxp",
        "Handle_Data.pic_TKM_NBY_PHP",
        "Handle_Data.pic_PAKH_NBY_PHP",
        "Handle_Data.pic_SOC2_NBY_PHP",
        "Handle_Data.pic_NV_NBY_PHP",
        "Cảnh báo tồn TKM",
        "Cảnh báo tồn PAKH",
        "Cảnh báo tồn SOC2",
        "Cảnh báo tồn tác vụ mức FT",
    ),
    (
        "VTH-VTY dây máy",
        "https://chat.whatsapp.com/Bb2dxwQy3GxBI7E8WU8JOh",
        "Handle_Data.pic_TKM_VTH_VTY",
        "Handle_Data.pic_PAKH_VTH_VTY",
        "Handle_Data.pic_SOC2_VTH_VTY",
        "Handle_Data.pic_NV_VTH_VTY",
        "Cảnh báo tồn TKM",
        "Cảnh báo tồn PAKH",
        "Cảnh báo tồn SOC2",
        "Cảnh báo tồn tác vụ mức FT",
    )
]

# browser variable
browser = BrowserManager()
whatsapp = WhatsAppBot()
outlook = OutLookBot()
zalo = ZaloBot()


# lấy dữ liệu DB lưu xuống excel
def getDB_to_excel_CDBR():
    try:
        # Kiểm tra kết nối OpenVPN
        if not on_openvpn():
            print("Warning: Kết nối OpenVPN thất bại.")
            return False

        time.sleep(5)  # Chờ VPN ổn định

        # Kết nối và lấy dữ liệu từ database
        if not connect_to_Db_CDBR():
            return False

        time.sleep(3)
        return True  # Thành công

    except Exception as e:
        print(f"Lỗi khi lấy dữ liệu: {e}")
        return False

    finally:
        # Đảm bảo tắt OpenVPN dù xảy ra lỗi hay không
        off_openvpn()


# Xử lý dữ liệu excel (Run macro)
def excel_process_CDBR():
    """
    QUÁ TRÌNH THỰC HIỆN EXCEL CỦA CDBR TRƯỚC KHI GỬI
    """
    try:
        # Mở Excel và thực hiện các macro xử lý dữ liệu
        if data_CDBR_tool_manager.open_file():
            print(f"Mở file {DATA_TOOL_MANAGEMENT_CDBR_PATH} thành công")
        else:
            print(f"Mở file {DATA_TOOL_MANAGEMENT_CDBR_PATH} thất bại")
            return False

        macro_name = [
            "Handle_Data.XoaDuLieu",
            "Handle_Data.DanDuLieu",
            "Handle_Data.rowHeight",
            "Handle_Data.autoFillFormulas",
            "Handle_Data.SortColumn_TKM",
            "Handle_Data.SortColumn_PAKH",
            "Handle_Data.SaveFile",
        ]

        for macro in macro_name:
            if data_CDBR_tool_manager.run_macro(macro):
                print(f"Macro {macro} executed successfully!")
            else:
                print(f"Macro {macro} failed to execute.")
                return False  # Nếu bất kỳ macro nào không chạy thành công, thoát ngay

        print("Tất cả các macro đã chạy thành công")
        return True

    except Exception as e:
        print(f"Không thể mở hoặc xử lý file Excel: {e}")
        return False

    finally:
        try:
            data_CDBR_tool_manager.save_file()  # Lưu file nếu có thay đổi
            data_CDBR_tool_manager.close_all_file()  # Đảm bảo đóng tất cả các file
            print("Đã lưu và đóng file thành công")
        except Exception as e:
            print(f"Không thể lưu hoặc đóng file: {e}")


# chạy hàm macro sao chép và dán vào ô tin nhắn gửi các huyện (ZALO)
def run_macro_and_send_message_CDBR_ZALO():
    """
    chạy hàm macro sao chép và dán vào ô tin nhắn gửi các huyện
    """
    browser.start_browser(CHROME_PROFILE_CDBR_PATH)
    zalo.driver = browser.driver

    try:
        # mở excel và run các macro xử lý dữ liệu
        data_CDBR_tool_manager.open_file()
        print(f"Mở file {DATA_TOOL_MANAGEMENT_CDBR_PATH} thành công")

        for group, link, macro1, macro2, macro3 in module_CDBR_ZALO_test:
            zalo.find_group_name(link)
            try:
                data_CDBR_tool_manager.run_macro(macro1)
                print("Run macro thành công")
                zalo.send_message_CDBR("Cảnh báo tồn TKM")

            except Exception as e:
                print(f"CĐBR: Lỗi xảy ra trong quá trình gửi tin nhắn: {e}")

            try:
                data_CDBR_tool_manager.run_macro(macro2)
                print("Run macro thành công")
                zalo.send_message_CDBR("Cảnh báo tồn PAKH")

            except Exception as e:
                print(f"CĐBR: Lỗi xảy ra trong quá trình gửi tin nhắn: {e}")

            try:
                data_CDBR_tool_manager.run_macro(macro3)
                print("Run macro thành công")
                zalo.send_message_CDBR("Cảnh báo tồn NV")

            except Exception as e:
                print(f"CĐBR: Lỗi xảy ra trong quá trình gửi tin nhắn: {e}")

        # lưu và đóng file excel
        data_CDBR_tool_manager.save_file()
        data_CDBR_tool_manager.close_all_file()
        browser.close()
        return True

    except Exception as e:
        print(f"không thể mở file excel: {e}")
        return False


# chạy hàm macro sao chép và dán vào ô tin nhắn gửi các huyện (WhatsApp)
def run_macro_and_send_message_CDBR_WsA():
    """
    chạy hàm macro sao chép và dán vào ô tin nhắn gửi các huyện (WHatsApp)
    """
    link_KTDayMay = "https://chat.whatsapp.com/8DYotLigdw19snJnSXNKKn"
    browser.start_browser(CHROME_PROFILE_CDBR_PATH)
    whatsapp.driver = browser.driver
    whatsapp.access_whatsapp()

    try:
        # mở excel và run các macro xử lý dữ liệu
        data_CDBR_tool_manager.open_file()
        print(f"Mở file {DATA_TOOL_MANAGEMENT_CDBR_PATH} thành công")

        temp = whatsapp.find_group_name(link_KTDayMay)
        retries = 0
        max_retries = 5
        while retries < max_retries:
            if temp:
                try:
                    date_clock = datetime.now().strftime("%H:%M:%S")
                    date_time = datetime.now().strftime("%d-%m-%Y")

                    CDBR_noti = f"Cảnh báo tồn CĐBR đến {date_clock} ngày {date_time}"

                    whatsapp.send_message(CDBR_noti)

                    data_CDBR_tool_manager.run_macro("Handle_data.pic_General")
                    print("Run macro thành công")
                    sleep(5)
                    whatsapp.send_message_CDBR("Cảnh báo tồn tác vụ mức cụm - huyện")
                    break

                except Exception as e:
                    print(f"CĐBR: Lỗi xảy ra trong quá trình gửi tin nhắn: {e}")

            else:
                retries += 1
                print(f"CĐBR: Lỗi xảy ra trong quá trình gửi tin nhắn: {e}")
                whatsapp.access_whatsapp()  # Hàm tải lại trang (giả định bạn có hàm này)
                temp = whatsapp.find_group_name(link_KTDayMay)  # Thử tìm lại nhóm

        if retries == max_retries and not temp:
            print(
                f"Đã thử {max_retries} lần nhưng không tìm thấy nhóm LAN KT Dây máy. Bỏ qua nhóm này."
            )

        for group, link, macro1, macro2, macro3, macro4, mess1, mess2, mess3, mess4 in module:
            temp = whatsapp.find_group_name(link)
            retries = 0
            max_retries = 5
            while retries < max_retries:
                if temp:

                    date_clock = datetime.now().strftime("%H:%M:%S")
                    date_time = datetime.now().strftime("%d-%m-%Y")

                    CDBR_noti = f"Cảnh báo tồn CĐBR đến {date_clock} ngày {date_time}"

                    whatsapp.send_message(CDBR_noti)
                    
                    try:
                        data_CDBR_tool_manager.run_macro(macro1)
                        print("Run macro thành công")
                        sleep(5)
                        whatsapp.send_message_CDBR(mess1)

                    except Exception as e:
                        print(f"CĐBR: Lỗi xảy ra trong quá trình gửi tin nhắn: {e}")

                    try:
                        data_CDBR_tool_manager.run_macro(macro2)
                        print("Run macro thành công")
                        sleep(5)
                        whatsapp.send_message_CDBR(mess2)

                    except Exception as e:
                        print(f"CĐBR: Lỗi xảy ra trong quá trình gửi tin nhắn: {e}")

                    try:
                        data_CDBR_tool_manager.run_macro(macro3)
                        print("Run macro thành công")
                        sleep(5)
                        whatsapp.send_message_CDBR(mess3)

                    except Exception as e:
                        print(f"CĐBR: Lỗi xảy ra trong quá trình gửi tin nhắn: {e}")

                    try:
                        data_CDBR_tool_manager.run_macro(macro4)
                        print("Run macro thành công")
                        sleep(5)
                        whatsapp.send_message_CDBR(mess4)

                    except Exception as e:
                        print(f"CĐBR: Lỗi xảy ra trong quá trình gửi tin nhắn: {e}")

                    break  # Thoát vòng lặp nếu tìm thấy nhóm và đã xử lý xong

                else:
                    retries += 1
                    print(
                        f"Không tìm thấy nhóm '{group}'. Thử lại lần {retries}/{max_retries}..."
                    )
                    whatsapp.access_whatsapp()  # Hàm tải lại trang (giả định bạn có hàm này)
                    temp = whatsapp.find_group_name(link)  # Thử tìm lại nhóm

            if retries == max_retries and not temp:
                print(
                    f"Đã thử {max_retries} lần nhưng không tìm thấy nhóm '{group}'. Bỏ qua nhóm này."
                )

        # lưu và đóng file excel
        data_CDBR_tool_manager.save_file()
        data_CDBR_tool_manager.close_all_file()
        browser.close()
        return True

    except Exception as e:
        print(f"không thể mở file excel: {e}")
        # lưu và đóng file excel
        data_CDBR_tool_manager.save_file()
        data_CDBR_tool_manager.close_all_file()
        browser.close()
        return False


# full quá trình xử lý của CDBR: lấy dữ liệu - xử lý - gửi dữ liệu (WhatsApp)
def auto_process_CDBR():
    """
    Quy trình chạy WhatsApp của CĐBR
    """
    try:
        if not getDB_to_excel_CDBR():
            print("Lỗi khi lấy dữ liệu, chờ đến tác vụ tiếp theo")
            return

        print("CĐBR: Lấy dữ liệu db về excel thành công!")

        try:
            for attempt in range(5):
                if excel_process_CDBR():
                    break
                print(f"CĐBR: Xử lý dữ liệu Excel thất bại, thử lần {attempt + 1}")
            else:
                print("CĐBR: Xử lý dữ liệu Excel thất bại sau 5 lần thử")
                return

            print("CĐBR: xử lý dữ liệu excel thành công")

            if SENDBY.upper() == "WHATSAPP":
                try:
                    if not run_macro_and_send_message_CDBR_WsA():
                        print("Gửi tin nhắn cho các huyện thất bại!!!")
                        return
                    print("gửi tin nhắn cho các huyện thành công!!!")

                except Exception as e:
                    print(f"CĐBR-Lỗi khi gửi tin nhắn: {e}")

            elif SENDBY.upper() == "ZALO":
                try:
                    if not run_macro_and_send_message_CDBR_ZALO():
                        print("Gửi tin nhắn cho các huyện thất bại!!!")
                        return
                    print("gửi tin nhắn cho các huyện thành công!!!")

                except Exception as e:
                    print(f"CĐBR-Lỗi khi gửi tin nhắn: {e}")

        except Exception as e:
            print(f"CĐBR-Lỗi khi xử lý dữ liệu excel: {e}")

    except Exception as e:
        print(f"CĐBR-Lỗi khi lấy dữ liệu: {e}")
