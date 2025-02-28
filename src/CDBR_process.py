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
            sleep(10)
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
        data_CDBR_tool_manager.open_file()

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


def run_macro_and_send_message_CDBR(platform):
    """
    Chạy macro sao chép và dán vào ô tin nhắn gửi các huyện (Zalo/WhatsApp)

    :param platform: 'zalo' hoặc 'whatsapp'
    """
    if platform not in ["zalo", "whatsapp"]:
        print("Lỗi: Nền tảng không hợp lệ. Chỉ chấp nhận 'zalo' hoặc 'whatsapp'.")
        return False

    messaging_service = zalo if platform == "zalo" else whatsapp

    if not browser.is_browser_open():
        browser.start_browser(CHROME_PROFILE_CDBR_PATH)
    messaging_service.driver = browser.driver

    try:
        df = pd.read_excel(DATA_CONFIG_CDBR_PATH, sheet_name="Sheet1", header=0)

        # Xử lý dòng đầu tiên
        first_row = df.iloc[:1]
        process_group(messaging_service, first_row, num_macro=2, num_mess=2)

        # Xử lý các dòng còn lại
        other_rows = df.iloc[1:]
        for index, row in other_rows.iterrows():
            process_single_group(messaging_service, row, num_macro=5, num_mess=4)

        return True

    except Exception as e:
        print(f"Lỗi: {e}")
        return False

    finally:
        data_CDBR_tool_manager.save_file()
        data_CDBR_tool_manager.close_all_file()


def process_group(service, row, num_macro, num_mess):
    """
    Xử lý từng nhóm trong danh sách
    """
    data_CDBR_tool_manager.open_file()

    if isinstance(row, pd.DataFrame):  # Nếu row là DataFrame, lặp qua từng dòng
        for _, single_row in row.iterrows():
            process_single_group(service, single_row, num_macro, num_mess)
    else:  # Nếu row chỉ là một Series, xử lý trực tiếp
        process_single_group(service, row, num_macro, num_mess)


def process_single_group(service, row, num_macro, num_mess):
    """
    Xử lý từng nhóm (từng dòng riêng lẻ)
    """

    link = (
        str(row.get("Link group", "")).strip()
        if pd.notna(row.get("Link group"))
        else None
    )

    if not link:
        return

    macros = [
        str(row.get(f"macro {i}", "")).strip()
        for i in range(1, num_macro + 1)  # Fix để lấy đủ macro
        if pd.notna(row.get(f"macro {i}")) and row.get(f"macro {i}", "").strip()
    ]

    messages = [
        str(row.get(f"mess {i}", "")).strip()
        for i in range(1, num_mess + 1)  # Fix để lấy đủ message
        if pd.notna(row.get(f"mess {i}")) and row.get(f"mess {i}", "").strip()
    ]
    
    temp = service.find_group_name(link)
    retries = 0
    max_retries = 5

    while retries < max_retries:
        if temp:
            for macro, message in zip(macros, messages):
                try:
                    if not data_CDBR_tool_manager.run_macro(macro):
                        continue
                    sleep(5)

                    if service == whatsapp:
                        service.send_message_CDBR(message)

                    elif service == zalo:
                        zalo_message_status = service.send_message_CDBR(message)
                        if zalo_message_status == "sent":
                            print("✅ Tin nhắn đã gửi đi và được nhận!")
                        elif zalo_message_status == "timeout":
                            print("⏰ Tin nhắn quá thời gian gửi!")
                        else:
                            print("❌ Tin nhắn không được gửi đi!")

                except Exception as e:
                    print(f"Lỗi khi chạy macro {macro} và gửi tin nhắn: {e}")
            break  # Thoát vòng lặp nếu tìm thấy nhóm và xử lý xong

        else:
            retries += 1
            print(f"Không tìm thấy nhóm. Thử lại lần {retries}/{max_retries}...")
            sleep(10)
            service.access_zalo() if service == zalo else service.access_whatsapp()
            temp = service.find_group_name(link)

    if retries == max_retries and not temp:
        print(f"Đã thử {max_retries} lần nhưng không tìm thấy nhóm. Bỏ qua nhóm này.")


# full quá trình xử lý của CDBR: lấy dữ liệu - xử lý - gửi dữ liệu (WhatsApp)
def auto_process_CDBR():
    """
    Quy trình chạy WhatsApp/Zalo của CĐBR
    """
    date_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    print(f"Bắt đầu chạy tiến trình CDBR vào lúc {date_time}")

    try:
        # Bước 1: Lấy dữ liệu từ DB vào Excel
        if not getDB_to_excel_CDBR():
            print("Lỗi khi lấy dữ liệu, chờ đến tác vụ tiếp theo")
            return
        print("CĐBR: Lấy dữ liệu DB về Excel thành công!")

        # Bước 2: Xử lý dữ liệu trong Excel (tối đa 5 lần thử)
        for attempt in range(5):
            if excel_process_CDBR():
                print("CĐBR: Xử lý dữ liệu Excel thành công")
                break
            print(f"CĐBR: Xử lý dữ liệu Excel thất bại, thử lần {attempt + 1}")
        else:
            print("CĐBR: Xử lý dữ liệu Excel thất bại sau 5 lần thử")
            return

        # Bước 3: Gửi tin nhắn qua WhatsApp hoặc Zalo
        if SENDBY.upper() in ["WHATSAPP", "ZALO"]:
            if not run_macro_and_send_message_CDBR(SENDBY.lower()):
                print(f"Gửi tin nhắn qua {SENDBY} thất bại!!!")
                return
            print(f"Gửi tin nhắn qua {SENDBY} thành công!!!")
        else:
            print(f"CĐBR: Phương thức gửi '{SENDBY}' không được hỗ trợ")

        #tắt browser sau khi chạy
        if browser.is_browser_open():
            browser.close()
            
    except Exception as e:
        print(f"CĐBR - Lỗi trong quá trình xử lý: {e}")

    finally:
        excel = win32com.client.GetActiveObject("Excel.Application")

        for wb in excel.Workbooks:
            if wb is not None:
                data_CDBR_tool_manager.close_all_file()
                break
