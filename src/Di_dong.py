from config import *
from database import *
from utils import *
from openVPN import *
from excel_handler import *
from browser import *
from GG_Sheet import *

import shutil
from sqlalchemy import text
from pynput.keyboard import Controller, Key
import schedule
import sys

# thay đôi môi trường tiếng Việt
sys.stdout.reconfigure(encoding="utf-8")

# excel variable
# |-- DI ĐỘNG
data_gnoc_raw_manager = ExcelManager(DATA_GNOC_RAW_PATH)
data_tool_manager = ExcelManager(DATA_TOOL_MANAGEMENT_PATH)
data_diDong_chatBot = ExcelManager(DATA_DIDONG_ChatBot_PATH)
# browser variable
browser = BrowserManager()
whatsapp = WhatsAppBot()
outlook = OutLookBot()
zalo = ZaloBot()
gnoc = GnocBot()

img_CNCT_path = CNCT_IMG_PATH / "tinh.jpg"  # đỉa chỉ gửi hình tỉnh


def getDB_to_excel(excel_gnoc_path):
    max_retries = 5
    retries = 0

    while retries < max_retries:
        try:
            # Kết nối OpenVPN
            if not on_openvpn():
                sleep(10)
                raise ConnectionError

            sleep(5)

            # Kết nối cơ sở dữ liệu
            connection = connect_to_db()

            if connection is None:
                raise ConnectionError("Kết nối Database thất bại.")
            print("Kết nối Database thành công")

            # Đọc file Excel "config", lấy dữ liệu từ "Sheet1"
            df = pd.read_excel(DATA_DiDong_CONFIG_PATH, sheet_name="Sheet1", header=0)
            # Đọc file Excel "Log", lấy dữ liệu từ "tinh"
            df2 = pd.read_excel(DATA_DiDong_CONFIG_PATH, sheet_name="tinh", header=0)

            # Lấy giá trị đầu tiên của cột "Tỉnh" (bỏ NaN nếu có)
            first_province = df["Tỉnh"].dropna().iloc[0]
            # Lọc dữ liệu theo giá trị đầu tiên của cột "Tỉnh"
            filtered_df = df2[df2["Tỉnh"] == first_province]

            # Lấy danh sách các nhóm từ cột "Nhóm" dựa trên dữ liệu đã lọc
            groups = filtered_df["Nhóm điều phối"].dropna().astype(str).tolist()

            # Chuyển danh sách các nhóm thành chuỗi SQL
            groups_sql = ", ".join(f"'{group}'" for group in groups)

            # groups_sql = ", ".join(df2["Nhóm điều phối"].dropna().astype(str))

            # Tạo câu truy vấn SQL
            query_pakh = f"""
                    SELECT *
                        FROM gnoc.gnoc_open_90d
                        WHERE `Nhóm điều phối` IN ({groups_sql})
                        AND `Hệ thống` IN (
                                        'TT',
                                        'CC_SCVT',
                                        'WFM-OTHERS',
                                        'ICMS',
                                        'WFM',
                                        'MR',
                                        'CO_DIEN',
                                        'WFM-FT',
                                        'VTNET_OS'
                                    )
                        AND `Loại công việc` NOT IN (
                            'VCC_VHKT_OS_VSMART',
                            'SAP_Nâng cấp trạm',
                            'VCC_QLTS Các công việc quản lý tài sản trong VHKT trong  OS',
                            'SAP_điều chuyển hạ cấp',
                            'SAP_điều chuyển nâng cấp',
                            'SAP_Hạ cấp thu hồi',
                            'SAP_Nâng cấp trạm ứng cứu thông tin'
                        )
                        AND `Nhân viên khởi tạo` NOT IN ('hongnt38');
                    """
            # Xuất dữ liệu ra Excel
            query_to_excel(connection, query_pakh, excel_gnoc_path)
            print("Lấy file database thành công!")
            return

        except ConnectionError as ce:
            print(f"Lỗi kết nối: {ce}")
        except Exception as e:
            print(f"Lỗi không xác định: {e}")
        finally:
            # Đảm bảo đóng kết nối database
            if "connection" in locals() and connection:
                connection.close()
                print("Đóng kết nối Database.")

            # Đảm bảo tắt VPN
            off_openvpn()

        # Tăng số lần thử và thời gian chờ
        retries += 1
        print(f"Thử lại lần thứ {retries} sau 5 giây...")
        sleep(5)

    print("Không thể hoàn thành tác vụ sau nhiều lần thử.")


def get_WO_dong():
    """
    Lấy dữ liệu đóng trên web gnoc
    """
    max_retries = 3
    retries = 0
    while retries < max_retries:
        try:
            delete_data_folder(DATA_WO_DONG_PATH)

            if not on_openvpn():
                raise ConnectionError

            if not browser.is_browser_open():
                browser.start_browser(CHROME_PROFILE_DI_DONG_PATH)

            gnoc.driver = browser.driver
            if not gnoc.access():
                return False

            sleep(5)

            if gnoc.get_WoDong():
                return True

        except ConnectionError as ce:
            print(f"Lỗi kết nối: {ce}")

        except Exception as e:
            print(f"lỗi khi lấy dữ liệu Wo đóng {e}")
            return False

        finally:
            browser.close()
            # Đảm bảo tắt VPN
            off_openvpn()

        # Tăng số lần thử và thời gian chờ
        retries += 1
        print(f"Thử lại lần thứ {retries} sau 5 giây...")
        sleep(5)

    print("không thành công sau nhiều lần thử!!!")


def push_data_GGsheet():
    try:
        data_diDong_chatBot.open_file()

        if not data_diDong_chatBot.run_macro("Module9.DanDuLieu"):
            return False

        if not data_diDong_chatBot.run_macro("Module9.DanDuLieu2"):
            return False

        print("chuyển dữ liệu thành công")

        # Danh sách macro cần chạy theo thứ tự
        macros = [
            ("Module9.RUN_MapTheoTram()", "Chạy số liệu cho MAP TRẠM"),
            ("Module9.RUN_MapMucNhanVien()", "Chạy số liệu cho NHÂN VIÊN"),
        ]

        for macro, description in macros:
            for attempt in range(3):
                if data_diDong_chatBot.run_macro(macro):
                    break
                print(
                    f"DATA chat bot: chạy lệnh {description} thất bại, thử lần {attempt + 1}"
                )
            else:
                print(
                    f"DATA chat bot: Xử lý VBA '{description}' thất bại sau 3 lần thử"
                )
                return False

        print("DATA chat bot: Đã xử lý VBA thành công!")

        # đẩy dữ liệu TRẠM lên GG SHEET
        excel_to_ggSheet(
            sheet_id=GG_SHEET_ID,
            new_worksheet_name="tram",
            excel_path=DATA_DIDONG_ChatBot_PATH,
            sheet_name_excel="MAP THEO TRAM",
        )

        # đẩy dữ liệu NHÂN VIÊN lên GG SHEET
        excel_to_ggSheet(
            sheet_id=GG_SHEET_ID,
            new_worksheet_name="FT",
            excel_path=DATA_DIDONG_ChatBot_PATH,
            sheet_name_excel="MAP Muc Nhan Vien",
        )

        return True

    except Exception as e:
        print(f"Lỗi khi đẩy dữ liệu lên ggSheet: {e}")
        return False

    finally:
        data_diDong_chatBot.save_file()
        data_diDong_chatBot.close_all_file()


def excel_transition_and_run_macro(excel_tool_manager: ExcelManager):
    """
    Chuyển dữ liệu từ file excel gnoc raw qua qua file tool để xử lý
    """
    try:
        excel_tool_manager.open_file()

        if not excel_tool_manager.run_macro("Module2.DanDuLieu"):
            return False

        print("chuyển dữ liệu thành công")

        # Danh sách macro cần chạy theo thứ tự
        macros = [
            ("Module1.PasteFormulasAndValuesTTH_DOC", "lọc TTH_DOC"),
            ("Module4.PasteFormulasAndValuesSimple_NhanVien", "lọc MapMucNhanVien"),
            ("Module2.pic_cum_huyen_loop", "xuất hình cụm huyện"),
            ("Module2.pic_TTH_doc", "xuất hình pic_TTH_doc"),
        ]

        for macro, description in macros:
            for attempt in range(3):
                if excel_tool_manager.run_macro(macro):
                    break
                print(
                    f"Di động: chạy lệnh {description} thất bại, thử lần {attempt + 1}"
                )
            else:
                print(f"Di động: Xử lý VBA '{description}' thất bại sau 3 lần thử")
                return False

        print("DI ĐỘNG: Đã xử lý VBA thành công!")
        return True

    except Exception as e:
        print(f"Lỗi khi xử lý VBA: {e}")
        return False

    finally:
        excel_tool_manager.save_file()
        excel_tool_manager.close_all_file()


# gửi thông báo cho cụm tỉnh (whatsApp)
def send_message_cnct_WSA():
    """
    gửi thông báo cấp tỉnh (whatsApp)
    """
    if not browser.is_browser_open():
        browser.start_browser(CHROME_PROFILE_DI_DONG_PATH)
    whatsapp.driver = browser.driver

    try:
        df = pd.read_excel(DATA_DiDong_CONFIG_PATH, sheet_name="Sheet1", header=0)

        for index, row in df.iloc[:1].iterrows():
            group_name = str(row["Tỉnh"]).strip() if not pd.isna(row["Tỉnh"]) else None
            link = (
                str(row["Link group"]).strip()
                if not pd.isna(row["Link group"])
                else None
            )
            message = row["Message"]
            img_name = row["img"]
            img_path = CNCT_IMG_PATH / f"{img_name}.jpg"

            if not img_name or not Path(img_path).is_file():
                print(
                    f"⚠️ Không tìm thấy ảnh {img_path}, bỏ qua {group_name}.", end="\n\n"
                )
                return False

            # nhập tên người dùng
            temp = whatsapp.find_group_name(link)
            retries = 0
            max_retries = 5
            while retries < max_retries:
                if temp:
                    send_mess_status = whatsapp.send_attached_img_message(
                        message, img_path
                    )

                    if send_mess_status:
                        sleep(5)
                        return True

                else:
                    retries += 1
                    print(
                        f"Không tìm thấy nhóm '{group_name}'. Thử lại lần {retries}/{max_retries}..."
                    )
                    whatsapp.access_whatsapp()  # Hàm tải lại trang (giả định bạn có hàm này)
                    temp = whatsapp.find_group_name(link)  # Thử tìm lại nhóm

    except Exception as e:
        print(f"khong tìm được tên {e}")
        return False


#  gửi thông báo cho huyện (WhatsApp)
def send_message_user_WSA():
    """
    Gửi thông báo cấp huyện - cho các cụm đội kỹ thuật (WhatsApp). Chỉ tag tên giám đốc huyện
    """
    # Khởi động trình duyệt
    if not browser.is_browser_open():
        browser.start_browser(CHROME_PROFILE_DI_DONG_PATH)
    whatsapp.driver = browser.driver
    success_list = []

    try:
        # Đọc file Excel "config", lấy dữ liệu từ "Sheet1"
        df = pd.read_excel(DATA_DiDong_CONFIG_PATH, sheet_name="Sheet1", header=0)

        for index, row in df.iloc[1:].iterrows():
            # Lấy thông tin từ từng dòng
            link = (
                str(row["Link group"]).strip()
                if not pd.isna(row["Link group"])
                else None
            )
            if not link:
                continue  # Bỏ qua nếu không có link nhóm

            cum_doi = row["Cụm đội KT"]
            ma_nhom = row["Mã nhóm"]
            img_name = row["img"]
            message = str(row["Message"])
            img_path = USER_IMG_PATH / f"{img_name}.jpg"

            if not img_name or not Path(img_path).is_file():
                print(f"⚠️ Không tìm thấy ảnh {img_path}, bỏ qua {cum_doi}.", end="\n\n")
                continue

            try:
                # Tìm nhóm
                retries = 0
                max_retries = 5

                while retries < max_retries:
                    group_found = whatsapp.find_group_name(link)

                    if group_found:
                        try:
                            # Gửi tin nhắn kèm hình ảnh
                            whatsapp.send_attached_img_message(message, img_path)
                            success_list.append(
                                cum_doi
                            )  # Thêm cụm đội vào danh sách thành công

                            print(f"Đã gửi thành công cho {cum_doi}.", end="\n\n")

                            break

                        except Exception as e:
                            print(f"Lỗi khi gửi tin nhắn cho {cum_doi}: {e}")

                    else:
                        retries += 1
                        print(
                            f"Không tìm thấy nhóm '{ma_nhom}'. Thử lại lần {retries}/{max_retries}..."
                        )
                        whatsapp.access_whatsapp()  # Tải lại trang
                        sleep(5)

            except Exception as e:
                print(f"Lỗi khi xử lý nhóm '{ma_nhom}': {e}")

    except Exception as e:
        print(f"Lỗi khi đọc file hoặc khởi tạo: {e}")

    finally:
        success_list_str = ", ".join(success_list)
        return success_list_str


# gửi thông báo cho tỉnh (ZALO)
def send_message_cnct_zalo():
    """
    gửi thông báo cấp tỉnh
    """
    if not browser.is_browser_open():
        browser.start_browser(CHROME_PROFILE_DI_DONG_PATH)

    zalo.driver = browser.driver

    try:
        # gửi tên nhóm theo cột cụm đội
        df = pd.read_excel(DATA_DiDong_CONFIG_PATH, sheet_name="Sheet1", header=0)
        for index, row in df.iloc[:1].iterrows():
            group_name = str(row["Tỉnh"]).strip() if not pd.isna(row["Tỉnh"]) else None
            link = (
                str(row["Link group"]).strip()
                if not pd.isna(row["Link group"])
                else None
            )

            message = row["Message"]
            img_name = row["img"]
            img_path = CNCT_IMG_PATH / f"{img_name}.jpg"

            zalo.find_group_name(link)

            status_message = zalo.send_attached_img_message(message, img_path, "all")

            if status_message == "sent":
                return "sent"
            elif status_message == "timeout":
                return "timeout"
            else:
                return "failed"

    except Exception as e:
        print(f"❌ Lỗi: {e}")
        return "failed"


# gửi thông báo cấp huyện (Zalo)
def send_message_user_with_TAG_zalo():
    """
    Gửi thông báo cấp huyện - các cụm đội kĩ thuật
    """
    if not browser.is_browser_open():
        browser.start_browser(CHROME_PROFILE_DI_DONG_PATH)
    zalo.driver = browser.driver

    success_list = []

    # gửi tên nhóm theo cột cụm đội
    df = pd.read_excel(DATA_DiDong_CONFIG_PATH, sheet_name="Sheet1", header=0)
    for index, row in df.iloc[1:].iterrows():
        link = (
            str(row["Link group"]).strip() if not pd.isna(row["Link group"]) else None
        )
        if not link:
            continue  # Bỏ qua nếu không có link nhóm

        cum_doi = row["Cụm đội KT"]
        message = str(row["Message"])
        img_name = row["img"]
        img_path = USER_IMG_PATH / f"{img_name}.jpg"

        try:
            # Kiểm tra xem hình ảnh có tồn tại không
            if not os.path.isfile(img_path):
                print(f"⚠️ Cảnh báo: Ảnh '{img_path}' không tồn tại! Bỏ qua hình ảnh.")
                continue  # Bỏ qua nếu không có hình

            zalo.find_group_name(link)
            sleep(2)

            # Gửi hình ảnh nếu có
            zalo.send_attached_img(img_path)
            sleep(1)

            # gửi tin nhắn và tag tên
            # Tìm ô message_box trong Zalo
            message_box = WebDriverWait(zalo.driver, 20).until(
                EC.presence_of_element_located((By.XPATH, XPATHS_ZALO["message_box"]))
            )
            message_box.click()
            message_box.send_keys(
                Keys.CONTROL, "a", Keys.BACKSPACE
            )  # xóa nội dung cũ nếu có
            message_box.send_keys(message)

            # tag ALL
            message_box.send_keys(" @")
            message_box.send_keys("all")
            sleep(1)
            message_box.send_keys(Keys.ARROW_DOWN)
            message_box.send_keys(Keys.ENTER)
            sleep(3)

            # Nhấn nút gửi tin nhắn
            send_button = WebDriverWait(zalo.driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, XPATHS_ZALO["send_button"]))
            )
            send_button.click()
            print(f"Tin nhắn gửi đến {cum_doi} thành công!!!")
            success_list.append(str(cum_doi))
            sleep(5)

        except Exception as e:
            print(f"⚠️ Lỗi khi gửi tin nhắn đến '{cum_doi}': {e}")

    return (
        ", ".join(success_list)
        if success_list
        else "Không có cụm đội nào nhận được tin nhắn"
    )


def process_whatsapp_notifications():
    """Hàm gửi tin nhắn cảnh báo qua whatsApp"""
    if browser.is_browser_open():
        browser.close()

    try:
        # gửi thông báo cấp CNCT
        status_process = send_message_cnct_WSA()

        if not status_process:
            print("Gửi tin nhắn cho Tỉnh thất bại!!!")

        else:
            print("Gửi tin nhắn cho Tỉnh thành công!!!")

    except Exception as e:
        print(f"Lỗi khi gửi tin nhắn tỉnh: {e}")

    try:
        # gửi thông báo cấp cụm đội kt
        list = send_message_user_WSA()
        print(f"Đã gửi tin nhắn cho các cụm huyện {list} thành công ")

    except Exception as e:
        print(f"Lỗi khi gửi tin nhắn huyện: {e}")
        print(e)


def process_zalo_notifications():
    """
    Hàm gửi tin nhắn cảnh báo qua Zalo
    """
    if browser.is_browser_open():
        browser.close()

    try:
        # gửi thông báo cấp CNCT
        status_process = send_message_cnct_zalo()

        if status_process == "sent":
            print("✅ Gửi tin nhắn cho tỉnh thành công!")
        elif status_process == "timeout":
            print("⏳ Vượt quá thời gian gửi tin nhắn (TỈNH)!")
        else:
            print("❌ Gửi tin nhắn cho tỉnh thất bại!")

    except Exception as e:
        print(f"⚠️ Lỗi trong quá trình gửi thông báo cho tỉnh: {e}")

    try:
        # gửi thông báo cấp cụm huyện có tag tên
        list = send_message_user_with_TAG_zalo()
        print(f"✅Đã gửi tin nhắn cho các cụm huyện {list} thành công ")

    except Exception as e:
        print(f"⚠️ Lỗi trong quá trình gửi thông báo cho các huyện: {e}")


def auto_process_diDong():
    """
    Quá trình gửi cảnh báo di động: lấy dữ liệu - xử lý dữ liệu - gửi tin nhắn
    """
    date_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    print(f"\nBắt đầu chạy tiến trình Di động vào lúc {date_time}")

    try:
        # Xóa dữ liệu cũ
        delete_data_folder(CNCT_IMG_PATH)
        delete_data_folder(USER_IMG_PATH)

        # lấy dữ liệu gnoc về xử lý
        getDB_to_excel(DATA_GNOC_RAW_PATH)

        for tempt in range(3):
            if get_WO_dong():
                print("Lấy dữ liệu WO dong thành công!")

                # Thử đẩy dữ liệu lên Google Sheet
                for attempt in range(3):
                    if push_data_GGsheet():
                        print("Dữ liệu đã được gửi lên Google Sheet thành công!")
                        break  # Nếu thành công, thoát vòng lặp nội
                    print(
                        f"Chat bot: Xử lý dữ liệu Excel thất bại, thử lần {attempt + 1}"
                    )
                else:
                    print(
                        "\nChat bot: Xử lý dữ liệu Excel thất bại sau 3 lần thử, không gửi dữ liệu lên GGsheet được\n"
                    )

                break  # Nếu lấy dữ liệu thành công, không chạy lại nữa

            print(f"Lấy dữ liệu WO dong thất bại, thử lần {tempt + 1}")
        else:
            print("\nLấy dữ liệu WO dong thất bại sau 3 lần thử\n")

        if not check_old_data_Didong(DATA_GNOC_RAW_PATH):
            print("Dữ liệu cũ, chờ đến tác vụ tiếp theo")
            return

        # # xử lý excel
        # for attempt in range(3):
        #     if excel_transition_and_run_macro(data_tool_manager):
        #         break
        #     print(f"DI động: Xử lý dữ liệu Excel thất bại, thử lần {attempt + 1}")

        # else:
        #     print("DI động: Xử lý dữ liệu Excel thất bại sau 3 lần thử")
        #     return

        # #gửi tin nhắn
        # if SENDBY.upper() == "WHATSAPP":
        #     process_whatsapp_notifications()

        # elif SENDBY.upper() == "ZALO":
        #     process_zalo_notifications()

        # if browser.is_browser_open():
        #     browser.close()

    except Exception as e:
        print(e)
