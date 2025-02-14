from config import *
from database import *
from utils import *
from openVPN import *
from excel_handler import *
from browser import *

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

# browser variable
browser = BrowserManager()
whatsapp = WhatsAppBot()
outlook = OutLookBot()
zalo = ZaloBot()

img_CNCT_path = CNCT_IMG_PATH / "tinh.jpg"  # đỉa chỉ gửi hình tỉnh


def getDB_to_excel(excel_gnoc_path):
    max_retries = 5
    retries = 0

    while retries < max_retries:
        try:
            # Kết nối OpenVPN
            if not on_openvpn():
                raise ConnectionError("Kết nối OpenVPN thất bại.")

            print("Kết nối OpenVPN thành công.")
            sleep(5)

            # Kết nối cơ sở dữ liệu
            connection = connect_to_db()
            if connection is None:
                raise ConnectionError("Kết nối Database thất bại.")
            print("Kết nối Database thành công")

            # Đọc file Excel, lấy dữ liệu từ sheet "Key"
            df = pd.read_excel(DATA_TOOL_MANAGEMENT_PATH, sheet_name="Key", header=0)
            df2 = pd.read_excel(DATA_TOOL_MANAGEMENT_PATH, sheet_name="tinh", header=0)

            # Lấy giá trị đầu tiên của cột "Tỉnh" (bỏ NaN nếu có)
            first_province = df["Tỉnh"].dropna().iloc[0]
            # Lọc dữ liệu theo giá trị đầu tiên của cột "Tỉnh"
            filtered_df = df2[df2["Tỉnh"] == first_province]

            # Lấy danh sách các nhóm từ cột "Nhóm" dựa trên dữ liệu đã lọc
            groups = filtered_df["Nhóm điều phối"].dropna().astype(str).tolist()

            # Chuyển danh sách các nhóm thành chuỗi SQL
            groups_sql = ", ".join(f"'{group}'" for group in groups)

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
                        );
                    """
            # print(query_pakh)
            print(f"Đang lấy dữ liệu của tỉnh {first_province}")
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
            print("Đã ngắt kết nối opvn")

        # Tăng số lần thử và thời gian chờ
        retries += 1
        print(f"Thử lại lần thứ {retries} sau 5 giây...")
        sleep(5)

    print("Không thể hoàn thành tác vụ sau nhiều lần thử.")


def excel_transition_and_run_macro(
    excel_gnoc_manager: ExcelManager, excel_tool_manager: ExcelManager
):
    """
    Chuyển dữ liệu từ file excel gnoc raw qua qua file tool để xử lý
    """
    try:
        excel_tool_manager.open_file()
        if not excel_tool_manager.run_macro("Module2.DanDuLieu"):
            return False

        print("chuyển dữ liệu thành công")

        try:
            # xử lý dữ liệu
            for attempt in range(3):
                if excel_tool_manager.run_macro(
                    "Module1.PasteFormulasAndValuesTTH_DOC"
                ):
                    break
                print(
                    f"Di động: chạy lệnh lọc TTH_DOC thất bại , thử lần {attempt + 1}"
                )
            else:
                print("Di động: Xử lý dữ liệu Excel thất bại sau 5 lần thử")
                return False

            for attempt in range(3):
                if excel_tool_manager.run_macro(
                    "Module4.PasteFormulasAndValuesSimple_NhanVien"
                ):
                    break
                print(
                    f"Di động: chạy lệnh lọc MapMucNhanVien thất bại , thử lần {attempt + 1}"
                )
            else:
                print("Di động: Xử lý dữ liệu Excel thất bại sau 5 lần thử")
                return False

            for attempt in range(3):
                if excel_tool_manager.run_macro("Module2.pic_cum_huyen_loop"):
                    break
                print(
                    f"Di động: chạy lệnh xuất hình cụm huyện thất bại , thử lần {attempt + 1}"
                )
            else:
                print("Di động: Xử lý dữ liệu Excel thất bại sau 5 lần thử")
                return False

            for attempt in range(3):
                if excel_tool_manager.run_macro("Module2.pic_TTH_doc"):
                    break
                print(
                    f"Di động: chạy lệnh xuất hình pic_TTH_doc thất bại , thử lần {attempt + 1}"
                )
            else:
                print("Di động: Xử lý dữ liệu Excel thất bại sau 5 lần thử")
                return False

            print("DI ĐỘNG: đã xử lý hàm vba thành công!!!")
            excel_tool_manager.save_file()
            excel_tool_manager.close_all_file()
            return True

        except Exception as e:
            print("Lỗi khi chạy hàm xử lý VBA: {e}")
            excel_tool_manager.save_file()
            excel_tool_manager.close_all_file()
            return False
    except Exception as e:
        print(f"Lỗi khi chuyển dữ liệu hoặc mở file: '{e}'")
        excel_tool_manager.save_file()
        excel_tool_manager.close_all_file()
        return False


# gửi thông báo cho cụm tỉnh (whatsApp)
def send_message_cnct():
    """
    gửi thông báo cấp tỉnh (whatsApp)
    """
    browser.start_browser(CHROME_PROFILE_DI_DONG_PATH)
    whatsapp.driver = browser.driver

    try:
        df = pd.read_excel(DATA_TOOL_MANAGEMENT_PATH, sheet_name="Key", header=0)

        for index, row in df.iloc[:1].iterrows():
            group_name = str(row["Tỉnh"]).strip() if not pd.isna(row["Tỉnh"]) else None
            link = (
                str(row["Link group"]).strip()
                if not pd.isna(row["Link group"])
                else None
            )
            message = row["Message"]
            img_path = CNCT_IMG_PATH / f"tinh.jpg"

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
                        sleep(10)
                        browser.close()
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
def send_message_user():
    """
    Gửi thông báo cấp huyện - cho các cụm đội kỹ thuật (WhatsApp). Chỉ tag tên giám đốc huyện
    """
    import os
    import pandas as pd
    from time import sleep

    # Khởi động trình duyệt
    browser.start_browser(CHROME_PROFILE_DI_DONG_PATH)
    whatsapp.driver = browser.driver
    success_list = []

    try:
        # Đọc file Excel
        df = pd.read_excel(DATA_TOOL_MANAGEMENT_PATH, sheet_name="Key", header=0)

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
            message = str(row["Message"])
            img_path = USER_IMG_PATH / f"{ma_nhom}.jpg"

            try:
                # Tìm nhóm
                temp = whatsapp.find_group_name(link)
                retries = 0
                max_retries = 5

                while retries < max_retries:
                    if temp:
                        try:
                            # Gửi tin nhắn kèm hình ảnh
                            whatsapp.send_attached_img_message(message, img_path)
                            success_list.append(
                                cum_doi
                            )  # Thêm cụm đội vào danh sách thành công
                            print(f"Đã gửi thành công cho {cum_doi}.")
                        except Exception as e:
                            print(f"Lỗi khi gửi tin nhắn cho {cum_doi}: {e}")
                        break
                    else:
                        retries += 1
                        print(
                            f"Không tìm thấy nhóm '{ma_nhom}'. Thử lại lần {retries}/{max_retries}..."
                        )
                        whatsapp.access_whatsapp()  # Tải lại trang
                        temp = whatsapp.find_group_name(link)  # Tìm lại nhóm

            except Exception as e:
                print(f"Lỗi khi xử lý nhóm '{ma_nhom}': {e}")

        sleep(5)

    except Exception as e:
        print(f"Lỗi khi đọc file hoặc khởi tạo: {e}")

    finally:
        browser.close()  # Đảm bảo trình duyệt được đóng
        success_list_str = ", ".join(success_list)
        return success_list_str


# gửi thông báo cho tỉnh (ZALO)
def send_message_cnct_zalo():
    """
    gửi thông báo cấp tỉnh
    Args:
        message (str): tin nhắn muốn gửi
        img_path (str): đường dẫn thư mục chứa ảnh muốn gửi
    """
    browser.start_browser(CHROME_PROFILE_DI_DONG_PATH)
    zalo.driver = browser.driver

    try:
        # gửi tên nhóm theo cột cụm đội
        df = pd.read_excel(DATA_TOOL_MANAGEMENT_PATH, sheet_name="Key", header=0)
        for index, row in df.iloc[:1].iterrows():
            group_name = str(row["Tỉnh"]).strip() if not pd.isna(row["Tỉnh"]) else None
            link = (
                str(row["Link group"]).strip()
                if not pd.isna(row["Link group"])
                else None
            )

            message = row["Message"]
            img_path = CNCT_IMG_PATH / f"tinh.jpg"

            zalo.find_group_name(link)
            try:
                if zalo.send_attached_img_message(message, img_path, "all"):
                    sleep(5)
                    browser.close()
                    return True
                else:
                    return False

            except Exception as e:
                print(f"khong gui duoc tin nhan {e}")
                return False

    except Exception as e:
        print(f"khong tìm thấy {group_name}: {e}")
        return False


# gửi thông báo cấp huyện (Zalo)
def send_message_user_with_TAG_zalo():
    """
    Gửi thông báo cấp huyện - các cụm đội kĩ thuật
    """
    browser.start_browser(CHROME_PROFILE_DI_DONG_PATH)
    zalo.driver = browser.driver
    success_list = []

    # gửi tên nhóm theo cột cụm đội
    df = pd.read_excel(DATA_TOOL_MANAGEMENT_PATH, sheet_name="Key", header=0)
    for index, row in df.iloc[1:].iterrows():
        link = (
            str(row["Link group"]).strip() if not pd.isna(row["Link group"]) else None
        )
        if not link:
            continue  # Bỏ qua nếu không có link nhóm

        cum_doi = row["Cụm đội KT"]
        ma_nhom = row["Mã nhóm"]
        message = str(row["Message"])
        img_path = USER_IMG_PATH / f"{ma_nhom}.jpg"

        try:
            zalo.find_group_name(link)
            sleep(2)

            try:
                zalo.send_attached_img(img_path)
                sleep(1)
                print("Gửi hình thành công!!!")

                try:
                    # gửi tin nhắn và tag tên
                    # Tìm ô message_box trong Zalo
                    message_box = WebDriverWait(zalo.driver, 20).until(
                        EC.presence_of_element_located(
                            (By.XPATH, XPATHS_ZALO["message_box"])
                        )
                    )
                    message_box.click()
                    print("Đã tìm thấy ô message_box.")
                    message_box.send_keys(message)
                    message_box.send_keys(" @")
                    message_box.send_keys("all")
                    sleep(1)
                    message_box.send_keys(Keys.ARROW_DOWN)
                    message_box.send_keys(Keys.ENTER)
                    sleep(3)
                    print("GÕ TIN NHẮN THÀNH CÔNG!!!")

                    try:
                        # Nhấn nút gửi tin nhắn
                        send_button = WebDriverWait(zalo.driver, 10).until(
                            EC.element_to_be_clickable(
                                (By.XPATH, XPATHS_ZALO["send_button"])
                            )
                        )
                        send_button.click()
                        print("Tin nhắn gửi thành công!!!")
                        success_list.append(str(cum_doi))
                        sleep(5)

                    except Exception as e:
                        print(f"Không gửi được tin nhắn {e}")

                except Exception as e:
                    print(f"Lỗi gửi tin nhắn: {e}")

            except Exception as e:
                print(f"Lỗi gửi hình: {e}")

        except Exception as e:
            print(f"khong tim được ô ser {e}")

    browser.close()
    try:
        success_list_str = ", ".join([str(item) for item in success_list])

    except Exception as e:
        print(e)

    return success_list_str


# gửi mail gồm hình ảnh và đính kèm file
def input_mail_outlook(folder_path, subject, context):
    """
    Gửi mail

    Args:
        folder_path (str): địa chỉ file đính kèm
        subject (str): tiêu đề mail
        context: nội dung mail
    """
    # khởi động trình duyệt
    browser = BrowserManager()
    browser.start_browser(CHROME_PROFILE_DI_DONG_PATH)
    outlook = OutLookBot()
    outlook.driver = browser.driver
    outlook.access_outlook()
    try:
        new_mail_button = outlook.find_new_mail_button()
        try:
            new_mail_button.click()
            print("mở mail mới thành công")
            try:
                # nhập tên người nhận
                df = pd.read_excel(folder_path, sheet_name="Contact")
                # tên người gửi mail
                for index, row in df.iterrows():
                    user_name = row["Mail GD Tinh"]
                    outlook.to_user(user_name)

                # tên người cc
                for index, row in df.iterrows():
                    user_name = row["Mail quan ly khu vuc"]
                    outlook.cc_user(user_name)

                try:
                    outlook.input_subject_mail(subject)
                    print("nhập tiêu đề mail thành công")

                    try:
                        # đính kèm file excel
                        outlook.send_attach_file(DATA_TOOL_MANAGEMENT_PATH)
                        print("đã đính kèm file excel")

                    except Exception as e:
                        print(f"không đính kèm được file: {e}")

                    try:
                        outlook.input_context_mail(context)
                        print("Hoàn tất nhập nội dung văn bản")

                        try:
                            context_box = outlook.find_context_box()
                            context_box.send_keys(Keys.ENTER)
                            # tìm và copy ảnh tổng BAO CAO
                            open_and_copy_img(img_CNCT_path)
                            sleep(2)
                            context_box.send_keys(Keys.ENTER)
                            sleep(1)
                            context_box.send_keys(Keys.ENTER)
                            print(f"dán ảnh thành công")

                            try:
                                sleep(5)
                                outlook.click_send_mail_button()
                                print("gửi mail thanh cong, tien hanh dong trinh duyet")

                            except Exception as e:
                                print(f"không bấm được nút gửi mail: {e}")

                        except Exception as e:
                            print(f"không dán được ảnh: {e}")

                    except Exception as e:
                        print(f"lỗi khi nhập nội dung văn bản: {e}")

                except Exception as e:
                    print(f"lỗi khi nhập tiêu đề mail: {e}")

            except Exception as e:
                print(f"lỗi khi tag tên người dùng: {e}")

        except Exception as e:
            print(f"lỗi khi mở mail: {e}")

    except Exception as e:
        print(e)

    finally:
        sleep(20)
        browser.close()
        print("Đã đóng trình duyệt")


def send_mail_process():
    try:
        date_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        subject = "Cảnh báo tiến độ thực hiện WO tháng 12"
        Context = f"Kính gửi BGĐ CNCT, GĐ trung tâm cụm huyện!P.VHKT báo cáo tiến độ thực hiện WO đến thời điểm: {date_time} "
        input_mail_outlook(DATA_TOOL_MANAGEMENT_PATH, subject, Context)
        return True
    except Exception as e:
        print(e)
        return False


# quá trình whatsAPP: lấy dữ liệu - xử lý dữ liệu - gửi tin nhắn
def auto_process_diDong():
    try:
        getDB_to_excel(DATA_GNOC_RAW_PATH)

        try:
            # xử lý excel
            for attempt in range(5):
                if excel_transition_and_run_macro(
                    data_gnoc_raw_manager, data_tool_manager
                ):
                    break
                print(f"CĐBR: Xử lý dữ liệu Excel thất bại, thử lần {attempt + 1}")
            else:
                print("CĐBR: Xử lý dữ liệu Excel thất bại sau 5 lần thử")
                return

            if SENDBY.upper() == "WHATSAPP":
                try:
                    # gửi thông báo cấp CNCT
                    status_process = send_message_cnct()
                    if not status_process:
                        print("Gửi tin nhắn cho Tỉnh thất bại!!!")
                    else:
                        browser.start_browser(CHROME_PROFILE_CDBR_PATH)
                        whatsapp.driver = browser.driver
                        whatsapp.send_Error_Notification(
                            PHONE_NUMBER, "Gửi tin nhắn cho Tỉnh thành công !!!"
                        )
                        browser.close()
                except Exception as e:
                    print(e)

                try:
                    # gửi thông báo cấp cụm đội kt
                    list = send_message_user()
                    nofication = f"Đã gửi tin nhắn cho các cụm huyện {list} thành công "
                    browser.start_browser(CHROME_PROFILE_CDBR_PATH)
                    whatsapp.driver = browser.driver
                    whatsapp.send_Error_Notification(PHONE_NUMBER, nofication)
                    browser.close()

                except Exception as e:
                    print(e)

            elif SENDBY.upper() == "ZALO":
                try:
                    # gửi thông báo cấp CNCT
                    status_process = send_message_cnct_zalo()
                    if not status_process:
                        print("Gửi tin nhắn cho Tỉnh thất bại!!!")

                    else:
                        print("Gửi tin nhắn cho Tỉnh thành công!!!")

                except Exception as e:
                    print(e)

                try:
                    # gửi thông báo cấp cụm huyện có tag tên
                    list = send_message_user_with_TAG_zalo()
                    nofication = f"Đã gửi tin nhắn cho các cụm huyện {list} thành công "
                    print(nofication)

                except Exception as e:
                    print(e)

        except Exception as e:
            print(e)

    except Exception as e:
        print(e)

    finally:
        if browser.driver:
            browser.driver.quit()
