import os
import sys
from time import sleep
import psutil
from bs4 import BeautifulSoup
from pynput.keyboard import Controller, Key
import pyperclip
from pywinauto import Desktop
from pywinauto.application import Application
from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException, TimeoutException
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
import time
from datetime import datetime, timedelta
from selenium.common.exceptions import StaleElementReferenceException


from config import *
from utils import *
from openVPN import *

# Thay đổi môi trường sang tiếng Việt
sys.stdout.reconfigure(encoding="utf-8")


# Các XPath cố định
# WHATSAPP
XPATHS_WHATSAPP = {
    "message_box": "//div[@contenteditable='true' and @data-tab='10']",
    "search_box": "//div[@contenteditable='true']",  # "search_box": "//div[@contenteditable='true']" dia chi cu
    "result_list": "//div[@role='grid']//span[@title]",
    "send_image_button": '//*[@id="app"]/div/div[3]/div[2]/div[2]/span/div/div/div/div[2]/div/div[2]/div[2]/div/div',
    "attached_button": '//*[@id="main"]/footer/div[1]/div/span/div/div[1]/div/button',
    "image_input": '//input[@type="file" and @accept="image/*"]',
    "send_button": '//*[@id="app"]/div/div[3]/div/div[2]/div[2]/span/div/div/div/div[2]/div/div[2]/div[2]/div/div',
    "audio_dow_icon": "//span[@data-icon='audio-download']",
    "caption_image": '//*[@id="app"]/div/div[3]/div/div[2]/div[2]/span/div/div/div/div[2]/div/div[1]/div[3]/div/div/div[2]/div[1]/div[1]/p',
    "group_title": '//*[@id="main"]/header/div[2]/div[1]/div/span',
    "join_group": '//*[@id="action-button"]',
    "use_web": '//*[@id="fallback_block"]/div/h4/a',
    "group_name_join_chat": '//*[@id="main_block"]/h3',
    "send_button_an": '//*[@id="main"]/footer/div[1]/div/span/div/div[2]/div[2]/button/span',
    "Doan_chat_heading": '//*[@id="app"]/div/div[3]/div/div[3]/header/header/div/div[1]/h1',  # Khi trang whatsapp load xong thì dòng chữ đoạn chat sẽ hồi xong
}
# OUTLOOK
XPATHS_OUTLOOK = {
    "icon_outLook": '//*[@id="ddea774c-382b-47d7-aab5-adc2139a802b"]/span',
    "new_mail_button": '//*[@id="114-group"]/div/div[1]/div/div/span/button[1]',
    "send_to_box": "/html/body/div[1]/div/div[2]/div/div[2]/div[2]/div/div/div/div[3]/div/div/div[3]/div[1]/div/div/div/div/div[3]/div[1]/div/div[3]/div/span/span[2]/div/div[1]",
    "send_cc_box": "/html/body/div[1]/div/div[2]/div/div[2]/div[2]/div/div/div/div[3]/div/div/div[3]/div[1]/div/div/div/div/div[3]/div[1]/div/div[4]/div/span/span[2]/div/div[1]",
    "add_subject_box": '//*[@id="docking_InitVisiblePart_0"]/div/div[3]/div[2]/span/input',
    "add_context_box": '//*[@id="editorParent_1"]/div',
    "send_mail_button": '//*[@id="splitButton-ro__primaryActionButton"]',
    "insert_button": "/html/body/div[1]/div/div[2]/div/div[2]/div[1]/div/div/div[1]/div[2]/div/div/div/div/span/div[1]/div/div/div[5]/div/button",
    "attach_file_button": "/html/body/div[1]/div/div[2]/div/div[2]/div[1]/div/div/div[2]/div[1]/div/div/div/div/div/div/div/div/div/div/div[1]/div/div/div/div[1]/button",
    "browse_this_computer": "/html/body/div[2]/div[3]/div/div/div/div/div/div/ul/li/div/ul/li[1]/button",
    "send_mail_button": "/html/body/div[1]/div/div[2]/div/div[2]/div[2]/div/div/div/div[3]/div/div/div[3]/div[1]/div/div/div/div/div[2]/div[1]/button[1]",  #'//*[@id="splitButton-ro__primaryActionButton"]',
}

XPATHS_ZALO = {
    "search_box": '//*[@id="contact-search-input"]',
    "message_box": '//*[@id="input_line_0"]',
    "send_button": '//*[@id="chat-input-container-id"]/div[2]/div[2]/div[2]', 
    "img_attached_button": '//*[@id="chat-box-bar-id"]/div[1]/ul/li[2]/div/i',
    "file_attached_button": '//*[@id="chat-box-bar-id"]/div[1]/ul/li[3]/div',
    "group_title": '//*[@id="header"]/div[1]/div[2]/div[1]',
    "button_join_group": '//*[@id="root"]/div[1]/main/div/div[1]/div/button',
    "send_status": '//*[@id="messageViewScroll"]/div[6]/div[2]/div/div[2]/span',
    "STR_RECEIVED": '[data-translate-inner="STR_RECEIVED"]',
    "STR_SENDING": 'data-translate-inner="STR_SENDING"',
    "Clock_24": 'class="fa fa-Clock_24_Line"',
    "suggestion_box": '/html/body/div[2]',
}

XPATHS_GNOC = {
    "QuanLyCongViec": '//*[@id="root"]/div/div[2]/div[3]/div/div/ul/li[2]/a',
    "FlagIcon": '//*[@id="idFormSearch"]/div[2]/div/div/div/div[2]/div/div[1]/div[3]/div[1]/div/div[3]/div/span[1]/span',
    "trangThai_mo_rong": '//*[@id="idFormSearch"]/div[2]/div/div/div/div[2]/div/div[1]/div[2]/div/div[5]/div/div[1]/div/div[2]/div[2]/div[2]/svg',
    "trangThai_X_button": '//*[@id="idFormSearch"]/div[2]/div/div/div/div[2]/div/div[1]/div[2]/div/div[5]/div/div[1]/div/div[2]/div[2]/div[1]/svg',
    "trangThai_lua_chon": '//*[@id="idFormSearch"]/div[2]/div/div/div/div[2]/div/div[1]/div[2]/div/div[5]/div/div[1]/div/div[1]/div[1]',
    "trangThai_dong": "/html/body/script[3]",
    "export_button": '//*[@id="idFormSearch"]/div[2]/div/div/div/div[1]/div/div[2]/div/button[4]',
    "tiepTuc_dang_nhap": '//*[@id="registerform"]/div/div[2]/input',
    "checkBox_30day": '//*[@id="registerform"]/div/div[1]/label/span',
    "userName": '//*[@id="username"]',
    "passWord": '//*[@id="password"]',
    "dangNhap_button": '//*[@id="submit"]',
}

#Fire fox 
XPATHS_LINK_KHO = {
    "userName" : '//*[@id="username"]',
    "passWord": '//*[@id="password"]',
    "login_button": '/html/body/div[2]/form/div[1]/div[5]/input[3]'
}
class BrowserManager:
    def __init__(self):
        self.driver = None

    def start_browser(self, profile_path):
        chrome_option = webdriver.ChromeOptions()
        chrome_option.add_argument(f"user-data-dir={profile_path}")
        self.driver = webdriver.Chrome(options=chrome_option)
        self.driver.maximize_window()
        print("mo trinh duyet thanh cong")

    def is_browser_open(self):
        """Kiểm tra trình duyệt có đang mở không."""
        try:
            return self.driver is not None and len(self.driver.window_handles) > 0
        except Exception:
            return False

    def open_url(self, url):
        self.driver.get(url)

    def switch_to_tab(self, tab_index):
        windows = self.driver.window_handles
        if tab_index < len(windows):
            self.driver.switch_to.window(windows[tab_index])
        else:
            raise Exception("Invalid tab index")

    def close(self):
        self.driver.quit()
        self.driver = None
        sleep(5)

    def element_is_present(self, selector, by=By.CSS_SELECTOR, timeout=10):
        """Kiểm tra phần tử có xuất hiện không"""
        try:
            WebDriverWait(self.driver, timeout).until(
                EC.presence_of_element_located((by, selector))
            )
            return True
        except:
            return False


# Lớp WhatsAppBot
class WhatsAppBot(BrowserManager):
    def access_whatsapp(self):
        self.open_url(WHATSAPP_URL)
        try:
            WebDriverWait(self.driver, 200).until(
                EC.presence_of_element_located(
                    (
                        By.XPATH,
                        XPATHS_WHATSAPP["Doan_chat_heading"],
                    )
                )
            )
            print("WhatsApp loaded successfully!")
            return True

        except Exception as e:
            print("Error loading WhatsApp:", e)
            return False

    def reload_web(self):
        if self.driver.refresh():  # Tải lại trang
            try:
                WebDriverWait(self.driver, 200).until(
                    EC.presence_of_element_located(
                        (
                            By.XPATH,
                            XPATHS_WHATSAPP["Doan_chat_heading"],
                        )
                    )
                )
                print("WhatsApp refresh successfully!")
                return True

            except Exception as e:
                print("Error loading WhatsApp:", e)
                return False

    def find_name(self, object_name):
        """
        Tìm tên cá nhân
        """
        try:
            # Tìm kiếm ô tìm kiếm nhóm
            search_box = WebDriverWait(self.driver, 100).until(
                EC.presence_of_element_located(
                    (By.XPATH, XPATHS_WHATSAPP["search_box"])
                )
            )

            # Nhập tên nhóm vào ô tìm kiếm
            search_box.click()
            search_box.send_keys(Keys.CONTROL, "a", Keys.BACKSPACE)
            search_box.send_keys(object_name)
            print("đã nhập tên nhóm")

            try:
                # Chờ danh sách kết quả xuất hiện
                results = WebDriverWait(self.driver, 100).until(
                    EC.presence_of_all_elements_located(
                        (By.XPATH, XPATHS_WHATSAPP["result_list"])
                    )
                )

                # Duyệt qua từng kết quả để tìm tên chính xác
                for result in results:
                    if result.get_attribute("title") == object_name:
                        # Cuộn đến phần tử và nhấp
                        self.driver.execute_script(
                            "arguments[0].scrollIntoView(true);", result
                        )
                        WebDriverWait(self.driver, 50).until(
                            EC.element_to_be_clickable(result)
                        ).click()

                        if self.check_group_name(self, object_name):
                            print(f"Đã tìm và mở nhóm '{object_name}' thành công!")
                            return True

                print(f"Không tìm thấy nhóm '{object_name}' với tên chính xác.")
                return False

            except Exception as e:
                # Nếu không tìm thấy danh sách kết quả, thử nhấn Enter
                search_box.send_keys(Keys.ENTER)
                print(
                    f"Không tìm thấy danh sách kết quả. Đã thử nhấn Enter để mở nhóm '{object_name}'."
                )

                if self.check_group_name(object_name):
                    print(f"Đã tìm và mở nhóm '{object_name}' thành công!")
                    return True

                print(f"Không tìm thấy nhóm '{object_name}' với tên chính xác.")
                return False

        except Exception as e:
            print(f"❌ Lỗi trong quá trình tìm nhóm: {e}")
            return False

    def find_group_name(self, link):
        """Mở liên kết nhóm và kiểm tra tên nhóm."""
        # tham gia vào nhóm bằng đường link
        self.open_url(link)

        try:
            # Lấy tên nhóm từ giao diện tham gia chat
            element = WebDriverWait(self.driver, 100).until(
                EC.presence_of_element_located(
                    (By.XPATH, XPATHS_WHATSAPP["group_name_join_chat"])
                )
            )
            group_name = element.text

            # Nhấn nút tham gia nhóm
            WebDriverWait(self.driver, 100).until(
                EC.presence_of_element_located(
                    (By.XPATH, XPATHS_WHATSAPP["join_group"])
                )
            ).click()

            try:
                # Tìm nút "Sử dụng WhatsApp Web"
                use_web = WebDriverWait(self.driver, 100).until(
                    EC.presence_of_element_located(
                        (By.XPATH, XPATHS_WHATSAPP["use_web"])
                    )
                )

                # Get the HTML of the 'use_web' element
                use_web_html = use_web.get_attribute("outerHTML")

                # Parse the HTML with BeautifulSoup
                soup = BeautifulSoup(use_web_html, "html.parser")

                # Find the 'a' tag and extract the href attribute
                href_value = soup.find("a")["href"]

                self.open_url(href_value)

                try:
                    # Chờ trang load xong
                    WebDriverWait(self.driver, 300).until(
                        EC.presence_of_element_located(
                            (
                                By.XPATH,
                                XPATHS_WHATSAPP["Doan_chat_heading"],
                            )
                        )
                    )

                    if self.check_group_name(group_name):
                        print("")
                        print(f"Mở nhóm [{group_name}] thành công!!!")
                        return True
                    else:
                        print("Mở nhóm thất bại")
                        return False

                except Exception as e:
                    print(f"⚠ Lỗi khi kiểm tra nhóm: {e}")

            except Exception as e:
                print(f"⚠ Lỗi khi mở WhatsApp Web: {e}")

        except Exception as e:
            print(f"⚠ Lỗi khi lấy tên nhóm: {e}")

        return False

    def check_group_name(self, group_name):
        """Xác định xem nhóm mở có đúng với tên mong muốn không."""
        try:
            # hàm xác định group đã mở đúng không?
            element = WebDriverWait(self.driver, 100).until(
                EC.presence_of_element_located(
                    (By.XPATH, XPATHS_WHATSAPP["group_title"])
                )
            )

            if element.text == group_name:
                return True

            else:
                print("Tên nhóm không đúng với tên đã nhập trong tìm kiếm")
                return False

        except Exception as e:
            print(f"⚠ Không có nhóm nào đang mở hoặc lỗi xảy ra: {e}")
            return False

    def send_message(self, message):
        """Gửi tinh nhắn văn bản"""
        # tìm ô tin nhắn
        message_box = WebDriverWait(self.driver, 100).until(
            EC.presence_of_element_located((By.XPATH, XPATHS_WHATSAPP["message_box"]))
        )
        message_box.click()
        message_box.send_keys(message)
        message_box.send_keys(Keys.ENTER)
        sleep(5)

        self.get_last_message_info()

    def send_attached_file(self, file_path: str) -> bool:
        """
        Gửi tệp đính kèm qua WhatsApp.

        :param file_path: Đường dẫn tới tệp cần gửi.
        :return: Trả về True nếu gửi thành công, ngược lại False.
        """

        try:
            if not os.path.isfile(file_path):
                raise FileNotFoundError(f"File không tồn tại: {file_path}")

            # chờ nút gửi xuất hiện (sau khi mở được gr/người cần nhắn tin)
            attached_button = WebDriverWait(self.driver, 200).until(
                EC.presence_of_element_located(
                    (By.XPATH, XPATHS_WHATSAPP["attached_button"])
                )
            )
            attached_button.click()
            print("Đã tìm thấy và nhấn nút đính kèm.")

            # Chọn nút tài liệu
            file_input = WebDriverWait(self.driver, 100).until(
                EC.presence_of_element_located(
                    (
                        By.XPATH,
                        '//input[@type="file"]',
                    )  # Đường dẫn thường dùng để gửi tài liệu
                )
            )
            sleep(2)

            absolute_path = os.path.abspath(
                file_path
            )  # Chuyển đường dẫn thành tuyệt đối
            file_input.send_keys(absolute_path)

            # nút gửi tin nhắn
            send_button = WebDriverWait(self.driver, 100).until(
                EC.element_to_be_clickable(
                    (
                        By.XPATH,
                        XPATHS_WHATSAPP["send_button"],
                    )
                )
            )
            send_button.click()
            sleep(5)

            self.get_last_message_info()

            return True

        except Exception as e:
            print(f"Lỗi khi chọn file đính kèm: {e}")
            return False

    # def send_attached_img_message(self, message, file_path, tag_name=None):
    #     """Gửi tin nhắn + hình cùng 1 lúc"""

    #     try:
    #         message_box = WebDriverWait(self.driver, 10).until(
    #             EC.element_to_be_clickable(
    #                 (
    #                     By.XPATH,
    #                     XPATHS_WHATSAPP["message_box"],
    #                 )
    #             )
    #         )
    #         message_box.click()
    #         # Xóa toàn bộ nội dung đang có
    #         message_box.send_keys(Keys.CONTROL, "a", Keys.BACKSPACE)

    #         # Nhập nội dung tin nhắn
    #         message_box.send_keys(message)

    #         # Nếu có tag tên
    #         if tag_name:
    #             message_box.send_keys(": @")
    #             message_box.send_keys(remove_accents(tag_name))
    #             sleep(3)
    #             message_box.send_keys(Keys.TAB)
    #             sleep(2)

    #         copy_image_to_clipboard(file_path)
    #         message_box.send_keys(Keys.CONTROL, "v")

    #         # nút gửi tin nhắn
    #         send_button = WebDriverWait(self.driver, 10).until(
    #             EC.element_to_be_clickable(
    #                 (
    #                     By.XPATH,
    #                     XPATHS_WHATSAPP["send_button"],
    #                 )
    #             )
    #         )
    #         send_button.click()
    #         print("gui thanh cong")
    #         sleep(3)
    #         return True

    #     except Exception as e:
    #             print(f"Lỗi khi gửi tin nhắn và hình cảnh báo: {e}")
    #             return False

    def send_attached_img_message(self, message, file_path) -> bool:
        """Gửi tin nhắn + hình cùng 1 lúc"""
        if not os.path.isfile(file_path):
            print(f"⚠️ Cảnh báo: Ảnh '{file_path}' không tồn tại! Bỏ qua hình ảnh.")
            return False  # Bỏ qua nếu không có hình

        try:
            message_box = WebDriverWait(self.driver, 30).until(
                EC.element_to_be_clickable(
                    (
                        By.XPATH,
                        XPATHS_WHATSAPP["message_box"],
                    )
                )
            )
            message_box.click()
            # Xóa toàn bộ nội dung đang có
            message_box.send_keys(Keys.CONTROL, "a", Keys.BACKSPACE)

            # Nhập nội dung tin nhắn
            message_box.send_keys(message)
            sleep(2)

            # gửi hình ảnh vào ô chat
            copy_image_to_clipboard(file_path)
            message_box.send_keys(Keys.CONTROL, "v")

            # nút gửi tin nhắn
            send_button = WebDriverWait(self.driver, 20).until(
                EC.element_to_be_clickable(
                    (
                        By.XPATH,
                        XPATHS_WHATSAPP["send_button"],
                    )
                )
            )
            send_button.click()
            sleep(5)

            # Kiểm tra tin nhắn đã được gửi chưa
            return self.get_last_message_info()

        except Exception as e:
            print(f"Lỗi khi gửi tin nhắn và hình cảnh báo: {e}")
            return False

    def send_img(self, file_path) -> bool:
        """Chỉ gửi hình"""
        try:
            message_box = WebDriverWait(self.driver, 30).until(
                EC.element_to_be_clickable(
                    (
                        By.XPATH,
                        XPATHS_WHATSAPP["message_box"],
                    )
                )
            )
            message_box.click()
            # Xóa toàn bộ nội dung đang có
            message_box.send_keys(Keys.CONTROL, "a", Keys.BACKSPACE)

            # gửi hình ảnh vào ô chat
            copy_image_to_clipboard(file_path)
            message_box.send_keys(Keys.CONTROL, "v")

            # nút gửi tin nhắn
            send_button = WebDriverWait(self.driver, 20).until(
                EC.element_to_be_clickable(
                    (
                        By.XPATH,
                        XPATHS_WHATSAPP["send_button"],
                    )
                )
            )
            send_button.click()
            sleep(5)

            # Kiểm tra tin nhắn đã được gửi chưa
            self.get_last_message_info()

            return True

        except Exception as e:
            print(f"Lỗi khi gửi hình ảnh: {e}")
            return False

    def send_message_CDBR(self, message):
        """Gửi tin nhắn riêng cho CDBR"""
        try:
            # tìm ô tin nhắn
            message_box = WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located(
                    (By.XPATH, XPATHS_WHATSAPP["message_box"])
                )
            )
            message_box.click()
            message_box.send_keys(Keys.CONTROL, "a", Keys.BACKSPACE)
            message_box.send_keys(message)
            message_box.send_keys(Keys.CONTROL, "v")
            sleep(5)

            xpaths = [XPATHS_WHATSAPP["send_button"], XPATHS_WHATSAPP["send_button_an"]]

            for xpath in xpaths:
                try:
                    send_button = WebDriverWait(self.driver, 5).until(
                        EC.presence_of_element_located((By.XPATH, xpath))
                    )
                    if send_button.is_displayed() and send_button.is_enabled():
                        send_button.click()
                        sleep(3)
                        return self.get_last_message_info()

                except:
                    pass

            return False

        except Exception as e:
            print(f"Lỗi trong quá trình gửi tin nhắn CDBR: {e}")
            return False

    def get_last_message_info(self):
        """Hàm lấy thông tin tin nhắn mới nhất"""
        try:
            messages = self.driver.find_elements(
                By.CSS_SELECTOR, "div.message-out"
            )  # Chỉ lấy tin nhắn do bạn gửi

            if not messages:
                print("❌ No sent messages found.")
                return False

            last_message = messages[-1]
            start_time = time.time()
            max_wait_time = 100

            while time.time() - start_time < max_wait_time:

                # Kiểm tra trạng thái tin nhắn
                if last_message.find_elements(
                    By.CSS_SELECTOR, 'span[data-icon="msg-dblcheck"]'
                ):
                    print("✅✅ Tin nhắn đã gửi đi và được nhận!")
                    return True

                elif last_message.find_elements(
                    By.CSS_SELECTOR, 'span[data-icon="msg-check"]'
                ):
                    print("✅ Tin nhắn đã gửi đi nhưng chưa được nhận.")
                    sleep(5)
                    return True

                sleep(2)  # Kiểm tra lại sau 2 giây

            print("⚠️ Quá thời gian chờ nhưng chưa có xác nhận nhận tin nhắn.")
            return False

        except Exception as e:
            print("❌ Lỗi khi lấy thông tin tin nhắn:", e)
            return False


# Lớp ZaloBot
class ZaloBot(BrowserManager):

    def access_zalo(self):
        self.open_url(ZALO_URL)
        try:
            WebDriverWait(self.driver, 200).until(
                EC.presence_of_element_located(
                    (
                        By.XPATH,
                        XPATHS_ZALO["search_box"],
                    )
                )
            )
            print("Zalo loaded successfully!")
            return True

        except Exception as e:
            print("Error loading Zalo:", e)
            return False

    def reload_web(self):
        if self.driver.refresh():  # Tải lại trang
            try:
                WebDriverWait(self.driver, 200).until(
                    EC.presence_of_element_located(
                        (
                            By.XPATH,
                            XPATHS_ZALO["search_box"],
                        )
                    )
                )
                print("Zalo refresh successfully!")
                return True

            except Exception as e:
                print("Error loading WhatsApp:", e)
                return False

    def find_name(self, object_name, xpath_address):
        """
        tìm tên người - nhóm theo tên và địa chỉ xpath

        Args:
            object_name (str): tên đối tượng cần tìm kiếm.
            xpath_address (str): địa chỉ xpath tương đương trên tìm kiếm
        """
        try:
            # Tìm kiếm ô tìm kiếm nhóm
            search_box = WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.XPATH, XPATHS_ZALO["search_box"]))
            )

            # Nhập tên nhóm vào ô tìm kiếm
            search_box.click()
            search_box.clear()
            search_box.send_keys(object_name)

            # # Chờ danh sách kết quả xuất hiện
            # results = WebDriverWait(self.driver, 20).until(
            #     EC.presence_of_all_elements_located(
            #         (By.XPATH, XPATHS_WHATSAPP["result_list"])
            #     )
            # )

            # tìm tên đúng theo xpath
            object_click = WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.XPATH, f"{xpath_address}"))
            )
            object_click.click()

            print(f"Tìm {object_name} thành công!!!")

        except Exception as e:
            print(f"Đã xảy ra lỗi: {e}")
            return False

    def find_group_name(self, link):
        """Tìm nhóm theo link mời vào nhóm"""
        try:
            self.open_url(link)
            # hàm xác định group đã mở đúng không?
            element = WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.XPATH, XPATHS_ZALO["group_title"]))
            )

            text_content = element.text

            print(f"\nGROUP [{text_content}] DANG MỞ")
            return True
        except Exception as e:
            print(e)
            return False

    def check_group_name(self, group_name):
        try:
            # hàm xác định group đã mở đúng không?
            element = WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.XPATH, XPATHS_ZALO["group_title"]))
            )

            text_content = element.text

            if text_content == group_name:
                return True
            else:
                return False

        except Exception as e:
            print(f"Không có group nào mở: {e}")

    def find_name_no_xpath(self, object_name):
        """
        tìm tên người - nhóm theo tên và địa chỉ xpath

        Args:
            object_name (str): tên đối tượng cần tìm kiếm.
            xpath_address (str): địa chỉ xpath tương đương trên tìm kiếm
        """
        try:
            # Tìm kiếm ô tìm kiếm nhóm
            search_box = WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.XPATH, XPATHS_ZALO["search_box"]))
            )

            # Nhập tên nhóm vào ô tìm kiếm
            search_box.click()
            search_box.clear()
            search_box.send_keys(object_name)

            search_box.send_keys(Keys.ENTER)
            print(f"Tìm {object_name} thành công!!!")

        except Exception as e:
            print(f"Đã xảy ra lỗi: {e}")
            return False

    def send_message(self, message):
        # tìm ô tin nhắn
        message_box = WebDriverWait(self.driver, 10).until(
            EC.presence_of_element_located((By.XPATH, XPATHS_ZALO["message_box"]))
        )
        message_box.click()
        message_box.send_keys(message)
        message_box.send_keys(Keys.ENTER)
        sleep(1)

        return self.check_last_message_status()

    def run_macro_and_send_message(self, driver, excel, macro, message):
        try:
            # excel.Application.Run(macro)
            print("Run Macro thành công!!!")

            try:
                # tìm ô tin nhắn
                message_box = WebDriverWait(self.driver, 10).until(
                    EC.presence_of_element_located(
                        (By.XPATH, XPATHS_ZALO["message_box"])
                    )
                )
                message_box.click()
                message_box.send_keys(message)
                message_box.send_keys(Keys.CONTROL, "V")

                try:
                    # tìm ô tin nhắn
                    send_button = WebDriverWait(self.driver, 10).until(
                        EC.presence_of_element_located(
                            (By.XPATH, XPATHS_ZALO["send_button"])
                        )
                    )
                    send_button.click()

                except Exception as e:
                    print(e)

            except Exception as e:
                print(e)

        except Exception as e:
            print(e)

    def find_message_box(self):
        return WebDriverWait(self.driver, 20).until(
                EC.presence_of_element_located((By.XPATH, XPATHS_ZALO["message_box"]))
            )
    
    def send_attached_img_message(self, message, file_path, tag_names=None):
        """
        gửi tin nhắn + gửi hình
        Args:
            message (str): tin nhắn muốn gửi
            img_path (str): đường dẫn thư mục chứa ảnh muốn gửi
        """
        try:
            # Kiểm tra đường dẫn ảnh
            if not os.path.isfile(file_path):
                raise FileNotFoundError(f"File không tồn tại: {file_path}")

            copy_image_to_clipboard(file_path)

            # Tìm ô message_box trong Zalo
            message_box = self.find_message_box()
            message_box.click()
            message_box.send_keys(Keys.CONTROL, "a", Keys.BACKSPACE)
            message_box.send_keys(Keys.CONTROL, "v")  # dán hình vào ô tin nhắn
            sleep(2)

            if tag_names is not None:
                if not isinstance(tag_names, list):
                    tag_names = [tag_names] 
                message_box.send_keys(message)

                for tag_name in tag_names:
                    try:
                        message_box = self.find_message_box()
                        message_box.send_keys(" @")
                        message_box.send_keys(tag_name)
                        sleep(2)
                        try: 
                            suggestion = WebDriverWait(self.driver, 5).until(
                                EC.presence_of_element_located((By.XPATH, XPATHS_ZALO["suggestion_box"]))
                            )
                            # message_box.send_keys(Keys.ARROW_DOWN)
                            message_box.send_keys(Keys.TAB)
                            message_box.send_keys(Keys.SPACE)
                            sleep(2)
                        except:
                            # print(f"⚠️ Không tìm thấy gợi ý cho {tag_name}, có thể không tag được.")
                            continue
                        
                    except StaleElementReferenceException:
                        # print(f"    ⚠️ Phần tử message_box bị mất, tìm lại...")
                        continue 
                    
            else:
                message_box.send_keys(message)

            # Nhấn nút gửi tin nhắn
            send_button = WebDriverWait(self.driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, XPATHS_ZALO["send_button"]))
            )
            send_button.click()
            sleep(2)

            return self.check_last_message_status()

        except FileNotFoundError as e:
            print(f"❌ Lỗi: {e}")
            return "failed"

        except Exception as e:
            print(f"⚠️ Lỗi không xác định khi gửi tin nhắn: {e}")
            return "failed"

    def send_img(self, message, file_path, tag_name=None):
        try:
            element = WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.XPATH, XPATHS_ZALO["message_box"]))
            )
        except Exception as e:
            print(e)
            return False

    def send_attached_img(self, file_path):
        try:
            # Kiểm tra đường dẫn ảnh
            if not os.path.isfile(file_path):
                raise FileNotFoundError(f"File không tồn tại: {file_path}")

            copy_image_to_clipboard(file_path)

            try:
                # Tìm ô message_box trong Zalo
                message_box = WebDriverWait(self.driver, 20).until(
                    EC.presence_of_element_located(
                        (By.XPATH, XPATHS_ZALO["message_box"])
                    )
                )
                message_box.click()
                message_box.send_keys(Keys.CONTROL, "a")
                message_box.send_keys(Keys.BACKSPACE)
                message_box.send_keys(Keys.CONTROL, "v")

            except Exception as e:
                print(e)

        except Exception as e:
            print(e)

    def send_file_zalo(self, file_path):
        try:
            # Nhấn vào nút đính kèm file (THAY XPATH CHO ĐÚNG)
            attach_button_xpath = '//*[@id="chat-box-bar-id"]/div[1]/ul/li[3]/div'  # XPath này có thể thay đổi theo Zalo
            attach_button = WebDriverWait(self.driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, attach_button_xpath))
            )
            attach_button.click()

            # Chờ input file xuất hiện
            file_input_xpath = (
                "/html/body/div[2]/div[2]/div/div/div/div"  # Xpath của input chọn file
            )
            file_input = WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.XPATH, file_input_xpath))
            )
            file_input.click()
            sleep(2)

            # desktop = Desktop(backend="win32")
            # for window in desktop.windows():
            #     print(f"🔍 {window.window_text()}")

            # Tương tác với cửa sổ file picker bằng pywinauto
            app = Application(backend="win32").connect(
                title="Open"
            )  # Thay bằng tiêu đề thực tế
            dialog = app.window(title_re="Open")  # Thay bằng tiêu đề thực tế

            dialog.set_focus()  # Đưa cửa sổ lên trước
            dialog["Edit"].set_text(
                r"C:\Users\Admin\Desktop\file.xlsx"
            )  # Nhập đường dẫn file
            dialog["Open"].click()  # Nhấn nút Open

            # Nhấn nút gửi (THAY XPATH CHO ĐÚNG)
            send_button = WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.XPATH, XPATHS_ZALO["send_button"]))
            )
            send_button.click()

            print("📂 File đã được gửi thành công!")

        except Exception as e:
            print(f"❌ Lỗi khi gửi file: {e}")

    def send_message_CDBR(self, message):
        try:
            # tìm ô tin nhắn
            message_box = WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.XPATH, XPATHS_ZALO["message_box"]))
            )
            message_box.click()
            message_box.send_keys(Keys.CONTROL, "a", Keys.BACKSPACE)

            message_box.send_keys(message)
            message_box.send_keys(Keys.CONTROL, "v")
            sleep(4)
            
            send_button = WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.XPATH, XPATHS_ZALO["send_button"]))
            )

            send_button.click()
            
            return self.check_last_message_status()

        except Exception as e:
            print(e)
            return "failed"

    def check_last_message_status(self):
        """Kiểm tra trạng thái tin nhắn sau khi gửi trên Zalo Web."""
        try:
            if self.element_is_present(
                '[data-translate-inner="STR_RECEIVED"]', By.CSS_SELECTOR, 20
            ):
                sleep(3)
                return "sent"

            start_time = time.time()
            max_wait_time = 100

            # Kiểm tra xem có phần tử hiển thị tin nhắn chưa được gửi đi không
            while self.element_is_present(
                '[data-translate-inner="STR_SENDING"]', By.CSS_SELECTOR, 1
            ) or self.element_is_present(".fa.fa-Clock_24_Line", By.CSS_SELECTOR, 1):

                sleep(1)
                if time.time() - start_time > max_wait_time:
                    return "timeout"

            return "sent"

        except Exception as e:
            return "failed"

    def get_last_message_info(self):
        """Lấy thông tin tin nhắn cuối cùng"""
        try:
            messages = self.driver.find_elements(
                By.CSS_SELECTOR, '[data-id="div_SentMsg_Text"]'
            )
            # Chỉ lấy tin nhắn do bạn gửi

            if not messages:
                print("❌ Không tìm thấy tin nhắn đã gửi.")
                return False

            last_message = messages[-1]

            message_text = last_message.text.strip()
            return message_text

        except Exception as e:
            print("❌ Lỗi khi lấy thông tin tin nhắn:", e)
            return None


class OutLookBot(BrowserManager):

    def access_outlook(self):
        self.open_url(OUT_LOOK_URL)
        try:
            WebDriverWait(self.driver, 200).until(
                EC.presence_of_element_located(
                    (
                        By.XPATH,
                        XPATHS_OUTLOOK["icon_outLook"],
                    )
                )
            )
            print("OutLook loaded successfully!")
        except Exception as e:
            print("Error loading OutLook:", e)

    # nhập tên người gửi
    def find_send_to_box(self):
        send_to_box = WebDriverWait(self.driver, 15).until(
            EC.presence_of_element_located(
                (
                    By.XPATH,
                    XPATHS_OUTLOOK["send_to_box"],
                )
            )
        )
        return send_to_box

    def to_user(self, user_name):
        send_to_box = self.find_send_to_box()
        send_to_box.click()
        sleep(1)
        try:
            send_to_box.send_keys(user_name)
            sleep(3)
            send_to_box.send_keys(Keys.TAB)
        except Exception as e:
            print(e)

    # nhập tên người cc
    def find_send_cc_box(self):
        send_cc_box = WebDriverWait(self.driver, 15).until(
            EC.presence_of_element_located(
                (
                    By.XPATH,
                    XPATHS_OUTLOOK["send_cc_box"],
                )
            )
        )
        return send_cc_box

    def cc_user(self, user_name):
        send_cc_box = self.find_send_cc_box()
        send_cc_box.click()
        sleep(1)
        try:
            send_cc_box.send_keys(user_name)
            sleep(3)
            send_cc_box.send_keys(Keys.TAB)
        except Exception as e:
            print(e)

    # tìm nút tạo thư mới
    def find_new_mail_button(self):
        new_mail_button = WebDriverWait(self.driver, 10).until(
            EC.presence_of_element_located(
                (By.XPATH, XPATHS_OUTLOOK["new_mail_button"])
            )
        )
        return new_mail_button

    # nhập tiêu đề thư
    def find_subject_box(self):
        subject_box = WebDriverWait(self.driver, 10).until(
            EC.presence_of_element_located(
                (By.XPATH, XPATHS_OUTLOOK["add_subject_box"])
            )
        )
        return subject_box

    def input_subject_mail(self, subject):
        subject_box = self.find_subject_box()
        subject_box.click()
        subject_box.send_keys(subject)

    # nhập nội dung văn bản
    def find_context_box(self):
        context_box = WebDriverWait(self.driver, 10).until(
            EC.presence_of_element_located(
                (By.XPATH, XPATHS_OUTLOOK["add_context_box"])
            )
        )
        return context_box

    def input_context_mail(self, context):
        context_box = self.find_context_box()
        context_box.click()
        context_box.send_keys(context)

    # tìm nút gửi
    def click_send_mail_button(self):
        send_mail_button = WebDriverWait(self.driver, 10).until(
            EC.presence_of_element_located(
                (By.XPATH, XPATHS_OUTLOOK["send_mail_button"])
            )
        )
        send_mail_button.click()
        return send_mail_button

    def send_attach_file(self, file_path):
        try:
            insert_button = WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located(
                    (By.XPATH, XPATHS_OUTLOOK["insert_button"])
                )
            )
            insert_button.click()
            sleep(1)

            attach_file_button = WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located(
                    (By.XPATH, XPATHS_OUTLOOK["attach_file_button"])
                )
            )
            attach_file_button.click()
            sleep(1)

            browse_this_computer = WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located(
                    (By.XPATH, XPATHS_OUTLOOK["browse_this_computer"])
                )
            )
            browse_this_computer.click()
            sleep(1)

            # Tương tác với cửa sổ file picker bằng pywinauto
            app = Application(backend="win32").connect(title_re="Open", found_index=0)
            dialog = app.window(title_re="Open")
            dialog.set_focus()  # Kích hoạt cửa sổ
            dialog["Edit"].type_keys(file_path)
            dialog["Open"].click()
            keyboard = Controller()
            keyboard.press(Key.enter)
            keyboard.release(Key.enter)

            sleep(3)
        except Exception as e:
            print(f"Lỗi trong quá trình đính kèm file: {e}")


class GnocBot(BrowserManager):
    def access(self, userName_gnoc, pwd_gnoc, otp_key):
        self.open_url(GNOC_URL)
        try:
            WebDriverWait(self.driver, 20).until(
                EC.presence_of_element_located(
                    (
                        By.XPATH,
                        XPATHS_GNOC["QuanLyCongViec"],
                    )
                )
            )
            print("Gnoc loaded successfully!")
            return True

        except Exception as e:
            print(f"Chưa đăng nhập Gnoc, tiến hành đăng nhập!!!")

            if not self.wait_and_send_keys("username", userName_gnoc):
                return False

            if not self.wait_and_send_keys("password", pwd_gnoc):
                return False

            if not Click_byImage("Dang_nhap_gnoc"):
                return False

            if find_Element_byImage("otp_pass"):
                print("nhập mã OTP")
                totp = pyotp.TOTP(otp_key).now()
                pyautogui.typewrite(totp)
                keyboard = Controller()
                keyboard.press(Key.enter)  # Press Enter
                keyboard.release(Key.enter)

            sleep(2)
            if find_Element_byImage("tiep_tuc_dang_nhap"):
                if not Click_byImage("tiep_tuc_dang_nhap"):
                    print("Thất bại khi cố bấm nút Tiếp tục")
                    return False

            try:
                WebDriverWait(self.driver, 200).until(
                    EC.presence_of_element_located(
                        (
                            By.XPATH,
                            XPATHS_GNOC["QuanLyCongViec"],
                        )
                    )
                )
                print("Đăng nhập GNOC thành công!!!")
                return True

            except Exception as e:
                print(f"Lỗi khi cố đăng nhập GNOC: {e}")
                return False

    def wait_and_send_keys(self, element_id, text, timeout=10):
        try:
            field = WebDriverWait(self.driver, timeout).until(
                EC.presence_of_element_located((By.ID, element_id))
            )
            field.send_keys(text)
            return True

        except Exception as e:
            print(f"Lỗi: Không tìm thấy phần tử {element_id}: {e}")
            return False

    def click_on(self, name_ele, sleep_time=2):
        try:
            element = WebDriverWait(self.driver, 10).until(
                EC.element_to_be_clickable(
                    (
                        By.XPATH,
                        XPATHS_GNOC[f"{name_ele}"],
                    )
                )
            )
            element.click()
            sleep(sleep_time)
            return True

        except Exception as e:
            print(f"lỗi khi cố click on {name_ele}: {e}")
            return False

    def get_WoDong(self):
        try:
            time_out = 100
            time_wait = 100
            wait = WebDriverWait(self.driver, time_wait)
            wait.until(
                EC.invisibility_of_element_located(
                    (By.XPATH, '//div[contains(text(), "Đang tải")]')
                )
            )
            sleep(2)

            if not Click_byImage("quanLyCongViec"):
                print("Lỗi không chọn được QLCV")
                return False
            
            wait = WebDriverWait(self.driver, time_wait)
            wait.until(
                EC.invisibility_of_element_located(
                    (By.XPATH, '//div[contains(text(), "Đang tải")]')
                )
            )
            sleep(2)
            if not Click_byImage("TrangThai_lua_chon"):
                print("Lỗi không chọn được TrangThai_lua_chon")
                return False

            pyautogui.typewrite(" Đóng")
            sleep(2)

            if not Click_byImage("dong"):
                return False

            if not Click_byImage("calender_icon"):
                return False

            new_date = (datetime.now() - timedelta(days=1)).strftime(
                "%d/%m/%Y 00:00:00"
            )

            sleep(2)
            input_element = self.driver.find_element(By.ID, "DateTimeInput_start")
            input_element.send_keys(Keys.CONTROL, "a", Keys.BACKSPACE)
            input_element.send_keys(new_date)
            input_element.send_keys(Keys.RETURN)
            sleep(2)

            if not Click_byImage("chon"):
                return False

            if not Click_byImage("tim_kiem"):
                return False

            wait = WebDriverWait(self.driver, time_wait)
            wait.until(
                EC.invisibility_of_element_located(
                    (By.XPATH, '//div[contains(text(), "Đang tải")]')
                )
            )

            export_button = WebDriverWait(self.driver, time_out).until(
                EC.element_to_be_clickable(
                    (
                        By.XPATH,
                        XPATHS_GNOC["export_button"],
                    )
                )
            )
            export_button.click()

            wait = WebDriverWait(self.driver, time_wait)
            wait.until(
                EC.invisibility_of_element_located(
                    (By.XPATH, '//div[contains(text(), "Đang tải")]')
                )
            )

            start_time = time.time()  # Lưu thời gian bắt đầu

            while time.time() - start_time < 200:  # Lặp trong 200 giây
                try:
                    Export_thanhCong_path = str(IMAGE_PATH / "Export_thanhCong.png")
                    Export_thanhCong_img = pyautogui.locateOnScreen(
                        Export_thanhCong_path, confidence=0.8
                    )

                    if Export_thanhCong_img:
                        if not Click_byImage("Keep"):
                            return False
                        else:
                            break
                        
                except pyautogui.ImageNotFoundException:
                    pass  # Nếu không tìm thấy, tiếp tục lặp

                sleep(1)  # Nghỉ 1 giây trước khi thử lại

            while time.time() - start_time < 200:  # Lặp trong 200 giây
                try:
                    dowload_comlate_img = str(IMAGE_PATH / "wo_dong_excel_done.png")
                    dowload_comlate = pyautogui.locateOnScreen(
                        dowload_comlate_img, confidence=0.8
                    )
                    
                    if dowload_comlate:
                        return True
                    
                except pyautogui.ImageNotFoundException:
                    pass  # Nếu không tìm thấy, tiếp tục lặp
                
                try:
                    keep_img = str(IMAGE_PATH / "keep.png")
                    keep_button = pyautogui.locateOnScreen(
                        keep_img, confidence=0.8
                    )
                    if keep_button: 
                        Click_byImage("Keep")
                        
                except pyautogui.ImageNotFoundException:
                    pass  # Nếu không tìm thấy, tiếp tục lặp

                sleep(5)

            print("❌ Hết thời gian chờ!")  # Debug khi timeout
            return False  # Hết thời gian mà vẫn chưa tìm thấy ảnh

        except Exception as e:
            print(f"Lỗi khi lấy dữ liệu gnoc web: {e}")
            return False

class FireFoxManager:
    def __init__(self):
        self.driver = None

    def start_browser(self, profile_path):
        firefox_option = webdriver.FirefoxOptions()
        firefox_option.profile = profile_path  # Chỉ định profile của Firefox

        self.driver = webdriver.Firefox(options=firefox_option)
        self.driver.maximize_window()
        print("Mở trình duyệt thành công")

    def open_url(self, url):
        self.driver.get(url)

    def close(self):
        self.driver.quit()
        self.driver = None
        print("Đóng trình duyệt fire fox")
        sleep(5)

class KHO(FireFoxManager):

    def access(self, userName, password):
        self.open_url(LINK_KHO)

        try:
            #Đăng nhập 
            if not self.wait_and_send_keys("username", userName):
                return False

            if not self.wait_and_send_keys("password", password):
                return False
            
            return True
        except Exception as e:
            print(f"Lỗi khi đăng nhập: {e}")
            return False
        
    def wait_and_send_keys(self, element_id, text, timeout=10):
        try:
            field = WebDriverWait(self.driver, timeout).until(
                EC.presence_of_element_located((By.ID, element_id))
            )
            field.send_keys(text)
            return True
        except Exception as e:
            print(f"Lỗi: Không tìm thấy phần tử {element_id}: {e}")
            return False
        

def open_sfive():
    try:
        # Mở SFive trên Windows
        subprocess.run(["cmd", "/c", "start sfive"], shell=True)
        time.sleep(5)  # Chờ SFive khởi động

        # Kiểm tra tiến trình SFive đã chạy chưa
        sfive_running = False
        for process in psutil.process_iter(attrs=["name"]):
            if "sfive" in process.info["name"].lower():
                sfive_running = True
                break
        
        if not sfive_running:
            print("Lỗi: SFive không chạy!")
            return False

        # Tìm cửa sổ SFive và maximize
        windows = gw.getWindowsWithTitle("SFive")
        if windows:
            for window in windows:
                window.maximize()
        else:
            print("Lỗi: Không tìm thấy cửa sổ SFive!")
            return False

        toolbar_path = str(IMAGE_PATH / "sfive" / "tool_bar.png")
        sFive_logo = str(IMAGE_PATH / "sfive" / "sFive_logo.png")

        start_time = time.time()
        max_wait_time = 20
        
        while time.time() - start_time < max_wait_time:
            try: 
                if find_Element_sFive(toolbar_path):
                    return True
                
                sFive_position = pyautogui.locateOnScreen(sFive_logo, confidence=0.8)
                if sFive_position: 
                    pyautogui.click(pyautogui.center(sFive_position))
                    
            except pyautogui.ImageNotFoundException:
                pass
        
        print("     Hết thời gian chờ mở sFive!!!")
        return False
    
    except Exception as e:
        print(f"Lỗi khi mở SFive: {e}")
        return False
    
def Login_bccs(link):
    webbar_position = [574,49]

    try:
        print("     Truy cập vào BCCS")
        pyautogui.click(webbar_position) 
        sleep(2)
        pyautogui.typewrite(LINK_KHO)
        sleep(5)
        pyautogui.press("enter") 


        login_button_path = str(IMAGE_PATH / "sfive" / "login.png")
        bccs_homepage = str(IMAGE_PATH / "sfive" / "BCCS_homepage.png")

        start_time = time.time()
        max_wait_time = 100
        
        while time.time() - start_time < max_wait_time:
            try: 
                sleep(5)
                login_button = pyautogui.locateOnScreen(login_button_path, confidence=0.8)

                if login_button:
                    print("     Chưa đăng nhập, tiến hành click nút đăng nhập!!!")
                    pyautogui.click(pyautogui.center(login_button))

                if find_Element_sFive(bccs_homepage):
                    print("    Đăng nhập BCCS thành công")
                    return True
            
            except pyautogui.ImageNotFoundException:
                pass

    except Exception as e:
        print(f"Lỗi khi lấy dữ liệu Wo dong: {e}")
        return False

def click_on(img_path):
    try:
        element = pyautogui.locateOnScreen(img_path, confidence=0.8)

        if not element:
            return False
        
        pyautogui.click(pyautogui.center(element))
        sleep(4)
        return True
    
    except pyautogui.ImageNotFoundException:
        return False

def process_WoDong():
    max_retries = 3
    retries = 0

    while retries < max_retries:
        try:
            if not on_openvpn():
                raise ConnectionError
            
            sleep(5)
            if open_sfive():
                if Login_bccs(LINK_KHO):
                    if get_Wo_Inventory():
                        return True

        except ConnectionError as ce:
            print(f"Lỗi kết nối: {ce}")
        
        except Exception as e:
            print(f"Lỗi không xác định: {e}")

        # finally:
        #     off_openvpn()

        # Tăng số lần thử và thời gian chờ
        retries += 1
        print(f"Thử lại lần thứ {retries} sau 5 giây...")
        sleep(5)

    print("Không thể hoàn thành tác vụ sau nhiều lần thử.")
    return False

def delete_All():
    pyautogui.hotkey("ctrl", "a")

    sleep(0.5)  # Chờ một chút để đảm bảo nội dung đã được chọn

    # Nhấn Delete để xóa nội dung
    pyautogui.press("delete")

    sleep(0.5)  

def get_Wo_Inventory():
    try:
        xuatExcelFile_position1 = [1496, 425]
        xuatExcelFile_position2 = [1471, 451]
        donVi = [428, 204]
        loai_giaoDich = [863, 202]


        inventory_path = str(IMAGE_PATH / "sfive" / "inventory.png")
        baoCaoChiTietTon_path = str(IMAGE_PATH / "sfive" / "baoCaoChiTietTon.png")
        xuatExcelButton_path = str(IMAGE_PATH / "sfive" / "xuatExcel.png")

        morong_path = str(IMAGE_PATH / "sfive" / "moRong.png")
        baocao_path = str(IMAGE_PATH / "sfive" / "baoCao.png")
        QuanLyBaoCaoOffline_path = str(IMAGE_PATH / "sfive" / "QuanLyBaoCaoOffline.png")
        timkiem_path = str(IMAGE_PATH / "sfive" / "timKiem.png")
        savebutton_path= str(IMAGE_PATH / "sfive" / "save.png")

        if not click_on(inventory_path):
            return False 
        
        if not click_on(baoCaoChiTietTon_path):
            return False 
        
        pyautogui.click(donVi[0], donVi[1])
        sleep(1)
        text = "CTCT_BDG -- Chi nhánh Kỹ Thuật Viettel BDG"

        # Sao chép vào clipboard
        pyperclip.copy(text)

        # Dán nội dung bằng Ctrl + V
        pyautogui.hotkey("ctrl", "v")
        pyautogui.press("enter") 

        pyautogui.click(loai_giaoDich[0], loai_giaoDich[1])
        sleep(1)
        text = "Kho đơn vị"

        # Sao chép vào clipboard
        pyperclip.copy(text)

        # Dán nội dung bằng Ctrl + V
        pyautogui.hotkey("ctrl", "v")
        pyautogui.press("enter") 
        xuatExcelButton = pyautogui.locateOnScreen(xuatExcelButton_path, confidence=0.8)

        if xuatExcelButton:
            pyautogui.click(pyautogui.center(xuatExcelButton))
            print("     Click nút xuất excel thành công!")

        pyautogui.click(loai_giaoDich[0], loai_giaoDich[1])
        sleep(1)
        delete_All()
        text = "Kho nhân viên"

        # Sao chép vào clipboard
        pyperclip.copy(text)

        # Dán nội dung bằng Ctrl + V
        pyautogui.hotkey("ctrl", "v")

        xuatExcelButton = None
        xuatExcelButton = pyautogui.locateOnScreen(xuatExcelButton_path, confidence=0.8)

        if xuatExcelButton:
            pyautogui.click(pyautogui.center(xuatExcelButton))
            print("     Click nút xuất excel thành công!")
        
        sleep(1)
        morong_button = pyautogui.locateOnScreen(morong_path, confidence=0.8)

        if morong_button:
            morong_button_center = pyautogui.center(morong_button)  # Lấy tọa độ trung tâm của nút
            pyautogui.moveTo(morong_button_center.x, morong_button_center.y, duration=0.5)  # Di chuyển chuột đến đó trong 0.5 giây

        sleep(1)
        baocao_button = pyautogui.locateOnScreen(baocao_path, confidence=0.8)

        if baocao_button:
            baocao_button_center = pyautogui.center(baocao_button)  # Lấy tọa độ trung tâm của nút
            pyautogui.moveTo(baocao_button_center.x, baocao_button_center.y, duration=0.5)  # Di chuyển chuột đến đó trong 0.5 giây
        
        sleep(1)
        QuanLyBaoCaoOffline_button = pyautogui.locateOnScreen(QuanLyBaoCaoOffline_path, confidence=0.8)

        if QuanLyBaoCaoOffline_button:
            pyautogui.click(pyautogui.center(QuanLyBaoCaoOffline_button))
            print("     Click nút QuanLyBaoCaoOffline_button")
        
        sleep(2)
        timKiem_button = pyautogui.locateOnScreen(timkiem_path, confidence=0.8)

        if timKiem_button:
            pyautogui.click(pyautogui.center(timKiem_button))
            print("     Click nút Tim kiem")

        sleep(10)
        
        # Xuat file 1
        pyautogui.click(xuatExcelFile_position1[0], xuatExcelFile_position1[1])
        sleep(2)
        delete_All()
        pyautogui.typewrite("E:\Auto\Auto_tool_offical\data\excel\data_CDBR\kho_dong")

        savebutton = pyautogui.locateOnScreen(savebutton_path, confidence=0.8)

        if savebutton:
            pyautogui.click(pyautogui.center(savebutton))
            print("     Click nút save")  

        # Xuat file 2
        pyautogui.click(xuatExcelFile_position2[0], xuatExcelFile_position2[1])
        sleep(2)

        delete_All()
        pyautogui.typewrite("E:\Auto\Auto_tool_offical\data\excel\data_CDBR\kho_dong")

        savebutton = None
        savebutton = pyautogui.locateOnScreen(savebutton_path, confidence=0.8)

        if savebutton:
            pyautogui.click(pyautogui.center(savebutton))
            print("     Click nút save")

        return True
    
    except Exception as e:
        print(f"lỗi khi lấy dữ liệu Kho: {e}")
        return False 
