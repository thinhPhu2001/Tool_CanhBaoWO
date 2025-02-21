import os
import sys
from time import sleep

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

from config import *
from utils import *

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
}


class BrowserManager:
    def __init__(self):
        self.driver = None

    def start_browser(self, profile_path):
        print("dang mo trinh duyet")
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
            print("gui thanh cong")
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
            message_box.send_keys(message)
            message_box.send_keys(Keys.CONTROL, "v")
            sleep(5)

            # nút gửi tin nhắn
            send_button = WebDriverWait(self.driver, 10).until(
                EC.element_to_be_clickable(
                    (
                        By.XPATH,
                        XPATHS_WHATSAPP["send_button"],
                    )
                )
            )
            send_button.click()
            print("gui thanh cong")
            sleep(3)

        except Exception as e:
            print(f"Lỗi trong quá trình gửi tin nhắn CDBR: {e}")

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

        except Exception as e:
            print(e)

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

    def send_attached_img_message(self, message, file_path, tag_name=None):
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
            message_box = WebDriverWait(self.driver, 20).until(
                EC.presence_of_element_located((By.XPATH, XPATHS_ZALO["message_box"]))
            )
            message_box.click()
            message_box.send_keys(Keys.CONTROL, "a", Keys.BACKSPACE)
            message_box.send_keys(Keys.CONTROL, "v")  # dán hình vào ô tin nhắn
            sleep(2)

            if tag_name:
                # ghi tin nhắn + tag all
                message_box.send_keys(message)
                message_box.send_keys(" @")
                message_box.send_keys(tag_name)
                sleep(2)
                message_box.send_keys(Keys.ARROW_DOWN)
                message_box.send_keys(Keys.ENTER)
                sleep(2)
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
            message_box.send_keys(message)
            message_box.send_keys(Keys.CONTROL, "v")
            sleep(2)
            message_box.send_keys(Keys.ENTER)
            sleep(5)
        except Exception as e:
            print(e)

    def check_last_message_status(self):
        """Kiểm tra trạng thái tin nhắn sau khi gửi trên Zalo Web."""
        try:
            if self.element_is_present(
                '[data-translate-inner="STR_RECEIVED"]', By.CSS_SELECTOR, 10
            ):
                return "sent"

            start_time = time.time()
            max_wait_time = 60

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
