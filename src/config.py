### CÀI ĐẶT CẤU HÌNH - VD: TÀI KHOẢN WHATSAPP, PHONE - NUMBER

from pathlib import Path
import glob
import pandas as pd

import sys

# thay đôi môi trường tiếng Việt
sys.stdout.reconfigure(encoding="utf-8")

# Đường dẫn URL
WHATSAPP_URL = "https://web.whatsapp.com/"
ZALO_URL = "https://id.zalo.me/account?continue=https%3A%2F%2Fchat.zalo.me%2F"
OUT_LOOK_URL = "https://outlook.office.com/mail/"
GNOC_URL = "http://10.255.58.201:9000/#/dashboard"

# Lấy đường dẫn thư mục gốc của project
project_dir = Path(__file__).resolve().parent.parent  # Thư mục cha của 'src'

# Đường dẫn thư mục
IMAGE_PATH = project_dir / "data" / "img"
EXCEL_PATH = project_dir / "data" / "excel"

# thư mục dữ liệu DI ĐỘNG
DATA_GNOC_RAW_PATH = EXCEL_PATH / "data_didong" / "data_GNOC_raw.xlsx"
DATA_TOOL_MANAGEMENT_PATH = EXCEL_PATH / "data_didong" / "Log_ton_GNOC_date.xlsb"
DATA_DiDong_CONFIG_PATH = EXCEL_PATH / "data_didong" / "config.xlsb"
DATA_DIDONG_ChatBot_PATH = (
    EXCEL_PATH / "data_didong" / "BDG_CANH BAO WO _BDG_ALL GNOC T01_2025.xlsb"
)
DATA_WO_DONG_PATH = EXCEL_PATH / "data_didong" / "WO_dong"

# thư mục dữ liệu CĐBR
DATA_GNOC_PAKH_PATH = EXCEL_PATH / "data_CDBR" / "Log ton PAKH.xlsx"
DATA_GNOC_TKM_PATH = EXCEL_PATH / "data_CDBR" / "Log ton TKM.xlsx"
DATA_GNOC_logGnoc_PATH = EXCEL_PATH / "data_CDBR" / "Log gnoc.xlsx"
DATA_CONFIG_CDBR_PATH = EXCEL_PATH / "data_CDBR" / "config.xlsx"
DATA_TOOL_MANAGEMENT_CDBR_PATH = EXCEL_PATH / "data_CDBR" / "CDBR.xlsm"

# LƯU HÌNH CẢNH BÁO DI ĐỘNG
CNCT_IMG_PATH = EXCEL_PATH / "img" / "tth"
USER_IMG_PATH = EXCEL_PATH / "img" / "cum_huyen"

SRC_PATH = project_dir / "src"
MAIN_PATH = project_dir / "src" / "Main_Auto.py"

# Credentials Email API
cre_path = project_dir / "credentials.json"

# Đọc file cấu hình OpenVPN
OPEN_VPN_CONFIG_PATH = project_dir / "config.txt"

# Định nghĩa từ điển để lưu trữ các giá trị cấu hình
config = {
    "my phone:": None,
    "pwd:": None,
    "opvn_profile:": None,
    "path:": None,
    "sendBY:": None,
    "didong:": None,
    "cdbr:": None,
    "server:": None,
    "GGSheet ID:": None,
    "user_name Gnoc:": None,
    "password Gnoc:": None,
    "otp gnoc:": None,
}

with open(OPEN_VPN_CONFIG_PATH, "r") as file:
    for line in file:
        for key in config.keys():
            if key in line:
                config[key] = line.replace(key, "").strip()
                break  # Dừng vòng lặp khi tìm thấy khóa phù hợp

# Gán giá trị từ từ điển cho các biến
PHONE_NUMBER = config["my phone:"]
OTP_SECRET = config["pwd:"]
OPEN_VPN_PROFILE_PATH = config["opvn_profile:"]
OPEN_VPN_PATH = config["path:"]
SENDBY = config["sendBY:"]
CHROME_PROFILE_DI_DONG_PATH = config["didong:"]
CHROME_PROFILE_CDBR_PATH = config["cdbr:"]
DB_SERVER = config["server:"]
GG_SHEET_ID = config["GGSheet ID:"]
USERNAME_GNOC = config["user_name Gnoc:"]
PASSWORD_GNOC = config["password Gnoc:"]
OTP_GNOC = config["otp gnoc:"]