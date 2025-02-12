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

# Lấy đường dẫn thư mục gốc của project
project_dir = Path(__file__).resolve().parent.parent  # Thư mục cha của 'src'

# Đường dẫn thư mục
IMAGE_PATH = project_dir / "data" / "img"
EXCEL_PATH = project_dir / "data" / "excel"

# thư mục dữ liệu DI ĐỘNG
DATA_GNOC_RAW_PATH = EXCEL_PATH / "data_didong" / "data_GNOC_raw.xlsx"
DATA_TOOL_MANAGEMENT_PATH = EXCEL_PATH / "data_didong" / "Log_ton_GNOC_date.xlsb"
DATA_CONFIG_PATH = EXCEL_PATH / "data_didong" / "config.xlsx"

# thư mục dữ liệu CĐBR
DATA_GNOC_PAKH_PATH = EXCEL_PATH / "data_CDBR" / "Log ton PAKH.xlsx"
DATA_GNOC_TKM_PATH = EXCEL_PATH / "data_CDBR" / "Log ton TKM.xlsx"
DATA_GNOC_logGnoc_PATH = EXCEL_PATH / "data_CDBR" / "Log gnoc.xlsx"
DATA_TOOL_MANAGEMENT_CDBR_PATH = EXCEL_PATH / "data_CDBR" / "TNH.xlsm"
 
# LƯU HÌNH CẢNH BÁO DI ĐỘNG
CNCT_IMG_PATH = EXCEL_PATH / "img" / "tth"
USER_IMG_PATH = EXCEL_PATH / "img" / "cum_huyen"

SRC_PATH = project_dir / "src"
MAIN_PATH = project_dir / "src" / "Main_Auto.py"

# # Tìm tất cả các file có đuôi .ovpn trong thư mục "opvn"
# ovpn_files = glob.glob(str(project_dir / "data" / "opvn" / "*.ovpn"))
# if ovpn_files is None:
#     print(f"Không tìm thấy file ovpn")
# else:
#     OPEN_VPN_PROFILE_PATH = Path(ovpn_files[0])
#     print(f"đường dẫn ovpn profile: {OPEN_VPN_PROFILE_PATH}")

# Đọc file cấu hình OpenVPN
OPEN_VPN_CONFIG_PATH = project_dir / "config.txt"
with open(OPEN_VPN_CONFIG_PATH, "r") as file:
    # Lặp qua các dòng của file ngay khi mở
    for line in file:
        if "my phone:" in line:
            PHONE_NUMBER = line.strip().replace("my phone:", "").strip()
            print(f"my phone: = {PHONE_NUMBER}")

        if "pwd:" in line:
            OTP_SECRET = line.strip().replace("pwd:", "").strip()
            print(f"OTP_SECRET = {OTP_SECRET}")

        if "opvn_profile:" in line:
            # OPEN_VPN_PROFILE_PATH = line.strip().replace("opvn_profile:", "").strip()
            OPEN_VPN_CONFIG_PATH = line.strip().replace("opvn_profile:", "").strip()
            print(f"opvn_profile = {OPEN_VPN_CONFIG_PATH}")

        if "path:" in line:
            OPEN_VPN_PATH = line.strip().replace("path:", "").strip()
            print(f"path: {OPEN_VPN_PATH}")

        if "sendBY:" in line:
            SENDBY = line.strip().replace("sendBY:", "").strip()
            print(f"SENDBY: {SENDBY}")

        if "didong:" in line:
            CHROME_PROFILE_DI_DONG_PATH = line.strip().replace("didong:", "").strip()
            print(f"CHROME_PROFILE_DI_DONG_PATH: {CHROME_PROFILE_DI_DONG_PATH}")

        if "cdbr:" in line:
            CHROME_PROFILE_CDBR_PATH = line.strip().replace("cdbr:", "").strip()
            print(f"CHROME_PROFILE_CDBR_PATH: {CHROME_PROFILE_CDBR_PATH}")

        if "server:" in line:
            DB_SERVER = line.strip().replace("server:", "").strip()
            print(f"DB server: {DB_SERVER}")
