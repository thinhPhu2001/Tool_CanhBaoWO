import pandas as pd
import pymysql
from sqlalchemy import create_engine
import traceback
import subprocess
import pyautogui
import pyotp
import time
from pynput.keyboard import Controller, Key
from time import sleep

from config import *


import sys

# Đảm bảo in tiếng Việt không bị lỗi
sys.stdout.reconfigure(encoding="utf-8")


def connect_to_db():
    """Hàm kết nối đến cơ sở dữ liệu MySQL."""
    pymysql.install_as_MySQLdb()
    try:
        engine = create_engine(DB_SERVER)
        connection = engine.connect()
        return connection
    except Exception as e:
        print("Lỗi kết nối Database:")
        traceback.print_exc()
        return None


def query_to_excel(connection, query, output_path):
    """Hàm thực hiện truy vấn SQL và xuất ra file Excel."""
    try:
        # Thực hiện truy vấn
        result = pd.read_sql(query, connection)

        # Xuất kết quả ra file Excel
        result.to_excel(output_path, index=False, engine="openpyxl")
        print(f"Dữ liệu đã được xuất ra file: {output_path}")
    except Exception as e:
        print("Lỗi khi thực hiện truy vấn hoặc xuất file:")
        traceback.print_exc()


# lấy dữ liệu CDBR
def connect_to_Db_CDBR():
    pymysql.install_as_MySQLdb()
    engine = create_engine(DB_SERVER)

    max_retries = 5  # Số lần thử lại tối đa
    retries = 0  # Đếm số lần thử
    connected = False

    while retries < max_retries and not connected:
        try:
            with engine.connect() as connection:
                print("Kết nối thành công, bắt đầu lấy dữ liệu")
                connected = True  # Đặt trạng thái thành công

                # Truy vấn và xử lý dữ liệu
                query_TKM = "select * from qlctkt.tkm_cdbr_open where `tỉnh/tp` = 'Tây Ninh' and `Dịch vụ` in ('SmartTV360 trả sau', 'FTTH', 'BoxTV360 trả sau', 'Camera', 'IPPhone', 'Multiscreen 2 chiều')"
                result_TKM = pd.read_sql(query_TKM, connection)
                result_TKM["ngày tạo công việc"] = pd.to_datetime(
                    result_TKM["ngày tạo công việc"]
                )
                result_TKM["ngày tạo công việc"] = result_TKM[
                    "ngày tạo công việc"
                ].dt.strftime("%d/%m/%Y %H:%M:%S")
                result_TKM.to_excel(DATA_GNOC_TKM_PATH, index=False, engine="openpyxl")
                print("Lấy dữ liệu TKM thành công!")

                query_PAKH = "SELECT * FROM bccs.pakh_ton where `Nguồn tiếp nhận` = 'Tây Ninh'"
                result_PAKH = pd.read_sql(query_PAKH, connection)
                result_PAKH.to_excel(
                    DATA_GNOC_PAKH_PATH, index=False, engine="openpyxl"
                )
                print("Lấy dữ liệu PAKH thành công!!!!!!")

                query_gnoc = "SELECT * FROM gnoc.gnoc_open_90d where `Loại công việc` = 'Chủ động xử lý port kém'"
                result_gnoc = pd.read_sql(query_gnoc, connection)
                result_gnoc.to_excel(
                    DATA_GNOC_logGnoc_PATH, index=False, engine="openpyxl"
                )
                print("Lấy dữ liệu log gnoc thành công!!!!!")

            return True

        except Exception as e:
            retries += 1
            print(
                f"Lỗi kết nối, thử lại lần {retries}/{max_retries}. Chi tiết lỗi: {e}"
            )
            time.sleep(5)  # Tạm dừng 5 giây trước khi thử lại

    if not connected:
        print("Không thể kết nối sau nhiều lần thử. Vui lòng kiểm tra hệ thống.")
        return False
