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
from datetime import datetime, timedelta

from config import *

import sys

# Đảm bảo in tiếng Việt không bị lỗi
sys.stdout.reconfigure(encoding="utf-8")

time_shift = {
    "First": {"start": "06:20"},
    "Second": {"start": "08:30"},
    "Third": {"start": "10:30"},
    "Fourth": {"start": "12:10"},
    "Fifth": {"start": "14:20"},
    "Sixth": {"start": "16:20"},
    "Seventh": {"start": "19:20"},
}


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

    tinh = pd.read_excel(DATA_CONFIG_CDBR_PATH, sheet_name="Sheet1", header=0).iloc[0][
        "Tỉnh"
    ]

    while retries < max_retries and not connected:
        try:
            with engine.connect() as connection:
                print("Kết nối thành công, bắt đầu lấy dữ liệu")
                connected = True  # Đặt trạng thái thành công

                # Truy vấn và xử lý dữ liệu
                query_TKM = f"select * from qlctkt.tkm_cdbr_open where `tỉnh/tp` = '{tinh}' and `Dịch vụ` in ('SmartTV360 trả sau', 'FTTH', 'BoxTV360 trả sau', 'Camera', 'IPPhone', 'Multiscreen 2 chiều')"
                result_TKM = pd.read_sql(query_TKM, connection)
                result_TKM["ngày tạo công việc"] = pd.to_datetime(
                    result_TKM["ngày tạo công việc"]
                )
                result_TKM["ngày tạo công việc"] = result_TKM[
                    "ngày tạo công việc"
                ].dt.strftime("%d/%m/%Y %H:%M:%S")
                result_TKM.to_excel(DATA_GNOC_TKM_PATH, index=False, engine="openpyxl")
                print("Lấy dữ liệu TKM thành công!")

                query_PAKH = (
                    f"SELECT * FROM bccs.pakh_ton where `Nguồn tiếp nhận` = '{tinh}'"
                )
                result_PAKH = pd.read_sql(query_PAKH, connection)
                result_PAKH.to_excel(
                    DATA_GNOC_PAKH_PATH, index=False, engine="openpyxl"
                )
                print("Lấy dữ liệu PAKH thành công!!!!!!")

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


def check_old_data(file_excel_path, time_geted):
    """
    Kiểm tra dữ liệu mới hay cũ
    Args:
        file_excel_path: đường dẫn đến file excel (dữ liệu gnoc)
        time_geted: thời gian trong quy định lấy dữ liệu là dữ liệu mới
    """
    current_time = datetime.now()
    # Read only the first row of the specified sheet
    df = pd.read_excel(file_excel_path, sheet_name="Sheet1", nrows=1)

    # Extract the 'last sync' value from the first row
    last_sync = df.at[0, "last sync"]

    time_difference = current_time - last_sync
    if time_difference > timedelta(hours=time_geted):
        return False
    else:
        return True


def trans_type_timeShift(start_time_str):
    """
    Chuyển đổi chuỗi thời gian bắt đầu của ca làm việc thành đối tượng datetime.

    Args:
        start_time_str (str): Thời gian bắt đầu của ca làm việc dưới dạng chuỗi "HH:MM".

    Returns:
        datetime: Đối tượng datetime tương ứng với thời gian bắt đầu trong ngày hiện tại.
    """
    # Lấy ngày hiện tại
    current_date = datetime.now().date()
    # Chuyển đổi chuỗi thời gian thành đối tượng time
    start_time = datetime.strptime(start_time_str, "%H:%M").time()
    # Kết hợp ngày hiện tại với thời gian bắt đầu để tạo đối tượng datetime
    return datetime.combine(current_date, start_time)


def Check_shift(timeShift):
    """
    Xử lý thông tin các ca làm việc.

    Args:
        timeShift (dict): Từ điển chứa thông tin các ca làm việc, với mỗi ca có thời gian bắt đầu
                          được biểu diễn dưới dạng chuỗi ký tự "HH:MM".
    """
    # Lấy thời gian hiện tại
    current_time = datetime.now()

    # Chuyển đổi thời gian bắt đầu của các ca thành datetime
    shift_times = {
        shift: trans_type_timeShift(info["start"]) for shift, info in timeShift.items()
    }

    # Danh sách các ca làm việc được sắp xếp theo thời gian bắt đầu
    sorted_shifts = sorted(shift_times.items(), key=lambda x: x[1])

    # Xác định ca làm việc hiện tại
    current_shift = None
    for i in range(len(sorted_shifts) - 1):
        shift_name, start_time = sorted_shifts[i]
        next_shift_name, next_start_time = sorted_shifts[i + 1]

        if start_time <= current_time < next_start_time:
            current_shift = shift_name
            break

    # Nếu thời gian hiện tại lớn hơn ca cuối cùng => Là ca cuối cùng
    if current_time >= sorted_shifts[-1][1]:
        current_shift = sorted_shifts[-1][0]

    # Tìm ca làm việc trước đó
    if current_shift:
        current_index = next(
            i for i, (name, _) in enumerate(sorted_shifts) if name == current_shift
        )
        previous_index = (current_index - 1) % len(
            sorted_shifts
        )  # Quay vòng về cuối nếu đang ở đầu
        previous_shift = sorted_shifts[previous_index][0]

        return previous_shift, current_shift
    else:
        print("Không có ca làm việc nào phù hợp.")
        return None


def check_old_data_Didong(file_excel_path):
    """
    Kiểm tra dữ liệu Di động có phải data cũ không

    Args:
        file_excel_path (str): đường dẫn tới file dữ liệu cần kiểm tra
    """
    try:
        previous_shift, current_shift = Check_shift(timeShift=time_shift)

        if previous_shift is None or current_shift is None:
            print("Không xác định được ca làm việc hiện tại hoặc trước đó.")
            return False

        time_previous_shift = trans_type_timeShift(time_shift[previous_shift]["start"])

        current_time = datetime.now()
        # Read only the first row of the specified sheet
        df = pd.read_excel(file_excel_path, sheet_name="Sheet1", nrows=1)

        # Extract the 'last sync' value from the first row
        last_sync = df.at[0, "last sync"]

        num_days = (last_sync.date() - current_time.date()).days

        current_shift_name = current_shift

        # So sánh chênh lệch ngày lấy dữ liệu
        if num_days <= -2:
            return False

        elif num_days == -1:
            if current_shift_name == "First":
                # so sánh thời gian lấy dữ liệu với ca chạy trước đó
                if last_sync.time() > time_previous_shift.time():
                    return True
                else:
                    return False
            else:
                return False

        elif num_days == 0:  # trùng ngày
            if current_shift_name == "First":
                return True

            # so sánh thời gian lấy dữ liệu với ca chạy trước đó
            if last_sync.time() > time_previous_shift.time():
                return True
            else:
                return False

    except Exception as e:
        print(f"Lỗi kiểm tra dữ liệu của di động (old/new): {e}")
        return False
