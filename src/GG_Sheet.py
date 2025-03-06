from config import *
from excel_handler import *

import gspread
from google.oauth2.service_account import Credentials
import pandas as pd
import sys

# thay đổi môi trường tiếng Việt
sys.stdout.reconfigure(encoding="utf-8")


def excel_to_ggSheet(sheet_id, new_worksheet_name, excel_path, sheet_name_excel):
    """
    Hàm chuyển dữ liệu excel dưới máy lên gg sheet
    Args:
        sheet_id: mã id của link ggsheet
        new_worksheet_name: tên sheet của ggsheet cần truyền dữ liệu vào
        excel_path: đường dẫn excel lấy dữ liệu đi
        sheet_name: tên sheet của file excel để lấy dữ liệu
    """
    # Kết nối Google Sheet
    scopes = ["https://www.googleapis.com/auth/spreadsheets"]
    creds = Credentials.from_service_account_file(cre_path, scopes=scopes)
    client = gspread.authorize(creds)

    workbook = client.open_by_key(sheet_id)

    # Đọc danh sách tên sheet
    worksheet_list = list(map(lambda x: x.title, workbook.worksheets()))

    # Kiểm tra tên sheet
    if new_worksheet_name not in worksheet_list:
        raise ValueError(
            f"Worksheet '{new_worksheet_name}' không tồn tại trong Google Sheet."
        )

    # Mở sheet mong muốn
    sheet = workbook.worksheet(new_worksheet_name)

    print("    Bắt đầu gửi dữ liệu từ excel lên ggSheet!!!")
    try:
        # Đọc dữ liệu từ tệp Excel
        df = pd.read_excel(excel_path, sheet_name=sheet_name_excel)
        # Thay thế NaN, infinity và -infinity
        df = df.replace([float("nan"), float("inf"), float("-inf")], "").fillna("")

        # Xử lý kiểu datetime và thay NaT bằng chuỗi rỗng
        for col in df.select_dtypes(include=["datetime64[ns]", "datetime"]).columns:
            df[col] = df[col].apply(
                lambda x: x.strftime("%Y-%m-%d %H:%M:%S") if pd.notnull(x) else ""
            )

        sheet.clear()  # Xóa dữ liệu cũ trong Google Sheet

        sheet.update(
            [df.columns.values.tolist()] + df.values.tolist()
        )  # Cập nhật dữ liệu

        print("Hoàn thành cập nhật dữ liệu.")
    except FileNotFoundError:
        print(f"Tệp Excel không tồn tại tại {excel_path}")
    except Exception as e:
        print(f"Lỗi: {e}")
