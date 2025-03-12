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

    print("\nBắt đầu gửi dữ liệu từ excel lên ggSheet!!!")
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

        print(f"    Hoàn thành cập nhật dữ liệu cho {sheet_name_excel}.")
        return True

    except FileNotFoundError:
        print(f"Tệp Excel không tồn tại tại {excel_path}")
        return False

    except Exception as e:
        print(f"Lỗi: {e}")
        return False

def ggSheet_to_excel(sheet_id, worksheet_name, excel_path, sheet_name_excel):
    """
    Hàm tải dữ liệu từ Google Sheets về tệp Excel có sẵn trên máy tính.
    Args:
        sheet_id: ID của Google Sheets.
        worksheet_name: Tên worksheet trong Google Sheets cần lấy dữ liệu.
        excel_path: Đường dẫn đến tệp Excel trên máy tính.
        sheet_name_excel: Tên sheet trong tệp Excel để ghi dữ liệu vào.
    """
    # Kết nối Google Sheets
    cre_path = SRC_PATH / "credentials.json"  # Đường dẫn đến tệp credentials.json
    scopes = ["https://www.googleapis.com/auth/spreadsheets.readonly"]
    creds = Credentials.from_service_account_file(cre_path, scopes=scopes)
    client = gspread.authorize(creds)

    # Mở Google Sheets và worksheet
    workbook = client.open_by_key(sheet_id)
    worksheet = workbook.worksheet(worksheet_name)

    # Lấy tất cả dữ liệu từ worksheet
    data = worksheet.get_all_values()

    # Chuyển dữ liệu thành DataFrame
    df = pd.DataFrame(data[1:], columns=data[0])

    # Ghi dữ liệu vào tệp Excel, ghi đè sheet nếu đã tồn tại
    with pd.ExcelWriter(
        excel_path, engine="openpyxl", mode="a", if_sheet_exists="replace"
    ) as writer:
        df.to_excel(writer, sheet_name=sheet_name_excel, index=False)