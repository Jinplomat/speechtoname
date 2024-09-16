import os
import openpyxl

# 1. 엑셀 파일에서 수정된 이름을 불러와 MP3 파일명 변경
def rename_mp3_files_from_excel(folder_path, excel_file_path):
    # 엑셀 파일 열기
    wb = openpyxl.load_workbook(excel_file_path)
    ws = wb.active

    # 엑셀 파일의 두 번째 행부터 데이터를 읽어옴
    for row in ws.iter_rows(min_row=2, values_only=True):
        original_file_name, new_name = row

        # 파일 경로 설정
        original_file_path = os.path.join(folder_path, original_file_name)

        # 새로운 이름이 존재하고 파일이 존재할 때만 파일명 변경
        if new_name and os.path.exists(original_file_path):
            # 새 파일명 설정 (확장자는 .mp3로 유지)
            new_file_path = os.path.join(folder_path, f"{new_name}.mp3")

            # 파일명 변경
            os.rename(original_file_path, new_file_path)
            print(f"파일명이 '{new_file_path}'로 변경되었습니다.")

# 예시: 엑셀 파일을 바탕화면에서 불러와 mp3STT 폴더의 파일명 변경
mp3_folder_path = os.path.expanduser("~/Desktop/MP3_Processing/mp3STT")
excel_file_path = os.path.expanduser("~/Desktop/mp3_name_extraction.xlsx")

rename_mp3_files_from_excel(mp3_folder_path, excel_file_path)
