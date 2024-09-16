import os
import re
import speech_recognition as sr
from pydub import AudioSegment
import openpyxl

# 1. MP3 파일에서 처음 10초 추출하는 함수
def extract_first_10_seconds(file_path):
    audio = AudioSegment.from_mp3(file_path)
    first_10_seconds = audio[:10000]  # 10초 추출
    first_10_seconds.export("first_10_sec.wav", format="wav")
    return "first_10_sec.wav"

# 2. Google STT API를 사용해 음성 텍스트 변환
def transcribe_audio(file_path):
    recognizer = sr.Recognizer()
    with sr.AudioFile(file_path) as source:
        audio = recognizer.record(source)
    text = recognizer.recognize_google(audio, language='ko-KR')
    return text

# 3. 변환된 텍스트에서 이름 추출
def extract_name(text):
    match = re.search(r'대한민국 홍익인간 (\w+)입니다', text)
    if match:
        return match.group(1)  # 이름 부분만 추출
    return None

# 4. MP3 파일에서 이름 추출 후 엑셀 파일에 저장
def process_mp3_files_to_excel(folder_path):
    # 엑셀 파일 생성
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Name Extraction"
    ws.append(["Original File Name", "Extracted Name"])  # 첫 번째 행에 헤더 추가

    # 폴더 내 MP3 파일 처리
    for filename in os.listdir(folder_path):
        if filename.endswith(".mp3"):
            file_path = os.path.join(folder_path, filename)
            
            # 처음 10초 추출 및 음성 인식
            wav_file = extract_first_10_seconds(file_path)
            transcribed_text = transcribe_audio(wav_file)
            name = extract_name(transcribed_text)
            
            # 엑셀에 파일명과 추출된 이름 저장
            ws.append([filename, name if name else "추출 실패"])

    # 엑셀 파일을 데스크탑에 저장
    excel_file_path = os.path.join(os.path.expanduser("~/Desktop"), "mp3_name_extraction.xlsx")
    wb.save(excel_file_path)
    print(f"엑셀 파일이 데스크탑에 저장되었습니다: {excel_file_path}")

# 예시: mp3STT 폴더 내 모든 MP3 파일 처리
mp3_folder_path = os.path.expanduser("~/Desktop/MP3_Processing/mp3STT")
process_mp3_files_to_excel(mp3_folder_path)
