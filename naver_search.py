import subprocess
import time
import pyperclip
from tkinter import Tk, Label, Entry, Button, StringVar
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.options import Options
from openpyxl import Workbook

# 키워드를 입력받는 tkinter GUI 창 설정
def get_keywords():
    def on_submit():
        user_input = keyword_entry.get()
        window.destroy()  # 입력 후 창 닫기
        process_keywords(user_input)
    
    # tkinter 창 생성
    window = Tk()
    window.title("키워드 입력창")

    # 레이블과 입력창
    label = Label(window, text="검색할 키워드를 입력하세요 (쉼표로 구분):")
    label.pack(padx=10, pady=10)
    
    keyword_entry = Entry(window, width=50)
    keyword_entry.pack(padx=10, pady=10)
    
    # 확인 버튼
    submit_button = Button(window, text="확인", command=on_submit)
    submit_button.pack(padx=10, pady=10)

    window.mainloop()

# 키워드를 처리하는 함수
def process_keywords(input_keywords):
    keywords = [keyword.strip() for keyword in input_keywords.split(",")]

    # 입력한 키워드를 파일에 저장 (텍스트 파일로 저장)
    with open("입력된_키워드.txt", "w", encoding="utf-8") as f:
        f.write("\n".join(keywords))
        print("입력된 키워드가 '입력된_키워드.txt' 파일에 저장되었습니다.")

    # 엑셀 파일 생성 (Workbook)
    wb = Workbook()
    ws = wb.active
    ws.title = "추천 검색어"
    ws.append(["추천 검색어"])  # 엑셀 파일에 제목 열 추가

    # Chrome 옵션 설정
    chrome_options = Options()

    # Chrome 드라이버 설정 및 실행
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=chrome_options)

    # 각 키워드에 대해 검색 및 추천 검색어 추출
    for keyword in keywords:
        # 네이버 검색창 접속
        driver.get("https://www.naver.com")

        # 검색창 찾기
        search_box = driver.find_element("name", "query")

        # 검색어 입력
        search_box.clear()  # 검색창을 비우기
        search_box.send_keys(keyword)

        # 약간의 딜레이 (자동 추천 검색어가 나오기까지 기다림)
        time.sleep(2)

        # 추천 검색어가 로드되었는지 확인하고 추출
        try:
            # 추천 검색어 목록에서 텍스트 가져오기 ('.kwd_txt' 클래스 사용)
            suggestions = driver.find_elements("css selector", '.kwd_lst .kwd_txt')

            if suggestions:
                # 추천 검색어 텍스트 목록 만들기
                suggestion_texts = [suggestion.text for suggestion in suggestions]

                # 엑셀에 키워드와 추천 검색어를 추가
                for suggestion in suggestion_texts:
                    ws.append([suggestion])

                # 복사해서 클립보드에도 복사해보기 (선택 사항)
                suggestion_texts_combined = f"검색어: {keyword}\n" + '\n'.join(suggestion_texts) + "\n\n"
                pyperclip.copy(suggestion_texts_combined)
                print(f"{keyword}에 대한 추천 검색어가 엑셀에 추가되었습니다.")
            else:
                print(f"{keyword}에 대한 추천 검색어가 없습니다.")
        except Exception as e:
            print(f"{keyword}에 대한 추천 검색어를 가져오는 동안 오류 발생: {e}")

        # 각 검색어 사이에 약간의 텀을 둠
        time.sleep(2)

    # 엑셀 파일 경로 설정
    excel_file_path = "추천검색어.xlsx"

    # 엑셀 파일 저장
    wb.save(excel_file_path)

    # 엑셀 파일을 자동으로 열기
    subprocess.Popen(["start", "excel", excel_file_path], shell=True)

    # 브라우저 종료
    driver.quit()

    print("모든 키워드에 대한 추천 검색어가 엑셀에 저장되고, 엑셀 파일이 열렸습니다.")


# 프로그램 실행
get_keywords()
