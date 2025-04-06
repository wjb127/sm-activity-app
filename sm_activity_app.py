import streamlit as st
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, Alignment
from datetime import datetime
import os

# Streamlit UI
st.title("🛠 SM Activity 기록 프로그램")

# 파일 선택 옵션
file_options = {
    "SM Activity - 대시보드": "data/SM_Activity_Dashboard.xlsx",
    "SM Activity - Plan": "data/SM_Activity_Plan.xlsx"
}

selected_file_name = st.selectbox(
    "작성할 문서 선택", 
    options=list(file_options.keys())
)

file_path = file_options[selected_file_name]
sheet_name = "SM Activity"  # 모든 파일에 동일한 시트 이름 사용

# 헤더 설정 (모든 파일 형식 동일)
headers = [
    "NO", "월", "구분", "작업유형", "TASK", "요청일", "작업일",
    "요청자", "IT", "CNS", "개발자", "내용", "결과"
]

# 파일 없으면 생성
if not os.path.exists(file_path):
    os.makedirs("data", exist_ok=True)
    wb = Workbook()
    ws = wb.active
    ws.title = sheet_name
    for col_num, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col_num, value=header)
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center")
    wb.save(file_path)

# 양식 (모든 파일 형식 동일)
with st.form("activity_form"):
    요청일 = st.date_input("요청일", value=datetime.today())
    작업일 = st.date_input("작업일", value=datetime.today())
    구분 = st.selectbox("구분", ["정기", "비정기"])
    작업유형 = st.selectbox("작업유형", [
        "조간점검", "재적재", "인프라 작업", "SI 지원", "ERRC",
        "CCB", "적재", "시스템 운영", "월정기작업", "인수인계"
    ])
    task = st.text_input("TASK 제목")
    요청자 = st.text_input("요청자")
    it = st.text_input("IT 담당자", value="한상욱")
    cns = st.text_input("CNS 담당자", value="이정인")
    개발자 = st.text_input("개발자", value="위승빈")
    내용 = st.text_area("내용")
    결과 = st.selectbox("결과", ["진행 중", "완료", "보류", "기타"])

    submitted = st.form_submit_button("추가하기")

    if submitted:
        # 엑셀 열기
        wb = load_workbook(file_path)
        ws = wb[sheet_name]
        new_row = ws.max_row + 1

        # 작성
        ws.cell(row=new_row, column=1, value=new_row - 1)  # NO
        ws.cell(row=new_row, column=2, value=요청일.strftime("%Y%m"))  # 월
        ws.cell(row=new_row, column=3, value=구분)
        ws.cell(row=new_row, column=4, value=작업유형)
        ws.cell(row=new_row, column=5, value=task)
        ws.cell(row=new_row, column=6, value=요청일.strftime("%Y-%m-%d"))
        ws.cell(row=new_row, column=7, value=작업일.strftime("%Y-%m-%d"))
        ws.cell(row=new_row, column=8, value=요청자)
        ws.cell(row=new_row, column=9, value=it)
        ws.cell(row=new_row, column=10, value=cns)
        ws.cell(row=new_row, column=11, value=개발자)
        ws.cell(row=new_row, column=12, value=내용)
        ws.cell(row=new_row, column=13, value=결과)
        
        # 정렬 기준 컬럼과 포맷
        sort_col_idx = 5  # 요청일 컬럼 인덱스
        date_format = "%Y-%m-%d"

        # 날짜 기준 정렬
        data = []
        for row in ws.iter_rows(min_row=2, values_only=True):
            if all(cell is None for cell in row):
                continue
            data.append(row)
        
        # 해당 컬럼 기준으로 정렬
        data.sort(key=lambda x: datetime.strptime(str(x[sort_col_idx]), date_format) if x[sort_col_idx] else datetime.min)

        # 다시 쓰기
        for i, row_data in enumerate(data, start=2):
            for j, value in enumerate(row_data, start=1):
                ws.cell(row=i, column=j, value=value)

        # 남은 행 초기화
        for row in range(len(data) + 2, ws.max_row + 1):
            for col in range(1, ws.max_column + 1):
                ws.cell(row=row, column=col, value=None)

        wb.save(file_path)
        st.success(f"✅ {selected_file_name} 파일에 성공적으로 추가되었고, 날짜 순으로 정렬되었습니다.")

# 다운로드 버튼
with open(file_path, "rb") as f:
    st.download_button(
        label=f"📥 {selected_file_name} 엑셀 다운로드",
        data=f,
        file_name=os.path.basename(file_path),
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
