import streamlit as st  # Streamlit 라이브러리 불러오기 - 웹 인터페이스 구축
from openpyxl import Workbook, load_workbook  # 엑셀 파일 처리를 위한 라이브러리
from openpyxl.styles import Font, Alignment  # 엑셀 셀 서식 지정용 스타일 클래스
from datetime import datetime  # 날짜 및 시간 처리를 위한 라이브러리
import os  # 파일 및 디렉토리 조작을 위한 라이브러리

# 세션 상태 초기화 - 요청일과 작업일 동기화를 위한 설정
if 'req_date' not in st.session_state:
    st.session_state.req_date = datetime.today()

if 'work_date' not in st.session_state:
    st.session_state.work_date = datetime.today()

# 요청일이 변경될 때 작업일도 자동으로 업데이트하는 콜백 함수
def update_work_date():
    st.session_state.work_date = st.session_state.req_date

# Streamlit UI - 웹 애플리케이션 제목 설정
st.title("🛠 SM Activity 기록 프로그램")

# 파일 선택 옵션 - 사용자가 선택할 수 있는 엑셀 파일 옵션 정의
file_options = {
    "SM Activity - 대시보드": "data/SM_Activity_Dashboard.xlsx",
    "SM Activity - Plan": "data/SM_Activity_Plan.xlsx"
}

# 사용자가 작업할 파일 선택을 위한 드롭다운 생성
selected_file_name = st.selectbox(
    "작성할 문서 선택", 
    options=list(file_options.keys())
)

# 선택된 파일 경로 설정
file_path = file_options[selected_file_name]
sheet_name = "SM Activity"  # 모든 파일에 동일한 시트 이름 사용

# 엑셀 파일 헤더 설정 (모든 파일 형식 동일)
headers = [
    "NO", "월", "구분", "작업유형", "TASK", "요청일", "작업일",
    "요청자", "IT", "CNS", "개발자", "내용", "결과"
]

# 선택한 파일이 없으면 새로 생성하는 로직
if not os.path.exists(file_path):
    os.makedirs("data", exist_ok=True)  # data 디렉토리가 없으면 생성
    wb = Workbook()  # 새 엑셀 워크북 생성
    ws = wb.active  # 활성 워크시트 가져오기
    ws.title = sheet_name  # 워크시트 이름 설정
    # 헤더 행 추가 및 스타일 적용
    for col_num, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col_num, value=header)
        cell.font = Font(bold=True)  # 헤더 텍스트 굵게 설정
        cell.alignment = Alignment(horizontal="center", vertical="center")  # 가운데 정렬
    wb.save(file_path)  # 파일 저장

# 폼 외부에 날짜 선택 UI 배치 (콜백 함수 사용 가능)
st.subheader("📅 날짜 설정")
col1, col2 = st.columns(2)
with col1:
    st.date_input("요청일 선택", key="req_date", on_change=update_work_date)
with col2:
    st.date_input("작업일 확인", key="work_date", disabled=True)

# SM Activity 입력 양식 생성 (모든 파일 형식 동일)
with st.form("activity_form"):
    # 각 필드 입력 UI 요소 생성
    st.subheader("📝 작업 정보 입력")
    
    구분 = st.selectbox("구분", ["정기", "비정기"])  # 작업 구분 선택
    # 작업 유형 선택 드롭다운
    작업유형 = st.selectbox("작업유형", [
        "조간점검", "재적재", "인프라 작업", "SI 지원", "ERRC",
        "CCB", "적재", "시스템 운영", "월정기작업", "인수인계"
    ])
    task = st.text_input("TASK 제목")  # 작업 제목 입력
    
    # 담당자 정보를 한 줄에 4개 컬럼으로 배치
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        요청자 = st.text_input("요청자")  # 요청자 입력
    with col2:
        it = st.text_input("IT 담당자", value="한상욱")  # IT 담당자 입력(기본값 설정)
    with col3:
        cns = st.text_input("CNS 담당자", value="이정인")  # CNS 담당자 입력(기본값 설정)
    with col4:
        개발자 = st.text_input("개발자", value="위승빈")  # 개발자 입력(기본값 설정)
    
    결과 = st.selectbox("결과", ["진행 중", "완료", "보류", "기타"])  # 작업 결과 상태 선택

    # 양식 제출 버튼 생성
    submitted = st.form_submit_button("추가하기")

    # 양식이 제출되면 실행되는 로직
    if submitted:
        요청일 = st.session_state.req_date  # 폼 외부에서 설정한 요청일 사용
        작업일 = st.session_state.work_date  # 폼 외부에서 설정한 작업일 사용
        
        # 엑셀 파일 열기
        wb = load_workbook(file_path)
        ws = wb[sheet_name]
        new_row = ws.max_row + 1  # 새로운 데이터를 추가할 행 번호 계산

        # 입력된 데이터를 엑셀에 작성
        ws.cell(row=new_row, column=1, value=new_row - 1)  # NO 자동 번호 부여
        ws.cell(row=new_row, column=2, value=요청일.strftime("%Y%m"))  # 월 정보 (YYYYMM 형식)
        ws.cell(row=new_row, column=3, value=구분)  # 구분 데이터 추가
        ws.cell(row=new_row, column=4, value=작업유형)  # 작업유형 데이터 추가
        ws.cell(row=new_row, column=5, value=task)  # TASK 제목 데이터 추가
        ws.cell(row=new_row, column=6, value=요청일.strftime("%Y-%m-%d"))  # 요청일 형식 변환 후 추가
        ws.cell(row=new_row, column=7, value=작업일.strftime("%Y-%m-%d"))  # 작업일 형식 변환 후 추가
        ws.cell(row=new_row, column=8, value=요청자)  # 요청자 데이터 추가
        ws.cell(row=new_row, column=9, value=it)  # IT 담당자 데이터 추가
        ws.cell(row=new_row, column=10, value=cns)  # CNS 담당자 데이터 추가
        ws.cell(row=new_row, column=11, value=개발자)  # 개발자 데이터 추가
        ws.cell(row=new_row, column=12, value=task)  # 내용 컬럼에 TASK 제목 그대로 사용
        ws.cell(row=new_row, column=13, value=결과)  # 결과 데이터 추가
        
        # 요청일 기준 정렬을 위한 설정
        sort_col_idx = 5  # 요청일 컬럼 인덱스 (6번째 컬럼, 0부터 시작하므로 5)
        date_format = "%Y-%m-%d"  # 날짜 형식

        # 시트의 모든 데이터를 읽어 리스트에 저장 (빈 행 제외)
        data = []
        for row in ws.iter_rows(min_row=2, values_only=True):
            if all(cell is None for cell in row):  # 모든 셀이 비어있으면 건너뛰기
                continue
            data.append(row)
        
        # 요청일 기준으로 데이터 정렬
        data.sort(key=lambda x: datetime.strptime(str(x[sort_col_idx]), date_format) if x[sort_col_idx] else datetime.min)

        # 정렬된 데이터를 다시 엑셀에 쓰기
        for i, row_data in enumerate(data, start=2):
            for j, value in enumerate(row_data, start=1):
                ws.cell(row=i, column=j, value=value)

        # 정렬 후 남은 행이 있으면 내용 삭제 (중복 방지)
        for row in range(len(data) + 2, ws.max_row + 1):
            for col in range(1, ws.max_column + 1):
                ws.cell(row=row, column=col, value=None)

        # 변경사항 저장 및 성공 메시지 표시
        wb.save(file_path)
        st.success(f"✅ {selected_file_name} 파일에 성공적으로 추가되었고, 날짜 순으로 정렬되었습니다.")

# 엑셀 파일 다운로드 버튼 생성
with open(file_path, "rb") as f:
    st.download_button(
        label=f"📥 {selected_file_name} 엑셀 다운로드",
        data=f,
        file_name=os.path.basename(file_path),
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
