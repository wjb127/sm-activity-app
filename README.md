# 🧾 SM Activity 자동화 기록 프로그램

업무 중 반복적으로 작성해야 하는 SM Activity 정리를 간편하게 입력하고 엑셀로 자동 저장해주는 Streamlit 기반의 웹 앱입니다.

## ✅ 기능
- 작성할 문서 선택 기능 (SM Activity - 대시보드 또는 SM Activity - Plan)
- 요청일, 구분, 작업유형, 요청자 정보를 입력하면 자동으로 엑셀에 저장
- 작성일자는 자동 기입
- 날짜 기준으로 자동 정렬
- 최신 엑셀 파일을 바로 다운로드 가능

## 📁 폴더 구조
```
sm_activity_app/
├── sm_activity_app.py                  # Streamlit 실행 파일
├── data/
│   ├── SM_Activity_Dashboard.xlsx      # SM Activity 대시보드용 엑셀 파일
│   └── SM_Activity_Plan.xlsx           # SM Activity 계획용 엑셀 파일
├── run_streamlit.bat                   # 윈도우 실행 배치 파일
└── README.md
```

## ▶️ 실행 방법

### 1. 가상환경 (선택)
```bash
python -m venv venv
.\venv\Scripts\activate
```

### 2. 필수 라이브러리 설치
```bash
pip install streamlit openpyxl
```

### 3. 앱 실행
```bash
python -m streamlit run sm_activity_app.py
```
또는 Windows 환경에서는 `run_streamlit.bat` 파일을 실행하면 됩니다.

실행 후 아래 주소로 접속:
```
http://localhost:8501
```

## 💡 실행 시 주의사항
- 엑셀 파일은 `data/` 폴더에 자동으로 생성됩니다.
- 문서 유형을 선택하여 저장할 파일을 선택할 수 있습니다:
  - SM Activity - 대시보드: data/SM_Activity_Dashboard.xlsx
  - SM Activity - Plan: data/SM_Activity_Plan.xlsx
- 두 파일 모두 동일한 양식을 가지고 있으며, 같은 형식으로 데이터가 저장됩니다.
- 기존 파일이 없으면 자동으로 생성되며, 기존 파일이 있으면 이어서 작성됩니다.
- 모든 데이터는 요청일 기준으로 자동 정렬됩니다.

## 🔒 개인정보 보호
Streamlit은 기본적으로 익명 사용 통계를 수집합니다. 원하지 않는 경우 아래 파일을 생성하여 사용 통계 수집을 비활성화할 수 있습니다:

```
%userprofile%/.streamlit/config.toml
```

```toml
[browser]
gatherUsageStats = false
``` 