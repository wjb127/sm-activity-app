#!/bin/bash

# 현재 디렉토리 기준으로 실행
cd "$(dirname "$0")"  # 스크립트가 위치한 디렉토리로 이동

# streamlit 앱 실행
python3 -m streamlit run sm_activity_app.py  # Streamlit을 사용하여 SM Activity 앱 시작

