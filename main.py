import streamlit as st
from streamlit_drawable_canvas import st_canvas
from PIL import Image
import io
import base64
import pandas as pd
from datetime import datetime
from streamlit_gsheets import GSheetsConnection

# 1. 시트 연결
conn = st.connection("gsheets", type=GSheetsConnection)

# 2. 서명 인코딩 함수
def process_signature(canvas_data):
    img = Image.fromarray(canvas_data.astype('uint8'), 'RGBA')
    # 크기 최적화 (직인은 클 필요가 없음)
    img.thumbnail((200, 200))
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    return base64.b64encode(buf.getvalue()).decode()

# 3. UI 구성
st.title("✍️ 교사 직인 등록 시스템")
teacher_name = st.text_input("교사 성함을 입력하세요")

canvas_result = st_canvas(
    fill_color="rgba(255, 255, 255, 0)",
    stroke_width=3,
    stroke_color="#000000",
    background_color="#eeeeee",
    height=200,
    width=200,
    drawing_mode="freedraw",
    key="teacher_sig"
)

if st.button("직인 저장하기"):
    if canvas_result.image_data is not None and teacher_name:
        sig_b64 = process_signature(canvas_result.image_data)
        
        # 시트에 저장할 데이터 생성
        new_row = pd.DataFrame([{
            "이름": teacher_name,
            "직인데이터": f"'{sig_b64}", # 문자열 강제 지정
            "등록일": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        }])
        
        # 데이터 업데이트 로직
        existing_data = conn.read(worksheet="교사직인")
        updated_data = pd.concat([existing_data, new_row], ignore_index=True)
        conn.update(worksheet="교사직인", data=updated_data)
        st.success("직인이 성공적으로 등록되었습니다!")
