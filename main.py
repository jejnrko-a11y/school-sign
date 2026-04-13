import streamlit as st
from streamlit_drawable_canvas import st_canvas
from streamlit_gsheets import GSheetsConnection
import pandas as pd
from datetime import datetime
from PIL import Image
import io
import base64

# --- 앱 설정 ---
st.set_page_config(page_title="경기기공 재구조화 동의서", layout="centered")

# 구글 시트 연결
conn = st.connection("gsheets", type=GSheetsConnection)

# --- 제목 및 안내 문구 ---
st.title("2026년 경기기계공업고등학교\n재구조화 교직원 동의서")

with st.container(border=True):
    st.subheader("📢 재구조화 사업 안내")
    # [여기에 내용을 채우시면 업데이트 됩니다]
    st.info("""
    본 안내문은 2026년 추진 예정인 경기기계공업고등학교 학교 재구조화 사업에 대한 
    교직원의 의견을 수렴하고 동의를 얻기 위한 목적으로 작성되었습니다.
    
    [내용 요약]
    1. 학과 개편 및 신설 추진 내용...
    2. 교육 환경 개선 및 실습동 리모델링...
    3. 교직원 역량 강화 및 배치 계획...
    
    위와 같은 학교의 변화와 발전을 위한 재구조화 사업에 대해 충분한 설명을 들었으며, 
    이에 동의하시는 선생님께서는 아래에 서명해 주시기 바랍니다.
    """)

# --- 데이터 불러오기 ---
try:
    # 시트 전체를 읽어옵니다.
    df = conn.read(worksheet="동의서양식", ttl=0)
    df = df.fillna("") # 빈 칸을 빈 문자열로 처리
    staff_list = df['성함'].tolist()
except Exception as e:
    st.error(f"시트를 불러오는 중 오류가 발생했습니다: {e}")
    st.stop()

# --- 입력 섹션 ---
st.markdown("### ✍️ 서명하기")
selected_name = st.selectbox("본인의 성함을 선택해 주세요", ["선택하세요"] + staff_list)

if selected_name != "선택하세요":
    # 덮어쓰기 여부 확인
    # 선택한 이름이 있는 행의 '서명' 열 확인
    user_row = df[df['성함'] == selected_name]
    existing_sig = user_row['서명'].values[0]

    if existing_sig:
        st.warning(f"⚠️ {selected_name} 선생님은 이미 서명을 제출하셨습니다. 다시 서명하시면 기존 서명을 새 서명으로 덮어씁니다.")
    else:
        st.success(f"✅ {selected_name} 선생님, 동의하신다면 아래 칸에 서명해 주세요.")

    # 서명 패드
    canvas_result = st_canvas(
        fill_color="rgba(255, 255, 255, 0)",
        stroke_width=3,
        stroke_color="#000000",
        background_color="#f0f2f6", # 서명란을 연한 회색으로 강조
        height=150,
        width=300,
        drawing_mode="freedraw",
        key="staff_sig",
    )

    if st.button("서명 제출하기", use_container_width=True):
        if canvas_result.image_data is not None:
            with st.spinner("서명을 안전하게 기록 중입니다..."):
                # 1. 서명 이미지 인코딩
                img = Image.fromarray(canvas_result.image_data.astype('uint8'), 'RGBA')
                img.thumbnail((200, 100)) # 시트 용량 최적화
                buf = io.BytesIO()
                img.save(buf, format="PNG")
                sig_b64 = base64.b64encode(buf.getvalue()).decode()

                # 2. 판다스 데이터프레임 업데이트
                row_idx = df.index[df['성함'] == selected_name].tolist()[0]
                df.at[row_idx, '서명'] = f"'{sig_b64}" # 문자열 보존을 위해 ' 추가
                
                # 3. 구글 시트 전체 업데이트
                conn.update(worksheet="동의서양식", data=df)
                
                st.balloons()
                st.success(f"감사합니다, {selected_name} 선생님! 서명이 정상적으로 수리되었습니다.")
                # 완료 후 화면 갱신을 원하시면 st.rerun() 사용 가능
        else:
            st.error("서명을 입력해 주세요.")
